

# -*- coding: utf-8 -*-

# arquivo: fase1_coleta_aps.py

import json

import re

import sys

import time

import getpass

import unicodedata

from pathlib import Path



from netmiko import ConnectHandler, NetmikoAuthenticationException, NetmikoTimeoutException

from docx import Document



INPUT_WLC_LIST   = Path("dispositivos.txt")

INPUT_DOCX       = Path("Access Points.docx")

OUTPUT_JSON      = Path("coleta_ap.json")



COMMAND_TIMEOUT  = 60      # + tolerante no 9800

GLOBAL_DELAY     = 1.2     # alivia a CLI

SESSION_LOG_DIR  = Path("./")  # mesma pasta



# Regex flexível do summary do 9800: tenta capturar Name, Model, IP, State

AP_SUMMARY_LINE_RE = re.compile(

    r"^(?P<name>\S[^\s]*)\s+\d+\s+(?P<model>\S+)\s+[0-9a-f\.\:]+\s+[0-9a-f\.\:]+\s+.*?(?P<ip>(?:\d{1,3}\.){3}\d{1,3})?\s+(?P<state>\S+)\s*$",

    re.IGNORECASE

)



DETAIL_FIELDS = {

    "ap_name"  : re.compile(r"^Cisco AP Name\s*:\s*(?P<val>.+)$", re.IGNORECASE),

    "ap_serial": re.compile(r"^AP Serial Number\s*:\s*(?P<val>.+)$", re.IGNORECASE),

    "ap_ip"    : re.compile(r"^IP Address\s*:\s*(?P<val>(?:\d{1,3}\.){3}\d{1,3})$", re.IGNORECASE),

    "ap_model" : re.compile(r"^AP Model\s*:\s*(?P<val>.+)$", re.IGNORECASE),

}



def normalize(s: str) -> str:

    s = (s or "").strip()

    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

    return s.lower()



def read_wlc_list():

    if not INPUT_WLC_LIST.exists():

        print(f"[ERRO] {INPUT_WLC_LIST} não encontrado.")

        sys.exit(1)

    wlcs = [ln.strip() for ln in INPUT_WLC_LIST.read_text(encoding="utf-8").splitlines()

            if ln.strip() and not ln.strip().startswith("#")]

    if not wlcs:

        print(f"[ERRO] Nenhum WLC em {INPUT_WLC_LIST}.")

        sys.exit(1)

    return wlcs



def read_hostnames_from_docx():

    if not INPUT_DOCX.exists():

        print(f"[ERRO] {INPUT_DOCX} não encontrado.")

        sys.exit(1)

    doc = Document(str(INPUT_DOCX))

    wanted = set()



    def is_header(cells):

        hdr = [normalize(c.text) for c in cells]

        return ("hostname" in " ".join(hdr)) and (("serial" in " ".join(hdr)) or ("ip address" in " ".join(hdr)))



    for tbl in doc.tables:

        if not tbl.rows:

            continue

        if is_header(tbl.rows[0].cells):

            # acha índice do Hostname

            hdr = [normalize(c.text) for c in tbl.rows[0].cells]

            try:

                idx_host = next(i for i,t in enumerate(hdr) if "hostname" in t)

            except StopIteration:

                continue

            # coleta hostnames das linhas existentes

            for r in range(1, len(tbl.rows)):

                hn = tbl.rows[r].cells[idx_host].text.strip()

                if hn:

                    wanted.add(hn)

    if not wanted:

        print("[ERRO] Não encontrei hostnames no DOCX.")

        sys.exit(1)

    return sorted(wanted)



def connect_wlc(host, username, password):

    session_log = str(SESSION_LOG_DIR / f"session_wlc_{host}.log")

    params = {

        "device_type": "cisco_xe",   # 9800 é IOS-XE

        "host": host,

        "username": username,

        "password": password,

        "fast_cli": False,

        "global_delay_factor": GLOBAL_DELAY,

        "timeout": 120,

        "session_log": session_log,

    }

    conn = ConnectHandler(**params)

    # sem rely no expect_string (evita pattern error)

    conn.send_command_timing("terminal length 0", strip_prompt=False, strip_command=False)

    return conn



def parse_ap_summary(output):

    ap_map = {}

    for line in output.splitlines():

        line = line.rstrip()

        if (not line) or line.startswith("-") or line.lower().startswith("ap name") or "Number of APs" in line:

            continue

        m = AP_SUMMARY_LINE_RE.search(line)

        if m:

            name = m.group("name").strip()

            ap_map[normalize(name)] = {

                "name": name,

                "ip": m.group("ip") or "",

                "model": (m.group("model") or "").strip(),

                "state": m.group("state") or "",

            }

        else:

            # fallback extremo: primeira coluna é o nome

            parts = line.split()

            if parts:

                name = parts[0]

                ap_map.setdefault(normalize(name), {"name": name, "ip": "", "model": "", "state": ""})

    return ap_map



def parse_ap_detail(ap_name, output):

    info = {"ap_name": ap_name, "ap_serial": "", "ap_ip": "", "ap_model": ""}

    for line in output.splitlines():

        line = line.strip()

        for key, regex in DETAIL_FIELDS.items():

            m = regex.match(line)

            if m:

                info[key] = m.group("val").strip()

    return info



def collect_from_wlc_for_targets(wlc, username, password, targets):

    data, errors, missing = [], [], []

    try:

        with connect_wlc(wlc, username, password) as conn:

            # 1) Lista resumo (um comando só)

            summary = conn.send_command_timing("show ap summary", delay_factor=2, max_loops=200,

                                               strip_prompt=False, strip_command=False)

            ap_map = parse_ap_summary(summary)  # nome normalizado -> {name,ip,model,state}



            # 2) Para cada AP do DOCX, tentar detalhe

            total = len(targets)

            print(f"[INFO] {wlc}: {total} AP(s) listados no DOCX.")

            for i, ap in enumerate(targets, start=1):

                ap_key = normalize(ap)

                pretty = ap

                seed = ap_map.get(ap_key, {"name": ap, "ip": "", "model": "", "state": ""})



                cmd = f"show ap name {seed['name']} config general"

                out = conn.send_command_timing(cmd, delay_factor=2.5, max_loops=400,

                                               strip_prompt=False, strip_command=False)

                # casos de erro comuns no 9800

                if "not found" in out.lower() or "invalid" in out.lower():

                    missing.append(pretty)

                    # mesmo assim, se o summary trouxe IP/Model, aproveita

                    item = {

                        "ap_name": pretty,

                        "ap_ip": seed.get("ip",""),

                        "ap_serial": "",

                        "ap_model": seed.get("model",""),

                        "wlc": wlc

                    }

                    data.append(item)

                else:

                    detail = parse_ap_detail(pretty, out)

                    # completar com summary

                    if not detail.get("ap_ip"):

                        detail["ap_ip"] = seed.get("ip","")

                    if not detail.get("ap_model"):

                        detail["ap_model"] = seed.get("model","")

                    detail["wlc"] = wlc

                    data.append(detail)



                if i % 20 == 0 or i == total:

                    print(f"[{wlc}] progresso: {i}/{total}")



    except (NetmikoAuthenticationException, NetmikoTimeoutException) as e:

        errors.append({"wlc": wlc, "stage": "connect", "error": str(e)})

    except Exception as e:

        errors.append({"wlc": wlc, "stage": "general", "error": str(e)})



    return data, errors, missing



def main():

    wlcs = read_wlc_list()

    targets = read_hostnames_from_docx()

    user = input("Usuário SSH: ").strip()

    pwd  = getpass.getpass("Senha SSH: ").strip()



    all_items, all_errors, all_missing = [], [], []

    t0 = time.time()

    for wlc in wlcs:

        print(f"[INFO] Coletando do WLC {wlc} ...")

        items, errs, missing = collect_from_wlc_for_targets(wlc, user, pwd, targets)

        all_items.extend(items)

        all_errors.extend(errs)

        all_missing.extend(missing)



    # Dedup por Hostname (mantém o primeiro - útil se houver mais de um WLC listado)

    seen = set()

    unique_items = []

    for it in all_items:

        key = normalize(it.get("ap_name",""))

        if not key or key in seen:

            continue

        seen.add(key)

        unique_items.append({

            "ap_name"  : it.get("ap_name",""),

            "ap_ip"    : it.get("ap_ip",""),

            "ap_serial": it.get("ap_serial",""),

            "ap_model" : it.get("ap_model",""),

            "wlc"      : it.get("wlc",""),

        })



    payload = {

        "coleta_tipo": "APs_WLC9800",

        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),

        "total_aps": len(unique_items),

        "itens": unique_items,

        "nao_encontrados_no_wlc": sorted(set(all_missing)),

        "erros": all_errors,

    }

    OUTPUT_JSON.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

    dt = time.time() - t0

    print(f"[OK] Coleta concluída: {len(unique_items)} AP(s). Tempo: {dt/60:.1f} min")

    if all_missing:

        print(f"[WARN] {len(set(all_missing))} hostname(s) do DOCX não encontrados no WLC (ver campo 'nao_encontrados_no_wlc' no JSON).")

    if all_errors:

        print(f"[WARN] Ocorreram {len(all_errors)} erro(s). Veja em {OUTPUT_JSON}")



if __name__ == "__main__":

    main()

