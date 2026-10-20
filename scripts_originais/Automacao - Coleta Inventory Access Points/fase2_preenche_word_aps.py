

# -*- coding: utf-8 -*-

# arquivo: fase2_preenche_word_aps.py

import json

import sys

import unicodedata

from pathlib import Path

from docx import Document

from docx.oxml.shared import OxmlElement, qn



INPUT_JSON = Path("coleta_ap.json")

TEMPLATE_DOCX = Path("Access Points.docx")

OUTPUT_DOCX = Path("Access_Points_Preenchido.docx")



# Se True, pinta de amarelo células vazias (Serial/IP) para evidenciar

HIGHLIGHT_EMPTY_CELLS = True



HEADER_ALIASES = {

    "hostname": ["hostname", "host name", "ap name", "apname", "nome do ap"],

    "serial": ["serial", "sn", "numero de serie", "número de série", "ap serial number"],

    "ip": ["ip", "ip address", "endereco ip", "endereço ip"],

    "model": ["model", "modelo", "ap model"],

    "location": ["location", "local", "site", "andar"],

}



def normalize(s: str) -> str:

    s = s or ""

    s = s.strip()

    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

    return s.lower()



def load_json_map():

    if not INPUT_JSON.exists():

        print(f"[ERRO] Arquivo {INPUT_JSON} não encontrado.")

        sys.exit(1)

    data = json.loads(INPUT_JSON.read_text(encoding="utf-8"))

    itens = data.get("itens", [])

    ap_map = {}

    for it in itens:

        key = normalize(it.get("ap_name", ""))

        if not key:

            continue

        ap_map[key] = {

            "serial": it.get("ap_serial", "") or "",

            "ip": it.get("ap_ip", "") or "",

            "model": it.get("ap_model", "") or "",

            "wlc": it.get("wlc", "") or "",

        }

    return ap_map



def index_headers(header_cells):

    hdr_norm = [normalize(c.text) for c in header_cells]

    idx_map = {}

    for role, aliases in HEADER_ALIASES.items():

        for alias in aliases:

            for i, txt in enumerate(hdr_norm):

                if alias in txt:

                    idx_map[role] = i

                    break

            if role in idx_map:

                break

    # exigimos ao menos hostname + (serial ou ip)

    ok = ("hostname" in idx_map) and (("serial" in idx_map) or ("ip" in idx_map))

    return idx_map, ok



def highlight_cell(cell, color="FFFF00"):  # amarelo

    # aplica <w:shd w:fill="FFFF00">

    tcPr = cell._tc.get_or_add_tcPr()

    shd = OxmlElement('w:shd')

    shd.set(qn('w:fill'), color)

    tcPr.append(shd)



def main():

    if not TEMPLATE_DOCX.exists():

        print(f"[ERRO] Arquivo base {TEMPLATE_DOCX} não encontrado.")

        sys.exit(1)



    ap_map = load_json_map()

    doc = Document(str(TEMPLATE_DOCX))



    total_tables = 0

    total_rows = 0

    filled_cells = 0

    not_found_hostnames = []



    for tbl in doc.tables:

        if not tbl.rows:

            continue

        idx_map, ok = index_headers(tbl.rows[0].cells)

        if not ok:

            continue  # pula tabelas que não têm o layout esperado

        total_tables += 1



        idx_hostname = idx_map.get("hostname")

        idx_serial   = idx_map.get("serial")

        idx_ip       = idx_map.get("ip")



        for r in range(1, len(tbl.rows)):

            total_rows += 1

            cells = tbl.rows[r].cells

            hn_raw = (cells[idx_hostname].text or "").strip()

            hn_key = normalize(hn_raw)

            if not hn_key:

                # linha sem hostname -> nada a fazer

                continue



            ap_info = ap_map.get(hn_key)

            if not ap_info:

                not_found_hostnames.append(hn_raw)

                # se quiser destacar as ausências, pinte as duas células:

                if HIGHLIGHT_EMPTY_CELLS:

                    if idx_serial is not None: highlight_cell(cells[idx_serial])

                    if idx_ip is not None:     highlight_cell(cells[idx_ip])

                continue



            # Preenche Serial/IP preservando model/location

            if idx_serial is not None:

                val = ap_info["serial"] or ""

                cells[idx_serial].text = val

                if HIGHLIGHT_EMPTY_CELLS and not val:

                    highlight_cell(cells[idx_serial])

                else:

                    filled_cells += 1



            if idx_ip is not None:

                val = ap_info["ip"] or ""

                cells[idx_ip].text = val

                if HIGHLIGHT_EMPTY_CELLS and not val:

                    highlight_cell(cells[idx_ip])

                else:

                    filled_cells += 1



    doc.save(str(OUTPUT_DOCX))

    print(f"[OK] Documento gerado: {OUTPUT_DOCX}")

    print(f"[INFO] Tabelas analisadas: {total_tables} | Linhas: {total_rows} | Células preenchidas: {filled_cells}")

    if not_found_hostnames:

        unique_missing = sorted(set(not_found_hostnames))

        print(f"[WARN] {len(unique_missing)} hostname(s) do DOCX não encontrados no JSON (não preenchidos):")

        for h in unique_missing:

            print(f"  - {h}")



if __name__ == "__main__":

    main()

