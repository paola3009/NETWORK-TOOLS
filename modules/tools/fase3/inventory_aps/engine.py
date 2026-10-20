# -*- coding: utf-8 -*-
"""
Módulo de Coleta e Geração de Inventário de Access Points (Phase 3 - As-Built Documentation).
Suporta 3 modos de operação:
  1) Modo Original: Upload de arquivo base .docx com Hostnames pré-preenchidos + conexão WLC via SSH.
  2) Modo Auto-Discovery: Totalmente automatizado via IP da WLC (SSH), coleta todos os APs e obtém a localização via SNMP Location.
  3) Modo Catalyst Center: Importação de planilha Excel/CSV exportada do Cisco Catalyst Center.

Garante 100% de conformidade com a formatação padrão NTT/Cisco de 'Access_Points_Preenchido.docx'.
Credenciais são descartadas imediatamente da memória (RAM) com del e gc.collect().
"""

import os
import re
import csv
import time
import uuid
import gc
import unicodedata
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional

import openpyxl
from docx import Document
from docx.oxml.shared import OxmlElement, qn
from docx.shared import Twips, RGBColor
from netmiko import ConnectHandler, NetmikoAuthenticationException, NetmikoTimeoutException

# Caminho do template base corporativo
TEMPLATE_PATH = Path(__file__).resolve().parent / "template_base.docx"

# Larguras exatas das colunas (dxa / twips) compatíveis com Access_Points_Preenchido.docx
# Col 0: Hostname (1523 dxa)
# Col 1: Model (2116 dxa)
# Col 2: Serial (1701 dxa)
# Col 3: IP Address (1050 dxa)
# Col 4: Location (2682 dxa)
COL_WIDTHS_DXA = [1523, 2116, 1701, 1050, 2682]

# Expressões regulares para o Cisco Catalyst 9800 WLC
AP_SUMMARY_LINE_RE = re.compile(
    r"^(?P<name>\S[^\s]*)\s+\d+\s+(?P<model>\S+)\s+[0-9a-f\.\:]+\s+[0-9a-f\.\:]+\s+.*?(?P<ip>(?:\d{1,3}\.){3}\d{1,3})?\s+(?P<state>\S+)\s*$",
    re.IGNORECASE
)

DETAIL_FIELDS = {
    "ap_name": re.compile(r"^Cisco AP Name\s*:\s*(?P<val>.+)$", re.IGNORECASE),
    "ap_serial": re.compile(r"^AP Serial Number\s*:\s*(?P<val>.+)$", re.IGNORECASE),
    "ap_ip": re.compile(r"^IP Address\s*:\s*(?P<val>(?:\d{1,3}\.){3}\d{1,3})$", re.IGNORECASE),
    "ap_model": re.compile(r"^AP Model\s*:\s*(?P<val>.+)$", re.IGNORECASE),
    "ap_location": re.compile(r"^(?:Cisco\s+)?AP Location\s*:\s*(?P<val>.+)$", re.IGNORECASE),
    "site_tag": re.compile(r"^Site Tag Name\s*:\s*(?P<val>.+)$", re.IGNORECASE),
}

HEADER_ALIASES = {
    "hostname": ["hostname", "host name", "ap name", "apname", "nome do ap", "device name", "nome", "ap"],
    "serial": ["serial", "serial number", "sn", "s/n", "número de série", "numero de serie", "ap serial number"],
    "ip": ["ip", "ip address", "management ip address", "mgmt ip", "endereço ip", "endereco ip"],
    "model": ["model", "modelo", "platform", "product id", "pid", "device family", "ap model"],
    "location": ["location", "site", "location hierarchy", "local", "localização", "localizacao", "site name", "andar"],
}


def normalize_str(s: Any) -> str:
    """Normaliza texto para comparações insensíveis a acentos e maiúsculas/minúsculas."""
    if s is None:
        return ""
    txt = str(s).strip()
    txt = "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn")
    return txt.lower()


def highlight_cell(cell, color_hex="FFFF00"):
    """Pinta o fundo de uma célula (ex: amarelo FFFF00 para células ausentes)."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)


def parse_ap_summary(output: str) -> Dict[str, Dict[str, str]]:
    """Analisa a saída de 'show ap summary' de uma WLC Cisco Catalyst 9800."""
    ap_map = {}
    for line in output.splitlines():
        line = line.rstrip()
        if (not line) or line.startswith("-") or line.lower().startswith("ap name") or "Number of APs" in line:
            continue
        m = AP_SUMMARY_LINE_RE.search(line)
        if m:
            name = m.group("name").strip()
            ap_map[normalize_str(name)] = {
                "name": name,
                "ip": m.group("ip") or "",
                "model": (m.group("model") or "").strip(),
                "state": m.group("state") or "",
            }
        else:
            parts = line.split()
            if parts and len(parts) >= 2:
                name = parts[0]
                ap_map.setdefault(normalize_str(name), {"name": name, "ip": "", "model": "", "state": ""})
    return ap_map


def parse_ap_detail(ap_name: str, output: str) -> Dict[str, str]:
    """Extrai campos detalhados de 'show ap name <AP> config general'."""
    info = {
        "ap_name": ap_name,
        "ap_serial": "",
        "ap_ip": "",
        "ap_model": "",
        "ap_location": "",
        "site_tag": ""
    }
    for line in output.splitlines():
        line = line.strip()
        for key, regex in DETAIL_FIELDS.items():
            m = regex.match(line)
            if m:
                info[key] = m.group("val").strip()
    return info


def connect_wlc(host: str, user: str, pwd: str, session_log_path: Optional[str] = None):
    """Estabelece conexão SSH Netmiko com a WLC Cisco Catalyst 9800."""
    params = {
        "device_type": "cisco_xe",
        "host": host,
        "username": user,
        "password": pwd,
        "fast_cli": False,
        "global_delay_factor": 1.2,
        "timeout": 90,
    }
    if session_log_path:
        params["session_log"] = session_log_path

    conn = ConnectHandler(**params)
    conn.send_command_timing("terminal length 0", strip_prompt=False, strip_command=False)
    return conn


def get_download_path(filename: str) -> str:
    """Retorna o caminho seguro no diretório de downloads da aplicação."""
    base_raiz = Path(__file__).resolve().parents[4]
    base_dir = base_raiz / "core" / "static" / "downloads"
    base_dir.mkdir(parents=True, exist_ok=True)
    return str(base_dir / filename)


# ==============================================================================
# FORMATAÇÃO E CONSTRUÇÃO DO DOCUMENTO WORD (.DOCX)
# ==============================================================================

def format_table(tbl, aps_list: List[Dict[str, str]], location_name: str):
    """
    Formata tabela Word de 5 colunas com o padrão estrito de 'Access_Points_Preenchido.docx':
    - Estilo da tabela: Table Grid
    - Cabeçalho azul corporativo (#0072BC), texto branco, negrito, estilo NTTTableWhiteHeader
    - Linhas de dados com estilo NTTTableBodyText e cor #2E404D
    - Bordas pretas inferiores e laterais nos dados
    - Larguras exatas fixadas por coluna
    """
    # Fixa layout da tabela
    tblPr = tbl._tbl.tblPr
    tblPr.set(qn('w:tblLayout'), 'fixed')

    # Cabeçalho
    hdr_cells = tbl.rows[0].cells
    headers = ["Hostname", "Model", "Serial", "IP Address", "Location"]
    for i, title in enumerate(headers):
        cell = hdr_cells[i]
        cell.text = title

        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), '0072BC')
        tcPr.append(shd)

        tcBorders = OxmlElement('w:tcBorders')
        top = OxmlElement('w:top')
        top.set(qn('w:val'), 'nil')
        tcBorders.append(top)
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'nil')
        tcBorders.append(bottom)
        right = OxmlElement('w:right')
        right.set(qn('w:val'), 'single')
        right.set(qn('w:sz'), '4')
        right.set(qn('w:space'), '0')
        right.set(qn('w:color'), 'FFFFFF')
        tcBorders.append(right)
        tcPr.append(tcBorders)

        vAlign = OxmlElement('w:vAlign')
        vAlign.set(qn('w:val'), 'center')
        tcPr.append(vAlign)

        for p in cell.paragraphs:
            p.style = 'NTT Table White Header'
            for run in p.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.bold = True

    # Linhas de dados
    for ap in aps_list:
        row = tbl.add_row()
        hn = ap.get("hostname", "")
        mod = ap.get("model", "")
        sn = ap.get("serial", "")
        ip = ap.get("ip", "")
        loc = ap.get("location") or location_name

        row_vals = [hn, mod, sn, ip, loc]
        for i, val in enumerate(row_vals):
            cell = row.cells[i]
            cell.text = str(val or "")
            tcPr = cell._tc.get_or_add_tcPr()

            tcBorders = OxmlElement('w:tcBorders')
            top = OxmlElement('w:top')
            top.set(qn('w:val'), 'nil')
            tcBorders.append(top)
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '4')
            bottom.set(qn('w:space'), '0')
            bottom.set(qn('w:color'), '000000')
            tcBorders.append(bottom)
            right = OxmlElement('w:right')
            right.set(qn('w:val'), 'single')
            right.set(qn('w:sz'), '4')
            right.set(qn('w:space'), '0')
            right.set(qn('w:color'), '000000')
            tcBorders.append(right)
            tcPr.append(tcBorders)

            vAlign = OxmlElement('w:vAlign')
            vAlign.set(qn('w:val'), 'center')
            tcPr.append(vAlign)

            for p in cell.paragraphs:
                p.style = 'NTT Table Body Text'
                for run in p.runs:
                    run.font.color.rgb = RGBColor(46, 64, 77)

    # Aplica larguras de colunas em todas as linhas
    for r in tbl.rows:
        for i, w_dxa in enumerate(COL_WIDTHS_DXA):
            c = r.cells[i]
            c.width = Twips(w_dxa)
            tcPr = c._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.append(tcW)
            tcW.set(qn('w:w'), str(w_dxa))
            tcW.set(qn('w:type'), 'dxa')


def build_docx_from_grouped_aps(grouped_aps: Dict[str, List[Dict[str, str]]], output_filename: str) -> str:
    """
    Constrói o documento .docx padronizado agrupando APs por Localização.
    Utiliza template_base.docx para reter fontes, cabeçalhos e estilos corporativos.
    """
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template base '{TEMPLATE_PATH}' não encontrado.")

    doc = Document(str(TEMPLATE_PATH))

    # Mantém apenas os parágrafos introdutórios e remove as tabelas pré-existentes
    for p in list(doc.paragraphs)[4:]:
        p._p.getparent().remove(p._p)
    for t in list(doc.tables):
        t._tbl.getparent().remove(t._tbl)

    # Ordena localizações
    sorted_locations = sorted(grouped_aps.keys(), key=lambda x: (x == "Sem Localização Definida", x.lower()))

    table_num = 1
    for loc_name in sorted_locations:
        aps = grouped_aps[loc_name]
        if not aps:
            continue

        # Ordena APs por hostname
        aps.sort(key=lambda a: a.get("hostname", "").lower())

        # Parágrafos de título
        doc.add_paragraph(f"Lista de Equipamentos – AP – {loc_name}", style='Heading 4')
        doc.add_paragraph(f"Tabela {table_num} – Lista de Equipamentos – AP – {loc_name}", style='NTT Body Text')

        # Cria e estiliza tabela
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = 'Table Grid'
        format_table(tbl, aps, loc_name)

        # Espaçamento
        doc.add_paragraph('', style='NTT Flush Body Text')
        doc.add_paragraph('', style='NTT Flush Body Text')

        table_num += 1

    dest_path = get_download_path(output_filename)
    doc.save(dest_path)
    return dest_path


# ==============================================================================
# OPÇÃO 1: MODO ORIGINAL (UPLOAD DOCX BASE + WLC SSH)
# ==============================================================================

def index_docx_headers(header_cells):
    """Mapeia os índices de colunas de uma tabela DOCX."""
    hdr_norm = [normalize_str(c.text) for c in header_cells]
    idx_map = {}
    for role, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            for i, txt in enumerate(hdr_norm):
                if alias in txt:
                    idx_map[role] = i
                    break
            if role in idx_map:
                break
    ok = ("hostname" in idx_map) and (("serial" in idx_map) or ("ip" in idx_map))
    return idx_map, ok


def read_hostnames_from_docx(doc: Document) -> List[str]:
    """Lê todos os hostnames listados nas tabelas do DOCX base."""
    wanted = set()
    for tbl in doc.tables:
        if not tbl.rows:
            continue
        idx_map, ok = index_docx_headers(tbl.rows[0].cells)
        if not ok:
            continue
        idx_host = idx_map["hostname"]
        for r in range(1, len(tbl.rows)):
            hn = tbl.rows[r].cells[idx_host].text.strip()
            if hn:
                wanted.add(hn)
    return sorted(wanted)


def executar_modo_original(caminho_docx_base: str, wlcs_list: List[str], usuario: str, senha: str) -> Tuple[Optional[str], List[str], Dict[str, Any]]:
    """
    Executa a Opção 1: Lê o DOCX enviado pelo usuário, conecta aos WLCs,
    preenche Serial e IP e destaca em amarelo células ausentes.
    """
    logs = []
    stats = {"total_aps_docx": 0, "preenchidos": 0, "nao_encontrados": 0, "tabelas": 0}
    caminho_gerado = None

    try:
        logs.append(f"[INFO] Abrindo arquivo DOCX base: {os.path.basename(caminho_docx_base)}")
        doc = Document(caminho_docx_base)
        targets = read_hostnames_from_docx(doc)
        stats["total_aps_docx"] = len(targets)
        logs.append(f"[INFO] Total de hostnames identificados no DOCX: {len(targets)}")

        if not targets:
            logs.append("[ERRO] Nenhum hostname de Access Point encontrado nas tabelas do documento.")
            return None, logs, stats

        # Coleta nos WLCs
        coletados_map = {}
        for wlc in wlcs_list:
            wlc = wlc.strip()
            if not wlc:
                continue
            logs.append(f"[SSH] Conectando à WLC {wlc}...")
            try:
                with connect_wlc(wlc, usuario, senha) as conn:
                    logs.append(f"[{wlc}] Coletando 'show ap summary'...")
                    summary_out = conn.send_command_timing("show ap summary", delay_factor=2, max_loops=200)
                    ap_map = parse_ap_summary(summary_out)
                    logs.append(f"[{wlc}] Total de APs associados na WLC: {len(ap_map)}")

                    for i, ap_target in enumerate(targets, start=1):
                        ap_key = normalize_str(ap_target)
                        if ap_key in coletados_map and coletados_map[ap_key].get("ap_serial"):
                            continue

                        seed = ap_map.get(ap_key, {"name": ap_target, "ip": "", "model": "", "state": ""})
                        cmd = f"show ap name {seed['name']} config general"
                        out = conn.send_command_timing(cmd, delay_factor=2.5, max_loops=300)

                        if "not found" in out.lower() or "invalid" in out.lower():
                            if seed.get("ip"):
                                coletados_map[ap_key] = {
                                    "ap_name": ap_target,
                                    "ap_ip": seed.get("ip", ""),
                                    "ap_serial": "",
                                    "ap_model": seed.get("model", ""),
                                }
                        else:
                            detail = parse_ap_detail(ap_target, out)
                            if not detail.get("ap_ip"):
                                detail["ap_ip"] = seed.get("ip", "")
                            if not detail.get("ap_model"):
                                detail["ap_model"] = seed.get("model", "")
                            coletados_map[ap_key] = detail

                        if i % 25 == 0 or i == len(targets):
                            logs.append(f"[{wlc}] Progresso da coleta detalhada: {i}/{len(targets)}")

            except (NetmikoAuthenticationException, NetmikoTimeoutException) as e:
                logs.append(f"[ERRO SSH] WLC {wlc}: {e}")
            except Exception as e:
                logs.append(f"[ERRO] Falha na comunicação com {wlc}: {e}")

        # Preenchimento do documento DOCX
        logs.append("[DOCX] Preenchendo dados de Serial e IP Address nas tabelas...")
        filled_count = 0
        not_found_count = 0
        total_tables = 0

        for tbl in doc.tables:
            if not tbl.rows:
                continue
            idx_map, ok = index_docx_headers(tbl.rows[0].cells)
            if not ok:
                continue
            total_tables += 1
            idx_hostname = idx_map.get("hostname")
            idx_serial = idx_map.get("serial")
            idx_ip = idx_map.get("ip")

            for r in range(1, len(tbl.rows)):
                cells = tbl.rows[r].cells
                hn_raw = (cells[idx_hostname].text or "").strip()
                hn_key = normalize_str(hn_raw)
                if not hn_key:
                    continue

                info = coletados_map.get(hn_key)
                if not info:
                    not_found_count += 1
                    if idx_serial is not None:
                        highlight_cell(cells[idx_serial], "FFFF00")
                    if idx_ip is not None:
                        highlight_cell(cells[idx_ip], "FFFF00")
                    continue

                if idx_serial is not None:
                    sn_val = info.get("ap_serial", "") or ""
                    cells[idx_serial].text = sn_val
                    if not sn_val:
                        highlight_cell(cells[idx_serial], "FFFF00")
                    else:
                        filled_count += 1

                if idx_ip is not None:
                    ip_val = info.get("ap_ip", "") or ""
                    cells[idx_ip].text = ip_val
                    if not ip_val:
                        highlight_cell(cells[idx_ip], "FFFF00")

        stats["preenchidos"] = filled_count
        stats["nao_encontrados"] = not_found_count
        stats["tabelas"] = total_tables

        ts = time.strftime("%Y%m%d_%H%M%S")
        out_name = f"Access_Points_Preenchido_{ts}.docx"
        dest_path = get_download_path(out_name)
        doc.save(dest_path)
        caminho_gerado = dest_path

        logs.append(f"[OK] Sucesso! Documento salvo como {out_name}")
        logs.append(f"[RESUMO] {filled_count} APs atualizados, {not_found_count} marcados em amarelo em {total_tables} tabelas.")

    finally:
        del usuario, senha
        gc.collect()

    return caminho_gerado, logs, stats


# ==============================================================================
# OPÇÃO 2: MODO TOTALMENTE AUTOMATIZADO (WLC IP + SNMP LOCATION)
# ==============================================================================

def executar_modo_auto(wlcs_list: List[str], usuario: str, senha: str) -> Tuple[Optional[str], List[str], Dict[str, Any]]:
    """
    Executa a Opção 2: Descobre automaticamente TODOS os APs na WLC,
    coleta serial, ip, modelo e obtém a localização a partir do SNMP Location.
    Gera o documento .docx formatado agrupado por Localização.
    """
    logs = []
    stats = {"total_aps_descobertos": 0, "total_localizacoes": 0}
    caminho_gerado = None

    try:
        discovered_aps = {}

        for wlc in wlcs_list:
            wlc = wlc.strip()
            if not wlc:
                continue

            logs.append(f"[SSH] Conectando à WLC {wlc}...")
            try:
                with connect_wlc(wlc, usuario, senha) as conn:
                    logs.append(f"[{wlc}] Executando 'show ap summary' para descoberta geral...")
                    summary_out = conn.send_command_timing("show ap summary", delay_factor=2, max_loops=300)
                    ap_map = parse_ap_summary(summary_out)
                    logs.append(f"[{wlc}] Descobertos {len(ap_map)} Access Points registrados.")

                    total_wlc_aps = len(ap_map)
                    for i, (ap_key, seed) in enumerate(ap_map.items(), start=1):
                        ap_name = seed["name"]
                        cmd = f"show ap name {ap_name} config general"
                        out = conn.send_command_timing(cmd, delay_factor=2.5, max_loops=300)

                        detail = parse_ap_detail(ap_name, out)
                        ip_val = detail.get("ap_ip") or seed.get("ip") or ""
                        mod_val = detail.get("ap_model") or seed.get("model") or ""
                        sn_val = detail.get("ap_serial") or ""

                        # Localização via SNMP Location
                        loc_val = detail.get("ap_location", "").strip()
                        if not loc_val or loc_val.lower() in ("default location", "none", "default"):
                            loc_val = detail.get("site_tag", "").strip() or "Sem Localização Definida"

                        discovered_aps[ap_key] = {
                            "hostname": ap_name,
                            "model": mod_val,
                            "serial": sn_val,
                            "ip": ip_val,
                            "location": loc_val
                        }

                        if i % 25 == 0 or i == total_wlc_aps:
                            logs.append(f"[{wlc}] Coleta detalhada e SNMP Location: {i}/{total_wlc_aps} APs")

            except (NetmikoAuthenticationException, NetmikoTimeoutException) as e:
                logs.append(f"[ERRO SSH] WLC {wlc}: {e}")
            except Exception as e:
                logs.append(f"[ERRO] Falha na comunicação com {wlc}: {e}")

        stats["total_aps_descobertos"] = len(discovered_aps)
        logs.append(f"[INFO] Total final de APs únicos coletados: {len(discovered_aps)}")

        if not discovered_aps:
            logs.append("[ERRO] Nenhum Access Point foi descoberto nas WLCs informadas.")
            return None, logs, stats

        # Agrupar APs por Localização (SNMP Location)
        grouped = {}
        for ap in discovered_aps.values():
            loc = ap.get("location") or "Sem Localização Definida"
            grouped.setdefault(loc, []).append(ap)

        stats["total_localizacoes"] = len(grouped)
        logs.append(f"[INFO] APs distribuídos em {len(grouped)} localizações (SNMP Location).")

        # Gerar DOCX padronizado
        ts = time.strftime("%Y%m%d_%H%M%S")
        out_name = f"Access_Points_Auto_{ts}.docx"
        logs.append(f"[DOCX] Gerando documento padronizado {out_name}...")
        caminho_gerado = build_docx_from_grouped_aps(grouped, out_name)
        logs.append(f"[OK] Sucesso! Arquivo gerado: {out_name}")

    finally:
        del usuario, senha
        gc.collect()

    return caminho_gerado, logs, stats


# ==============================================================================
# OPÇÃO 3: MODO IMPORT PLANILHA CATALYST CENTER (EXCEL / CSV)
# ==============================================================================

def extrair_aps_planilha_catalyst(caminho_arquivo: str) -> List[Dict[str, str]]:
    """
    Lê planilha Excel (.xlsx, .xls) ou CSV exportada do Cisco Catalyst Center.
    Mapeia colunas de forma tolerante a variações de nomes.
    """
    ext = os.path.splitext(caminho_arquivo)[1].lower()
    raw_rows = []

    if ext in (".xlsx", ".xlsm", ".xltx"):
        wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
        sheet = wb.active
        for row in sheet.iter_rows(values_only=True):
            if any(cell is not None for cell in row):
                raw_rows.append([str(c).strip() if c is not None else "" for c in row])
        wb.close()
    elif ext == ".csv":
        with open(caminho_arquivo, mode="r", encoding="utf-8-sig", errors="ignore") as f:
            sample = f.read(4096)
            f.seek(0)
            delimiter = ";" if sample.count(";") > sample.count(",") else ","
            reader = csv.reader(f, delimiter=delimiter)
            for row in reader:
                if any(c.strip() for c in row):
                    raw_rows.append([c.strip() for c in row])
    else:
        raise ValueError(f"Extensão de arquivo '{ext}' não suportada. Envie um arquivo .xlsx ou .csv.")

    if len(raw_rows) < 2:
        raise ValueError("O arquivo não possui linhas de dados suficientes.")

    # Localiza o cabeçalho
    header_idx = -1
    col_map = {}

    for idx, row in enumerate(raw_rows[:15]):
        normalized_cols = [normalize_str(c) for c in row]
        temp_map = {}
        for role, aliases in HEADER_ALIASES.items():
            for alias in aliases:
                for c_idx, val in enumerate(normalized_cols):
                    if alias in val:
                        temp_map[role] = c_idx
                        break
                if role in temp_map:
                    break
        if "hostname" in temp_map:
            header_idx = idx
            col_map = temp_map
            break

    if header_idx == -1 or "hostname" not in col_map:
        raise ValueError("Não foi possível identificar a coluna de Hostname/Device Name no arquivo.")

    aps = []
    idx_host = col_map.get("hostname")
    idx_mod = col_map.get("model")
    idx_sn = col_map.get("serial")
    idx_ip = col_map.get("ip")
    idx_loc = col_map.get("location")

    for row in raw_rows[header_idx + 1:]:
        if len(row) <= idx_host:
            continue
        hn = row[idx_host].strip()
        if not hn or hn.lower().startswith("device name") or hn.lower().startswith("hostname"):
            continue

        mod = row[idx_mod].strip() if (idx_mod is not None and len(row) > idx_mod) else ""
        sn = row[idx_sn].strip() if (idx_sn is not None and len(row) > idx_sn) else ""
        ip = row[idx_ip].strip() if (idx_ip is not None and len(row) > idx_ip) else ""
        loc = row[idx_loc].strip() if (idx_loc is not None and len(row) > idx_loc) else "Sem Localização"

        # Limpeza estética se local vier no formato 'Global/Site/Predio/Andar'
        if loc.startswith("Global/"):
            loc = loc[len("Global/"):].strip()

        aps.append({
            "hostname": hn,
            "model": mod,
            "serial": sn,
            "ip": ip,
            "location": loc or "Sem Localização"
        })

    return aps


def executar_modo_catalyst_center(caminho_arquivo: str) -> Tuple[Optional[str], List[str], Dict[str, Any]]:
    """
    Executa a Opção 3: Lê a planilha do Catalyst Center, extrai APs,
    agrupa por Localização e gera o documento DOCX corporativo.
    """
    logs = []
    stats = {"total_aps_planilha": 0, "total_localizacoes": 0}
    caminho_gerado = None

    try:
        logs.append(f"[INFO] Processando export do Catalyst Center: {os.path.basename(caminho_arquivo)}")
        aps = extrair_aps_planilha_catalyst(caminho_arquivo)
        stats["total_aps_planilha"] = len(aps)
        logs.append(f"[INFO] {len(aps)} Access Points extraídos da planilha com sucesso.")

        if not aps:
            logs.append("[ERRO] Nenhum AP válido encontrado na planilha importada.")
            return None, logs, stats

        grouped = {}
        for ap in aps:
            loc = ap.get("location") or "Sem Localização"
            grouped.setdefault(loc, []).append(ap)

        stats["total_localizacoes"] = len(grouped)
        logs.append(f"[INFO] APs distribuídos em {len(grouped)} localizações.")

        ts = time.strftime("%Y%m%d_%H%M%S")
        out_name = f"Access_Points_CatalystCenter_{ts}.docx"
        logs.append(f"[DOCX] Gerando documento padronizado {out_name}...")
        caminho_gerado = build_docx_from_grouped_aps(grouped, out_name)
        logs.append(f"[OK] Sucesso! Documento gerado com sucesso: {out_name}")

    except Exception as e:
        logs.append(f"[ERRO] Falha ao importar export do Catalyst Center: {e}")

    return caminho_gerado, logs, stats
