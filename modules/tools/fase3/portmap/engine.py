"""
Módulo de automação para geração de Port-Map de Edges (Switches Cisco/IOS).
Suporta dois modos operacionais:
  Modo 1: Com Planilha Base (preenchimento automático de IPs e portas CDP)
  Modo 2: Auto-Discovery recursivo via COREs (crawler de cascateamento, orientação Topológica A/B,
          filtro estrito de switches de acesso e coleta de Midia Type via SFP).
"""

import copy
import gc
import json
import os
import re
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Optional, Set, Tuple

from netmiko import ConnectHandler
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd


# -------------------------------------------------------------
# Utilitários de Normalização e Validação de Hostnames
# -------------------------------------------------------------

# Padrão estrito de Switch de Acesso (Lado B):
# Começa obrigatoriamente com SW1 + 4 dígitos numéricos + 3 letras finais (ex: SW10901IFO, SW12345ICA)
PADRAO_SWITCH_ACESSO_B = re.compile(r"^SW1\d{4}[A-Za-z]{3}$")


def strip_domain(h: Optional[str]) -> str:
    """Remove FQDN e espaços em branco."""
    if not h or not isinstance(h, str):
        return ""
    h = h.strip()
    return h.split(".")[0] if "." in h else h


def label(h: Optional[str]) -> str:
    """Rótulo sem domínio em MAIÚSCULAS para comparações consistentes."""
    return strip_domain(h).upper()


def is_switch_acesso_b(h: Optional[str]) -> bool:
    """Valida se o hostname atende estritamente à regra de switch de acesso final (LADO B)."""
    lbl = label(h)
    return bool(PADRAO_SWITCH_ACESSO_B.match(lbl))


def is_switch_generico(h: Optional[str]) -> bool:
    """Identifica se o hostname pertence a um switch na rede (CORE, Distribuição ou Acesso)."""
    lbl = label(h)
    return lbl.startswith("SW") or "SWITCH" in lbl


def short_ifname(name: str) -> str:
    """Normaliza nome de interface Cisco para notação curta (ex: GigabitEthernet1/0/1 -> Gi1/0/1)."""
    if not name:
        return ""
    name = name.strip().rstrip(",")
    mapping = [
        (r"(?i)TwentyFiveGig(?:abitEthernet|E|Ethernet)", "Twe"),
        (r"(?i)FortyGig(?:abitEthernet|E|Ethernet)", "Fo"),
        (r"(?i)HundredGig(?:abitEthernet|E|Ethernet)", "Hu"),
        (r"(?i)TenGig(?:abitEthernet|E|Ethernet)", "Te"),
        (r"(?i)Gig(?:abitEthernet|E|Ethernet)", "Gi"),
        (r"(?i)FastEthernet", "Fa"),
        (r"(?i)Ethernet", "Eth"),
        (r"(?i)Port-Channel|PortChannel", "Po"),
        (r"(?i)Loopback", "Lo"),
        (r"(?i)MgmtEthernet|Management", "Ma"),
    ]
    for pat, abbr in mapping:
        name = re.sub(pat, abbr, name)
    name = re.sub(r"^([A-Za-z]+)\s+(\S+)$", r"\1\2", name)
    return name


# -------------------------------------------------------------
# Parsers de Comandos Cisco IOS
# -------------------------------------------------------------

def parse_cdp_detail(output: str) -> List[Dict[str, str]]:
    """Extrai vizinhos CDP, IP, porta local e porta remota."""
    neighbors = []
    if not output:
        return neighbors

    blocks = re.split(r"(?=^Device ID:\s*)", output, flags=re.M)
    for blk in blocks:
        blk = blk.strip()
        if not blk or not blk.startswith("Device ID:"):
            continue

        dev = re.search(r"^Device ID:\s*(.+)", blk, re.M)
        neighbor = strip_domain(dev.group(1).strip()) if dev else None

        ipm = re.search(r"IP address:\s*([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)", blk)
        loc = re.search(r"Interface:\s*([^\s,]+(?:\s+\S+)?),", blk)
        port = re.search(r"Port ID \(outgoing port\):\s*(.+)", blk)

        neighbors.append({
            "neighbor": neighbor,
            "neighbor_ip": ipm.group(1) if ipm else None,
            "local_intf": short_ifname(loc.group(1)) if loc else None,
            "neighbor_intf": short_ifname(port.group(1)) if port else None
        })
    return neighbors


def parse_loopback(output: str) -> Optional[str]:
    """Extrai IP da Loopback0."""
    m = re.search(r"(\d+\.\d+\.\d+\.\d+)", output)
    return m.group(1) if m else None


def parse_snmp_location(output: str) -> str:
    """Extrai localização do SNMP (ex: snmp-server location ...)."""
    if not output:
        return ""
    for line in output.splitlines():
        line = line.strip()
        if "snmp-server location" in line.lower():
            partes = line.split("location", 1)
            if len(partes) > 1:
                return partes[1].strip().strip('"')
        elif "syslocation" in line.lower():
            partes = line.split(":", 1)
            if len(partes) > 1:
                return partes[1].strip()
    return ""


def parse_inventory_sfps(output: str) -> Dict[str, str]:
    """
    Analisa o output de 'show inventory' para mapear o modelo de transceptor SFP por interface.
    Ex: NAME: "TenGigabitEthernet1/1/1", DESCR: ...
        PID: SFP-10G-LR
    Retorna dicionário: { 'Te1/1/1': 'SFP-10G-LR', ... }
    """
    sfps = {}
    if not output:
        return sfps

    # Separa por blocos NAME: ... PID: ...
    padrao = re.compile(r'NAME:\s*"([^"]+)"[^\n]*\n(?:[^\n]*\n)?PID:\s*([^\s,]+)', re.IGNORECASE)
    for m in padrao.finditer(output):
        raw_name = m.group(1).strip()
        pid = m.group(2).strip()

        # Filtra apenas se for transceiver/SFP/GLC/QSFP
        pid_upper = pid.upper()
        if any(tag in pid_upper for tag in ["SFP", "GLC", "QSFP", "X2", "XENPAK"]):
            intf_curta = short_ifname(raw_name)
            sfps[intf_curta] = pid
            # Mapeia também sem abreviação por garantia
            sfps[raw_name] = pid

    return sfps


# -------------------------------------------------------------
# Conexão SSH via Netmiko com descarte seguro de senha
# -------------------------------------------------------------

def coletar_switch(host: str, usuario: str, senha_ssh: str) -> Dict:
    """
    Conecta no switch, coleta CDP, Loopback, SNMP Location e Inventário de SFPs.
    A senha existe apenas durante a chamada e é expurgada logo após.
    """
    device = {
        "device_type": "cisco_ios",
        "host": host,
        "username": usuario,
        "password": senha_ssh,
        "fast_cli": False,
        "timeout": 15,
        "conn_timeout": 10,
    }

    resultado = {
        "host": host,
        "status": "erro",
        "loopback0_ip": None,
        "snmp_location": "",
        "cdp": [],
        "sfps": {},
        "msg": ""
    }

    conn = None
    try:
        conn = ConnectHandler(**device)
        conn.send_command_timing("terminal length 0")

        cdp_out = conn.send_command("show cdp neighbors detail")
        loop_out = conn.send_command("show ip interface brief | include Loopback0")
        snmp_out = conn.send_command("show run | include snmp-server location")
        inv_out = conn.send_command("show inventory")

        resultado["loopback0_ip"] = parse_loopback(loop_out)
        resultado["snmp_location"] = parse_snmp_location(snmp_out)
        resultado["cdp"] = parse_cdp_detail(cdp_out)
        resultado["sfps"] = parse_inventory_sfps(inv_out)
        resultado["status"] = "ok"
    except Exception as e:
        resultado["msg"] = str(e)
    finally:
        if conn is not None:
            try:
                conn.disconnect()
            except Exception:
                pass
        # Descarte seguro imediato de credenciais deste escopo
        device.clear()
        del device

    return resultado


# -------------------------------------------------------------
# Formatação Rigorosa e Mesclagem de Células no Excel (OpenPyXL)
# -------------------------------------------------------------

def _header_to_col_index(ws, header_name: str) -> Optional[int]:
    for col in range(1, ws.max_column + 1):
        val = ws[f"{get_column_letter(col)}1"].value
        if val and str(val).strip().lower() == header_name.strip().lower():
            return col
    return None


def aplicar_estilo_base(excel_file: str):
    """
    Aplica rigorosamente o estilo padrão da planilha base:
    - Cabeçalho: Azul Royal (#0070C0), Fonte Aptos Narrow 9pt Negrito Branca
    - Dados: Fonte Arial 9pt, Bordas Finas (#D9D9D9), Alinhamento Vertical Centralizado
    - Largura fixa das colunas conforme a planilha modelo original
    """
    wb = load_workbook(excel_file)
    ws = wb.active

    # Larguras oficiais extraídas da planilha modelo
    larguras = {
        1: 12.0,   # Hostname A
        2: 15.0,   # IP Address A
        3: 12.0,   # Interface A
        4: 16.0,   # Midia Type A
        5: 14.0,   # Hostname B
        6: 15.0,   # IP Address B
        7: 12.0,   # Interface B
        8: 16.0,   # Midia Type B
        9: 35.0,   # Prédio
        10: 10.0,  # Andar
        11: 14.0,  # Bastidor
    }
    for col, w in larguras.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    header_fill = PatternFill(start_color="FF0070C0", end_color="FF0070C0", fill_type="solid")
    header_font = Font(name="Aptos Narrow", size=9, bold=True, color="FFFFFF")
    data_font = Font(name="Arial", size=9, bold=False, color="000000")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    # Formatar cabeçalho
    for col in range(1, 12):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = thin_border

    # Formatar linhas de dados
    for row in range(2, ws.max_row + 1):
        for col in range(1, 12):
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            cell.border = thin_border
            # Prédio alinhado à esquerda; demais colunas centralizadas
            if col == 9:
                cell.alignment = align_left
            else:
                cell.alignment = align_center

    wb.save(excel_file)


def merge_blocks_by_key(excel_file: str, key_col_name: str, cols_to_merge: List[str]):
    """Mescla células de blocos consecutivos na planilha final."""
    wb = load_workbook(excel_file)
    ws = wb.active

    key_col_idx = _header_to_col_index(ws, key_col_name)
    if key_col_idx is None:
        wb.save(excel_file)
        return

    merge_col_indices = []
    for name in cols_to_merge:
        idx = _header_to_col_index(ws, name)
        if idx is not None:
            merge_col_indices.append(idx)

    start = 2
    while start <= ws.max_row:
        key_val = ws.cell(row=start, column=key_col_idx).value
        end = start + 1
        while end <= ws.max_row and ws.cell(row=end, column=key_col_idx).value == key_val:
            end += 1

        if end - start > 1 and key_val is not None:
            ws.merge_cells(start_row=start, start_column=key_col_idx,
                           end_row=end - 1, end_column=key_col_idx)
            for c in merge_col_indices:
                ws.merge_cells(start_row=start, start_column=c,
                               end_row=end - 1, end_column=c)
        start = end

    wb.save(excel_file)


# -------------------------------------------------------------
# Índices de Busca CDP Bidirecional e SFPs
# -------------------------------------------------------------

def build_indexes(dados: Dict[str, Dict]):
    pairs_by_name = {}
    pairs_by_ip = {}
    loopbacks = {}
    locations = {}
    sfps_map = {}

    for host_key, obj in dados.items():
        src = label(host_key)
        loopbacks[src] = obj.get("loopback0_ip") or ""
        locations[src] = obj.get("snmp_location") or ""
        sfps_map[src] = obj.get("sfps") or {}

        for n in obj.get("cdp", []):
            neigh_name = n.get("neighbor") or ""
            neigh_ip = n.get("neighbor_ip") or ""
            li = n.get("local_intf") or ""
            ri = n.get("neighbor_intf") or ""
            dst = label(neigh_name)

            if src and dst:
                pairs_by_name[(src, dst)] = (li, ri)
            if src and neigh_ip:
                pairs_by_ip[(src, neigh_ip)] = (li, ri)

    return pairs_by_name, pairs_by_ip, loopbacks, locations, sfps_map


# -------------------------------------------------------------
# MODO 1: Processamento com Planilha Base
# -------------------------------------------------------------

def executar_modo_planilha(caminho_planilha_base: str, usuario: str, senha_ssh: str, max_threads: int = 8) -> Tuple[str, List[str]]:
    """Lê a planilha base informada, conecta nos switches e preenche as portas, IPs e Midia Type."""
    logs = []
    df = pd.read_excel(caminho_planilha_base)

    if "Hostname B" in df.columns:
        df["Hostname B"] = df["Hostname B"].ffill()

    hosts_a = set(df["Hostname A"].dropna().astype(str).unique()) if "Hostname A" in df.columns else set()
    hosts_b = set(df["Hostname B"].dropna().astype(str).unique()) if "Hostname B" in df.columns else set()
    hosts = {strip_domain(h) for h in (hosts_a | hosts_b) if strip_domain(h)}

    logs.append(f"Iniciando coleta paralela de {len(hosts)} switches identificados na planilha base.")

    dados_coleta = {}
    with ThreadPoolExecutor(max_workers=max_threads) as executor:
        futuros = {executor.submit(coletar_switch, h, usuario, senha_ssh): h for h in hosts}
        for fut in as_completed(futuros):
            h = futuros[fut]
            try:
                res = fut.result()
                dados_coleta[h] = res
                if res["status"] == "ok":
                    logs.append(f"✔ OK: {h} → {len(res['cdp'])} vizinhos CDP, {len(res['sfps'])} SFPs identificados")
                else:
                    logs.append(f"✖ Falha em {h}: {res['msg']}")
            except Exception as e:
                logs.append(f"✖ Erro ao executar {h}: {e}")

    pairs_by_name, pairs_by_ip, loopbacks, locations, sfps_map = build_indexes(dados_coleta)

    # Forçar colunas como texto
    for col in ["IP Address A", "IP Address B", "Interface A", "Interface B", "Midia Type", "Midia Type.1"]:
        if col in df.columns:
            df[col] = df[col].astype("object")

    for idx, row in df.iterrows():
        hostA = str(row.get("Hostname A") or "")
        hostB = str(row.get("Hostname B") or "")

        lblA = label(hostA)
        lblB = label(hostB)

        df.at[idx, "IP Address A"] = loopbacks.get(lblA, df.at[idx, "IP Address A"] if "IP Address A" in df.columns else "")
        df.at[idx, "IP Address B"] = loopbacks.get(lblB, df.at[idx, "IP Address B"] if "IP Address B" in df.columns else "")

        # Apenas se for switch preenche interface
        if not (is_switch_generico(hostA) and is_switch_generico(hostB)):
            continue

        ipA = str(df.at[idx, "IP Address A"] or "")
        ipB = str(df.at[idx, "IP Address B"] or "")

        # 1) A -> B por nome
        la, nb = pairs_by_name.get((lblA, lblB), (None, None))
        # 2) B -> A por nome (espelhado)
        if not la and not nb:
            li, ri = pairs_by_name.get((lblB, lblA), (None, None))
            if li or ri:
                la, nb = ri, li
        # 3) A -> B por IP
        if not la and not nb and ipB:
            la, nb = pairs_by_ip.get((lblA, ipB), (None, None))
        # 4) B -> A por IP (espelhado)
        if not la and not nb and ipA:
            li, ri = pairs_by_ip.get((lblB, ipA), (None, None))
            if li or ri:
                la, nb = ri, li

        if la:
            df.at[idx, "Interface A"] = la
            # Midia Type A a partir do show inventory
            sfp_a = sfps_map.get(lblA, {}).get(la, "")
            if sfp_a:
                df.at[idx, "Midia Type"] = sfp_a

        if nb:
            df.at[idx, "Interface B"] = nb
            # Midia Type B a partir do show inventory
            sfp_b = sfps_map.get(lblB, {}).get(nb, "")
            if sfp_b:
                df.at[idx, "Midia Type.1"] = sfp_b

    # Salva arquivo gerado
    out_id = str(uuid.uuid4())[:8]
    base_raiz = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
    out_dir = os.path.join(base_raiz, "core", "static", "downloads")
    os.makedirs(out_dir, exist_ok=True)
    caminho_out = os.path.join(out_dir, f"Port-Map-Preenchido-{out_id}.xlsx")
    df.to_excel(caminho_out, index=False)

    # Aplica estilização idêntica à planilha base
    aplicar_estilo_base(caminho_out)

    # Mescla células de Hostname B e metadados
    merge_blocks_by_key(
        excel_file=caminho_out,
        key_col_name="Hostname B",
        cols_to_merge=["IP Address B", "Prédio", "Andar", "Bastidor"]
    )

    logs.append(f"Planilha final gerada com sucesso: {os.path.basename(caminho_out)}")

    # Limpeza de memória
    usuario = None
    senha_ssh = None
    gc.collect()

    return caminho_out, logs


# -------------------------------------------------------------
# MODO 2: Auto-Discovery com Orientação Topológica (CORE/Distribuição = LADO A; Acesso = LADO B)
# -------------------------------------------------------------

def executar_modo_autodiscovery(cores_iniciais: List[str], usuario: str, senha_ssh: str, max_threads: int = 8) -> Tuple[str, List[str]]:
    """
    Auto-Discovery com cálculo de distância (BFS) a partir dos COREs:
    - O LADO A é sempre o switch mais próximo do CORE (ou o próprio CORE / switch cascateado primário).
    - O LADO B é estritamente o switch de acesso final (SW1 + 4 dígitos + 3 letras, ex: SW10901IFO).
    - Midia Type (colunas D e H) preenchido via modelo da SFP (show inventory).
    - Formatação estética e mesclagem 100% idênticas à planilha base.
    """
    logs = []
    switches_visitados: Set[str] = set()
    fila_switches: Set[str] = {strip_domain(c) for c in cores_iniciais if strip_domain(c)}
    distancias_ao_core: Dict[str, int] = {label(c): 0 for c in fila_switches}
    dados_coleta: Dict[str, Dict] = {}

    logs.append(f"Iniciando Auto-Discovery a partir dos COREs: {', '.join(fila_switches)}")

    # Crawler topológico (hop-by-hop para mapear toda a topologia e calcular a distância até o CORE)
    distancia_atual = 0
    while fila_switches:
        switches_atuais = list(fila_switches - switches_visitados)
        if not switches_atuais:
            break

        logs.append(f"--- Nível {distancia_atual + 1}: Coletando {len(switches_atuais)} switch(es) ---")

        novos_descobertos = set()
        with ThreadPoolExecutor(max_workers=max_threads) as executor:
            futuros = {executor.submit(coletar_switch, h, usuario, senha_ssh): h for h in switches_atuais}
            for fut in as_completed(futuros):
                h = futuros[fut]
                lbl_h = label(h)
                switches_visitados.add(lbl_h)
                try:
                    res = fut.result()
                    dados_coleta[h] = res
                    if res["status"] == "ok":
                        logs.append(f"✔ OK: {h} (Loopback: {res['loopback0_ip'] or 'N/A'}, SFPs: {len(res['sfps'])}, Location: '{res['snmp_location'] or 'N/A'}')")
                        for n in res["cdp"]:
                            vizinho = strip_domain(n.get("neighbor") or "")
                            lbl_viz = label(vizinho)
                            if vizinho and is_switch_generico(vizinho):
                                # Registra a menor distância ao CORE
                                if lbl_viz not in distancias_ao_core:
                                    distancias_ao_core[lbl_viz] = distancias_ao_core.get(lbl_h, 0) + 1
                                if lbl_viz not in switches_visitados:
                                    novos_descobertos.add(vizinho)
                    else:
                        logs.append(f"✖ Falha em {h}: {res['msg']}")
                except Exception as e:
                    logs.append(f"✖ Erro ao processar {h}: {e}")

        fila_switches = novos_descobertos
        distancia_atual += 1

    logs.append(f"Topologia completa mapeada! Total de {len(dados_coleta)} switches acessados.")

    # Constrói índices gerais
    pairs_by_name, pairs_by_ip, loopbacks, locations, sfps_map = build_indexes(dados_coleta)

    linhas_planilha = []
    pares_processados = set()

    for host_key, info in dados_coleta.items():
        src = label(host_key)
        for n in info.get("cdp", []):
            viz = label(n.get("neighbor") or "")

            # Ignora dispositivos que não sejam switches
            if not is_switch_generico(viz):
                continue

            # Evita duplicidade de enlaces bidirecionais
            par_id = tuple(sorted([src, viz]))
            if par_id in pares_processados:
                continue

            # REGRA 1 & 2: Determinação Topológica de Lado A e Lado B
            # Lado A deve ser o mais próximo do CORE (menor distância no grafo BFS)
            # Lado B deve ser o switch de acesso mais ao final da rede.
            dist_src = distancias_ao_core.get(src, 999)
            dist_viz = distancias_ao_core.get(viz, 999)

            if dist_src < dist_viz:
                sw_a_lbl, sw_b_lbl = src, viz
                nome_a, nome_b = host_key, n.get("neighbor") or viz
            elif dist_viz < dist_src:
                sw_a_lbl, sw_b_lbl = viz, src
                nome_a, nome_b = n.get("neighbor") or viz, host_key
            else:
                # Se ambos têm a mesma distância, prioriza como Lado B quem atender ao padrão SW1...
                if is_switch_acesso_b(viz) and not is_switch_acesso_b(src):
                    sw_a_lbl, sw_b_lbl = src, viz
                    nome_a, nome_b = host_key, n.get("neighbor") or viz
                else:
                    sw_a_lbl, sw_b_lbl = src, viz
                    nome_a, nome_b = host_key, n.get("neighbor") or viz

            # FILTRO ESTRITO DO LADO B:
            # O switch da ponta B deve ser estritamente do padrão SW1 + 4 dígitos + 3 letras (ex: SW10901IFO)
            if not is_switch_acesso_b(sw_b_lbl):
                continue

            pares_processados.add(par_id)

            # Busca interfaces locais e remotas correspondentes (A e B)
            la, nb = pairs_by_name.get((sw_a_lbl, sw_b_lbl), (None, None))
            if not la and not nb:
                li, ri = pairs_by_name.get((sw_b_lbl, sw_a_lbl), (None, None))
                if li or ri:
                    la, nb = ri, li  # Inverte para manter la = porta de A, nb = porta de B

            # Fallback caso venha apenas pelo CDP do host atual
            if not la and sw_a_lbl == src:
                la = n.get("local_intf") or ""
            if not nb and sw_b_lbl == viz:
                nb = n.get("neighbor_intf") or ""

            # REGRA 4: Obter Midia Type (SFP) de cada ponta
            midia_a = sfps_map.get(sw_a_lbl, {}).get(la, "") if la else ""
            midia_b = sfps_map.get(sw_b_lbl, {}).get(nb, "") if nb else ""

            # Fallback seguro para modelo padrão caso porta não utilize SFP óptica removível
            if not midia_a:
                midia_a = "SFP-10/25G-LR-S" if ("Te" in la or "Twe" in la) else "1000BASE-T"
            if not midia_b:
                midia_b = "SFP-10/25G-LR-S" if ("Te" in nb or "Twe" in nb) else "1000BASE-T"

            linhas_planilha.append({
                "Hostname A": nome_a,
                "IP Address A": loopbacks.get(sw_a_lbl, ""),
                "Interface A": la or "",
                "Midia Type": midia_a,
                "Hostname B": nome_b,
                "IP Address B": loopbacks.get(sw_b_lbl, n.get("neighbor_ip") or ""),
                "Interface B": nb or "",
                "Midia Type.1": midia_b,
                "Prédio": locations.get(sw_b_lbl, "") or locations.get(sw_a_lbl, ""),
                "Andar": "",
                "Bastidor": ""
            })

    # Ordena alfabeticamente por Hostname B e depois Hostname A (padrão da planilha base)
    linhas_planilha.sort(key=lambda x: (str(x["Hostname B"]), str(x["Hostname A"])))

    df = pd.DataFrame(linhas_planilha)

    # Garante todas as colunas mesmo se vazio
    colunas_obrigatorias = [
        "Hostname A", "IP Address A", "Interface A", "Midia Type",
        "Hostname B", "IP Address B", "Interface B", "Midia Type.1",
        "Prédio", "Andar", "Bastidor"
    ]
    for c in colunas_obrigatorias:
        if c not in df.columns:
            df[c] = ""
    df = df[colunas_obrigatorias]

    # Salva arquivo gerado
    out_id = str(uuid.uuid4())[:8]
    base_raiz = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
    out_dir = os.path.join(base_raiz, "core", "static", "downloads")
    os.makedirs(out_dir, exist_ok=True)
    caminho_out = os.path.join(out_dir, f"Port-Map-AutoDiscovery-{out_id}.xlsx")
    df.to_excel(caminho_out, index=False)

    # Aplica estilo idêntico ao modelo original
    aplicar_estilo_base(caminho_out)

    # Mescla de blocos por Hostname B
    merge_blocks_by_key(
        excel_file=caminho_out,
        key_col_name="Hostname B",
        cols_to_merge=["IP Address B", "Prédio", "Andar", "Bastidor"]
    )

    logs.append(f"Planilha de Auto-Discovery gerada com sucesso: {os.path.basename(caminho_out)} ({len(df)} enlaces mapeados)")

    # Limpeza de memória
    usuario = None
    senha_ssh = None
    gc.collect()

    return caminho_out, logs
