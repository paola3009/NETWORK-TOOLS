import openpyxl
import os
import re
from typing import Dict, Any, List, Tuple
from copy import copy
from openpyxl.styles import PatternFill, Font
try:
    from .config_store import get_site_config, load_configs
except ImportError:
    from config_store import get_site_config, load_configs

# Estilo de célula vermelha para alertas manuais
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
RED_FONT = Font(name="Calibri", size=10, color="9C0006", bold=True)

# Pools combinados de dados e voz (notebooks_desktops / telefoneIP_gateway_ATA sempre andam juntas)
COMBINED_DATA_VOICE_POOL = "notebooks_desktops / telefoneIP_gateway_ATA"

def normalize_vlan(val: Any) -> str:
    """Normaliza o valor de VLAN para string limpa sem casas decimais."""
    if val is None:
        return ""
    s = str(val).strip()
    if not s or s.lower() in ["none", "null"]:
        return ""
    try:
        if s.endswith(".0"):
            s = s[:-2]
        if s.isdigit():
            return str(int(s))
        if s.replace(".", "", 1).isdigit():
            return str(int(float(s)))
    except Exception:
        pass
    return s

def parse_vlan_list(vlan_str: Any) -> List[str]:
    """Converte uma string de VLANs (ex: '110,111' ou '110, 111') em lista de VLANs normalizadas."""
    if not vlan_str:
        return []
    parts = re.split(r'[,;\s]+', str(vlan_str))
    return [normalize_vlan(p) for p in parts if normalize_vlan(p)]

def normalize_pool_name(pool_name: str) -> str:
    """
    Se o pool for notebooks_desktops ou telefoneIP_gateway_ATA,
    garante que notebooks_desktops / telefoneIP_gateway_ATA sempre andem juntas.
    """
    p_clean = pool_name.strip()
    if p_clean in ["notebooks_desktops", "telefoneIP_gateway_ATA", COMBINED_DATA_VOICE_POOL]:
        return COMBINED_DATA_VOICE_POOL
    return p_clean

def get_model_port_count(model_name: str) -> int:
    """Retorna o número de portas base do modelo de switch."""
    model_upper = model_name.upper().strip()
    if "24UX" in model_upper or "24P" in model_upper or "24" in model_upper:
        return 24
    elif "48U" in model_upper or "48P" in model_upper or "48" in model_upper:
        return 48
    return 48

def has_redundant_power_supply(model_name: str) -> bool:
    """Retorna True se o modelo possui fonte redundante (ex: C9300-24UX)."""
    return "24UX" in model_name.upper()

class SwitchMember:
    def __init__(self, member_index: int, model: str):
        self.member_index = member_index
        self.model = model.strip()
        self.max_ports = get_model_port_count(model)
        self.has_redundant_power = has_redundant_power_supply(model)
        self.used_ports: List[int] = []

    def get_next_available_port(self) -> int:
        for p in range(1, self.max_ports + 1):
            if p not in self.used_ports:
                return p
        return -1

    def allocate_port(self, port_num: int) -> str:
        if port_num in self.used_ports:
            raise ValueError(f"Porta {port_num} já alocada no switch {self.member_index}")
        self.used_ports.append(port_num)
        return f"{self.member_index}/0/{port_num}"

class SwitchStack:
    def __init__(self, new_hostname: str, stack_models: List[str]):
        self.new_hostname = new_hostname
        self.members: List[SwitchMember] = [
            SwitchMember(idx + 1, model) for idx, model in enumerate(stack_models)
        ]

    def allocate_ap_port(self) -> Tuple[str, str]:
        ux_members = [m for m in self.members if m.has_redundant_power]
        for m in ux_members:
            p = m.get_next_available_port()
            if p != -1:
                formatted_port = m.allocate_port(p)
                return self.new_hostname, formatted_port

        for m in self.members:
            p = m.get_next_available_port()
            if p != -1:
                formatted_port = m.allocate_port(p)
                return self.new_hostname, formatted_port

        raise RuntimeError("Não há portas livres na pilha de switches para alocar o AP!")

    def allocate_camera_port(self, target_member_index: int) -> Tuple[str, str]:
        member = self.members[target_member_index - 1]
        p = member.get_next_available_port()
        if p != -1:
            formatted_port = member.allocate_port(p)
            return self.new_hostname, formatted_port
        
        for m in self.members:
            p = m.get_next_available_port()
            if p != -1:
                formatted_port = m.allocate_port(p)
                return self.new_hostname, formatted_port
        raise RuntimeError("Não há portas livres na pilha de switches para alocar Câmera!")

    def allocate_general_port(self, target_member_index: int) -> Tuple[str, str]:
        member = self.members[target_member_index - 1]
        p = member.get_next_available_port()
        if p != -1:
            formatted_port = member.allocate_port(p)
            return self.new_hostname, formatted_port

        for m in self.members:
            p = m.get_next_available_port()
            if p != -1:
                formatted_port = m.allocate_port(p)
                return self.new_hostname, formatted_port

        raise RuntimeError("Pilha de switches sem portas disponíveis para migração!")


def is_camera_row(row_data: Dict[str, Any], site_config: Dict[str, Any]) -> bool:
    """
    Identifica se uma linha é uma Câmera.
    Checa se a regra mapeada pelo site (VLAN Replace, VLAN Direta ou Vendor) possui device_type == 'camera'
    ou se o nome do pool possui 'camera' / 'cameras'.
    Também faz busca secundária por palavras-chave em descrição/vendor/ponto_rede.
    """
    vlan_dados = normalize_vlan(row_data.get("vlan_dados"))
    vendor = str(row_data.get("vendor", "") or "").strip().lower()

    # 1. Checa VLAN Replace
    vlan_replace = site_config.get("vlan_replace_pools", {})
    norm_vlan_replace = {normalize_vlan(k): v for k, v in vlan_replace.items()}
    if vlan_dados in norm_vlan_replace:
        rule = norm_vlan_replace[vlan_dados]
        dev_type = str(rule.get("device_type", "") or "").lower()
        pool_name = str(rule.get("pool_name", "") or "").lower()
        if dev_type in ["camera", "câmera", "cam"] or "camera" in pool_name or "cameras" in pool_name:
            return True

    # 2. Checa VLAN Direta
    vlan_direct = site_config.get("vlan_direct_pools", {})
    norm_vlan_direct = {normalize_vlan(k): v for k, v in vlan_direct.items()}
    if vlan_dados in norm_vlan_direct:
        rule = norm_vlan_direct[vlan_dados]
        if isinstance(rule, dict):
            dev_type = str(rule.get("device_type", "") or "").lower()
            pool_name = str(rule.get("pool_name", "") or "").lower()
            if dev_type in ["camera", "câmera", "cam"] or "camera" in pool_name or "cameras" in pool_name:
                return True
        elif isinstance(rule, str):
            if "camera" in rule.lower() or "cameras" in rule.lower():
                return True

    # 3. Checa Vendor Pools
    vendor_pools = site_config.get("vendor_pools", [])
    if vendor:
        for vp in vendor_pools:
            keywords = vp.get("keywords", [])
            if any(k.lower() in vendor for k in keywords):
                dev_type = str(vp.get("device_type", "") or vp.get("service_type", "") or "").lower()
                pool_name = str(vp.get("pool_name", "") or "").lower()
                if dev_type in ["camera", "câmera", "cam"] or "camera" in pool_name or "cameras" in pool_name:
                    return True

    # 4. Busca Secundária por Palavras-Chave
    keywords = site_config.get("camera_keywords", ["camera", "cam", "hikvision", "axis", "intelbras", "vivotek", "dahua", "cctv"])
    search_text = " ".join([
        str(row_data.get("descricao", "") or ""),
        str(row_data.get("vendor", "") or ""),
        str(row_data.get("ponto_rede", "") or "")
    ]).lower()
    for k in keywords:
        pattern = r'(?:\b|_)' + re.escape(k.lower()) + r'(?:\b|_)'
        if re.search(pattern, search_text):
            return True

    return False

def is_ap_row(row_data: Dict[str, Any], site_config: Dict[str, Any]) -> bool:
    """
    IMPORTANTE: APs NÃO SÃO buscados na planilha enviada pelo cliente.
    Os APs são informados exclusivamente pelo campo 'Novos APs a Adicionar' no formulário.
    """
    return False

def determine_ip_pool(
    row_data: Dict[str, Any],
    site_config: Dict[str, Any],
    default_vlan_dados: str = "",
    is_ap: bool = False
) -> Tuple[str, str, bool]:
    """
    Determina o IP Pool, a Observação (apenas se alerta manual) e se deve pintar a célula de vermelho.
    Retorna: (ip_pool, observacao, is_red_flag)
    """
    # 1. Se for um AP criado pelo formulário do usuário
    if is_ap:
        ap_pool = site_config.get("ap_pool", "ap_INFRA_VN")
        return ap_pool, "", False

    vendor = str(row_data.get("vendor", "") or "").strip()
    mac = str(row_data.get("mac", "") or "").strip()
    vlan_dados_raw = row_data.get("vlan_dados")
    vlan_voz_raw = row_data.get("vlan_voz")

    vlan_dados = normalize_vlan(vlan_dados_raw)
    target_default_vlans = parse_vlan_list(default_vlan_dados)
    if not target_default_vlans and site_config.get("default_vlan_dados"):
        target_default_vlans = parse_vlan_list(site_config.get("default_vlan_dados"))

    has_mac = bool(mac and mac.lower() not in ["none", "null", ""])

    # 1. PRIORITY 1: Checa Regras por Replace de VLAN (normalizando as chaves)
    vlan_replace = site_config.get("vlan_replace_pools", {})
    norm_vlan_replace = {normalize_vlan(k): v for k, v in vlan_replace.items()}
    if vlan_dados in norm_vlan_replace:
        rule = norm_vlan_replace[vlan_dados]
        rep_vlan = normalize_vlan(rule.get("replace_vlan", vlan_dados))
        raw_pool = rule.get("pool_name", f"POOL_VLAN_{rep_vlan}")
        pool_name = normalize_pool_name(raw_pool)
        return pool_name, "", False

    # 2. PRIORITY 2: Checa Regras Diretas por VLAN (normalizando as chaves)
    vlan_direct = site_config.get("vlan_direct_pools", {})
    norm_vlan_direct = {normalize_vlan(k): v for k, v in vlan_direct.items()}
    if vlan_dados in norm_vlan_direct:
        rule = norm_vlan_direct[vlan_dados]
        if isinstance(rule, dict):
            raw_pool = rule.get("pool_name", f"POOL_VLAN_{vlan_dados}")
        else:
            raw_pool = str(rule)
        pool_name = normalize_pool_name(raw_pool)
        return pool_name, "", False

    # 3. PRIORITY 3: Se a porta TEM Vendor conhecido cadastrado nas regras de marca do site
    vendor_pools = site_config.get("vendor_pools", [])
    if vendor:
        for vp in vendor_pools:
            keywords = vp.get("keywords", [])
            if any(k.lower() in vendor.lower() for k in keywords):
                raw_pool = vp.get("pool_name", "notebooks_desktops")
                pool_name = normalize_pool_name(raw_pool)
                return pool_name, "", False

    # 4. Se a porta NÃO TEM MAC e a VLAN de dados BATE com alguma das VLANs padrão de dados informadas (ex: "110,111")
    if not has_mac and target_default_vlans and (vlan_dados in target_default_vlans):
        return COMBINED_DATA_VOICE_POOL, "", False

    # 5. Se o Vendor for desconhecido (existe Vendor na planilha mas não bateu em NENHUMA regra de VLAN nem de Vendor)
    if vendor:
        unidentified_msg = "Dispositivo não identificado, analisar manualmente"
        return unidentified_msg, unidentified_msg, True

    # 6. Se a porta NÃO TEM MAC e a VLAN de dados NÃO bateu com nenhuma regra explícita (ex: VLAN 761)
    if not has_mac:
        vlan_unk_msg = "VLAN DESCONHECIDA, VERIFICAR MANUALMENTE"
        return vlan_unk_msg, vlan_unk_msg, True

    # 7. Fallback para casos sem NENHUMA regra
    unidentified_msg = "Dispositivo não identificado, analisar manualmente"
    return unidentified_msg, unidentified_msg, True

def find_header_row_and_columns(ws) -> Tuple[int, Dict[str, int]]:
    for r in range(1, min(ws.max_row + 1, 25)):
        row_values = [str(ws.cell(row=r, column=c).value or "").strip().lower() for c in range(1, ws.max_column + 1)]
        
        matches = 0
        if any("migra" in v for v in row_values): matches += 1
        if any("interface" in v or "origem" in v for v in row_values): matches += 1
        if any("vlan" in v for v in row_values): matches += 1
        if any("vendor" in v or "marca" in v for v in row_values): matches += 1

        if matches >= 3:
            col_map = {}
            for col_idx in range(1, ws.max_column + 1):
                val = str(ws.cell(row=r, column=col_idx).value or "").strip().lower()
                if not val:
                    continue
                if "migra" in val and "migrar" not in col_map:
                    col_map["migrar"] = col_idx
                elif ("interface" in val or "porta" in val) and "uptime" not in val and "oferta" not in val and "porta" not in col_map:
                    col_map["porta"] = col_idx
                elif "status" in val and "status" not in col_map:
                    col_map["status"] = col_idx
                elif ("description" in val or "descri" in val) and "descricao" not in col_map:
                    col_map["descricao"] = col_idx
                elif "mac" in val and "ip" not in val and "mac" not in col_map:
                    col_map["mac"] = col_idx
                elif ("ip address" in val or "ip do mac" in val) and "ip_mac" not in col_map:
                    col_map["ip_mac"] = col_idx
                elif ("vendor" in val or "marca" in val or "fabricante" in val) and "vendor" not in col_map:
                    col_map["vendor"] = col_idx
                elif ("vlan de dados" in val or "vlan dados" in val or "vlan data" in val or val == "vlan") and "vlan_dados" not in col_map:
                    col_map["vlan_dados"] = col_idx
                elif ("vlan de voz" in val or "vlan voz" in val or "vlan voice" in val or "voice vlan" in val) and "vlan_voz" not in col_map:
                    col_map["vlan_voz"] = col_idx
                elif "uptime" in val and "uptime" not in col_map:
                    col_map["uptime"] = col_idx
                elif ("ponto de rede" in val or "ponto" in val) and "ponto_rede" not in col_map:
                    col_map["ponto_rede"] = col_idx
                elif ("origem sw" in val or "hostname" in val) and "hostname" not in col_map:
                    col_map["hostname"] = col_idx

            return r, col_map

    return 1, {
        "hostname": 1, "porta": 2, "status": 3, "descricao": 4, "mac": 5,
        "ip_mac": 6, "vlan_dados": 7, "vlan_voz": 8, "uptime": 9, "ponto_rede": 11,
        "migrar": 12, "vendor": 15
    }

def update_excel_tables(ws, header_row: int):
    """
    Atualiza todas as tabelas formatadas do Excel (ws.tables) para expandir a faixa (ref)
    da coluna A até a coluna U (21) e até ws.max_row, inserindo as novas colunas no XML da tabela.
    Isso elimina o erro de reparo do Excel (/xl/tables/table1.xml) ao abrir a planilha.
    """
    if not hasattr(ws, "tables") or not ws.tables:
        return

    for t_name in list(ws.tables.keys()):
        try:
            tbl = ws.tables[t_name]
            start_col = "A"
            if hasattr(tbl, "ref") and tbl.ref:
                match = re.match(r'^([A-Z]+)', str(tbl.ref))
                if match:
                    start_col = match.group(1)

            tbl.ref = f"{start_col}{header_row}:U{ws.max_row}"

            existing_cols = list(tbl.tableColumns) if hasattr(tbl, "tableColumns") else []
            col_id = len(existing_cols) + 1

            for col_idx in range(len(existing_cols) + 1, 22):
                col_name = str(ws.cell(row=header_row, column=col_idx).value or f"Coluna_{col_idx}")
                tbl.tableColumns.append(TableColumn(id=col_id, name=col_name))
                col_id += 1
        except Exception as e:
            print(f"Aviso ao atualizar tabela {t_name}: {e}")


def process_de_para_excel(
    file_path: str,
    output_path: str,
    new_hostname: str,
    stack_models: List[str],
    new_aps_count: int,
    site_id: str = "default",
    default_vlan_dados: str = ""
) -> Dict[str, Any]:
    """
    Processa a planilha DE PARA Excel, aplicando todas as regras de distribuição e IP Pool.
    """
    site_config = get_site_config(site_id)
    all_configs = load_configs()

    merged_config = dict(site_config)
    merged_direct = dict(site_config.get("vlan_direct_pools", {}))
    merged_replace = dict(site_config.get("vlan_replace_pools", {}))
    merged_vendor = list(site_config.get("vendor_pools", []))

    for s_id, s_conf in all_configs.items():
        for k, v in s_conf.get("vlan_direct_pools", {}).items():
            if k not in merged_direct:
                merged_direct[k] = v
        for k, v in s_conf.get("vlan_replace_pools", {}).items():
            if k not in merged_replace:
                merged_replace[k] = v
        for vp in s_conf.get("vendor_pools", []):
            if vp not in merged_vendor:
                merged_vendor.append(vp)

    merged_config["vlan_direct_pools"] = merged_direct
    merged_config["vlan_replace_pools"] = merged_replace
    merged_config["vendor_pools"] = merged_vendor

    site_config = merged_config

    vlan_dados_final = default_vlan_dados.strip() or str(site_config.get("default_vlan_dados", "")).strip()

    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    header_row, col_map = find_header_row_and_columns(ws)

    ws.cell(row=header_row, column=18, value="Switch novo (hostname)")
    ws.cell(row=header_row, column=19, value="Porta")
    ws.cell(row=header_row, column=20, value="IP pool")
    ws.cell(row=header_row, column=21, value="Observação")

    col_migrar = col_map.get("migrar", 12)
    col_hostname = col_map.get("hostname", 1)
    col_porta = col_map.get("porta", 2)
    col_status = col_map.get("status", 3)
    col_desc = col_map.get("descricao", 4)
    col_mac = col_map.get("mac", 5)
    col_ip_mac = col_map.get("ip_mac", 6)
    col_vlan_dados = col_map.get("vlan_dados", 7)
    col_vlan_voz = col_map.get("vlan_voz", 8)
    col_ponto = col_map.get("ponto_rede", 11)
    col_vendor = col_map.get("vendor", 15)

    rows_to_migrate: List[Dict[str, Any]] = []
    
    for row_idx in range(header_row + 1, ws.max_row + 1):
        raw_migrar = ws.cell(row=row_idx, column=col_migrar).value
        migrar_val = str(raw_migrar or "").strip().lower()

        if migrar_val in ["sim", "s", "true", "1"] or raw_migrar is True:
            row_data = {
                "row_idx": row_idx,
                "hostname": ws.cell(row=row_idx, column=col_hostname).value,
                "porta": ws.cell(row=row_idx, column=col_porta).value,
                "status": ws.cell(row=row_idx, column=col_status).value,
                "descricao": ws.cell(row=row_idx, column=col_desc).value,
                "mac": ws.cell(row=row_idx, column=col_mac).value,
                "ip_mac": ws.cell(row=row_idx, column=col_ip_mac).value,
                "vendor": ws.cell(row=row_idx, column=col_vendor).value,
                "vlan_dados": ws.cell(row=row_idx, column=col_vlan_dados).value,
                "vlan_voz": ws.cell(row=row_idx, column=col_vlan_voz).value,
                "ponto_rede": ws.cell(row=row_idx, column=col_ponto).value,
            }
            rows_to_migrate.append(row_data)

    stack = SwitchStack(new_hostname, stack_models)

    ap_rows: List[Dict[str, Any]] = []
    camera_rows: List[Dict[str, Any]] = []
    general_rows: List[Dict[str, Any]] = []

    for r in rows_to_migrate:
        if is_ap_row(r, site_config):
            ap_rows.append(r)
        elif is_camera_row(r, site_config):
            camera_rows.append(r)
        else:
            general_rows.append(r)

    # Insere as novas linhas no TOPO da tabela se novos APs forem solicitados
    if new_aps_count > 0:
        sample_format_row = header_row + 1
        ws.insert_rows(header_row + 1, amount=new_aps_count)
        source_format_row = header_row + 1 + new_aps_count

        for i in range(1, new_aps_count + 1):
            top_r_idx = header_row + i
            for c_idx in range(1, 22):
                cell = ws.cell(row=top_r_idx, column=c_idx)
                ref_cell = ws.cell(row=source_format_row, column=c_idx)
                if ref_cell.has_style:
                    cell.font = copy(ref_cell.font)
                    cell.border = copy(ref_cell.border)
                    cell.fill = copy(ref_cell.fill)
                    cell.number_format = copy(ref_cell.number_format)
                    cell.alignment = copy(ref_cell.alignment)

    results_summary = {
        "total_linhas_lidas": ws.max_row - header_row,
        "total_migradas": len(rows_to_migrate),
        "total_aps": len(ap_rows) + new_aps_count,
        "total_cameras": len(camera_rows),
        "total_geral": len(general_rows),
        "novos_aps_adicionados": new_aps_count,
        "alocacoes": []
    }

    def apply_cell_result(row_i: int, sw: str, prt: str, pool: str, obs: str, is_red: bool):
        c_sw = ws.cell(row=row_i, column=18, value=sw)
        c_prt = ws.cell(row=row_i, column=19, value=prt)
        c_pool = ws.cell(row=row_i, column=20, value=pool)
        c_obs = ws.cell(row=row_i, column=21, value=obs)

        if is_red:
            c_pool.fill = RED_FILL
            c_pool.font = RED_FONT
            c_obs.fill = RED_FILL
            c_obs.font = RED_FONT

    row_offset = new_aps_count if new_aps_count > 0 else 0

    # 1. Novos APs solicitados pelo formulário -> C9300-24UX
    if new_aps_count > 0:
        ap_pool = site_config.get("ap_pool", "ap_INFRA_VN")
        for i in range(1, new_aps_count + 1):
            sw_name, new_port = stack.allocate_ap_port()
            top_row_idx = header_row + i

            ws.cell(row=top_row_idx, column=col_hostname, value=new_hostname)
            ws.cell(row=top_row_idx, column=col_desc, value=f"Novo Access Point #{i}")
            ws.cell(row=top_row_idx, column=col_migrar, value="SIM")
            
            apply_cell_result(top_row_idx, sw_name, new_port, ap_pool, "", False)

            results_summary["alocacoes"].append({
                "linha": top_row_idx,
                "descricao": f"Novo Access Point #{i}",
                "tipo": "Novo AP",
                "switch_novo": sw_name,
                "porta_nova": new_port,
                "ip_pool": ap_pool,
                "observacao": "",
                "is_unidentified": False
            })

    # 2. APs encontrados na planilha (se houver)
    for r in ap_rows:
        sw_name, new_port = stack.allocate_ap_port()
        ip_pool, obs, is_red = determine_ip_pool(r, site_config, vlan_dados_final, is_ap=True)
        obs_full = obs if is_red else ""
        
        target_r_idx = r["row_idx"] + row_offset
        apply_cell_result(target_r_idx, sw_name, new_port, ip_pool, obs_full, is_red)

        results_summary["alocacoes"].append({
            "linha": target_r_idx,
            "descricao": r["descricao"],
            "tipo": "AP",
            "switch_novo": sw_name,
            "porta_nova": new_port,
            "ip_pool": ip_pool,
            "observacao": obs_full,
            "is_unidentified": is_red
        })

    # 3. Câmeras -> Balanceadas na pilha
    num_switches = len(stack.members)
    if num_switches > 0 and camera_rows:
        base_cams_per_sw = len(camera_rows) // num_switches
        remainder = len(camera_rows) % num_switches

        cams_quota = [base_cams_per_sw] * num_switches

        ux_indices = [idx for idx, m in enumerate(stack.members) if m.has_redundant_power]
        non_ux_indices = [idx for idx, m in enumerate(stack.members) if not m.has_redundant_power]
        
        priority_indices = ux_indices + non_ux_indices
        for i in range(remainder):
            target_idx = priority_indices[i % len(priority_indices)]
            cams_quota[target_idx] += 1

        cam_index = 0
        for sw_idx, quota in enumerate(cams_quota, start=1):
            for _ in range(quota):
                if cam_index >= len(camera_rows):
                    break
                r = camera_rows[cam_index]
                sw_name, new_port = stack.allocate_camera_port(sw_idx)
                ip_pool, obs, is_red = determine_ip_pool(r, site_config, vlan_dados_final)
                obs_full = obs if is_red else ""

                target_r_idx = r["row_idx"] + row_offset
                apply_cell_result(target_r_idx, sw_name, new_port, ip_pool, obs_full, is_red)

                results_summary["alocacoes"].append({
                    "linha": target_r_idx,
                    "descricao": r["descricao"],
                    "tipo": "Câmera",
                    "switch_novo": sw_name,
                    "porta_nova": new_port,
                    "ip_pool": ip_pool,
                    "observacao": obs_full,
                    "is_unidentified": is_red
                })
                cam_index += 1

    # 4. Demais Dispositivos (Pontos Gerais) -> BALANCEADOS ENTRE OS SWITCHES FORA O C9300-24UX
    non_ux_members = [m for m in stack.members if not m.has_redundant_power]
    if non_ux_members:
        general_candidate_members = non_ux_members
    else:
        general_candidate_members = stack.members

    num_candidates = len(general_candidate_members)
    if num_candidates > 0 and general_rows:
        base_gen_per_sw = len(general_rows) // num_candidates
        gen_remainder = len(general_rows) % num_candidates

        gen_quotas = [base_gen_per_sw] * num_candidates
        for i in range(gen_remainder):
            gen_quotas[i] += 1

        gen_index = 0
        for cand_idx, quota in enumerate(gen_quotas):
            candidate_member = general_candidate_members[cand_idx]
            sw_idx = candidate_member.member_index

            for _ in range(quota):
                if gen_index >= len(general_rows):
                    break
                r = general_rows[gen_index]
                sw_name, new_port = stack.allocate_general_port(sw_idx)
                ip_pool, obs, is_red = determine_ip_pool(r, site_config, vlan_dados_final)
                obs_full = obs if is_red else ""

                target_r_idx = r["row_idx"] + row_offset
                apply_cell_result(target_r_idx, sw_name, new_port, ip_pool, obs_full, is_red)

                results_summary["alocacoes"].append({
                    "linha": target_r_idx,
                    "descricao": r["descricao"],
                    "tipo": "Geral",
                    "switch_novo": sw_name,
                    "porta_nova": new_port,
                    "ip_pool": ip_pool,
                    "observacao": obs_full,
                    "is_unidentified": is_red
                })
                gen_index += 1

    # Atualiza o XML das tabelas formatadas do Excel
    update_excel_tables(ws, header_row)

    wb.save(output_path)
    return results_summary
