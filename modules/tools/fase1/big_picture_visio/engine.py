# -*- coding: utf-8 -*-
"""
Motor Universal de Descoberta e Mapeamento de Topologia Multi-Vendor.
(Phase 1 - Assessment Big Picture Visio).

Suporta nativamente:
  - Huawei (VRP, CloudEngine, Switches Campus S5700/S6700/S12700, etc.)
  - Cisco (IOS, IOS-XE, NX-OS)
  - Datacom (DmOS, DmSwitch, DM4000, DM4100, DM4200, DM4300, DM4700, DM2100)
  - HP / Aruba (Comware, ProCurve, ArubaOS-S, ArubaOS-CX)
  - Juniper (Junos) & Dell (OS9, OS10)
  - Auto-Detecção inteligente com probe dinâmico por switch

Regras de Negócio:
  1. Access Points (APs) NUNCA são acessados via SSH e NUNCA são desenhados na topologia do Visio.
  2. APs conectados são contabilizados por switch para preencher a coluna 'Total de APs' na planilha Excel.
  3. Apenas Switches, Roteadores e WLCs são mapeados, conectados e catalogados.
  4. Credenciais são descartadas imediatamente da memória após a execução (del e gc.collect()).
"""

import os
import re
import time
import gc
import uuid
from typing import Dict, List, Set, Tuple, Any, Optional
from datetime import datetime

from netmiko import ConnectHandler, NetmikoAuthenticationException, NetmikoTimeoutException

from modules.tools.fase1.big_picture_visio import excel_exporter
from modules.tools.fase1.big_picture_visio import visio_exporter

# Expressões regulares para detecção de Access Points (Cisco, Huawei, Aruba, etc.)
AP_PLATFORM_RE = re.compile(
    r"(AIR-|C91\d\d|AP\d\d|CAP\d\d|Cisco\s+AP|Access\s*Point|AIR-AP|AIR-CAP|IW-|Catalyst\s+91|"
    r"AirEngine|AP40\d\d|AP50\d\d|AP60\d\d|AP70\d\d|AP80\d\d|WA\d\d|Huawei\s+AP|"
    r"ArubaOS\s*\(MODEL:\s*\d+|ArubaAP|IAP-\d+|Instant\s+AP|AP-\d\d\d)",
    re.IGNORECASE
)
AP_HOSTNAME_RE = re.compile(
    r"^(AP|WAP|CAP|AIR|AP-|_AP|AP_|[A-Z0-9]+-AP-)",
    re.IGNORECASE
)
AP_CAPABILITY_RE = re.compile(
    r"(WLAN\s+Access\s+Point|Translan\s+Bridge|wlan-access-point|wlan)",
    re.IGNORECASE
)

# Expressões regulares para Roteadores e WLCs
ROUTER_PLATFORM_RE = re.compile(
    r"(ISR|ASR|CSR|C8000|C8200|C8300|C8500|19\d\d|29\d\d|39\d\d|43\d\d|44\d\d|"
    r"AR\d\d|NetEngine|NE\d\d|Router|DM8\d\d|MX\d\d)",
    re.IGNORECASE
)
WLC_PLATFORM_RE = re.compile(
    r"(WLC|AIR-CT|9800|5520|3504|8540|2504|5508|AC6\d\d\d|AC\s+Controller|Aruba7\d\d\d)",
    re.IGNORECASE
)


def normalizar_hostname(nome: str) -> str:
    """Limpa símbolos de prompt Cisco (#, >), Huawei (<, >, [, ]) e domínios FQDN."""
    if not nome:
        return ""
    # Remove caracteres de prompt no início e fim
    nome = re.sub(r"^[<\[]+|[>\]#]+$", "", nome.strip())
    # Remove domínio
    nome = nome.split(".")[0].strip()
    return nome.upper()


def is_access_point(device_id: str, platform: str, capabilities: str) -> bool:
    """Verifica se o vizinho CDP/LLDP é um Access Point."""
    if AP_PLATFORM_RE.search(platform):
        return True
    if AP_HOSTNAME_RE.search(device_id):
        return True
    if AP_CAPABILITY_RE.search(capabilities) and "Switch" not in capabilities and "Router" not in capabilities and "bridge router" not in capabilities:
        return True
    return False


def extrair_modelo_e_portas(pid_str: str) -> Tuple[str, int]:
    """
    Estima a oferta de portas por membro de stack com base no modelo.
    Suporta Cisco (ex: C9300-48U), Huawei (ex: S5720-52X), Datacom (ex: DM4370-48X),
    HP/Aruba (ex: 2930F-48G), Juniper (ex: EX4300-48T) e Dell (ex: N3048P).
    """
    pid_str = pid_str.upper().strip()
    portas = 0

    if re.search(r"[-_ ]?(48|52)[A-Z0-9-]*", pid_str):
        portas = 48
    elif re.search(r"[-_ ]?(24|28)[A-Z0-9-]*", pid_str):
        portas = 24
    elif re.search(r"[-_ ]?(32)[A-Z0-9-]*", pid_str):
        portas = 32
    elif re.search(r"[-_ ]?(16)[A-Z0-9-]*", pid_str):
        portas = 16
    elif re.search(r"[-_ ]?(8|10)[A-Z0-9-]*", pid_str):
        portas = 8

    return pid_str, portas


# ==============================================================================
# PARSERS UNIVERSAIS (LLDP & CDP)
# ==============================================================================

def parse_generic_lldp_detail(saida: str) -> List[Dict[str, str]]:
    """
    Parser universal de LLDP suportando:
      - Huawei (display lldp neighbor verbose)
      - Cisco (show lldp neighbors detail)
      - Datacom (show lldp neighbors detail)
      - HP / Aruba (show lldp info remote-device detail)
    """
    vizinhos = []

    # 1. Tenta formato Huawei: "10GE1/0/1 has 1 neighbor(s):"
    if "has" in saida and "neighbor" in saida:
        blocos = re.split(r"(\S+)\s+has\s+\d+\s+neighbor", saida)
        for i in range(1, len(blocos), 2):
            local_port = blocos[i].strip()
            corpo = blocos[i + 1]

            sys_name = ""
            m_name = re.search(r"(?:System name|SysName|Device ID)\s*:\s*([^\r\n]+)", corpo, re.IGNORECASE)
            if m_name:
                sys_name = m_name.group(1).strip()
            else:
                m_chid = re.search(r"Chassis ID\s*:\s*([^\r\n]+)", corpo, re.IGNORECASE)
                if m_chid:
                    sys_name = m_chid.group(1).strip()

            port_id = ""
            m_pid = re.search(r"(?:Port ID|PortId|Port name)\s*:\s*([^\r\n]+)", corpo, re.IGNORECASE)
            if m_pid:
                port_id = m_pid.group(1).strip()

            ip = ""
            m_ip = re.search(r"Management address[^\r\n:]*:\s*([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)", corpo, re.IGNORECASE)
            if m_ip:
                ip = m_ip.group(1).strip()

            desc = ""
            m_desc = re.search(r"System description\s*:\s*(.*?)(System capabilities|Management address|Expired time|$)", corpo, re.IGNORECASE | re.DOTALL)
            if m_desc:
                desc = " ".join(m_desc.group(1).split())

            cap = ""
            m_cap = re.search(r"System capabilities enabled\s*:\s*([^\r\n]+)", corpo, re.IGNORECASE)
            if m_cap:
                cap = m_cap.group(1).strip()

            if sys_name:
                vizinhos.append({
                    "device_id": sys_name,
                    "ip": ip,
                    "platform": desc,
                    "capabilities": cap,
                    "local_port": local_port,
                    "remote_port": port_id,
                    "protocol": "LLDP"
                })
        if vizinhos:
            return vizinhos

    # 2. Formato Cisco, Datacom, HP/Aruba
    blocos = re.split(r"(?:Local Port\s*:|Local Interface\s*:|------------------------------------------------|-------------------------+)", saida)
    for b in blocos:
        if not b.strip():
            continue

        sys_name = ""
        m_name = re.search(r"(?:SysName|System Name|Device ID)\s*:\s*([^\r\n]+)", b, re.IGNORECASE)
        if m_name:
            sys_name = m_name.group(1).strip()
        else:
            m_chid = re.search(r"Chassis\s*Id\s*:\s*([^\r\n]+)", b, re.IGNORECASE)
            if m_chid:
                sys_name = m_chid.group(1).strip()

        local_p = ""
        m_lp = re.search(r"^\s*([^\r\n,]+)", b)
        if m_lp:
            local_p = m_lp.group(1).strip()

        port_id = ""
        m_pid = re.search(r"(?:PortId|Port ID|Port id|Port ID\s*\([^)]*\))\s*:\s*([^\r\n]+)", b, re.IGNORECASE)
        if m_pid:
            port_id = m_pid.group(1).strip()

        ip = ""
        m_ip = re.search(r"(?:Management Address|Management Addresses:\s*IP:\s*|IP address|IP)\s*:\s*([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)", b, re.IGNORECASE)
        if m_ip:
            ip = m_ip.group(1).strip()

        desc = ""
        m_desc = re.search(r"(?:System Descr|System Description|Platform)\s*:\s*([^\r\n]+)", b, re.IGNORECASE)
        if m_desc:
            desc = m_desc.group(1).strip()

        cap = ""
        m_cap = re.search(r"(?:System Capabilities Enabled|System Capabilities|Capabilities Enabled|Capabilities)\s*:\s*([^\r\n]+)", b, re.IGNORECASE)
        if m_cap:
            cap = m_cap.group(1).strip()

        if sys_name:
            vizinhos.append({
                "device_id": sys_name,
                "ip": ip,
                "platform": desc,
                "capabilities": cap,
                "local_port": local_p,
                "remote_port": port_id,
                "protocol": "LLDP"
            })

    return vizinhos


def parse_cisco_cdp_detail(saida: str) -> List[Dict[str, str]]:
    """Analisa 'show cdp neighbors detail' de switches Cisco e HP ProCurve."""
    vizinhos = []
    blocos = re.split(r"-------------------------", saida)

    for b in blocos:
        if not b.strip():
            continue

        dev_id = ""
        m_dev = re.search(r"Device ID:\s*([^\r\n]+)", b, re.IGNORECASE)
        if m_dev:
            dev_id = m_dev.group(1).strip()

        ip = ""
        m_ip = re.search(r"IP address:\s*([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)", b, re.IGNORECASE)
        if m_ip:
            ip = m_ip.group(1).strip()

        platform = ""
        m_plat = re.search(r"Platform:\s*([^,\r\n]+)", b, re.IGNORECASE)
        if m_plat:
            platform = m_plat.group(1).strip()

        capabilities = ""
        m_cap = re.search(r"Capabilities:\s*([^\r\n]+)", b, re.IGNORECASE)
        if m_cap:
            capabilities = m_cap.group(1).strip()

        local_port = ""
        remote_port = ""
        m_ports = re.search(r"Interface:\s*([^,\r\n]+),\s*Port ID\s*\([^)]*\):\s*([^\r\n]+)", b, re.IGNORECASE)
        if m_ports:
            local_port = m_ports.group(1).strip()
            remote_port = m_ports.group(2).strip()

        if dev_id:
            vizinhos.append({
                "device_id": dev_id,
                "ip": ip,
                "platform": platform,
                "capabilities": capabilities,
                "local_port": local_port,
                "remote_port": remote_port,
                "protocol": "CDP"
            })

    return vizinhos


def parse_huawei_lldp_brief(saida: str) -> List[Dict[str, str]]:
    """
    Analisa a saída tabular de 'display lldp neighbor brief' de switches Huawei:
    Local Intf         Neighbor Dev             Neighbor Intf             Exptime(s)
    XGE0/0/1           CGNAT-RJ2                sfp-sfpplus1              90
    100GE0/0/4         SW-CORE-ROCA             100GE0/0/2                118
    100GE0/0/5         PE00-ARISTA-7060x-100... Eth7/1                    92
    """
    vizinhos = []
    linhas = saida.splitlines()
    in_table = False

    for l in linhas:
        l_str = l.strip()
        if not l_str:
            continue
        if "Local Intf" in l_str and "Neighbor Dev" in l_str:
            in_table = True
            continue
        if in_table:
            if l_str.startswith("---") or l_str.startswith("==="):
                continue
            if l_str.startswith("<") or l_str.startswith("["):
                break

            parts = l_str.split()
            if len(parts) >= 4:
                local_intf = parts[0]
                neigh_dev = parts[1]
                neigh_intf = " ".join(parts[2:-1]).strip()

                vizinhos.append({
                    "device_id": neigh_dev,
                    "local_port": local_intf,
                    "remote_port": neigh_intf,
                    "ip": "",
                    "platform": "Dispositivo de Rede",
                    "capabilities": "bridge router",
                    "protocol": "LLDP"
                })

    return vizinhos


def classificar_dispositivo(dev_id: str, platform: str, capabilities: str, is_seed_core: bool = False) -> str:
    """Classifica a função do equipamento na topologia."""
    if is_seed_core:
        return "Core"
    dev_upper = dev_id.upper()

    if WLC_PLATFORM_RE.search(platform) or WLC_PLATFORM_RE.search(dev_id):
        return "WLC"
    if ROUTER_PLATFORM_RE.search(platform) or (("Router" in capabilities) and ("Switch" not in capabilities)):
        return "Router"
    if any(k in dev_upper for k in ["CGNAT", "PPPOE", "BNG", "BRAS", "FIREWALL", "FW", "RTR", "ROUTER", "BORDER", "EDGE", "PE-", "PE0", "PE1", "PE2", "P0"]):
        return "Router"
    if "CORE" in dev_upper:
        return "Core"
    return "Switch"


class BigPictureCrawler:
    """
    Rastreador e analisador recursivo de topologia multicamadas.
    Compatível com Cisco, Huawei, Datacom, HP/Aruba, Juniper e Dell.
    """

    def __init__(
        self,
        core_ips: List[str],
        usuario: str,
        senha: str,
        secret: str = "",
        vendor_preference: str = "auto",  # 'auto', 'huawei', 'cisco', 'datacom', 'hp'
        porta: int = 22,
        max_depth: int = 10,
        timeout: int = 10
    ):
        self.core_ips = [ip.strip() for ip in core_ips if ip.strip()]
        self.usuario = usuario
        self.senha = senha
        self.secret = secret or senha
        self.vendor_preference = vendor_preference.lower()
        self.porta = porta
        self.max_depth = max_depth
        self.timeout = timeout

        self.devices: Dict[str, Dict[str, Any]] = {}
        self.links: List[Dict[str, Any]] = []
        self.visited_ips: Set[str] = set()
        self.visited_hostnames: Set[str] = set()
        self.logs: List[str] = []

    def log(self, mensagem: str):
        hora = datetime.now().strftime("%H:%M:%S")
        txt = f"[{hora}] {mensagem}"
        self.logs.append(txt)

    def executar(self) -> Dict[str, Any]:
        """Executa a descoberta recursiva de todos os níveis de cascateamento."""
        vendor_desc = "Universal Auto-Detect (Cisco, Huawei, HP, Datacom...)" if self.vendor_preference == "auto" else self.vendor_preference.upper()
        self.log(f"Iniciando varredura Big Picture. Perfil: {vendor_desc}")
        self.log(f"Cores semente informados: {', '.join(self.core_ips)}")
        self.log("Regra de negócio ativa: Access Points serão filtrados e NÃO receberão conexões SSH.")

        fila: List[Tuple[str, str, int]] = []
        for ip_core in self.core_ips:
            fila.append((ip_core, "Core", 0))

        while fila:
            ip_alvo, role_sugerido, nivel = fila.pop(0)

            if ip_alvo in self.visited_ips:
                continue
            if nivel > self.max_depth:
                self.log(f"Limite máximo de profundidade ({self.max_depth} saltos) atingido para {ip_alvo}. Ignorando.")
                continue

            self.visited_ips.add(ip_alvo)

            self.log(f"--> [Nível {nivel}] Conectando ao switch {ip_alvo} ({role_sugerido})...")
            dados_switch = self._coletar_switch(ip_alvo, role_sugerido, nivel)

            if not dados_switch:
                self.log(f"Falha de conexão com {ip_alvo}. Prosseguindo para próximos equipamentos.")
                continue

            hostname_canonico = dados_switch["id"]
            self.devices[hostname_canonico] = dados_switch
            self.visited_hostnames.add(hostname_canonico)

            # Processa vizinhos para expandir a topologia
            for viz in dados_switch.get("vizinhos_validos", []):
                viz_ip = viz.get("ip")
                viz_id = viz.get("hostname_canonico")
                viz_role = viz.get("role")

                self.links.append({
                    "source_id": hostname_canonico,
                    "source_port": viz.get("local_port", ""),
                    "target_id": viz_id,
                    "target_port": viz.get("remote_port", ""),
                    "media_type": viz.get("protocol", "LLDP")
                })

                if viz_id not in self.devices:
                    self.devices[viz_id] = {
                        "id": viz_id,
                        "hostname": viz.get("device_id", viz_id),
                        "ip": viz_ip or "N/D",
                        "role": viz_role,
                        "modelo": viz.get("platform", "Desconhecido"),
                        "stack_count": 1,
                        "total_portas": 0,
                        "total_aps": 0,
                        "nivel": nivel + 1,
                        "status": "Descoberto via Vizinho"
                    }

                # Se for Switch ou Core com IP alcançável ainda não visitado, coloca na fila
                if viz_role in ("Switch", "Core") and viz_ip and (viz_ip not in self.visited_ips):
                    fila.append((viz_ip, viz_role, nivel + 1))

        self.log(f"Varredura concluída! Total de dispositivos de rede mapeados: {len(self.devices)}")
        self.log(f"Total de enlaces (links) entre equipamentos: {len(self.links)}")

        # Gera Planilha Excel e Diagrama Visio
        switches_para_planilha = [
            dev for dev in self.devices.values()
            if dev.get("role") in ("Core", "Switch")
        ]

        self.log("Gerando Planilha Excel com colunas de capacidade, stack e contagem de APs...")
        caminho_excel = excel_exporter.gerar_planilha_inventario_switches(switches_para_planilha, links=self.links)
        self.log(f"Planilha Excel gerada: {os.path.basename(caminho_excel)}")

        self.log("Gerando Topologia Microsoft Visio (.vsdx)...")
        caminho_visio = visio_exporter.gerar_topologia_visio(
            list(self.devices.values()),
            self.links
        )
        self.log(f"Diagrama Visio (.vsdx) gerado: {os.path.basename(caminho_visio)}")

        stats = {
            "total_dispositivos": len(self.devices),
            "total_switches": len(switches_para_planilha),
            "total_routers": sum(1 for d in self.devices.values() if d.get("role") == "Router"),
            "total_wlcs": sum(1 for d in self.devices.values() if d.get("role") == "WLC"),
            "total_aps_conectados": sum(d.get("total_aps", 0) for d in switches_para_planilha),
            "total_portas_ofertadas": sum(d.get("total_portas", 0) for d in switches_para_planilha),
            "total_links": len(self.links)
        }

        return {
            "success": True,
            "excel_path": caminho_excel,
            "visio_path": caminho_visio,
            "devices": list(self.devices.values()),
            "links": self.links,
            "stats": stats,
            "logs": self.logs
        }

    def _coletar_switch(self, ip: str, role_sugerido: str, nivel: int) -> Optional[Dict[str, Any]]:
        """
        Conecta ao switch detectando automaticamente o fabricante:
        Huawei, Cisco, Datacom, HP/Aruba, Juniper ou Dell.
        """
        # Lista ordenada de drivers SSH para tentativa
        if self.vendor_preference == "huawei":
            drivers = ["huawei", "cisco_ios"]
        elif self.vendor_preference == "cisco":
            drivers = ["cisco_ios", "huawei"]
        elif self.vendor_preference == "datacom":
            drivers = ["cisco_ios", "huawei"]
        elif self.vendor_preference == "hp":
            drivers = ["hp_comware", "hp_procurve", "aruba_aoscx", "cisco_ios"]
        else:
            # Auto: começa com huawei e cisco_ios
            drivers = ["huawei", "cisco_ios", "hp_procurve", "juniper_junos"]

        net_connect = None
        driver_ativo = None

        for d in drivers:
            params = {
                "device_type": d,
                "host": ip,
                "username": self.usuario,
                "password": self.senha,
                "secret": self.secret,
                "port": self.porta,
                "timeout": self.timeout,
                "session_timeout": self.timeout * 2,
                "global_delay_factor": 0.8,
            }
            try:
                net_connect = ConnectHandler(**params)
                driver_ativo = d
                break
            except (NetmikoAuthenticationException, NetmikoTimeoutException):
                # Se for timeout ou credencial errada, não fica tentando os outros
                return None
            except Exception:
                continue

        if not net_connect:
            return None

        try:
            raw_prompt = net_connect.find_prompt()
            hostname = normalizar_hostname(raw_prompt)

            # Probe para detecção precisa do Sistema Operacional / Fabricante
            vendor = self._identificar_fabricante(net_connect, raw_prompt, driver_ativo)

            self.log(f"Switch identificado: '{hostname}' ({ip}) -> Fabricante: {vendor}")

            if vendor in ("HUAWEI", "HP_COMWARE"):
                return self._coletar_huawei(net_connect, ip, hostname, role_sugerido, nivel)
            elif vendor == "DATACOM":
                return self._coletar_datacom(net_connect, ip, hostname, role_sugerido, nivel)
            elif vendor in ("HP_PROCURVE", "ARUBA"):
                return self._coletar_hp_aruba(net_connect, ip, hostname, role_sugerido, nivel)
            else:
                # Padrão Cisco IOS / NX-OS
                return self._coletar_cisco(net_connect, ip, hostname, role_sugerido, nivel)

        except Exception as err:
            self.log(f"Erro ao coletar dados de [{ip}]: {err}")
            return None
        finally:
            try:
                net_connect.disconnect()
            except Exception:
                pass

    def _identificar_fabricante(self, net_connect, raw_prompt: str, driver: str) -> str:
        """Determina o fabricante através do prompt e resposta a comandos de versão."""
        # Indicadores fortes no prompt
        if "<" in raw_prompt or ">" in raw_prompt or "[" in raw_prompt:
            # Huawei VRP ou HP Comware
            try:
                out_disp = net_connect.send_command("display version", read_timeout=10)
                if "HUAWEI" in out_disp.upper() or "VRP" in out_disp.upper():
                    return "HUAWEI"
                if "COMWARE" in out_disp.upper() or "H3C" in out_disp.upper():
                    return "HP_COMWARE"
            except Exception:
                return "HUAWEI"
            return "HUAWEI"

        # Se o prompt tiver #, envia show version para identificar Cisco, Datacom, Aruba
        try:
            out_ver = net_connect.send_command("show version", read_timeout=10)
            out_upper = out_ver.upper()
            if "DATACOM" in out_upper or "DMOS" in out_upper or "DMSWITCH" in out_upper:
                return "DATACOM"
            if "ARUBA" in out_upper or "PROCURVE" in out_upper:
                return "ARUBA"
            if "CISCO" in out_upper:
                return "CISCO"
            if "JUNOS" in out_upper:
                return "JUNIPER"
        except Exception:
            pass

        return "CISCO" if driver == "cisco_ios" else "HUAWEI"

    def _coletar_huawei(self, net_connect, ip: str, hostname: str, role_sugerido: str, nivel: int) -> Dict[str, Any]:
        """Executa comandos VRP no switch Huawei ou HP Comware."""
        try:
            net_connect.send_command("screen-length 0 temporary", read_timeout=10)
        except Exception:
            pass

        out_version = net_connect.send_command("display version", read_timeout=20)
        out_stack = net_connect.send_command("display stack", read_timeout=15)
        out_device = net_connect.send_command("display device", read_timeout=15)

        # 1. Stack e Membros
        membros_stack = []
        for line in out_stack.splitlines():
            m = re.match(r"^\s*(\d+)\s+(Master|Standby|Member)\s+([0-9a-fA-F.-]+)\s+\d+\s+(\S+)", line, re.IGNORECASE)
            if m:
                membros_stack.append(m.group(4).strip())

        stack_count = len(membros_stack) if membros_stack else 1
        modelos_encontrados = list(membros_stack)

        if not modelos_encontrados:
            m_ver = re.search(r"HUAWEI\s+([A-Za-z0-9_-]+)", out_version, re.IGNORECASE)
            if m_ver:
                modelos_encontrados.append(m_ver.group(1))
            else:
                m_sub = re.search(r"Version\s+\S+\s*\(([A-Za-z0-9_-]+)", out_version)
                if m_sub:
                    modelos_encontrados.append(m_sub.group(1))

        # 2. Portas
        soma_portas = 0
        pids_limpos = []
        for m_str in modelos_encontrados:
            m_clean, p_count = extrair_modelo_e_portas(m_str)
            pids_limpos.append(m_clean)
            if p_count > 0:
                soma_portas += p_count

        if soma_portas == 0:
            soma_portas = 48 * stack_count

        modelo_final = ", ".join(sorted(set(pids_limpos))) if pids_limpos else "Huawei S-Series"

        # 3. Vizinhos LLDP - Executa display lldp neighbor brief (formato padrão do Huawei S6730/S5730)
        self.log(f"Coletando vizinhos LLDP em '{hostname}' via 'display lldp neighbor brief'...")
        out_brief = net_connect.send_command("display lldp neighbor brief", read_timeout=25)
        vizinhos_brief = parse_huawei_lldp_brief(out_brief)

        # Executa display lldp neighbor para obter IPs de gerência e detalhes adicionais
        out_detail = net_connect.send_command("display lldp neighbor", read_timeout=30)
        vizinhos_detail = parse_generic_lldp_detail(out_detail)

        # Mapa de IP, portas e capacidades por hostname e por porta local
        detalhes_map_host = {}
        detalhes_map_port = {}
        for d in vizinhos_detail:
            d_id = normalizar_hostname(d.get("device_id", ""))
            d_port = d.get("local_port", "").strip().upper()
            if d_id:
                detalhes_map_host[d_id] = d
            if d_port:
                detalhes_map_port[d_port] = d

        # Também consulta tabela ARP para associar IP caso o LLDP não forneça Management Address
        arp_map = {}
        try:
            out_arp = net_connect.send_command("display arp", read_timeout=20)
            for line in out_arp.splitlines():
                m_arp = re.search(r"([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)\s+[0-9a-fA-F.-]+\s+(?:\d+|-)\s+\S+\s+(\S+)", line)
                if m_arp:
                    arp_ip, arp_intf = m_arp.group(1), m_arp.group(2)
                    arp_map[arp_intf.strip().upper()] = arp_ip
        except Exception:
            pass

        # Combina os vizinhos do brief com os detalhes do detail e ARP
        vizinhos_brutos = []
        if vizinhos_brief:
            for vb in vizinhos_brief:
                dev_canonico = normalizar_hostname(vb.get("device_id", ""))
                lport = vb.get("local_port", "").strip().upper()

                # Tenta enriquecer primeiro por porta local, depois por hostname
                det = detalhes_map_port.get(lport) or detalhes_map_host.get(dev_canonico)
                if det:
                    # Desfaz eventual truncamento de hostname feito pelo brief (ex: reticências ...)
                    det_dev = det.get("device_id", "")
                    if det_dev and not det_dev.endswith("..."):
                        vb["device_id"] = det_dev
                    if det.get("ip"):
                        vb["ip"] = det.get("ip")
                    if det.get("platform"):
                        vb["platform"] = det.get("platform")
                    if det.get("capabilities"):
                        vb["capabilities"] = det.get("capabilities")
                    if det.get("remote_port") and not vb.get("remote_port"):
                        vb["remote_port"] = det.get("remote_port")

                # Se ainda não tem IP, consulta o ARP na porta local
                if not vb.get("ip"):
                    ip_arp = arp_map.get(lport)
                    if ip_arp:
                        vb["ip"] = ip_arp
                        self.log(f"IP de '{vb['device_id']}' localizado via ARP na porta {lport}: {ip_arp}")

                vizinhos_brutos.append(vb)
        elif vizinhos_detail:
            vizinhos_brutos = vizinhos_detail

        self.log(f"Vizinhos brutos LLDP detectados em '{hostname}': {len(vizinhos_brutos)}")

        total_aps = 0
        vizinhos_validos = []
        vizinhos_processados = set()

        for v in vizinhos_brutos:
            v_dev_id = v.get("device_id", "")
            v_plat = v.get("platform", "")
            v_cap = v.get("capabilities", "")

            # AP FILTER: conta e não acessa nem desenha no Visio
            if is_access_point(v_dev_id, v_plat, v_cap):
                total_aps += 1
                continue

            v_canonico = normalizar_hostname(v_dev_id)
            if not v_canonico or v_canonico == hostname:
                continue

            par_id = (v_canonico, v.get("local_port"))
            if par_id in vizinhos_processados:
                continue
            vizinhos_processados.add(par_id)

            v_role = classificar_dispositivo(v_canonico, v_plat, v_cap)
            v["hostname_canonico"] = v_canonico
            v["role"] = v_role
            vizinhos_validos.append(v)

        self.log(f"OK [Huawei]: '{hostname}' coletado. Stack: {stack_count} membro(s) | Oferta Portas: {soma_portas} | APs conectados: {total_aps}")

        return {
            "id": hostname,
            "hostname": hostname,
            "ip": ip,
            "role": role_sugerido,
            "modelo": modelo_final,
            "stack_count": stack_count,
            "total_portas": soma_portas,
            "total_aps": total_aps,
            "nivel": nivel,
            "status": "Acessado com Sucesso (Huawei)",
            "vizinhos_validos": vizinhos_validos
        }

    def _coletar_datacom(self, net_connect, ip: str, hostname: str, role_sugerido: str, nivel: int) -> Dict[str, Any]:
        """Executa comandos DmOS no switch Datacom."""
        try:
            net_connect.send_command("terminal length 0", read_timeout=10)
        except Exception:
            pass

        out_ver = net_connect.send_command("show version", read_timeout=15)
        out_sys = net_connect.send_command("show system", read_timeout=15)

        # Procura Modelo Datacom (DM4000, DM4100, DM4200, DM4370, DM4770, etc.)
        modelo = "Datacom Switch"
        m_dm = re.search(r"DM\d{4}[A-Za-z0-9_-]*", out_ver + " " + out_sys)
        if m_dm:
            modelo = m_dm.group(0)

        m_clean, portas = extrair_modelo_e_portas(modelo)
        if portas == 0:
            portas = 48

        # LLDP Datacom
        out_lldp = net_connect.send_command("show lldp neighbors detail", read_timeout=30)
        vizinhos_brutos = parse_generic_lldp_detail(out_lldp)

        total_aps = 0
        vizinhos_validos = []
        vizinhos_processados = set()

        for v in vizinhos_brutos:
            v_dev_id = v.get("device_id", "")
            v_plat = v.get("platform", "")
            v_cap = v.get("capabilities", "")

            if is_access_point(v_dev_id, v_plat, v_cap):
                total_aps += 1
                continue

            v_canonico = normalizar_hostname(v_dev_id)
            if not v_canonico or v_canonico == hostname:
                continue

            par_id = (v_canonico, v.get("local_port"))
            if par_id in vizinhos_processados:
                continue
            vizinhos_processados.add(par_id)

            v_role = classificar_dispositivo(v_canonico, v_plat, v_cap)
            v["hostname_canonico"] = v_canonico
            v["role"] = v_role
            vizinhos_validos.append(v)

        self.log(f"OK [Datacom]: '{hostname}' coletado. Modelo: {modelo} | Oferta Portas: {portas} | APs conectados: {total_aps}")

        return {
            "id": hostname,
            "hostname": hostname,
            "ip": ip,
            "role": role_sugerido,
            "modelo": modelo,
            "stack_count": 1,
            "total_portas": portas,
            "total_aps": total_aps,
            "nivel": nivel,
            "status": "Acessado com Sucesso (Datacom)",
            "vizinhos_validos": vizinhos_validos
        }

    def _coletar_hp_aruba(self, net_connect, ip: str, hostname: str, role_sugerido: str, nivel: int) -> Dict[str, Any]:
        """Executa comandos ProCurve / ArubaOS no switch HP."""
        try:
            net_connect.send_command("no page", read_timeout=10)
        except Exception:
            pass

        out_sys = net_connect.send_command("show system-information", read_timeout=15)
        out_stack = net_connect.send_command("show stacking", read_timeout=15)

        # Modelo
        modelo = "HP / Aruba Switch"
        m_mod = re.search(r"System Description\s*:\s*([^\r\n]+)", out_sys)
        if m_mod:
            modelo = m_mod.group(1).strip()

        stack_count = 1
        m_st = re.findall(r"Member ID\s*:\s*\d+", out_stack)
        if m_st:
            stack_count = len(m_st)

        m_clean, portas = extrair_modelo_e_portas(modelo)
        if portas == 0:
            portas = 48 * stack_count

        out_lldp = net_connect.send_command("show lldp info remote-device detail", read_timeout=35)
        vizinhos_brutos = parse_generic_lldp_detail(out_lldp)

        total_aps = 0
        vizinhos_validos = []
        vizinhos_processados = set()

        for v in vizinhos_brutos:
            v_dev_id = v.get("device_id", "")
            v_plat = v.get("platform", "")
            v_cap = v.get("capabilities", "")

            if is_access_point(v_dev_id, v_plat, v_cap):
                total_aps += 1
                continue

            v_canonico = normalizar_hostname(v_dev_id)
            if not v_canonico or v_canonico == hostname:
                continue

            par_id = (v_canonico, v.get("local_port"))
            if par_id in vizinhos_processados:
                continue
            vizinhos_processados.add(par_id)

            v_role = classificar_dispositivo(v_canonico, v_plat, v_cap)
            v["hostname_canonico"] = v_canonico
            v["role"] = v_role
            vizinhos_validos.append(v)

        self.log(f"OK [HP/Aruba]: '{hostname}' coletado. Stack: {stack_count} | Oferta Portas: {portas} | APs conectados: {total_aps}")

        return {
            "id": hostname,
            "hostname": hostname,
            "ip": ip,
            "role": role_sugerido,
            "modelo": modelo,
            "stack_count": stack_count,
            "total_portas": portas,
            "total_aps": total_aps,
            "nivel": nivel,
            "status": "Acessado com Sucesso (HP/Aruba)",
            "vizinhos_validos": vizinhos_validos
        }

    def _coletar_cisco(self, net_connect, ip: str, hostname: str, role_sugerido: str, nivel: int) -> Dict[str, Any]:
        """Executa comandos IOS/IOS-XE no switch Cisco."""
        try:
            net_connect.send_command("terminal length 0", read_timeout=10)
        except Exception:
            pass

        out_switch = net_connect.send_command("show switch", read_timeout=15)
        out_inv = net_connect.send_command("show inventory", read_timeout=20)

        stack_count = 1
        if "Ready" in out_switch or "Active" in out_switch:
            linhas_membros = [l for l in out_switch.splitlines() if re.match(r"^\s*\*?\d+\s+", l)]
            if linhas_membros:
                stack_count = len(linhas_membros)

        modelos_encontrados = []
        soma_portas = 0

        pids = re.findall(r"PID:\s*([A-Za-z0-9_-]+)", out_inv)
        for p in pids:
            p_clean, p_portas = extrair_modelo_e_portas(p)
            if p_portas > 0:
                modelos_encontrados.append(p_clean)
                soma_portas += p_portas

        if not modelos_encontrados:
            m_model = re.search(r"Model Number\s*:\s*([A-Za-z0-9_-]+)", out_inv)
            if m_model:
                p_clean, p_portas = extrair_modelo_e_portas(m_model.group(1))
                modelos_encontrados.append(p_clean)
                soma_portas = p_portas * stack_count
            else:
                modelos_encontrados = ["Cisco Switch"]
                soma_portas = 48 * stack_count

        modelo_final = ", ".join(sorted(set(modelos_encontrados)))
        if soma_portas == 0:
            soma_portas = 48 * stack_count

        out_cdp = net_connect.send_command("show cdp neighbors detail", read_timeout=30)
        vizinhos_cdp = parse_cisco_cdp_detail(out_cdp)

        out_lldp = net_connect.send_command("show lldp neighbors detail", read_timeout=30)
        vizinhos_lldp = parse_generic_lldp_detail(out_lldp)

        todos_vizinhos = vizinhos_cdp + vizinhos_lldp

        total_aps = 0
        vizinhos_validos = []
        vizinhos_processados = set()

        for v in todos_vizinhos:
            v_dev_id = v.get("device_id", "")
            v_plat = v.get("platform", "")
            v_cap = v.get("capabilities", "")

            if is_access_point(v_dev_id, v_plat, v_cap):
                total_aps += 1
                continue

            v_canonico = normalizar_hostname(v_dev_id)
            if not v_canonico or v_canonico == hostname:
                continue

            par_id = (v_canonico, v.get("local_port"))
            if par_id in vizinhos_processados:
                continue
            vizinhos_processados.add(par_id)

            v_role = classificar_dispositivo(v_canonico, v_plat, v_cap)
            v["hostname_canonico"] = v_canonico
            v["role"] = v_role
            vizinhos_validos.append(v)

        self.log(f"OK [Cisco]: '{hostname}' coletado. Stack: {stack_count} membro(s) | Oferta Portas: {soma_portas} | APs conectados: {total_aps}")

        return {
            "id": hostname,
            "hostname": hostname,
            "ip": ip,
            "role": role_sugerido,
            "modelo": modelo_final,
            "stack_count": stack_count,
            "total_portas": soma_portas,
            "total_aps": total_aps,
            "nivel": nivel,
            "status": "Acessado com Sucesso (Cisco)",
            "vizinhos_validos": vizinhos_validos
        }


def executar_assessment_visio(
    cores_input: List[str],
    usuario: str,
    senha: str,
    secret: str = "",
    vendor_preference: str = "auto",
    porta: int = 22,
    max_depth: int = 10,
    timeout: int = 10
) -> Dict[str, Any]:
    """Ponto de entrada chamado pela rota Flask."""
    crawler = BigPictureCrawler(
        core_ips=cores_input,
        usuario=usuario,
        senha=senha,
        secret=secret,
        vendor_preference=vendor_preference,
        porta=porta,
        max_depth=max_depth,
        timeout=timeout
    )

    try:
        resultado = crawler.executar()
        return resultado
    finally:
        del usuario
        del senha
        del secret
        del crawler.usuario
        del crawler.senha
        del crawler.secret
        gc.collect()
