"""
Módulo de cálculo de sub-redes IPv4 e IPv6.
Usa a biblioteca nativa ipaddress do Python (altamente confiável, precisa e sem dependências externas).
Inclui análise hierárquica reversa de super-redes (Supernet / Parent Networks) e calculadora de sub-divisões.
"""

import ipaddress


def calcular_subrede(entrada_ip: str, supernet_prefixo: int = None):
    """
    Calcula informações detalhadas da sub-rede para um IP com máscara (CIDR ou decimal).
    Inclui lista completa de redes pai / super-redes (ex: saber dentro de qual /24, /19, /16 ela está).
    """
    entrada_limpa = entrada_ip.strip()
    if not entrada_limpa:
        raise ValueError("Please provide an IP address and mask (e.g. 10.255.60.0/29 or 192.168.1.1/24).")

    # Suporta formato com espaço (ex: "10.255.60.0 255.255.255.248")
    if " " in entrada_limpa:
        partes = entrada_limpa.split(None, 1)
        entrada_limpa = f"{partes[0]}/{partes[1]}"

    try:
        interface = ipaddress.ip_interface(entrada_limpa)
    except ValueError as e:
        raise ValueError(f"Invalid IP address or netmask format: {e}")

    rede = interface.network
    versao = rede.version
    prefixo_atual = rede.prefixlen

    if versao == 4:
        mascara = str(rede.netmask)
        wildcard = str(rede.hostmask)
        prefixo = f"/{prefixo_atual}"
        rede_ip = str(rede.network_address)
        broadcast = str(rede.broadcast_address)
        total_ips = rede.num_addresses

        if prefixo_atual == 32:
            primeiro_ip = str(rede.network_address)
            ultimo_ip = str(rede.network_address)
            hosts_utilizaveis = 1
        elif prefixo_atual == 31:
            primeiro_ip = str(rede.network_address)
            ultimo_ip = str(rede.broadcast_address)
            hosts_utilizaveis = 2
        else:
            primeiro_ip = str(rede.network_address + 1)
            ultimo_ip = str(rede.broadcast_address - 1)
            hosts_utilizaveis = total_ips - 2

        # Tipo de IP (Privado, Público, Loopback, Multicast)
        ip_obj = interface.ip
        tipo_ip = "Public (Internet Routable)"
        if ip_obj.is_private:
            tipo_ip = "Private (RFC 1918)"
        elif ip_obj.is_loopback:
            tipo_ip = "Loopback"
        elif ip_obj.is_multicast:
            tipo_ip = "Multicast (Class D)"
        elif ip_obj.is_link_local:
            tipo_ip = "Link-Local (APIPA)"

        # Representação binária da máscara
        mascara_bin = ".".join(f"{int(octeto):08b}" for octeto in mascara.split("."))

        # Hierarquia de Redes Superiores (Supernets / Parent Networks)
        # Permite ver exatamente em qual /24, /23, /22, /19, /16, /8 essa subrede está contida
        supernets = []
        for p in range(prefixo_atual - 1, 7, -1):  # De prefixo atual - 1 até /8
            sup_rede = ipaddress.ip_network(f"{rede.network_address}/{p}", strict=False)
            supernets.append({
                "prefixo": f"/{p}",
                "prefixo_num": p,
                "network": str(sup_rede.network_address),
                "cidr_completo": f"{sup_rede.network_address}/{p}",
                "netmask": str(sup_rede.netmask),
                "primeiro_ip": str(sup_rede.network_address + 1) if p <= 30 else str(sup_rede.network_address),
                "ultimo_ip": str(sup_rede.broadcast_address - 1) if p <= 30 else str(sup_rede.broadcast_address),
                "total_ips": f"{sup_rede.num_addresses:,}",
                "total_subredes_deste_tamanho": f"{(2 ** (prefixo_atual - p)):,}"
            })

        return {
            "valido": True,
            "versao": 4,
            "ip_digitado": str(interface.ip),
            "cidr": prefixo,
            "prefixo_num": prefixo_atual,
            "network": rede_ip,
            "netmask": mascara,
            "wildcard": wildcard,
            "broadcast": broadcast,
            "primeiro_ip": primeiro_ip,
            "ultimo_ip": ultimo_ip,
            "total_ips": f"{total_ips:,}",
            "hosts_utilizaveis": f"{hosts_utilizaveis:,}",
            "tipo_ip": tipo_ip,
            "mascara_bin": mascara_bin,
            "supernets": supernets,
        }
    else:
        # IPv6
        supernets = []
        for p in [64, 56, 48, 32]:
            if prefixo_atual > p:
                sup_rede = ipaddress.ip_network(f"{rede.network_address}/{p}", strict=False)
                supernets.append({
                    "prefixo": f"/{p}",
                    "prefixo_num": p,
                    "network": str(sup_rede.network_address),
                    "cidr_completo": f"{sup_rede.network_address}/{p}",
                    "netmask": str(sup_rede.netmask),
                    "primeiro_ip": str(sup_rede.network_address),
                    "ultimo_ip": str(sup_rede.network_address + (sup_rede.num_addresses - 1)),
                    "total_ips": f"{sup_rede.num_addresses:,}",
                    "total_subredes_deste_tamanho": "N/A"
                })

        return {
            "valido": True,
            "versao": 6,
            "ip_digitado": str(interface.ip),
            "cidr": f"/{prefixo_atual}",
            "prefixo_num": prefixo_atual,
            "network": str(rede.network_address),
            "netmask": str(rede.netmask),
            "wildcard": "N/A (IPv6)",
            "broadcast": "N/A (IPv6 Multicast)",
            "primeiro_ip": str(rede.network_address),
            "ultimo_ip": str(rede.network_address + (rede.num_addresses - 1)),
            "total_ips": f"{rede.num_addresses:,}",
            "hosts_utilizaveis": f"{rede.num_addresses:,}",
            "tipo_ip": "Private (ULA)" if interface.ip.is_private else "Global / Other",
            "mascara_bin": "N/A",
            "supernets": supernets,
        }
