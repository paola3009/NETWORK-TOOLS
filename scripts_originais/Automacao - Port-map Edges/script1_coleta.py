

import pandas as pd

import json

import getpass

from netmiko import ConnectHandler

import re



# --------------------------------------------

# Remove FQDN e mantém apenas o rótulo

# --------------------------------------------

def strip_domain(h):

    return h.split('.')[0] if isinstance(h, str) else h



# --------------------------------------------

# Normaliza interface para versão curta

# --------------------------------------------

def short_ifname(name: str) -> str:

    if not name:

        return name

    name = name.strip().rstrip(',')



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



# --------------------------------------------

# Parser do CDP DETAIL

# --------------------------------------------

def parse_cdp_detail(output: str):

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



# --------------------------------------------

# Parser do Loopback

# --------------------------------------------

def parse_loopback(output: str):

    m = re.search(r"(\d+\.\d+\.\d+\.\d+)", output)

    return m.group(1) if m else None



# --------------------------------------------

# Script principal

# --------------------------------------------

def main():

    planilha = "Port-Map-Base.xlsx"

    df = pd.read_excel(planilha)



    hosts = set(df["Hostname A"].dropna().unique()) | set(df["Hostname B"].dropna().unique())

    hosts = {strip_domain(h) for h in hosts}



    print(f"\nColetando CDP DETAIL de {len(hosts)} switches.\n")



    usuario = input("Usuário SSH: ")

    senha = getpass.getpass("Senha SSH: ")



    resultados = {}



    for raw_host in sorted(hosts):

        host = strip_domain(raw_host)

        print(f"\nConectando em {host}...")



        device = {

            "device_type": "cisco_ios",

            "host": host,

            "username": usuario,

            "password": senha,

        }



        try:

            conn = ConnectHandler(**device)



            conn.send_command_timing("terminal length 0")



            cdp_out = conn.send_command("show cdp neighbors detail")

            loop_out = conn.send_command("show ip interface brief | include Loopback0")



            conn.disconnect()



            resultados[host] = {

                "loopback0_ip": parse_loopback(loop_out),

                "cdp": parse_cdp_detail(cdp_out)

            }



            print(f"✔ OK: {host} → {len(resultados[host]['cdp'])} vizinhos")



        except Exception as e:

            print(f"✖ Falha em {host}: {e}")

            resultados[host] = {

                "loopback0_ip": None,

                "cdp": []

            }



    with open("dados_portmap.json", "w", encoding="utf-8") as f:

        json.dump(resultados, f, indent=2, ensure_ascii=False)



    print("\n✔ Arquivo 'dados_portmap.json' gerado com sucesso!\n")



if __name__ == "__main__":

    main()

