import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

def create_sample_de_para(filename: str = "DE_PARA_Exemplo.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DE PARA SWITCH"

    headers = [
        "Hostname",                 # A (1)
        "Porta",                    # B (2)
        "Status",                   # C (3)
        "Descrição",                # D (4)
        "MAC aprendido",            # E (5)
        "IP do MAC",                # F (6)
        "Vendor",                   # G (7)
        "VLAN de dados",            # H (8)
        "VLAN de voz",              # I (9)
        "Uptime da interface",      # J (10)
        "Ponto de rede conectado",   # K (11)
        "Modelo Legado",            # L (12)
        "Localização",              # M (13)
        "Patch Panel",              # N (14)
        "Porta Patch Panel",        # O (15)
        "Observação Legado",        # P (16)
        "Migrar"                    # Q (17)
    ]

    ws.append(headers)

    # Dados de exemplo cobrindo todas as regras do usuário:
    # - APs (devem ir pro C9300-24UX)
    # - Câmeras (devem ser distribuídas entre os switches)
    # - Impressoras (Seiko Epson, Argox) -> IP pool impressora
    # - Workstations (Dell, AsRock) -> IP pool notebook/desktop
    # - VLANs com Pools diretos (108, 1940, 162)
    # - VLANs com Replace (749->1030, 750->1031, 751->1032)
    sample_rows = [
        # Linha 1: AP em VLAN 108
        ["SW-LEGADO-01", "Gi1/0/1", "connected", "AP-ANDAR-01", "00:1A:2B:3C:4D:5E", "10.108.0.15", "Cisco AP", "108", "", "45 days", "PTO-01", "WS-C2960X", "Rack 1", "PP-A", "01", "", "Sim"],
        # Linha 2: AP em VLAN 108
        ["SW-LEGADO-01", "Gi1/0/2", "connected", "AP-ANDAR-02", "00:1A:2B:3C:4D:5F", "10.108.0.16", "Aruba", "108", "", "45 days", "PTO-02", "WS-C2960X", "Rack 1", "PP-A", "02", "", "Sim"],
        
        # 7 Câmeras (para testar divisão entre switches com 1 switch redundante 24UX)
        ["SW-LEGADO-01", "Gi1/0/3", "connected", "CAM-RECEPCAO-01", "00:11:22:33:44:55", "192.168.1.10", "Hikvision", "1940", "", "10 days", "CAM-PTO-01", "WS-C2960X", "Rack 1", "PP-A", "03", "", "Sim"],
        ["SW-LEGADO-01", "Gi1/0/4", "connected", "CAM-CORREDOR-02", "00:11:22:33:44:56", "192.168.1.11", "Axis", "1940", "", "10 days", "CAM-PTO-02", "WS-C2960X", "Rack 1", "PP-A", "04", "", "Sim"],
        ["SW-LEGADO-01", "Gi1/0/5", "connected", "CAM-ESTACIONAMENTO-03", "00:11:22:33:44:57", "192.168.1.12", "Intelbras", "1940", "", "10 days", "CAM-PTO-03", "WS-C2960X", "Rack 1", "PP-A", "05", "", "Sim"],
        ["SW-LEGADO-01", "Gi1/0/6", "connected", "CAM-ENTRADA-04", "00:11:22:33:44:58", "192.168.1.13", "Vivotek", "1940", "", "10 days", "CAM-PTO-04", "WS-C2960X", "Rack 1", "PP-A", "06", "", "Sim"],
        ["SW-LEGADO-01", "Gi1/0/7", "connected", "CAM-GALPAO-05", "00:11:22:33:44:59", "192.168.1.14", "Hikvision", "1940", "", "10 days", "CAM-PTO-05", "WS-C2960X", "Rack 1", "PP-A", "07", "", "Sim"],
        ["SW-LEGADO-01", "Gi1/0/8", "connected", "CAM-ALMOXARIFADO-06", "00:11:22:33:44:5A", "192.168.1.15", "Axis", "1940", "", "10 days", "CAM-PTO-06", "WS-C2960X", "Rack 1", "PP-A", "08", "", "Sim"],
        ["SW-LEGADO-01", "Gi1/0/9", "connected", "CAM-TI-07", "00:11:22:33:44:5B", "192.168.1.16", "Intelbras", "1940", "", "10 days", "CAM-PTO-07", "WS-C2960X", "Rack 1", "PP-A", "09", "", "Sim"],

        # Impressoras (Vendors Seiko Epson, Argox)
        ["SW-LEGADO-01", "Gi1/0/10", "connected", "IMP-ETIQUETAS-01", "00:22:33:44:55:66", "10.10.20.5", "Seiko Epson", "162", "", "12 days", "PTO-IMP-01", "WS-C2960X", "Rack 1", "PP-A", "10", "", "Sim"],
        ["SW-LEGADO-01", "Gi1/0/11", "connected", "IMP-LOGISTICA-02", "00:22:33:44:55:67", "10.10.20.6", "Argox", "162", "", "12 days", "PTO-IMP-02", "WS-C2960X", "Rack 1", "PP-A", "11", "", "Sim"],

        # Workstations / Desktops (Vendors Dell, AsRock)
        ["SW-LEGADO-01", "Gi1/0/12", "connected", "DESK-ADM-01", "00:33:44:55:66:77", "10.10.10.100", "Dell", "162", "", "2 days", "PTO-WORK-01", "WS-C2960X", "Rack 1", "PP-A", "12", "", "Sim"],
        ["SW-LEGADO-01", "Gi1/0/13", "connected", "NOTE-DEV-02", "00:33:44:55:66:78", "10.10.10.101", "AsRock", "162", "", "5 days", "PTO-WORK-02", "WS-C2960X", "Rack 1", "PP-A", "13", "", "Sim"],

        # VLANs com Replace (749->1030, 750->1031, 751->1032)
        ["SW-LEGADO-01", "Gi1/0/14", "connected", "SERVICO-VLAN-749", "00:44:55:66:77:88", "10.749.0.10", "Generic", "749", "", "30 days", "PTO-SRV-01", "WS-C2960X", "Rack 1", "PP-A", "14", "", "Sim"],
        ["SW-LEGADO-01", "Gi1/0/15", "connected", "SERVICO-VLAN-750", "00:44:55:66:77:89", "10.750.0.10", "Generic", "750", "", "30 days", "PTO-SRV-02", "WS-C2960X", "Rack 1", "PP-A", "15", "", "Sim"],
        ["SW-LEGADO-01", "Gi1/0/16", "connected", "SERVICO-VLAN-751", "00:44:55:66:77:8A", "10.751.0.10", "Generic", "751", "", "30 days", "PTO-SRV-03", "WS-C2960X", "Rack 1", "PP-A", "16", "", "Sim"],

        # Linha que NÃO deve ser migrada (Migrar = Não)
        ["SW-LEGADO-01", "Gi1/0/17", "notconnect", "PORTA-DESATIVADA", "", "", "", "1", "", "0 days", "PTO-OFF", "WS-C2960X", "Rack 1", "PP-A", "17", "Porta desativada", "Não"],
    ]

    for row in sample_rows:
        ws.append(row)

    # Estilização do cabeçalho
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(filename)
    print(f"Planilha de exemplo criada com sucesso: {os.path.abspath(filename)}")

if __name__ == "__main__":
    create_sample_de_para()
