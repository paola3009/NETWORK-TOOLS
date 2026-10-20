import json
import os
from typing import Dict, Any, List

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "site_configs.json")

# Pools Padrão Globais presentes em todos os sites
STANDARD_POOLS = [
    "notebooks_desktops",
    "impressoras_plotters_scanners",
    "telefoneIP_gateway_ATA",
    "1250_ATA",
    "ap_INFRA_VN"
]

def load_configs() -> Dict[str, Any]:
    """Lê estritamente o arquivo JSON de configurações de sites sem alterar nada do usuário."""
    if not os.path.exists(CONFIG_FILE):
        default_data = {
            "default": {
                "site_id": "default",
                "site_name": "Padrão (Geral)",
                "ap_pool": "ap_INFRA_VN",
                "default_vlan_dados": "",
                "vlan_direct_pools": {
                    "108": "POOL_VLAN_108",
                    "162": "POOL_VLAN_162",
                    "1940": "POOL_VLAN_1940"
                },
                "vlan_replace_pools": {
                    "749": {"replace_vlan": "1030", "pool_name": "POOL_VLAN_1030"},
                    "750": {"replace_vlan": "1031", "pool_name": "POOL_VLAN_1031"},
                    "751": {"replace_vlan": "1032", "pool_name": "POOL_VLAN_1032"}
                },
                "vendor_pools": [
                    {
                        "keywords": ["seiko", "epson", "argox", "zebra"],
                        "pool_name": "impressoras_plotters_scanners",
                        "service_type": "Impressora"
                    },
                    {
                        "keywords": ["dell", "asrock", "lenovo", "hp", "intel", "apple"],
                        "pool_name": "notebooks_desktops",
                        "service_type": "Desktop/Notebook"
                    }
                ],
                "camera_keywords": ["camera", "cam", "hikvision", "axis", "intelbras", "vivotek", "dahua", "cctv"],
                "ap_keywords": ["ap", "access point", "aruba", "ruckus", "cisco ap", "meraki", "aironet"]
            }
        }
        save_configs(default_data)
        return default_data

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            configs = json.load(f)
            return configs
    except Exception as e:
        print(f"Erro ao carregar configuracoes: {e}")
        return {}

def save_configs(configs: Dict[str, Any]) -> None:
    """Salva o dicionário de configurações no arquivo JSON exatamente como está."""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(configs, f, indent=4, ensure_ascii=False)

def get_site_config(site_id: str) -> Dict[str, Any]:
    """Retorna a configuração exata de um site cadastrado pelo usuário."""
    configs = load_configs()
    if site_id in configs:
        return configs[site_id]
    if "default" in configs:
        return configs["default"]
    if configs:
        first_key = list(configs.keys())[0]
        return configs[first_key]
    return {}

def update_site_config(site_id: str, config_data: Dict[str, Any]) -> None:
    """Atualiza ou cria a configuração de um site mantendo a integridade."""
    configs = load_configs()
    config_data["site_id"] = site_id
    configs[site_id] = config_data
    save_configs(configs)

def delete_site_config(site_id: str) -> bool:
    """Exclui um site da configuração (exceto 'default')."""
    if site_id == "default":
        return False
    configs = load_configs()
    if site_id in configs:
        del configs[site_id]
        save_configs(configs)
        return True
    return False
