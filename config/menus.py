"""
Estrutura de menus da interface.

Hierarquia:  GRUPO (ex.: TOOLS, STANDARD DOCUMENTS, KNOWLEDGE BASE)  ->  MENU  ->  SUBMENUS (ferramentas/itens)

Para adicionar submenus, inclua itens na lista "submenus" do menu desejado:
    {"id": "identificador_unico", "titulo": "Nome exibido"}
"""

GRUPOS = {
    "tools": {
        "titulo": "TOOLS",
        "icone": "tools",
        "menus": {
            "assessment": {
                "titulo": "Phase 1 - Discovery / Assessment",
                "submenus": [
                    {"id": "assessment_visio", "titulo": "Assessment Big Picture Visio", "descricao": "Mapeamento multicamadas via CDP/LLDP com exportação para Microsoft Visio (.vsdx) e planilha de capacidade de switches e APs."},
                    {"id": "assessment_core", "titulo": "Assessment CORE (L2 e L3)"},
                ],
            },
            "rollout": {
                "titulo": "Phase 2 - Pre/Post Rollout",
                "submenus": [
                    {"id": "elaboracao_de_para", "titulo": "Elaboração DE PARA"},
                    {"id": "finalizacao_para", "titulo": "Finalização \"PARA\""},
                    {"id": "report_migracoes", "titulo": "Report de Migrações"},
                ],
            },
            "asbuilt": {
                "titulo": "Phase 3 - As-Built Documentation",
                "submenus": [
                    {"id": "inventory_aps", "titulo": "Automacao - Coleta Inventory Access Points"},
                    {"id": "inventory_edges", "titulo": "Automacao - Coleta Inventory Edges"},
                    {"id": "portmap_edges", "titulo": "Automacao - Port-map Edges"},
                ],
            },
            "subnet_calc": {
                "titulo": "Subnet Calculator",
                "icone": "calculator",
                "sem_numero": True,
                "submenus": [],
            },
        },
    },
    "standard_documents": {
        "titulo": "STANDARD DOCUMENTS",
        "icone": "file-text",
        "menus": {
            "templates_docs": {
                "titulo": "Document Templates",
                "submenus": [],
            },
            "procedures": {
                "titulo": "Standard Operating Procedures",
                "submenus": [],
            },
        },
    },
    "knowledge_base": {
        "titulo": "KNOWLEDGE BASE",
        "icone": "book-open",
        "menus": {
            "kb_articles": {
                "titulo": "Articles & Guides",
                "submenus": [],
            },
            "troubleshooting": {
                "titulo": "Troubleshooting & Playbooks",
                "submenus": [],
            },
        },
    },
}

# Índice plano de todos os menus (usado pelas rotas como whitelist). Não editar.
MENUS = {mid: m for g in GRUPOS.values() for mid, m in g["menus"].items()}
