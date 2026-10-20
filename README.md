# Network Tools — Plataforma de Automação e Engenharia de Redes

Plataforma web modular desenvolvida em Python e Flask para automação, descoberta (assessment), documentação e operações de engenharia de redes multi-fabricante (**Cisco, Huawei, HP/Aruba, Datacom, Juniper, Dell**).

---

## 🚀 Principais Módulos e Funcionalidades

### 1. Fase 1: Assessment Big Picture Visio (`/tool/assessment_visio`)
* **Descoberta Recursiva Multi-Hop**: Conexão aos switches Core e varredura em cascata de todos os níveis de vizinhos via **CDP** e **LLDP**.
* **Suporte Multi-Vendor Nativo**: Detecção automática e comandos sob medida para Cisco (IOS/IOS-XE/NX-OS), Huawei (VRP Campus S-Series e CloudEngine), HP/Aruba (Comware, ProCurve, AOS-CX) e Datacom (DmOS).
* **Filtro Inteligente de Access Points**: Identifica APs via CDP/LLDP para contagem de capacidade sem tentar conexão SSH nem poluir o diagrama principal.
* **Exportação para Microsoft Visio (`.vsdx`)**: Criação automática do diagrama hierárquico com shapes específicos por papel (Core, Switch de Acesso, Router, WLC), identificação de interfaces e enlaces entre dispositivos.
* **Exportação para Excel (`.xlsx`)**: Planilha com modelo, stack, oferta total de portas, contagem de APs e aba dedicada com todos os enlaces e conexões mapeadas.

### 2. Rollout e Migração De-Para
* Análise automatizada de configurações legadas e mapeamento inteligente de interfaces (De-Para) para novos switches e pilhas.

### 3. Calculadora de Sub-redes IPv4
* Cálculo de blocos CIDR, máscaras, faixas úteis, broadcast e planejamento de endereçamento IP.

### 4. Base de Conhecimento e Documentação Padrão
* Repositório centralizado de manuais operacionais (KBs) com exportação nos formatos **PDF** e **Word (`.docx`)**.

---

## 🛠️ Tecnologias Utilizadas

* **Backend**: Python 3.10+ / Flask
* **Engenharia de Redes / SSH**: Netmiko
* **Geração de Documentos**:
  * `vsdx`: Manipulação e geração nativa de diagramas Microsoft Visio
  * `openpyxl`: Planilhas executivas Microsoft Excel
  * `python-docx`: Documentos Microsoft Word
  * `reportlab`: Relatórios e PDFs técnicos
* **Frontend**: HTML5, CSS3 corporativo responsivo e SVG dinâmico

---

## 📦 Instalação e Execução

### 1. Clonar o Repositório
```bash
git clone https://github.com/<seu-usuario>/<seu-repositorio>.git
cd "NETWORK TOOLS"
```

### 2. Criar e Ativar Ambiente Virtual (Recomendado)
```bash
python -m venv venv

# Windows (PowerShell):
.\venv\Scripts\Activate.ps1

# Linux / Mac:
source venv/bin/activate
```

### 3. Instalar Dependências
```bash
pip install -r requirements.txt
```

### 4. Iniciar a Aplicação
```bash
python main.py
```
Acesse no seu navegador: **`http://127.0.0.1:5000`**

---

## 🔒 Boas Práticas de Segurança
* **Credenciais**: Credenciais informadas na interface web (usuário/senha/enable) são utilizadas em tempo de execução e descartadas da memória imediatamente após a coleta.
* **Arquivos Gerados**: O arquivo `.gitignore` previne a submissão de topologias com dados sensíveis de clientes (arquivos em `core/static/downloads/`).
