"""
Módulo de armazenamento de documentos e tópicos de base de conhecimento.
Gerencia arquivos e metadados com segurança (validação de extensões, sanitização e isolamento de caminhos).
"""

import json
import os
import re
import uuid
from datetime import datetime
import bleach
import markdown

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DATA_DIR = os.path.join(BASE_DIR, "data")
DOCS_DIR = os.path.join(DATA_DIR, "standard_documents")
KB_DIR = os.path.join(DATA_DIR, "knowledge_base")

DOCS_INDEX_FILE = os.path.join(DOCS_DIR, "index.json")
KB_INDEX_FILE = os.path.join(KB_DIR, "index.json")

# Extensões permitidas para anexos (Excel, Word, PowerPoint, Visio, PDF, Texto, Imagens)
ALLOWED_EXTENSIONS = {
    ".xlsx", ".xls", ".csv",
    ".docx", ".doc",
    ".pptx", ".ppt",
    ".vsdx", ".vsd",
    ".pdf", ".txt",
    ".png", ".jpg", ".jpeg"
}

# Tags e atributos permitidos para sanitização de HTML renderizado a partir de Markdown (estilo GitHub)
ALLOWED_TAGS = [
    "h1", "h2", "h3", "h4", "h5", "h6", "p", "strong", "b", "em", "i", "u", "s", "mark", "font",
    "ul", "ol", "li", "blockquote", "code", "pre", "table", "thead", "tbody", "tr", "th", "td",
    "hr", "br", "span", "a", "img"
]
ALLOWED_ATTRIBUTES = {
    "a": ["href", "title", "target", "rel"],
    "code": ["class"],
    "span": ["class"],
    "font": ["color", "size", "face"],
    "mark": ["class"],
    "u": ["class"],
    "p": ["class"],
    "img": ["src", "alt", "title", "class", "width", "height"]
}


def _carregar_json(caminho):
    if not os.path.isfile(caminho):
        return []
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _salvar_json(caminho, dados):
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    temp_caminho = f"{caminho}.tmp"
    with open(temp_caminho, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)
    os.replace(temp_caminho, caminho)


def sanitizar_nome_arquivo(nome):
    """Sanitiza o nome do arquivo para prevenir directory traversal ou caracteres maliciosos."""
    nome = os.path.basename(nome).strip()
    nome = re.sub(r'[^a-zA-Z0-9_\-\.\(\)\s]', '_', nome)
    return nome if nome else "document"


def extensao_permitida(nome):
    ext = os.path.splitext(nome)[1].lower()
    return ext in ALLOWED_EXTENSIONS


# ---------------- STANDARD DOCUMENTS ----------------

def listar_documentos(categoria=None):
    docs = _carregar_json(DOCS_INDEX_FILE)
    if categoria:
        return [d for d in docs if d.get("categoria") == categoria]
    return docs


def salvar_documento(file_obj, titulo, categoria, descricao=""):
    nome_original = sanitizar_nome_arquivo(file_obj.filename)
    if not extensao_permitida(nome_original):
        raise ValueError("File extension not permitted.")

    ext = os.path.splitext(nome_original)[1].lower()
    doc_id = str(uuid.uuid4())
    nome_armazenado = f"{doc_id}{ext}"
    caminho_final = os.path.join(DOCS_DIR, nome_armazenado)

    file_obj.save(caminho_final)
    tamanho = os.path.getsize(caminho_final)

    doc_info = {
        "id": doc_id,
        "titulo": titulo.strip() or nome_original,
        "categoria": categoria,
        "descricao": descricao.strip(),
        "nome_original": nome_original,
        "arquivo": nome_armazenado,
        "tamanho": tamanho,
        "data": datetime.now().strftime("%Y-%m-%d %H:%M")
    }

    docs = _carregar_json(DOCS_INDEX_FILE)
    docs.insert(0, doc_info)
    _salvar_json(DOCS_INDEX_FILE, docs)
    return doc_info


def obter_caminho_documento(doc_id):
    docs = _carregar_json(DOCS_INDEX_FILE)
    for d in docs:
        if d["id"] == doc_id:
            caminho = os.path.join(DOCS_DIR, d["arquivo"])
            if os.path.isfile(caminho):
                return caminho, d["nome_original"]
    return None, None


# ---------------- KNOWLEDGE BASE ----------------

def listar_topicos(categoria=None):
    topicos = _carregar_json(KB_INDEX_FILE)
    if categoria:
        return [t for t in topicos if t.get("categoria") == categoria]
    return topicos


def obter_topico(topico_id):
    topicos = _carregar_json(KB_INDEX_FILE)
    for t in topicos:
        if t["id"] == topico_id:
            # Renderizar Markdown com sanitização
            raw_md = t.get("conteudo_md", "")
            html_bruto = markdown.markdown(
                raw_md,
                extensions=["fenced_code", "tables", "codehilite", "nl2br"]
            )
            html_seguro = bleach.clean(
                html_bruto,
                tags=ALLOWED_TAGS,
                attributes=ALLOWED_ATTRIBUTES,
                protocols=["http", "https", "data"],
                strip=True
            )
            topico_copia = dict(t)
            topico_copia["conteudo_html"] = html_seguro
            return topico_copia
    return None


def criar_topico(titulo, categoria, conteudo_md, autor="Admin", anexos_arquivos=None):
    topico_id = str(uuid.uuid4())
    anexos_salvos = []

    if anexos_arquivos:
        for f in anexos_arquivos:
            if not f or not f.filename:
                continue
            nome_original = sanitizar_nome_arquivo(f.filename)
            if not extensao_permitida(nome_original):
                continue
            ext = os.path.splitext(nome_original)[1].lower()
            anexo_id = str(uuid.uuid4())
            nome_armazenado = f"kb_{anexo_id}{ext}"
            caminho_anexo = os.path.join(KB_DIR, nome_armazenado)
            f.save(caminho_anexo)
            tamanho = os.path.getsize(caminho_anexo)
            anexos_salvos.append({
                "id": anexo_id,
                "nome_original": nome_original,
                "arquivo": nome_armazenado,
                "tamanho": tamanho
            })

    novo_topico = {
        "id": topico_id,
        "titulo": titulo.strip(),
        "categoria": categoria,
        "autor": autor.strip() or "User",
        "conteudo_md": conteudo_md,
        "anexos": anexos_salvos,
        "data": datetime.now().strftime("%Y-%m-%d %H:%M")
    }

    topicos = _carregar_json(KB_INDEX_FILE)
    topicos.insert(0, novo_topic := novo_topico)
    _salvar_json(KB_INDEX_FILE, topicos)
    return novo_topico


def obter_caminho_anexo_kb(anexo_id):
    topicos = _carregar_json(KB_INDEX_FILE)
    for t in topicos:
        for a in t.get("anexos", []):
            if a["id"] == anexo_id:
                caminho = os.path.join(KB_DIR, a["arquivo"])
                if os.path.isfile(caminho):
                    return caminho, a["nome_original"]
    return None, None


def excluir_documento(doc_id):
    """Exclui o arquivo físico e remove do índice."""
    docs = _carregar_json(DOCS_INDEX_FILE)
    novos_docs = []
    excluido = False
    for d in docs:
        if d["id"] == doc_id:
            caminho = os.path.join(DOCS_DIR, d["arquivo"])
            if os.path.isfile(caminho):
                try:
                    os.remove(caminho)
                except OSError:
                    pass
            excluido = True
        else:
            novos_docs.append(d)
    if excluido:
        _salvar_json(DOCS_INDEX_FILE, novos_docs)
    return excluido


def excluir_topico(topico_id):
    """Exclui o tópico e todos os seus anexos físicos do disco."""
    topicos = _carregar_json(KB_INDEX_FILE)
    novos_topicos = []
    excluido = False
    for t in topicos:
        if t["id"] == topico_id:
            for a in t.get("anexos", []):
                caminho = os.path.join(KB_DIR, a["arquivo"])
                if os.path.isfile(caminho):
                    try:
                        os.remove(caminho)
                    except OSError:
                        pass
            excluido = True
        else:
            novos_topicos.append(t)
    if excluido:
        _salvar_json(KB_INDEX_FILE, novos_topicos)
    return excluido


def atualizar_topico(topico_id, titulo, conteudo_md, autor, novos_anexos=None, anexos_remover_ids=None):
    """Atualiza o título, autor, conteúdo e anexos de um tópico existente."""
    topicos = _carregar_json(KB_INDEX_FILE)
    topico_atualizado = None
    remover_ids = set(anexos_remover_ids or [])

    for t in topicos:
        if t["id"] == topico_id:
            t["titulo"] = titulo
            t["conteudo_md"] = conteudo_md
            t["autor"] = autor
            t["data_atualizado"] = datetime.now().strftime("%Y-%m-%d %H:%M")

            # Remove anexos marcados
            if remover_ids:
                mantidos = []
                for a in t.get("anexos", []):
                    if a["id"] in remover_ids:
                        caminho = os.path.join(KB_DIR, a["arquivo"])
                        if os.path.isfile(caminho):
                            try:
                                os.remove(caminho)
                            except OSError:
                                pass
                    else:
                        mantidos.append(a)
                t["anexos"] = mantidos

            # Adiciona novos anexos
            if novos_anexos:
                for f in novos_anexos:
                    if not f or not f.filename:
                        continue
                    nome_original = sanitizar_nome_arquivo(f.filename)
                    if not extensao_permitida(nome_original):
                        continue
                    ext = os.path.splitext(nome_original)[1].lower()
                    anexo_id = str(uuid.uuid4())
                    nome_armazenado = f"kb_{anexo_id}{ext}"
                    caminho_anexo = os.path.join(KB_DIR, nome_armazenado)
                    f.save(caminho_anexo)
                    tamanho = os.path.getsize(caminho_anexo)
                    t.setdefault("anexos", []).append({
                        "id": anexo_id,
                        "nome_original": nome_original,
                        "arquivo": nome_armazenado,
                        "tamanho": tamanho
                    })

            topico_atualizado = t
            break

    if topico_atualizado:
        _salvar_json(KB_INDEX_FILE, topicos)
    return topico_atualizado

