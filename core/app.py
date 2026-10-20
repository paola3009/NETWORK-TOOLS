"""
NETWORK TOOLS - Interface web local (somente coleta de informaÃ§Ãµes).
Executar: python app.py  ->  http://127.0.0.1:5000
"""

import gc
import hashlib
import hmac
import os
import re
import sys
import uuid

# Garante que a raiz do projeto esteja no sys.path mesmo se executado de dentro da pasta core/
_RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)

from flask import Flask, abort, jsonify, redirect, render_template, request, send_file, url_for

from config import branding
from config.menus import GRUPOS, MENUS
from modules import storage
from modules.tools.subnet_calc import calc as subnet_calc
from modules.tools.fase3.portmap import engine as portmap_engine
from modules.tools.fase3.inventory_aps import engine as inventory_aps_engine
from modules.tools.fase2.finalizacao_para import processor as de_para_processor
from modules.tools.fase2.finalizacao_para import config_store as de_para_config_store
from modules.tools.fase2.finalizacao_para import generate_sample_excel as de_para_sample
from modules.tools.fase1.big_picture_visio import executar_assessment_visio
from modules.tools.kb_export import exportar_topico_docx, exportar_topico_pdf

HOST = "127.0.0.1"   # Acesso somente local
PORT = 5000
HOSTS_PERMITIDOS = {f"127.0.0.1:{PORT}", f"localhost:{PORT}"}
LOGO_EXTENSOES = {".png", ".svg", ".webp", ".jpg", ".jpeg"}

# Hash PBKDF2-HMAC-SHA256 seguro da senha de exclusão (salt : derivacao)
# Protege contra hardcode em texto plano e timing attacks
_SALT_HEX = "2e4228e16d614fee1c02d8a02748819c"
_HASH_HEX = "95fe617130308c27e3aec61a9ca1ca290e27c6599224d56486bea5ff94c058a1"


def verificar_senha_exclusao(senha_informada: str) -> bool:
    """Verifica se a senha de exclusão é válida com comparação em tempo constante."""
    if not senha_informada:
        return False
    derivada = hashlib.pbkdf2_hmac(
        "sha256",
        senha_informada.encode("utf-8"),
        bytes.fromhex(_SALT_HEX),
        100000
    ).hex()
    return hmac.compare_digest(derivada, _HASH_HEX)

app = Flask(__name__)
app.config.update(
    MAX_CONTENT_LENGTH=50 * 1024 * 1024,  # Permite uploads de documentos e anexos de até 50 MB
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Strict",
)


@app.before_request
def validar_host():
    """Bloqueia requisiÃ§Ãµes com Host diferente do local (proteÃ§Ã£o contra DNS rebinding)."""
    if request.host not in HOSTS_PERMITIDOS:
        abort(400)


@app.after_request
def cabecalhos_seguranca(resp):
    resp.headers["Content-Security-Policy"] = (
        "default-src 'none'; style-src 'self' 'unsafe-inline'; script-src 'self'; connect-src 'self'; img-src 'self' data:; "
        "form-action 'self'; frame-ancestors 'none'; base-uri 'none'"
    )
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["Pragma"] = "no-cache"
    resp.headers.pop("Server", None)
    return resp


def _carregar_marca():
    """Monta nome/logo a partir de branding.py. Logo invÃ¡lida ou ausente -> Ã­cone padrÃ£o."""
    nome = " ".join(p for p in (branding.NOME_SISTEMA, branding.NOME_EMPRESA.strip()) if p)
    logo = None
    arq = branding.LOGO_ARQUIVO.strip()
    if (
        arq
        and arq == os.path.basename(arq)                       # impede path traversal
        and os.path.splitext(arq)[1].lower() in LOGO_EXTENSOES   # sÃ³ imagens
        and os.path.isfile(os.path.join(app.static_folder, "branding", arq))
    ):
        logo = f"branding/{arq}"
    return {"nome": nome, "logo": logo}


MARCA = _carregar_marca()


@app.context_processor
def injetar_menus():
    css = os.path.join(app.static_folder, "style.css")
    return {"grupos": GRUPOS, "menus": MENUS, "marca": MARCA, "css_v": int(os.path.getmtime(css))}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/menu/<menu_id>", methods=["GET", "POST"])
def menu(menu_id):
    item = MENUS.get(menu_id)  # Só aceita IDs existentes na whitelist
    if item is None:
        abort(404)
    numero = list(MENUS).index(menu_id) + 1

    # Calculadora de Sub-redes (em TOOLS, sem numeração)
    if menu_id == "subnet_calc":
        resultado = None
        erro_msg = None
        ip_input = ""
        if request.method == "POST":
            ip_input = request.form.get("ip_input", "").strip()
            try:
                resultado = subnet_calc.calcular_subrede(ip_input)
            except ValueError as e:
                erro_msg = str(e)
        return render_template(
            "subnet_calc.html",
            menu_id=menu_id,
            menu=item,
            resultado=resultado,
            erro_msg=erro_msg,
            ip_input=ip_input,
            numero=numero
        )

    # Standard Documents (templates_docs, procedures)
    if menu_id in ("templates_docs", "procedures"):
        docs = storage.listar_documentos(categoria=menu_id)
        return render_template("documents.html", menu_id=menu_id, menu=item, docs=docs, numero=numero)

    # Knowledge Base (kb_articles, troubleshooting)
    if menu_id in ("kb_articles", "troubleshooting"):
        topicos = storage.listar_topicos(categoria=menu_id)
        return render_template("kb_list.html", menu_id=menu_id, menu=item, topicos=topicos, numero=numero)

    return render_template("menu.html", menu_id=menu_id, menu=item, numero=numero)


# ---------------- ROTAS STANDARD DOCUMENTS ----------------

@app.route("/documents/<menu_id>/upload", methods=["POST"])
def upload_doc(menu_id):
    if menu_id not in ("templates_docs", "procedures"):
        abort(400)
    arquivo = request.files.get("arquivo")
    if not arquivo or not arquivo.filename:
        abort(400)
    titulo = request.form.get("titulo", "").strip()
    descricao = request.form.get("descricao", "").strip()

    try:
        storage.salvar_documento(arquivo, titulo=titulo, categoria=menu_id, descricao=descricao)
    except ValueError:
        abort(400)

    return redirect(url_for("menu", menu_id=menu_id))


@app.route("/documents/download/<doc_id>")
def download_doc(doc_id):
    caminho, nome_original = storage.obter_caminho_documento(doc_id)
    if not caminho:
        abort(404)
    return send_file(caminho, as_attachment=True, download_name=nome_original)


@app.route("/documents/<menu_id>/delete/<doc_id>", methods=["POST"])
def excluir_doc(menu_id, doc_id):
    if menu_id not in ("templates_docs", "procedures"):
        abort(400)
    senha = request.form.get("senha_exclusao", "")
    if not verificar_senha_exclusao(senha):
        return render_template("erro.html", codigo=403, mensagem="Invalid administrative password. Deletion unauthorized."), 403

    storage.excluir_documento(doc_id)
    return redirect(url_for("menu", menu_id=menu_id))


# ---------------- ROTAS KNOWLEDGE BASE ----------------

@app.route("/kb/<menu_id>/new", methods=["GET", "POST"])
def novo_topico_kb(menu_id):
    item = MENUS.get(menu_id)
    if not item or menu_id not in ("kb_articles", "troubleshooting"):
        abort(404)

    if request.method == "POST":
        titulo = request.form.get("titulo", "").strip()
        autor = request.form.get("autor", "").strip()
        conteudo_md = request.form.get("conteudo_md", "").strip()
        anexos = request.files.getlist("anexos")

        if not titulo or not conteudo_md:
            abort(400)

        topico = storage.criar_topico(
            titulo=titulo,
            categoria=menu_id,
            conteudo_md=conteudo_md,
            autor=autor,
            anexos_arquivos=anexos
        )
        return redirect(url_for("ver_topico_kb", menu_id=menu_id, topico_id=topico["id"]))

    return render_template("kb_new.html", menu_id=menu_id, menu=item)


@app.route("/kb/<menu_id>/view/<topico_id>")
def ver_topico_kb(menu_id, topico_id):
    item = MENUS.get(menu_id)
    if not item or menu_id not in ("kb_articles", "troubleshooting"):
        abort(404)
    topico = storage.obter_topico(topico_id)
    if not topico or topico.get("categoria") != menu_id:
        abort(404)
    return render_template("kb_view.html", menu_id=menu_id, menu=item, topico=topico)


@app.route("/kb/download/<anexo_id>")
def download_kb_anexo(anexo_id):
    caminho, nome_original = storage.obter_caminho_anexo_kb(anexo_id)
    if not caminho:
        abort(404)
    return send_file(caminho, as_attachment=True, download_name=nome_original)


@app.route("/kb/<menu_id>/delete/<topico_id>", methods=["POST"])
def excluir_topico_kb(menu_id, topico_id):
    if menu_id not in ("kb_articles", "troubleshooting"):
        abort(400)
    senha = request.form.get("senha_exclusao", "")
    if not verificar_senha_exclusao(senha):
        return render_template("erro.html", codigo=403, mensagem="Invalid administrative password. Deletion unauthorized."), 403

    storage.excluir_topico(topico_id)
    return redirect(url_for("menu", menu_id=menu_id))


# Armazena tokens temporários de autorização de edição (evita pedir senha duas vezes)
_EDIT_TOKENS = set()

@app.route("/kb/<menu_id>/edit/<topico_id>", methods=["POST"])
def editar_topico_kb(menu_id, topico_id):
    """Valida a senha de administrador e abre o formulário de edição do artigo."""
    item = MENUS.get(menu_id)
    if not item or menu_id not in ("kb_articles", "troubleshooting"):
        abort(404)
    senha = request.form.get("senha_edicao", "")
    if not verificar_senha_exclusao(senha):
        return render_template("erro.html", codigo=403, mensagem="Invalid administrative password. Edit unauthorized."), 403

    topico = storage.obter_topico(topico_id)
    if not topico or topico.get("categoria") != menu_id:
        abort(404)

    token = uuid.uuid4().hex
    _EDIT_TOKENS.add(token)

    return render_template("kb_edit.html", menu_id=menu_id, menu=item, topico=topico, edit_token=token)


@app.route("/kb/<menu_id>/update/<topico_id>", methods=["POST"])
def salvar_edicao_topico_kb(menu_id, topico_id):
    """Salva as alterações feitas no artigo após validação do token."""
    item = MENUS.get(menu_id)
    if not item or menu_id not in ("kb_articles", "troubleshooting"):
        abort(404)

    token = request.form.get("edit_token", "")
    if token not in _EDIT_TOKENS:
        abort(403)
    _EDIT_TOKENS.discard(token)

    titulo = request.form.get("titulo", "").strip()
    autor = request.form.get("autor", "").strip()
    conteudo_md = request.form.get("conteudo_md", "").strip()
    novos_anexos = request.files.getlist("anexos")
    remover_ids = request.form.getlist("remover_anexos")

    if not titulo or not conteudo_md:
        abort(400)

    storage.atualizar_topico(
        topico_id=topico_id,
        titulo=titulo,
        conteudo_md=conteudo_md,
        autor=autor,
        novos_anexos=novos_anexos,
        anexos_remover_ids=remover_ids
    )

    return redirect(url_for("ver_topico_kb", menu_id=menu_id, topico_id=topico_id))


@app.route("/kb/<menu_id>/export/docx/<topico_id>")
def exportar_kb_docx(menu_id, topico_id):
    """Exporta o artigo da Base de Conhecimento para Microsoft Word (.docx)."""
    topico = storage.obter_topico(topico_id)
    if not topico or topico.get("categoria") != menu_id:
        abort(404)
    caminho_docx = exportar_topico_docx(topico)
    nome_download = os.path.basename(caminho_docx)
    return send_file(caminho_docx, as_attachment=True, download_name=nome_download)


@app.route("/kb/<menu_id>/export/pdf/<topico_id>")
def exportar_kb_pdf(menu_id, topico_id):
    """Exporta o artigo da Base de Conhecimento para PDF (.pdf)."""
    topico = storage.obter_topico(topico_id)
    if not topico or topico.get("categoria") != menu_id:
        abort(404)
    caminho_pdf = exportar_topico_pdf(topico)
    nome_download = os.path.basename(caminho_pdf)
    return send_file(caminho_pdf, as_attachment=True, download_name=nome_download)



@app.route("/kb/upload_image", methods=["POST"])
def upload_imagem_kb():
    """Recebe imagem colada da área de transferência ou arrastada via AJAX e salva na pasta da KB."""
    arquivo = request.files.get("image")
    if not arquivo or not arquivo.filename:
        abort(400)
    nome_original = storage.sanitizar_nome_arquivo(arquivo.filename)
    ext = os.path.splitext(nome_original)[1].lower()
    if ext not in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"):
        abort(400)

    img_id = str(uuid.uuid4())[:8]
    nome_armazenado = f"kb_img_{img_id}{ext}"
    caminho = os.path.join(storage.KB_DIR, nome_armazenado)
    os.makedirs(storage.KB_DIR, exist_ok=True)
    arquivo.save(caminho)

    url_imagem = url_for("exibir_imagem_kb", nome_imagem=nome_armazenado)
    return {
        "success": True,
        "url": url_imagem,
        "nome": nome_original,
        "filename": nome_armazenado
    }


@app.route("/kb/image/<nome_imagem>")
def exibir_imagem_kb(nome_imagem):
    """Exibe imagem colada/inserida inline no artigo da Knowledge Base."""
    nome_imagem = os.path.basename(nome_imagem)
    ext = os.path.splitext(nome_imagem)[1].lower()
    if ext not in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"):
        abort(400)
    caminho = os.path.join(storage.KB_DIR, nome_imagem)
    if not os.path.isfile(caminho):
        abort(404)
    return send_file(caminho)


@app.route("/tool/<tool_id>")
def executar_ferramenta(tool_id):
    """Encaminha para a página da ferramenta correspondente."""
    if tool_id == "assessment_visio":
        return render_template("tool_assessment_visio.html", resultado=None, form_data=None, visio_url=None, excel_url=None)
    if tool_id == "portmap_edges":
        return render_template("tool_portmap.html", logs=None, download_url=None)
    if tool_id == "inventory_aps":
        return render_template("tool_inventory_aps.html", logs=None, download_url=None, stats=None)
    if tool_id == "finalizacao_para":
        return render_template("tool_finalizacao_para.html")

    # Ferramentas ainda não integradas
    return render_template(
        "erro.html",
        codigo="Under Construction",
        mensagem=f"A ferramenta '{tool_id}' será integrada na próxima etapa."
    )


@app.route("/action/portmap", methods=["POST"])
def executar_portmap_action():
    """Recebe as credenciais em memória, executa a coleta e descarta usuário/senha."""
    modo = request.form.get("modo", "")
    usuario = request.form.get("usuario", "").strip()
    senha = request.form.get("senha", "")

    if not usuario or not senha:
        abort(400)

    caminho_gerado = None
    logs = []

    try:
        if modo == "autodiscovery":
            cores_raw = request.form.get("cores_input", "")
            # Separa por vírgula, ponto-e-vírgula ou espaço
            cores = [c.strip() for c in re.split(r"[,;\s]+", cores_raw) if c.strip()]
            if not cores:
                abort(400)
            caminho_gerado, logs = portmap_engine.executar_modo_autodiscovery(cores, usuario, senha)

        elif modo == "planilha":
            arquivo = request.files.get("planilha_base")
            if not arquivo or not arquivo.filename:
                abort(400)
            # Salva temporariamente a planilha base na pasta temporária
            temp_name = f"base_{uuid.uuid4().hex[:8]}.xlsx"
            temp_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "data", temp_name)
            os.makedirs(os.path.dirname(temp_path), exist_ok=True)
            arquivo.save(temp_path)

            try:
                caminho_gerado, logs = portmap_engine.executar_modo_planilha(temp_path, usuario, senha)
            finally:
                if os.path.isfile(temp_path):
                    try:
                        os.remove(temp_path)
                    except OSError:
                        pass
        else:
            abort(400)

    except Exception as e:
        logs.append(f"Erro inesperado durante a execução: {e}")

    finally:
        # Descarte seguro imediato de variáveis com a senha
        usuario = None
        senha = None
        gc.collect()

    download_url = None
    if caminho_gerado and os.path.isfile(caminho_gerado):
        nome_arq = os.path.basename(caminho_gerado)
        download_url = url_for("baixar_resultado_portmap", nome_arquivo=nome_arq)

    return render_template("tool_portmap.html", logs=logs, download_url=download_url)


@app.route("/downloads/portmap/<nome_arquivo>")
def baixar_resultado_portmap(nome_arquivo):
    """Download seguro do arquivo Excel de resultado (com validação anti path-traversal)."""
    nome_arquivo = os.path.basename(nome_arquivo)
    if not nome_arquivo.endswith(".xlsx"):
        abort(400)
    caminho = os.path.join(os.path.abspath(os.path.dirname(__file__)), "static", "downloads", nome_arquivo)
    if not os.path.isfile(caminho):
        abort(404)
    return send_file(caminho, as_attachment=True, download_name=nome_arquivo)


# ---------------- ROTAS INVENTORY ACCESS POINTS ----------------

@app.route("/action/inventory_aps", methods=["POST"])
def executar_inventory_aps_action():
    """Recebe parâmetros, executa a coleta/importação e descarta credenciais em memória."""
    modo = request.form.get("modo", "")
    usuario = request.form.get("usuario", "").strip() if modo in ("auto", "original") else None
    senha = request.form.get("senha", "") if modo in ("auto", "original") else None

    caminho_gerado = None
    logs = []
    stats = {}

    try:
        if modo == "auto":
            if not usuario or not senha:
                abort(400)
            wlcs_raw = request.form.get("wlcs_input", "")
            wlcs = [w.strip() for w in re.split(r"[,;\s]+", wlcs_raw) if w.strip()]
            if not wlcs:
                abort(400)
            caminho_gerado, logs, stats = inventory_aps_engine.executar_modo_auto(wlcs, usuario, senha)

        elif modo == "original":
            if not usuario or not senha:
                abort(400)
            arquivo = request.files.get("docx_base")
            if not arquivo or not arquivo.filename:
                abort(400)
            wlcs_raw = request.form.get("wlcs_input", "")
            wlcs = [w.strip() for w in re.split(r"[,;\s]+", wlcs_raw) if w.strip()]
            if not wlcs:
                abort(400)

            temp_name = f"base_ap_{uuid.uuid4().hex[:8]}.docx"
            temp_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "data", temp_name)
            os.makedirs(os.path.dirname(temp_path), exist_ok=True)
            arquivo.save(temp_path)

            try:
                caminho_gerado, logs, stats = inventory_aps_engine.executar_modo_original(temp_path, wlcs, usuario, senha)
            finally:
                if os.path.isfile(temp_path):
                    try:
                        os.remove(temp_path)
                    except OSError:
                        pass

        elif modo == "catalyst_center":
            arquivo = request.files.get("arquivo_catalyst")
            if not arquivo or not arquivo.filename:
                abort(400)
            ext = os.path.splitext(arquivo.filename)[1].lower()
            if ext not in (".xlsx", ".xls", ".csv"):
                abort(400)

            temp_name = f"cat_ap_{uuid.uuid4().hex[:8]}{ext}"
            temp_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "data", temp_name)
            os.makedirs(os.path.dirname(temp_path), exist_ok=True)
            arquivo.save(temp_path)

            try:
                caminho_gerado, logs, stats = inventory_aps_engine.executar_modo_catalyst_center(temp_path)
            finally:
                if os.path.isfile(temp_path):
                    try:
                        os.remove(temp_path)
                    except OSError:
                        pass
        else:
            abort(400)

    except Exception as e:
        logs.append(f"Erro inesperado durante a execução: {e}")

    finally:
        usuario = None
        senha = None
        gc.collect()

    download_url = None
    if caminho_gerado and os.path.isfile(caminho_gerado):
        nome_arq = os.path.basename(caminho_gerado)
        download_url = url_for("baixar_resultado_inventory_aps", nome_arquivo=nome_arq)

    return render_template("tool_inventory_aps.html", logs=logs, download_url=download_url, stats=stats)


@app.route("/downloads/inventory_aps/<nome_arquivo>")
def baixar_resultado_inventory_aps(nome_arquivo):
    """Download seguro do arquivo Word (.docx) com validação anti path-traversal."""
    nome_arquivo = os.path.basename(nome_arquivo)
    if not nome_arquivo.endswith(".docx"):
        abort(400)
    caminho = os.path.join(os.path.abspath(os.path.dirname(__file__)), "static", "downloads", nome_arquivo)
    if not os.path.isfile(caminho):
        abort(404)
    return send_file(caminho, as_attachment=True, download_name=nome_arquivo)


# ---------------- ROTAS PROJETO DE PARA (FINALIZAÇÃO "PARA") ----------------

@app.route("/api/sites", methods=["GET"])
def get_de_para_sites():
    """Retorna a lista de todos os sites cadastrados."""
    configs = de_para_config_store.load_configs()
    return jsonify(configs)


@app.route("/api/sites/<site_id>", methods=["GET"])
def get_de_para_site(site_id):
    """Retorna a configuração de um site específico."""
    config = de_para_config_store.get_site_config(site_id)
    return jsonify(config)


@app.route("/api/sites/<site_id>", methods=["POST"])
def save_de_para_site(site_id):
    """Salva ou atualiza a configuração de um site mantendo a consistência do JSON."""
    data = request.json
    if not data:
        return jsonify({"error": "Dados inválidos"}), 400
    de_para_config_store.update_site_config(site_id, data)
    return jsonify({"status": "success", "site_id": site_id})


@app.route("/api/sites/<site_id>", methods=["DELETE"])
def remove_de_para_site(site_id):
    """Exclui a configuração de um site customizado (exceto o 'default')."""
    if site_id == "default":
        return jsonify({"error": "O perfil Padrão não pode ser excluído"}), 400
    success = de_para_config_store.delete_site_config(site_id)
    if success:
        return jsonify({"status": "success"})
    return jsonify({"error": "Site não encontrado"}), 404


@app.route("/api/process-excel", methods=["POST"])
def process_de_para_excel():
    """Processa a planilha DE PARA e gera o arquivo com colunas R a U preenchidas."""
    if 'excel_file' not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    file = request.files['excel_file']
    if file.filename == '':
        return jsonify({"error": "Nenhum arquivo selecionado"}), 400

    if not file.filename.endswith('.xlsx'):
        return jsonify({"error": "O arquivo precisa estar no formato .xlsx"}), 400

    new_hostname = request.form.get("new_hostname", "SW-NOVO-01").strip()
    site_id = request.form.get("site_id", "default").strip()
    default_vlan_dados = request.form.get("default_vlan_dados", "").strip()

    try:
        new_aps_count = int(request.form.get("new_aps_count", "0"))
    except ValueError:
        new_aps_count = 0

    stack_models = request.form.getlist("stack_models[]")
    if not stack_models:
        stack_models = ["C9300-48U", "C9300-24UX", "C9300-24P"]

    file_uid = uuid.uuid4().hex[:8]
    original_filename = re.sub(r'[^A-Za-z0-9_.-]', '_', file.filename)

    upload_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "data", "temp_uploads")
    download_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "static", "downloads")
    os.makedirs(upload_dir, exist_ok=True)
    os.makedirs(download_dir, exist_ok=True)

    input_filepath = os.path.join(upload_dir, f"{file_uid}_{original_filename}")
    file.save(input_filepath)

    output_filename = f"DE_PARA_MIGRADO_{new_hostname}_{file_uid}.xlsx"
    output_filepath = os.path.join(download_dir, output_filename)

    try:
        summary = de_para_processor.process_de_para_excel(
            file_path=input_filepath,
            output_path=output_filepath,
            new_hostname=new_hostname,
            stack_models=stack_models,
            new_aps_count=new_aps_count,
            site_id=site_id,
            default_vlan_dados=default_vlan_dados
        )

        return jsonify({
            "status": "success",
            "filename": output_filename,
            "summary": summary
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

    finally:
        if os.path.isfile(input_filepath):
            try:
                os.remove(input_filepath)
            except OSError:
                pass


@app.route("/api/download/<filename>", methods=["GET"])
def download_de_para_file(filename):
    """Download do arquivo DE PARA migrado (com validação anti path-traversal)."""
    filename = os.path.basename(filename)
    if not filename.endswith(".xlsx"):
        abort(400)
    download_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "static", "downloads")
    caminho = os.path.join(download_dir, filename)
    if not os.path.isfile(caminho):
        abort(404)
    return send_file(caminho, as_attachment=True, download_name=filename)


@app.route("/api/generate-sample", methods=["POST"])
def generate_de_para_sample():
    """Gera planilha de exemplo para testes do DE PARA."""
    filename = "DE_PARA_Exemplo.xlsx"
    download_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "static", "downloads")
    os.makedirs(download_dir, exist_ok=True)
    filepath = os.path.join(download_dir, filename)
    de_para_sample.create_sample_de_para(filepath)
    return jsonify({
        "status": "success",
        "filename": filename,
        "download_url": url_for('download_de_para_file', filename=filename)
    })


# ---------------- ROTAS ASSESSMENT BIG PICTURE VISIO ----------------

@app.route("/action/assessment_visio", methods=["POST"])
def executar_assessment_visio_action():
    """Recebe parâmetros, executa a varredura recursiva da topologia e descarta credenciais."""
    core1_ip = request.form.get("core1_ip", "").strip()
    core2_ip = request.form.get("core2_ip", "").strip()
    usuario = request.form.get("usuario", "").strip()
    senha = request.form.get("senha", "")
    secret = request.form.get("secret", "").strip()

    if not core1_ip or not usuario or not senha:
        abort(400)

    try:
        porta = int(request.form.get("porta", 22) or 22)
    except (ValueError, TypeError):
        porta = 22

    try:
        max_depth = int(request.form.get("max_depth", 10) or 10)
    except (ValueError, TypeError):
        max_depth = 10

    try:
        timeout = int(request.form.get("timeout", 10) or 10)
    except (ValueError, TypeError):
        timeout = 10

    vendor_preference = request.form.get("vendor_preference", "auto").strip().lower() or "auto"

    form_data = {
        "core1_ip": core1_ip,
        "core2_ip": core2_ip,
        "usuario": usuario,
        "vendor_preference": vendor_preference,
        "porta": porta,
        "max_depth": max_depth,
        "timeout": timeout
    }

    cores_input = [core1_ip]
    if core2_ip:
        cores_input.append(core2_ip)

    resultado = None
    visio_url = None
    excel_url = None

    try:
        resultado = executar_assessment_visio(
            cores_input=cores_input,
            usuario=usuario,
            senha=senha,
            secret=secret,
            vendor_preference=vendor_preference,
            porta=porta,
            max_depth=max_depth,
            timeout=timeout
        )

        if resultado.get("visio_path") and os.path.isfile(resultado["visio_path"]):
            visio_name = os.path.basename(resultado["visio_path"])
            visio_url = url_for("baixar_arquivo_visio", nome_arquivo=visio_name)

        if resultado.get("excel_path") and os.path.isfile(resultado["excel_path"]):
            excel_name = os.path.basename(resultado["excel_path"])
            excel_url = url_for("baixar_planilha_assessment", nome_arquivo=excel_name)

    except Exception as e:
        resultado = {
            "success": False,
            "devices": [],
            "links": [],
            "stats": {},
            "logs": [f"Erro crítico durante a execução do Assessment: {e}"]
        }

    finally:
        usuario = None
        senha = None
        secret = None
        gc.collect()

    return render_template(
        "tool_assessment_visio.html",
        resultado=resultado,
        form_data=form_data,
        visio_url=visio_url,
        excel_url=excel_url
    )


@app.route("/downloads/visio/<nome_arquivo>")
def baixar_arquivo_visio(nome_arquivo):
    """Download seguro do diagrama Microsoft Visio gerado."""
    nome_arquivo = os.path.basename(nome_arquivo)
    if not nome_arquivo.endswith(".vsdx"):
        abort(400)
    caminho = os.path.join(os.path.abspath(os.path.dirname(__file__)), "static", "downloads", nome_arquivo)
    if not os.path.isfile(caminho):
        abort(404)
    return send_file(caminho, as_attachment=True, download_name=nome_arquivo)


@app.route("/downloads/assessment_planilha/<nome_arquivo>")
def baixar_planilha_assessment(nome_arquivo):
    """Download seguro da planilha de capacidade e inventário gerada."""
    nome_arquivo = os.path.basename(nome_arquivo)
    if not nome_arquivo.endswith(".xlsx"):
        abort(400)
    caminho = os.path.join(os.path.abspath(os.path.dirname(__file__)), "static", "downloads", nome_arquivo)
    if not os.path.isfile(caminho):
        abort(404)
    return send_file(caminho, as_attachment=True, download_name=nome_arquivo)


@app.errorhandler(400)
@app.errorhandler(404)
@app.errorhandler(405)
@app.errorhandler(413)
@app.errorhandler(500)
def erro(e):
    # Mensagem genÃ©rica: nÃ£o expÃµe detalhes internos
    codigo = getattr(e, "code", 500)
    return render_template("erro.html", codigo=codigo), codigo


if __name__ == "__main__":
    app.run(host=HOST, port=PORT, debug=False)
