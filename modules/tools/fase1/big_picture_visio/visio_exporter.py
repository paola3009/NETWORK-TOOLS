# -*- coding: utf-8 -*-
"""
Módulo de Exportação para Microsoft Visio (.vsdx).
Gera diagramas de topologia de rede executivos contendo:
  - Formas diferenciadas por função (Core, Switch de Acesso, Router, WLC).
  - Rótulos com Hostname, IP de Gerência e Modelo.
  - Conexões (links) automáticas com identificação das interfaces (Portas Local <---> Remota).
  - Layout hierárquico automatizado em camadas (Cores no topo, cascateamento em níveis).
"""

import os
import re
from pathlib import Path
from typing import Dict, List, Any, Tuple
from datetime import datetime

import vsdx

BASE_DIR = Path(__file__).resolve().parents[4]
DOWNLOADS_DIR = BASE_DIR / "core" / "static" / "downloads"
TEMPLATE_VSDX = Path(__file__).resolve().parent / "template_visio.vsdx"


def gerar_topologia_visio(
    nos: List[Dict[str, Any]],
    links: List[Dict[str, Any]],
    nome_arquivo: str = None
) -> str:
    """
    Gera o arquivo .vsdx da topologia a partir dos nós (dispositivos) e links descobertos.
    Retorna o caminho absoluto do arquivo .vsdx gerado.
    """
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    if not nome_arquivo:
        data_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"Topologia_Rede_Visio_{data_str}.vsdx"

    caminho_final = str(DOWNLOADS_DIR / nome_arquivo)

    if not TEMPLATE_VSDX.is_file():
        raise FileNotFoundError(f"Template base do Visio não encontrado: {TEMPLATE_VSDX}")

    vf = vsdx.VisioFile(str(TEMPLATE_VSDX))
    page = vf.pages[0]
    page.set_name("Topologia de Rede")

    # Localiza protótipos de shapes no template
    rect_proto = None
    circle_proto = None
    conn_proto = None

    for s in page.child_shapes:
        txt = (s.text or "").strip().upper()
        if "RECTANGLE" in txt and not rect_proto:
            rect_proto = s
        elif "CIRCLE" in txt and not circle_proto:
            circle_proto = s
        elif "STRAIGHT_CONNECTOR" in txt and not conn_proto:
            conn_proto = s

    if not rect_proto or not conn_proto:
        # Fallback para primeiras formas disponíveis se o texto diferir
        shapes_list = list(page.child_shapes)
        rect_proto = shapes_list[0]
        conn_proto = shapes_list[2] if len(shapes_list) > 2 else shapes_list[0]

    # 1. Agrupamento em Camadas / Níveis Hierárquicos
    camadas: Dict[int, List[Dict[str, Any]]] = {}
    for no in nos:
        lvl = no.get("nivel", 1)
        role = no.get("role", "Switch")
        if role == "Core":
            lvl = 0
        elif role == "Router":
            lvl = 0
        elif role == "WLC":
            lvl = 1

        camadas.setdefault(lvl, []).append(no)

    niveis_ordenados = sorted(camadas.keys())
    max_nos_por_camada = max(len(camadas[lvl]) for lvl in niveis_ordenados) if camadas else 1

    # 2. Configura Dimensões da Página do Visio
    largura_shape = 2.4
    altura_shape = 1.1
    espacamento_x = 3.2
    espacamento_y = 2.4

    page_width = max(16.0, (max_nos_por_camada + 1) * espacamento_x)
    total_camadas = len(niveis_ordenados)
    page_height = max(12.0, (total_camadas + 2) * espacamento_y)

    page.width = page_width
    page.height = page_height

    # 3. Cálculo das Coordenadas (X, Y) para cada nó
    # No Visio, Y=0 é o rodapé e Y=height é o topo
    shape_map: Dict[str, Any] = {}   # id -> vsdx.Shape
    coord_map: Dict[str, Tuple[float, float]] = {}

    y_atual = page_height - 2.0

    for lvl in niveis_ordenados:
        nos_camada = camadas[lvl]
        qtd = len(nos_camada)
        largura_total_camada = (qtd - 1) * espacamento_x
        x_inicio = (page_width - largura_total_camada) / 2.0

        for i, no in enumerate(nos_camada):
            x = x_inicio + (i * espacamento_x)
            y = y_atual

            dev_id = no.get("id") or no.get("hostname")
            role = no.get("role", "Switch")
            hostname = no.get("hostname", "DISPOSITIVO")
            ip = no.get("ip", "")
            modelo = no.get("modelo", "")
            stack = no.get("stack_count", 1)

            # Escolhe o protótipo
            proto = circle_proto if (role == "Router" and circle_proto) else rect_proto
            novo_shape = proto.copy(page)
            novo_shape.width = largura_shape
            novo_shape.height = altura_shape
            novo_shape.x = x
            novo_shape.y = y

            # Texto informativo no nó
            linhas_texto = [hostname]
            if ip:
                linhas_texto.append(f"IP: {ip}")
            if modelo and modelo != "Desconhecido":
                detalhe_modelo = f"Mod: {modelo}"
                if stack > 1:
                    detalhe_modelo += f" (Stack {stack})"
                linhas_texto.append(detalhe_modelo)
            else:
                linhas_texto.append(f"[{role.upper()}]")

            novo_shape.text = "\n".join(linhas_texto)

            # Cores de acordo com o papel na rede
            if role == "Core":
                novo_shape.fill_color = "#1E40AF"  # Azul escuro Core
            elif role == "Router":
                novo_shape.fill_color = "#0D9488"  # Verde Petróleo Router
            elif role == "WLC":
                novo_shape.fill_color = "#7C3AED"  # Roxo WLC
            else:
                novo_shape.fill_color = "#0284C7"  # Azul Acesso

            shape_map[dev_id] = novo_shape
            coord_map[dev_id] = (x, y)

        y_atual -= espacamento_y

    # 4. Criação dos Conectores (Links) com Interfaces
    links_adicionados = set()

    for link in links:
        src_id = link.get("source_id") or link.get("source")
        dst_id = link.get("target_id") or link.get("target")

        # Evita duplicar o mesmo enlace bidirecional
        link_par = tuple(sorted([src_id, dst_id]))
        src_port = link.get("source_port", "")
        dst_port = link.get("target_port", "")
        link_chave = (link_par[0], link_par[1], src_port, dst_port)

        if link_chave in links_adicionados:
            continue
        links_adicionados.add(link_chave)

        coord_src = coord_map.get(src_id)
        coord_dst = coord_map.get(dst_id)

        if coord_src and coord_dst:
            conn = conn_proto.copy(page)
            conn.set_start_and_finish(coord_src, coord_dst)

            # Rótulo nas interfaces do link
            label_link = ""
            if src_port and dst_port:
                label_link = f"{src_port}  <--->  {dst_port}"
            elif src_port:
                label_link = src_port
            elif dst_port:
                label_link = dst_port

            if label_link:
                conn.text = label_link

    # 5. Remove as formas protótipos de exemplo originais da página
    for s in list(page.child_shapes):
        txt_orig = (s.text or "").strip().upper()
        if txt_orig in ["RECTANGLE", "CONNECTED_SHAPE", "STRAIGHT_CONNECTOR", "CURVED_CONNECTOR", "CIRCLE", "LINE"]:
            try:
                s.remove()
            except Exception:
                pass

    vf.save_vsdx(caminho_final)
    return caminho_final
