"""
Módulo de Ferramentas - Projeto Agentes IA Autônomos
Responsável pela geração de documentos, leitura local, buscas na web e APIs.
"""

import os
import re
import time
import json
import requests
import pandas as pd
import urllib.parse
from fpdf import FPDF
from bs4 import BeautifulSoup

# --- Imports para Word ---
from docx import Document
from docx.shared import Pt as DocxPt, RGBColor as DocxRGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH 

# --- Imports para PowerPoint ---
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor

# --- Imports para Excel ---
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows

# ==========================================
# 0. SISTEMA DE CACHE E PESQUISA (COM VALIDADE DE 24H)
# ==========================================
_cache_busca = {}
_cache_pagina = {}

def carregar_cache():
    """Carrega os caches do disco. Se tiverem mais de 24h, são apagados."""
    global _cache_busca, _cache_pagina
    agora = time.time()
    
    def ler_arquivo_cache(nome_arquivo):
        if os.path.exists(nome_arquivo):
            # Calcula a idade do arquivo em segundos (86400 seg = 24 horas)
            idade_arquivo = agora - os.path.getmtime(nome_arquivo)
            if idade_arquivo > 86400:
                print(f"🧹 Limpando cache antigo (mais de 24h): {nome_arquivo}")
                os.remove(nome_arquivo)
                return {} # Retorna vazio para a IA buscar dados frescos
            
            with open(nome_arquivo, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    _cache_busca = ler_arquivo_cache("cache_busca.json")
    _cache_pagina = ler_arquivo_cache("cache_pagina.json")

def salvar_cache():
    """Salva os caches no disco para persistência entre execuções."""
    with open("cache_busca.json", "w", encoding="utf-8") as f:
        json.dump(_cache_busca, f, ensure_ascii=False, indent=2)
    with open("cache_pagina.json", "w", encoding="utf-8") as f:
        json.dump(_cache_pagina, f, ensure_ascii=False, indent=2)

carregar_cache()


# ==========================================
# 1. FERRAMENTAS DE CRIAÇÃO (DOCUMENTOS)
# ==========================================

def criar_apresentacao_com_ia(dados_apresentacao: dict, nome_arquivo: str = "apresentacao.pptx") -> str:
    print(f"🖥️ Montando Apresentação Avançada: {nome_arquivo}")
    prs = Presentation()
    slides_dados = dados_apresentacao.get("slides", [])
    
    for i, slide_info in enumerate(slides_dados):
        titulo = slide_info.get("titulo", f"Slide {i+1}")
        texto = slide_info.get("texto", "")
        prompt_imagem = slide_info.get("prompt_imagem", "")
        
        print(f"   - Processando slide {i+1}: {titulo}")
        
        if i == 0:
            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)
            title = slide.shapes.title
            subtitle = slide.placeholders[1]
            title.text = titulo
            title.text_frame.paragraphs[0].font.size = PptxPt(44)
            title.text_frame.paragraphs[0].font.bold = True
            title.text_frame.paragraphs[0].font.color.rgb = PptxRGBColor(0, 51, 102)
            subtitle.text = texto
        else:
            slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(slide_layout)
            title_shape = slide.shapes.title
            title_shape.text = titulo
            title_shape.text_frame.paragraphs[0].font.size = PptxPt(32)
            title_shape.text_frame.paragraphs[0].font.bold = True
            title_shape.text_frame.paragraphs[0].font.color.rgb = PptxRGBColor(0, 51, 102)
            
            caminho_imagem = None
            if prompt_imagem and prompt_imagem.lower() != "nenhuma":
                for tentativa in range(3):
                    try:
                        # Para os slides, o formato 800x600 (4:3) funciona melhor
                        url_img = f"https://image.pollinations.ai/prompt/{prompt_imagem}?width=800&height=600&nologo=true&model=flux"
                        response = requests.get(url_img, timeout=45)
                        if response.status_code == 200:
                            caminho_imagem = f"temp_slide_{i}.jpg"
                            with open(caminho_imagem, "wb") as f:
                                f.write(response.content)
                            break
                    except:
                        time.sleep(2)
            
            if caminho_imagem:
                pic = slide.shapes.add_picture(caminho_imagem, PptxInches(0.5), PptxInches(2.0), width=PptxInches(4.5))
                txBox = slide.shapes.add_textbox(PptxInches(5.2), PptxInches(2.0), PptxInches(4.3), PptxInches(4.5))
                tf = txBox.text_frame
                tf.word_wrap = True
                p = tf.add_paragraph()
                p.text = texto
                p.font.size = PptxPt(18)
                try: os.remove(caminho_imagem)
                except: pass
            else:
                txBox = slide.shapes.add_textbox(PptxInches(1.0), PptxInches(2.0), PptxInches(8.0), PptxInches(4.5))
                tf = txBox.text_frame
                tf.word_wrap = True
                p = tf.add_paragraph()
                p.text = texto
                p.font.size = PptxPt(20)

    caminho_final = os.path.join(os.getcwd(), nome_arquivo)
    prs.save(caminho_final)
    return f"Apresentação '{nome_arquivo}' criada com sucesso!"

def _adicionar_texto_com_negrito(paragrafo, texto):
    partes = re.split(r'(\*\*.*?\*\*)', texto)
    for parte in partes:
        if parte.startswith('**') and parte.endswith('**') and len(parte) > 4:
            run = paragrafo.add_run(parte[2:-2])
            run.bold = True
        else:
            paragrafo.add_run(parte)

def criar_documento_word(texto: str, nome_arquivo: str = "relatorio.docx", imagens_para_inserir: list = None) -> str:
    print(f"📝 Gerando documento Word profissional: {nome_arquivo}")
    
    texto = re.sub(r'\s*[—–]\s*', ' - ', texto)
    texto = re.sub(r'\s*-\s*,', ',', texto)
    texto = re.sub(r'\s*-\s*\.', '.', texto)
    
    doc = Document()
    linhas = texto.splitlines()

    imagens_colocadas = 0
    imagens_inseridas_por_tag = []

    for linha in linhas:
        linha_limpa = linha.strip()
        if not linha_limpa: continue

        if linha_limpa.replace('-', '') == '' or linha_limpa.replace('*', '') == '':
            continue

        # == CAÇADOR DE IMAGENS ==
        match_tag = re.search(r'\[IMAGEM:\s*(.+?)\]', linha_limpa, re.IGNORECASE)
        match_md = re.search(r'!\[.*?\]\((.+?)\)', linha_limpa)
        
        img_path = None
        if match_tag:
            img_path = match_tag.group(1).strip()
        elif match_md:
            img_path = match_md.group(1).strip()
            
        if img_path:
            if os.path.exists(img_path):
                print(f"   - Inserindo imagem no meio do texto: {img_path}")
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = DocxPt(12)
                p_img.paragraph_format.space_after = DocxPt(12)
                
                r_img = p_img.add_run()
                # NOVO: Inserindo em 5.5 polegadas para ter margens perfeitas na folha A4
                r_img.add_picture(img_path, width=Inches(5.5))
                
                imagens_colocadas += 1
                imagens_inseridas_por_tag.append(img_path)
            else:
                print(f"   - Aviso: Imagem '{img_path}' solicitada não foi encontrada.")
            continue 

        # == PROCESSAMENTO DE TÍTULOS E SUBTÍTULOS ==
        match_heading = re.match(r'^(#+)\s+(.*)', linha_limpa)
        if match_heading:
            nivel_word = min(len(match_heading.group(1)), 9) 
            texto_titulo = match_heading.group(2).replace('**', '')
            h = doc.add_heading(texto_titulo, level=nivel_word)
            
            h.paragraph_format.space_before = DocxPt(18)
            h.paragraph_format.space_after = DocxPt(12)
            
            if nivel_word == 1:
                h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                h.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            continue

        # == PROCESSAMENTO DE LISTAS (Bolinhas) ==
        if re.match(r'^[-*]\s+', linha_limpa):
            p = doc.add_paragraph(style='List Bullet')
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 
            texto_lista = re.sub(r'^[-*]\s+', '', linha_limpa)
            _adicionar_texto_com_negrito(p, texto_lista)
            continue

        # == PROCESSAMENTO DE LISTAS (Números) ==
        if re.match(r'^\d+\.\s+', linha_limpa):
            p = doc.add_paragraph(style='List Number')
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            texto_lista = re.sub(r'^\d+\.\s+', '', linha_limpa)
            _adicionar_texto_com_negrito(p, texto_lista)
            continue

        # == PROCESSAMENTO DE PARÁGRAFOS NORMAIS ==
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 
        p.paragraph_format.first_line_indent = Cm(1.25)
        _adicionar_texto_com_negrito(p, linha_limpa)
        
        for palavra in ['importante:', 'atenção:', 'conclusão:']:
            if palavra in linha_limpa.lower():
                if len(p.runs) > 0:
                    p.runs[0].font.color.rgb = DocxRGBColor(200, 0, 0)
                break

    # == PLANO B (FALLBACK) ==
    if imagens_para_inserir:
        imagens_faltantes = [img for img in imagens_para_inserir if img not in imagens_inseridas_por_tag]
        if imagens_faltantes:
            doc.add_paragraph()
            h_fallback = doc.add_heading('Anexos e Ilustrações Adicionais', level=1)
            h_fallback.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            for img in imagens_faltantes:
                if os.path.exists(img):
                    print(f"   - Inserindo imagem via Fallback (Final do Doc): {img}")
                    p_fallback_img = doc.add_paragraph()
                    p_fallback_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_fallback_img.paragraph_format.space_before = DocxPt(12)
                    p_fallback_img.add_run().add_picture(img, width=Inches(5.5))
                    
                    p_legenda = doc.add_paragraph(f"Ilustração: {img}")
                    p_legenda.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    imagens_colocadas += 1

    caminho_final = os.path.join(os.getcwd(), nome_arquivo)
    doc.save(caminho_final)
    
    return f"Documento Word '{nome_arquivo}' salvo com sucesso! {imagens_colocadas} imagens inseridas."

class RelatorioPDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.set_text_color(0, 51, 102) 
        self.cell(0, 10, 'Relatório Gerado por IA Autônoma', 0, 1, 'R')
        self.line(10, 20, 200, 20) 
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

def criar_pdf(texto: str, nome_arquivo: str = "documento.pdf") -> str:
    print(f"📄 Gerando PDF executivo: {nome_arquivo}")
    pdf = RelatorioPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    texto_limpo = texto.replace('**', '').replace('##', '').replace('#', '')
    for linha in texto_limpo.split('\n'):
        linha = linha.strip()
        if not linha:
            pdf.ln(5)
            continue
        if len(linha) < 60 and not linha.endswith('.'):
            pdf.set_font('Arial', 'B', 14)
            pdf.set_text_color(0, 51, 102)
            pdf.cell(0, 10, linha.encode('latin-1', 'replace').decode('latin-1'), ln=True)
            pdf.ln(2)
        else:
            pdf.set_font('Arial', '', 11)
            pdf.set_text_color(0, 0, 0)
            pdf.multi_cell(0, 6, linha.encode('latin-1', 'replace').decode('latin-1'))

    caminho_final = os.path.join(os.getcwd(), nome_arquivo)
    pdf.output(caminho_final)
    return f"PDF '{nome_arquivo}' criado com sucesso."

def criar_planilha_excel(dados: list, nome_arquivo: str = "dados.xlsx") -> str:
    print(f"📊 Gerando Excel Inteligente: {nome_arquivo}")
    if not dados: return "Erro: A lista de dados fornecida está vazia."

    df = pd.DataFrame(dados)
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório IA"

    for r in dataframe_to_rows(df, index=False, header=True): ws.append(r)

    cor_fundo = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') 
    fonte_branca = Font(bold=True, color='FFFFFF')
    for cell in ws[1]:
        cell.fill = cor_fundo
        cell.font = fonte_branca
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for col in ws.columns:
        tamanho_max = max((len(str(c.value)) for c in col if c.value), default=10)
        ws.column_dimensions[col[0].column_letter].width = tamanho_max + 2

    try:
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Análise de Dados"
        chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=len(df)+1, max_col=2), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=len(df)+1))
        ws.add_chart(chart, f"{chr(65 + len(df.columns) + 1)}2")
    except Exception as e:
        print(f"      ⚠️ Sem gráfico: {e}")

    caminho_final = os.path.join(os.getcwd(), nome_arquivo)
    wb.save(caminho_final)
    return f"Planilha '{nome_arquivo}' gerada com sucesso."

def gerar_imagem(prompt: str, nome_arquivo: str = "imagem_gerada.jpg") -> str:
    print(f"🎨 Gerando imagem: {nome_arquivo}")
    try:
        prompt_turbinado = f"{prompt}, award-winning photography, highly detailed, 8k resolution, ultra-realistic, cinematic lighting"
        prompt_codificado = urllib.parse.quote(prompt_turbinado)
        
        # NOVO: Imagens Widescreen (1024x576) para não quebrar a página do Word!
        url_img = f"https://image.pollinations.ai/prompt/{prompt_codificado}?width=1024&height=576&nologo=true&model=flux"
        response = requests.get(url_img, timeout=45)
        
        if response.status_code == 200:
            caminho_final = os.path.join(os.getcwd(), nome_arquivo)
            with open(caminho_final, "wb") as f:
                f.write(response.content)
            return f"Imagem gerada e salva com sucesso em: {caminho_final}"
        return "Erro: Falha no servidor de imagens."
    except Exception as e:
        return f"Erro ao gerar imagem: {e}"


# ==========================================
# 2. FERRAMENTAS DE PESQUISA E LEITURA (CRÍTICAS)
# ==========================================

def ler_arquivo(caminho: str) -> str:
    print(f"📖 Lendo arquivo: {caminho}")
    if not os.path.exists(caminho):
        return f"Erro: O arquivo '{caminho}' não foi encontrado no sistema."
    
    ext = caminho.split('.')[-1].lower()
    try:
        if ext == 'txt' or ext == 'csv':
            with open(caminho, 'r', encoding='utf-8') as f:
                return f.read()
                
        elif ext == 'xlsx' or ext == 'xls':
            df = pd.read_excel(caminho)
            return f"Conteúdo da Planilha Excel:\n{df.to_string()}"
            
        elif ext == 'docx':
            doc = Document(caminho)
            return "\n".join([p.text for p in doc.paragraphs])
            
        elif ext == 'pdf':
            try:
                from PyPDF2 import PdfReader
                reader = PdfReader(caminho)
                texto = "\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
                return texto if texto else "O PDF não contém texto legível."
            except ImportError:
                return "Erro interno: A biblioteca 'PyPDF2' não está instalada. Execute 'pip install PyPDF2'."
        else:
            return f"Formato de arquivo '.{ext}' não suportado para leitura direta."
    except Exception as e:
        return f"Erro ao tentar ler o arquivo: {e}"

def buscar_web(query: str) -> str:
    print(f"🔍 Buscando na Web: {query}")
    if query in _cache_busca: return _cache_busca[query]
    
    try:
        from ddgs import DDGS
        resultados = DDGS().text(query, max_results=5)
        if not resultados: return "Nenhum resultado encontrado."
        
        texto_resultado = "\n".join([f"Título: {r['title']}\nLink: {r['href']}\nResumo: {r['body']}\n" for r in resultados])
        _cache_busca[query] = texto_resultado
        salvar_cache()
        return texto_resultado
    except ImportError:
        return "Erro: Instale a biblioteca ddgs (pip install ddgs)."
    except Exception as e:
        return f"Erro na busca: {e}"

def abrir_pagina(url: str) -> str:
    print(f"🌐 Lendo página: {url}")
    if url in _cache_pagina: return _cache_pagina[url]
    
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        resposta = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(resposta.text, 'html.parser')
        
        for script in soup(["script", "style", "nav", "footer", "header"]):
            script.extract()
            
        texto = soup.get_text(separator=' ', strip=True)
        texto = texto[:8000] 
        
        _cache_pagina[url] = texto
        salvar_cache()
        return texto
    except Exception as e:
        return f"Erro ao acessar a página: {e}"

def cotacao_moeda(moeda: str) -> str:
    moeda = moeda.upper().strip()
    print(f"💱 Buscando cotação: {moeda}")
    try:
        url = f"https://economia.awesomeapi.com.br/last/{moeda}-BRL"
        req = requests.get(url)
        dados = req.json()
        chave = f"{moeda}BRL"
        valor = dados[chave]["bid"]
        data = dados[chave]["create_date"]
        return f"Cotação {moeda}/BRL: R$ {valor} (Atualizado em: {data})"
    except Exception as e:
        return f"Erro ao buscar cotação. Verifique se a moeda é válida (USD, EUR, BTC). Erro: {e}"


# ==========================================
# 3. REGISTRO GERAL NO DICIONÁRIO
# ==========================================
FERRAMENTAS = {
    "criar_apresentacao_com_ia": criar_apresentacao_com_ia,
    "criar_documento_word": criar_documento_word,
    "criar_pdf": criar_pdf,
    "criar_planilha_excel": criar_planilha_excel,
    "gerar_imagem": gerar_imagem,
    
    "ler_arquivo": ler_arquivo,
    "buscar_web": buscar_web,
    "abrir_pagina": abrir_pagina,
    "cotacao_moeda": cotacao_moeda
}
