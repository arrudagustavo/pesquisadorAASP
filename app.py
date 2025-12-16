import os
import re
import json
import time
import zipfile
from flask import Flask, render_template, request, send_file, flash, jsonify
from werkzeug.utils import secure_filename
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import google.generativeai as genai

app = Flask(__name__)
app.secret_key = 'segredo_juridico_absoluto'

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

current_status = {"message": "Aguardando...", "percent": 0}

def update_status(msg, percent):
    global current_status
    current_status["message"] = msg
    current_status["percent"] = percent

def setup_gemini(api_key):
    if api_key:
        genai.configure(api_key=api_key)

# ==============================================================================
# 1. SANITIZAÇÃO (Mantida igual - Essencial para o 'GLAUCIO')
# ==============================================================================
def sanitize_docx_xml(filepath):
    try:
        with zipfile.ZipFile(filepath, 'r') as zin:
            xml_content = zin.read('word/document.xml').decode('utf-8')

        # Remove tags que sujam a leitura
        xml_content = re.sub(r'<w:proofErr[^>]*/>', '', xml_content)
        xml_content = re.sub(r'<w:noBreakHyphen[^>]*/>', '', xml_content)
        xml_content = re.sub(r'<w:softHyphen[^>]*/>', '', xml_content)
        xml_content = re.sub(r'<w:lang[^>]*/>', '', xml_content)
        xml_content = re.sub(r'<w:lastRenderedPageBreak[^>]*/>', '', xml_content)

        new_filepath = filepath.replace('.docx', '_clean.docx')
        
        with zipfile.ZipFile(filepath, 'r') as zin:
            with zipfile.ZipFile(new_filepath, 'w') as zout:
                for item in zin.infolist():
                    if item.filename == 'word/document.xml':
                        zout.writestr(item, xml_content)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        return new_filepath
    except Exception as e:
        print(f"Erro sanitização: {e}")
        return filepath

# ==============================================================================
# 2. EXTRATOR MECÂNICO (REGEX) - Atualizado com Agravo de Instrumento
# ==============================================================================
def extract_hardcoded_blocks(full_text):
    names_found = []
    
    # Padroniza texto
    text = full_text.replace('\r', '\n')
    
    # 1. Captura NOMES (Blocos de Partes e Advogados)
    # A estratégia é pegar tudo entre os rótulos conhecidos
    partes_blocks = re.findall(r'Parte\(s\):(.*?)(?:Advogado\(s\)|Intimação|Processo:)', text, re.DOTALL | re.IGNORECASE)
    adv_blocks = re.findall(r'Advogado\(s\)(.*?)(?:ID\s+\d+|Intimação|Processo:|Publ\.|Certifico|$)', text, re.DOTALL | re.IGNORECASE)

    raw_list = partes_blocks + adv_blocks
    
    for block in raw_list:
        clean_block = block.replace('\n', ' ').strip()
        # Quebra onde tiver "OAB" para separar advogados colados
        parts = re.split(r'OAB\s+[A-Z]{2}-?\d+', clean_block)
        
        for part in parts:
            p = part.strip()
            # Limpa pontuação das pontas
            p = p.strip('.,;:-')
            
            # Filtros de qualidade (tamanho mínimo, não ser número puro)
            if len(p) > 3 and not p.lower().startswith('oab') and not re.match(r'^[0-9\-\.\/\s]+$', p):
                names_found.append(p)

    return names_found

def extract_processes_regex(full_text):
    processes_found = []
    
    # Padrão 1: CNJ Comum (0000000-00.0000.0.00.0000)
    regex_cnj = r"\d{7}[\s.-]?\d{2}[\s.]?\d{4}[\s.]?\d[\s.]?\d{2}[\s.]?\d{4}"
    
    # Padrão 2: Agravo de Instrumento / Numeração Antiga com Barra (1.0000.24.175224-5/004)
    # Explicação: Sequência de dígitos e pontos, traço, digito, barra, 3 ou 4 dígitos
    regex_agravo = r"\d[\d\.]+\-\d\/\d{3,4}"
    
    # Padrão 3: Numeração sequencial longa (apenas segurança)
    regex_long = r"[0-9]{15,25}"

    # Combina todas as buscas
    matches_cnj = re.findall(regex_cnj, full_text)
    matches_agravo = re.findall(regex_agravo, full_text)
    matches_long = re.findall(regex_long, full_text)
    
    processes_found.extend(matches_cnj)
    processes_found.extend(matches_agravo)
    processes_found.extend(matches_long)
    
    return processes_found

# ==============================================================================
# 3. MOTOR DE HIGHLIGHT (Reconstrução de Parágrafo)
# ==============================================================================
def apply_highlight_brute_force(doc, terms):
    # Ordena por tamanho (Decrescente) para evitar destacar "Silva" dentro de "Silva Souza"
    terms = sorted(list(set(terms)), key=len, reverse=True)
    count = 0
    
    for para in doc.paragraphs:
        original_text = para.text
        original_text_lower = original_text.lower()
        
        # Verifica se algum termo está neste parágrafo
        found_terms_in_para = []
        for term in terms:
            if term.lower() in original_text_lower:
                found_terms_in_para.append(term)
        
        if not found_terms_in_para:
            continue

        # Se achou, vamos reconstruir o parágrafo
        # Mapeia onde estão os destaques: (início, fim)
        highlights_map = []
        for term in found_terms_in_para:
            start = 0
            while True:
                idx = original_text_lower.find(term.lower(), start)
                if idx == -1:
                    break
                
                # Checa colisão com destaques já marcados
                is_overlap = False
                for h_start, h_end in highlights_map:
                    # Se o novo termo começa dentro ou termina dentro de um existente
                    if (idx >= h_start and idx < h_end) or (idx + len(term) > h_start and idx + len(term) <= h_end):
                        is_overlap = True
                        break
                    # Se o novo termo engloba totalmente um existente (remove o menor)
                    if idx <= h_start and (idx + len(term)) >= h_end:
                        highlights_map.remove((h_start, h_end))
                        is_overlap = False # Agora não é overlap, é substituição
                        break
                
                if not is_overlap:
                    highlights_map.append((idx, idx + len(term)))
                    count += 1
                
                start = idx + 1
        
        if not highlights_map:
            continue
            
        highlights_map.sort()
        
        # Reconstrói o parágrafo visualmente
        para.clear()
        current_cursor = 0
        
        for start, end in highlights_map:
            # Texto normal antes do destaque
            if start > current_cursor:
                para.add_run(original_text[current_cursor:start])
            
            # Texto Destacado
            run = para.add_run(original_text[start:end])
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            run.font.bold = True
            
            current_cursor = end
            
        # Resto do texto
        if current_cursor < len(original_text):
            para.add_run(original_text[current_cursor:])
            
    return count

@app.route('/progress')
def progress():
    return jsonify(current_status)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # API Key é opcional agora se o documento for padrão AASP
        api_key = request.form.get('api_key')
        file = request.files.get('file')

        if not file:
            return jsonify({"error": "Envie um arquivo"}), 400

        try:
            setup_gemini(api_key)
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)

            # 1. SANITIZAÇÃO (Crucial para nomes quebrados)
            update_status("Higienizando XML do Word...", 10)
            clean_filepath = sanitize_docx_xml(filepath)
            
            doc = Document(clean_filepath)
            full_text = "\n".join([p.text for p in doc.paragraphs])

            # 2. EXTRAÇÃO 100% PYTHON (Mais rápida e precisa para AASP)
            update_status("Extraindo Nomes e Processos (Modo Turbo Python)...", 30)
            
            # Pega nomes via Regex Estrutural
            names_found = extract_hardcoded_blocks(full_text)
            
            # Pega processos via Regex Matemático (Incluindo Agravo)
            processes_found = extract_processes_regex(full_text)
            
            # Adiciona termos manuais
            names_found.append("EDUARDO TAKEMI DUTRA DOS SANTOS KATAOKA")
            names_found.append("EDUARDO TAKEMI KATAOKA")
            
            # Junta tudo
            all_terms = list(set(names_found + processes_found))
            
            # (Opcional) Se tiver API Key e quiser usar IA para limpar mais ainda, poderia chamar aqui
            # Mas como o usuário disse que ficou rápido e bom, vamos confiar no Python.
            
            update_status(f"Aplicando destaque em {len(all_terms)} termos...", 70)
            
            count = apply_highlight_brute_force(doc, all_terms)
            
            update_status("Salvando arquivo...", 90)
            output_filename = f"FINAL_{filename}"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            doc.save(output_path)
            
            update_status("Concluído!", 100)
            return jsonify({"status": "success", "download_url": f"/download/{output_filename}", "count": count})

        except Exception as e:
            return jsonify({"error": str(e)}), 500

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(UPLOAD_FOLDER, filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)