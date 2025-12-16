import os
import re
import json
import time
import zipfile
import traceback
from flask import Flask, render_template, request, send_file, flash, jsonify
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Pt # Importante para o tamanho da fonte
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
# 1. VALIDADOR V26 (COM BLOQUEIO DE OAB DE CABEÇALHO)
# ==============================================================================
def validate_final_term(text):
    if not text: return False
    t = text.strip()
    u = t.upper()
    
    # Regras Básicas
    if len(t) < 3: return False
    if len(t) > 200: return False 
    if re.match(r'^[0-9\-\.\/\s]+$', t): return False 
    if re.match(r'^\d+[\.\-]\s', t): return False # Índices "1. "

    # === BLOQUEIO DE CABEÇALHO ESPECÍFICO ===
    # Bloqueia "OAB: 12345" ou apenas "OAB"
    if u.startswith("OAB"): return False
    
    # Datas
    days = ["SEGUNDA-FEIRA", "TERÇA-FEIRA", "QUARTA-FEIRA", "QUINTA-FEIRA", "SEXTA-FEIRA", "SÁBADO", "DOMINGO"]
    for day in days:
        if day in u: return False
    if re.search(r'\d{1,2}\s+DE\s+[A-ZÇ]+\s+DE\s+\d{4}', u): return False

    # Blocklist
    BLOCK_PHRASES = [
        "GERADO EM", "ASSOCIADO:", "ASSOCIADO",
        "D J E N", "TJRJ", "TJMG", "TJSP", "STJ", "TRT",
        "DISPONIBILIZAÇÃO", "PUBLICAÇÃO", "ARQUIVO:",
        "DIÁRIO ELETRÔNICO", "DIÁRIO DA JUSTIÇA",
        "TIPO DE COMUNICAÇÃO", "MEIO:", "DATA DE",
        "PODER JUDICIÁRIO", "JUSTIÇA DE PRIMEIRA",
        "MINISTÉRIO PÚBLICO", "DEFENSORIA PÚBLICA",
        "ADMINISTRADORA JUDICIAL", "REPRESENTANTE DO MINISTÉRIO",
        "VARA CÍVEL", "VARA EMPRESARIAL", "COMARCA DE", "CARTÓRIO",
        "JUIZ DE DIREITO", "ESCRIVÃ", "DIRETOR DE SECRETARIA",
        "CENTRAL PARA PROCEDER", "ID ", "AUTOS"
    ]

    for block in BLOCK_PHRASES:
        if block in u: return False

    EXACT_BLOCKS = ["DECISÃO", "SENTENÇA", "DESPACHO", "VISTOS", "ÓRGÃO:", "ADVOGADO(S)", "INTIMAÇÃO"]
    if u in EXACT_BLOCKS: return False
        
    return True

# ==============================================================================
# 2. LEITURA
# ==============================================================================
def get_full_text_from_docx(doc):
    full_text = []
    for p in doc.paragraphs:
        full_text.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    full_text.append(p.text)
    return "\n".join(full_text)

# ==============================================================================
# 3. SANITIZAÇÃO
# ==============================================================================
def sanitize_docx_xml(filepath):
    try:
        with zipfile.ZipFile(filepath, 'r') as zin:
            xml_content = zin.read('word/document.xml').decode('utf-8')

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
# 4. EXTRATOR MECÂNICO
# ==============================================================================
def extract_hardcoded_blocks(full_text):
    names_found = []
    text = full_text.replace('\r', '\n')
    
    stop_markers = r'(?:Advogado\(s\)|Intimação|Processo:|Poder Judiciário|Tribunal|Data de)'
    
    partes_regex = rf'Parte\(s\):(.*?){stop_markers}'
    partes_blocks = re.findall(partes_regex, text, re.DOTALL | re.IGNORECASE)
    
    adv_regex = rf'Advogado\(s\)(.*?)(?:ID\s+\d+|Intimação|Processo:|Publ\.|Certifico|$)'
    adv_blocks = re.findall(adv_regex, text, re.DOTALL | re.IGNORECASE)

    raw_list = partes_blocks + adv_blocks
    
    for block in raw_list:
        parts = re.split(r'(?:OAB\s*(?:[A-Z]{2})?-?\s*\d+|\n|;|\s{2,})', block, flags=re.IGNORECASE)
        for part in parts:
            p = part.strip().strip('.,;:- ')
            if validate_final_term(p):
                names_found.append(p)

    return names_found

def extract_processes_regex(full_text):
    processes_found = []
    regex_cnj = r"\d{7}[\s.-]?\d{2}[\s.]?\d{4}[\s.]?\d[\s.]?\d{2}[\s.]?\d{4}"
    regex_agravo = r"\d[\d\.]+\-\d\/\d{3,4}"
    regex_long = r"[0-9]{15,25}"

    processes_found.extend(re.findall(regex_cnj, full_text))
    processes_found.extend(re.findall(regex_agravo, full_text))
    processes_found.extend(re.findall(regex_long, full_text))
    
    return processes_found

# ==============================================================================
# 5. AUDITORIA IA
# ==============================================================================
def audit_missing_entities(full_text, found_names, model_name="gemini-2.0-flash"):
    model = genai.GenerativeModel(model_name)
    chunk_size = 30000 
    overlap = 1000
    chunks = []
    start = 0
    while start < len(full_text):
        end = min(start + chunk_size, len(full_text))
        chunks.append(full_text[start:end])
        start += (chunk_size - overlap)

    new_names = []

    for i, chunk in enumerate(chunks):
        local_context = [n for n in found_names if n in chunk]
        context_str = json.dumps(local_context[:50], ensure_ascii=False)

        try:
            prompt = f"""
            Auditor Jurídico. Extraia o que FALTOU.
            JÁ TENHO: {context_str}
            
            BUSQUE: Empresas, Pessoas, Consórcios.
            IGNORE: "Poder Judiciário", "Gerado em", "Associado", "Ministério Público", "OAB".

            Retorne JSON: {{"missed": ["NOME 1"]}}
            TEXTO: {chunk}
            """
            
            response = model.generate_content(prompt, generation_config={"temperature": 0.1})
            clean_json = response.text.replace("```json", "").replace("```", "").strip()
            data = json.loads(clean_json)
            
            missed = data.get("missed", [])
            if missed:
                valid_missed = [n for n in missed if validate_final_term(n)]
                new_names.extend(valid_missed)
            
            pct = 30 + int((i / len(chunks)) * 50)
            update_status(f"🤖 Auditoria IA: Parte {i+1}/{len(chunks)}...", pct)
            time.sleep(0.5)

        except:
            continue

    return new_names

# ==============================================================================
# 6. MOTOR DE HIGHLIGHT V26 (PRESERVA A FONTE ORIGINAL)
# ==============================================================================
def apply_highlight_brute_force(doc, terms):
    terms = sorted(list(set(terms)), key=len, reverse=True)
    count = 0
    
    all_paragraphs = []
    all_paragraphs.extend(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paragraphs.extend(cell.paragraphs)

    for para in all_paragraphs:
        original_text = para.text
        original_text_lower = original_text.lower()
        
        found_terms_in_para = []
        for term in terms:
            if term.lower() in original_text_lower:
                found_terms_in_para.append(term)
        
        if not found_terms_in_para: continue

        highlights_map = []
        for term in found_terms_in_para:
            start = 0
            while True:
                idx = original_text_lower.find(term.lower(), start)
                if idx == -1: break
                
                is_overlap = False
                for h_start, h_end in highlights_map:
                    if (idx >= h_start and idx < h_end) or (idx + len(term) > h_start and idx + len(term) <= h_end):
                        is_overlap = True
                        break
                    if idx <= h_start and (idx + len(term)) >= h_end:
                        highlights_map.remove((h_start, h_end))
                        is_overlap = False 
                        break
                
                if not is_overlap:
                    highlights_map.append((idx, idx + len(term)))
                    count += 1
                start = idx + 1
        
        if not highlights_map: continue
        highlights_map.sort()
        
        # === CAPTURA DE FONTE ORIGINAL ===
        # Tenta pegar a fonte do primeiro run do parágrafo
        font_name = None
        font_size = None
        if para.runs:
            font_name = para.runs[0].font.name
            font_size = para.runs[0].font.size
        # =================================

        para.clear()
        current_cursor = 0
        
        for start, end in highlights_map:
            # Texto Normal Antes
            if start > current_cursor:
                run = para.add_run(original_text[current_cursor:start])
                if font_name: run.font.name = font_name
                if font_size: run.font.size = font_size
            
            # Texto Grifado
            run = para.add_run(original_text[start:end])
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            run.font.bold = True
            # Reaplica a fonte original também no texto grifado
            if font_name: run.font.name = font_name
            if font_size: run.font.size = font_size
            
            current_cursor = end
            
        # Texto Normal Depois
        if current_cursor < len(original_text):
            run = para.add_run(original_text[current_cursor:])
            if font_name: run.font.name = font_name
            if font_size: run.font.size = font_size
            
    return count

@app.route('/progress')
def progress():
    return jsonify(current_status)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        api_key = request.form.get('api_key')
        use_ai_audit = request.form.get('use_ai_audit') == 'true'
        file = request.files.get('file')

        if not file: return jsonify({"error": "Envie um arquivo"}), 400

        try:
            if api_key: setup_gemini(api_key)
            
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)

            update_status("Sanitizando XML...", 5)
            clean_filepath = sanitize_docx_xml(filepath)
            
            doc = Document(clean_filepath)
            update_status("Lendo documento...", 10)
            full_text = get_full_text_from_docx(doc)

            update_status("Extraindo dados...", 20)
            names_found = extract_hardcoded_blocks(full_text)
            processes_found = extract_processes_regex(full_text)
            
            names_found.append("EDUARDO TAKEMI DUTRA DOS SANTOS KATAOKA")
            names_found.append("EDUARDO TAKEMI KATAOKA")

            if use_ai_audit and api_key:
                update_status("Auditoria IA...", 30)
                base_list = list(set(names_found))
                extra_names = audit_missing_entities(full_text, base_list)
                if extra_names:
                    names_found.extend(extra_names)
                    update_status(f"IA: +{len(extra_names)} nomes.", 80)
            else:
                update_status("Processando...", 80)

            all_raw = list(set(names_found + processes_found))
            
            all_terms = []
            for t in all_raw:
                if re.match(r'^\d', t) or validate_final_term(t):
                    all_terms.append(t)

            update_status(f"Grifando {len(all_terms)} termos...", 90)
            count = apply_highlight_brute_force(doc, all_terms)
            
            update_status("Salvando...", 98)
            base_name = filename.replace('.docx', '')
            output_filename = f"GRIFADO_{base_name}_{int(time.time())}.docx"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            doc.save(output_path)
            
            update_status("Concluído!", 100)
            return jsonify({"status": "success", "download_url": f"/download/{output_filename}", "count": count})

        except Exception as e:
            traceback.print_exc()
            return jsonify({"error": str(e)}), 500

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(UPLOAD_FOLDER, filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)