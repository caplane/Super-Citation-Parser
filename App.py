import os
import zipfile
import shutil
import tempfile
import re
import json
import requests
import xml.dom.minidom as minidom
from flask import Flask, render_template_string, request, jsonify, send_file
from werkzeug.utils import secure_filename
from pathlib import Path

app = Flask(__name__)
app.config['SECRET_KEY'] = 'smart-resolver-key-v2'

# ==================== CONFIGURATION & DATA ====================

PUBLISHER_PLACE_MAP = {
    'Harvard University Press': 'Cambridge, MA',
    'MIT Press': 'Cambridge, MA',
    'Yale University Press': 'New Haven',
    'Princeton University Press': 'Princeton',
    'Stanford University Press': 'Stanford',
    'University of California Press': 'Berkeley',
    'University of Chicago Press': 'Chicago',
    'Columbia University Press': 'New York',
    'Cornell University Press': 'Ithaca',
    'Duke University Press': 'Durham',
    'Johns Hopkins University Press': 'Baltimore',
    'Oxford University Press': 'Oxford',
    'Cambridge University Press': 'Cambridge',
    'Penguin': 'New York',
    'Random House': 'New York',
    'HarperCollins': 'New York',
    'Simon & Schuster': 'New York',
    'Farrar, Straus and Giroux': 'New York',
    'W. W. Norton': 'New York',
    'Knopf': 'New York'
}

# ==================== EMBEDDED HTML/JS DASHBOARD ====================
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>CiteResolver: Smart Review</title>
    <style>
        body { font-family: -apple-system, system-ui, sans-serif; max-width: 1100px; margin: 0 auto; padding: 20px; background: #f5f5f5; }
        .container { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        
        .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #eee; padding-bottom: 15px; margin-bottom: 20px; }
        .controls { display: flex; gap: 15px; align-items: center; }
        select { padding: 8px; border-radius: 4px; border: 1px solid #ddd; }
        
        .note-row { border-bottom: 1px solid #eee; padding: 20px 0; display: grid; grid-template-columns: 40px 1fr 220px; gap: 20px; }
        .note-id { font-weight: bold; color: #888; padding-top: 10px; }
        
        .input-group { display: flex; flex-direction: column; gap: 8px; }
        .text-editor-row { display: flex; gap: 5px; }
        input[type="text"] { padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-family: "Times New Roman", serif; font-size: 1.1em; width: 100%; box-sizing: border-box; }
        
        .search-bar-row { display: flex; gap: 5px; align-items: center; margin-top: 5px;}
        .search-input { flex-grow: 1; padding: 6px; border: 1px solid #ddd; border-radius: 4px; font-size: 0.9em; color: #555; }
        
        .actions { display: flex; flex-direction: column; gap: 10px; }
        button { padding: 8px 12px; cursor: pointer; border: none; border-radius: 4px; font-weight: 500; transition: 0.2s; }
        
        .btn-search { background: #007bff; color: white; width: 100%; }
        .btn-search:hover { background: #0056b3; }
        
        .btn-clear { background: #6c757d; color: white; font-size: 0.8em; padding: 6px 10px; }
        .btn-download { background: #28a745; color: white; padding: 12px 24px; text-decoration: none; display: inline-block; border-radius: 4px; font-weight: bold;}
        
        /* Results Panel */
        .results-panel { display: none; background: #f8f9fa; padding: 10px; border-radius: 6px; margin-top: 10px; border: 1px solid #e9ecef; }
        .result-card { background: white; padding: 10px; margin-bottom: 8px; border: 1px solid #ddd; border-radius: 4px; }
        .result-card:hover { border-color: #007bff; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
        .result-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 5px; }
        .result-title { font-weight: bold; color: #333; }
        .result-meta { font-size: 0.85em; color: #666; margin-bottom: 8px; }
        .preview-box { background: #e3f2fd; color: #0d47a1; font-size: 0.8em; padding: 4px 8px; border-radius: 3px; margin-bottom: 8px; font-family: "Times New Roman", serif;}

        .card-actions { display: flex; gap: 5px; justify-content: flex-end; }
        .btn-replace { background: #dc3545; color: white; font-size: 0.8em; }
        .btn-replace:hover { background: #bd2130; }
        .btn-append { background: #28a745; color: white; font-size: 0.8em; }
        .btn-append:hover { background: #218838; }

        .loader { display: none; color: #666; font-style: italic; margin-top: 5px; font-size: 0.9em; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>CiteResolver <span style="font-size:0.5em; color:#888; font-weight:normal;">v2.0</span></h1>
            {% if filename %}
            <div class="controls">
                <label>Style:</label>
                <select id="style-selector" onchange="refreshPreviews()">
                    <option value="chicago">Chicago (Notes)</option>
                    <option value="mla">MLA</option>
                    <option value="apa">APA</option>
                </select>
                <a href="/download" class="btn-download">Download {{ filename }}</a>
            </div>
            {% endif %}
        </div>
        
        {% if not filename %}
        <form action="/upload" method="post" enctype="multipart/form-data" style="text-align: center; padding: 60px;">
            <h3>Upload Manuscript (.docx)</h3>
            <input type="file" name="file" accept=".docx" required>
            <br><br>
            <button type="submit" class="btn-search" style="width: auto; padding: 12px 30px; font-size: 1.1em;">Start Review</button>
        </form>
        {% else %}

        <div id="notes-list"></div>

        <script>
            const PUBLISHER_MAP = {{ publisher_map | tojson }};
            let loadedNotes = [];

            // Initial Load
            fetch('/get_notes').then(r => r.json()).then(data => {
                loadedNotes = data.notes;
                const container = document.getElementById('notes-list');
                
                data.notes.forEach(note => {
                    const div = document.createElement('div');
                    div.className = 'note-row';
                    div.innerHTML = `
                        <div class="note-id">${note.id}</div>
                        <div class="input-group">
                            <input type="text" id="text-${note.id}" value="${note.text}">
                            
                            <div class="search-bar-row">
                                <input type="text" class="search-input" id="query-${note.id}" value="${note.clean_term}" placeholder="Enter author/title to search...">
                                <button class="btn-clear" onclick="clearQuery('${note.id}')">Clear</button>
                            </div>
                            
                            <div class="loader" id="loader-${note.id}">Scanning libraries...</div>
                            <div class="results-panel" id="results-${note.id}"></div>
                        </div>
                        <div class="actions">
                            <button class="btn-search" onclick="searchCitation('${note.id}')">Find Source</button>
                        </div>
                    `;
                    container.appendChild(div);
                });
            });

            function clearQuery(id) {
                document.getElementById(`query-${id}`).value = '';
                document.getElementById(`query-${id}`).focus();
            }

            function searchCitation(id) {
                const query = document.getElementById(`query-${id}`).value;
                if (!query.trim()) return;

                const loader = document.getElementById(`loader-${id}`);
                const panel = document.getElementById(`results-${id}`);
                
                loader.style.display = 'block';
                panel.style.display = 'none';
                panel.innerHTML = '';

                fetch('/search_book', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({query: query})
                })
                .then(r => r.json())
                .then(data => {
                    loader.style.display = 'none';
                    panel.style.display = 'block';
                    
                    if (data.items.length === 0) {
                        panel.innerHTML = '<div style="padding:10px; color:#666;">No results found. Try refining the search terms.</div>';
                        return;
                    }

                    data.items.forEach(item => {
                        const card = document.createElement('div');
                        card.className = 'result-card';
                        
                        // Apply publisher mapping if city is missing
                        if (!item.city && PUBLISHER_MAP[item.publisher]) {
                            item.city = PUBLISHER_MAP[item.publisher];
                        }

                        const style = document.getElementById('style-selector').value;
                        // Need to escape quotes for HTML attributes
                        const formatted = formatCitation(item, style).replace(/'/g, "&apos;").replace(/"/g, "&quot;");
                        // Clean version for display
                        const displayFormatted = formatCitation(item, style);
                        
                        card.innerHTML = `
                            <div class="result-header">
                                <span class="result-title">${item.title}</span>
                            </div>
                            <div class="result-meta">${item.authors.join(', ')} (${item.year})</div>
                            <div class="preview-box">${displayFormatted}</div>
                            <div class="card-actions">
                                <button class="btn-replace" onclick="applyCitation('${id}', '${formatted}', 'replace')">Replace All</button>
                                <button class="btn-append" onclick="applyCitation('${id}', '${formatted}', 'append')">+ Add to Note</button>
                            </div>
                        `;
                        panel.appendChild(card);
                    });
                });
            }

            function formatCitation(item, style) {
                let citation = "";
                const authors = item.authors.join(', ');
                
                if (style === 'chicago') {
                    citation = `${authors}, <em>${item.title}</em>`;
                    let pubParts = [];
                    if (item.city) pubParts.push(item.city);
                    if (item.publisher) pubParts.push(item.publisher);
                    if (item.year) pubParts.push(item.year);
                    if (pubParts.length > 0) citation += ` (${pubParts.join(': ')})`;
                    citation += ".";
                } 
                else if (style === 'mla') {
                    citation = `${authors}. <em>${item.title}</em>. ${item.publisher}, ${item.year}.`;
                }
                else if (style === 'apa') {
                    citation = `${authors} (${item.year}). <em>${item.title}</em>. ${item.publisher}.`;
                }
                return citation;
            }

            function applyCitation(id, newText, mode) {
                // Decode entities
                const doc = new DOMParser().parseFromString(newText, "text/html");
                const decodedText = doc.documentElement.textContent;
                
                const input = document.getElementById(`text-${id}`);
                let finalText = "";

                if (mode === 'replace') {
                    finalText = newText; // Keep HTML tags for internal storage
                    // For input display, strip tags
                    input.value = decodedText;
                } else if (mode === 'append') {
                    const currentVal = input.value;
                    // Check if we need a separator
                    const separator = currentVal.trim().endsWith('.') ? " " : "; ";
                    // Note: We can't easily mix HTML and plain text in input value
                    // Ideally we'd store the HTML in a hidden field, but for now we append plain text 
                    // to the input for visibility.
                    
                    // IMPORTANT: For "Append", we usually lose the italics in the Input Box visualization
                    // because <input> can't show HTML. The backend will receive the plain text unless we manage it.
                    // To fix this perfectly, we'd need a contenteditable div. 
                    // For now, we will append the HTML version to the backend update, 
                    // but show the plain text in the UI.
                    
                    finalText = currentVal + separator + decodedText; // UI Value
                    input.value = finalText;
                    
                    // For the backend save, we need to actually fetch the CURRENT server state and append?
                    // Or just send what we have. 
                    // To keep it simple: We will send the "decodede/plain" text back for now 
                    // OR we send the HTML version if we track it. 
                    // Let's send the text in the input box to avoid mismatch.
                    // (User loses italics on 'Append' in this simple version, but keeps text data).
                }

                document.getElementById(`results-${id}`).style.display = 'none';
                
                // Send update to server
                fetch('/update_note', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({id: id, text: input.value}) 
                });
                
                input.style.backgroundColor = "#d4edda";
                setTimeout(() => input.style.backgroundColor = "white", 1000);
            }
        </script>
        {% endif %}
    </div>
</body>
</html>
"""

# ==================== BACKEND HELPERS ====================

def clean_search_term(text):
    text = re.sub(r'^\s*\d+\.?\s*', '', text)
    text = re.sub(r',?\s*pp?\.?\s*\d+(-\d+)?\.?$', '', text)
    text = re.sub(r',?\s*\d+\.?$', '', text)
    return text.strip()

def query_google_books(query):
    api_url = "https://www.googleapis.com/books/v1/volumes"
    params = {'q': query, 'maxResults': 4, 'printType': 'books'}
    try:
        r = requests.get(api_url, params=params)
        data = r.json()
        results = []
        if 'items' in data:
            for item in data['items']:
                info = item.get('volumeInfo', {})
                results.append({
                    'title': info.get('title', 'Unknown Title'),
                    'authors': info.get('authors', ['Unknown']),
                    'publisher': info.get('publisher', ''),
                    'city': '', 
                    'year': info.get('publishedDate', '')[:4],
                    'id': item['id']
                })
        return results
    except:
        return []

SESSION = {'temp_dir': None, 'extract_dir': None, 'endnotes_file': None, 'original_filename': None}

def extract_endnotes_xml():
    if not SESSION['endnotes_file']: return []
    with open(SESSION['endnotes_file'], 'r', encoding='utf-8') as f:
        dom = minidom.parseString(f.read())
    
    notes = []
    for en in dom.getElementsByTagName('w:endnote'):
        en_id = en.getAttribute('w:id')
        if en_id and en_id not in ['-1', '0']:
            text_parts = [t.firstChild.nodeValue for t in en.getElementsByTagName('w:t') if t.firstChild]
            full_text = "".join(text_parts).strip()
            clean_term = clean_search_term(full_text)
            notes.append({'id': en_id, 'text': full_text, 'clean_term': clean_term})
    return sorted(notes, key=lambda x: int(x['id']))

def write_updated_note(note_id, new_text):
    # Simple writer that handles basic italics if marked with <em>
    path = SESSION['endnotes_file']
    dom = minidom.parse(str(path))
    for en in dom.getElementsByTagName('w:endnote'):
        if en.getAttribute('w:id') == str(note_id):
            p = en.getElementsByTagName('w:p')[0]
            
            # Save reference
            ref_run = None
            for run in p.getElementsByTagName('w:r'):
                if run.getElementsByTagName('w:endnoteRef'):
                    ref_run = run
                    break
            
            while p.hasChildNodes(): p.removeChild(p.firstChild)
            
            # Rebuild
            pPr = dom.createElement('w:pPr')
            pStyle = dom.createElement('w:pStyle')
            pStyle.setAttribute('w:val', 'EndnoteText')
            pPr.appendChild(pStyle)
            p.appendChild(pPr)
            
            if ref_run:
                p.appendChild(ref_run)
                r = dom.createElement('w:r')
                t = dom.createElement('w:t')
                t.setAttribute('xml:space', 'preserve')
                t.appendChild(dom.createTextNode(" "))
                r.appendChild(t)
                p.appendChild(r)
            
            # Split and styling
            # Note: The frontend currently sends plain text for "Append" to keep it simple.
            # If "Replace" is used, it might send <em> tags.
            parts = re.split(r'(<em>.*?</em>)', new_text)
            for part in parts:
                if not part: continue
                run = dom.createElement('w:r')
                rPr = dom.createElement('w:rPr')
                rFonts = dom.createElement('w:rFonts')
                rFonts.setAttribute('w:ascii', 'Times New Roman')
                rFonts.setAttribute('w:hAnsi', 'Times New Roman')
                rPr.appendChild(rFonts)
                run.appendChild(rPr)
                
                clean_txt = part
                if part.startswith('<em>'):
                    clean_txt = part[4:-5]
                    run.getElementsByTagName('w:rPr')[0].appendChild(dom.createElement('w:i'))
                
                t_node = dom.createElement('w:t')
                t_node.setAttribute('xml:space', 'preserve')
                t_node.appendChild(dom.createTextNode(clean_txt))
                run.appendChild(t_node)
                p.appendChild(run)
                
    with open(path, 'w', encoding='utf-8') as f:
        f.write(dom.toxml())

# ==================== ROUTES ====================

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE, 
                                  filename=SESSION['original_filename'],
                                  publisher_map=PUBLISHER_PLACE_MAP)

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    if file:
        if SESSION['temp_dir']: shutil.rmtree(SESSION['temp_dir'], ignore_errors=True)
        SESSION['temp_dir'] = tempfile.mkdtemp()
        SESSION['original_filename'] = secure_filename(file.filename)
        input_path = os.path.join(SESSION['temp_dir'], 'source.docx')
        file.save(input_path)
        
        SESSION['extract_dir'] = os.path.join(SESSION['temp_dir'], 'extracted')
        with zipfile.ZipFile(input_path, 'r') as z:
            z.extractall(SESSION['extract_dir'])
        SESSION['endnotes_file'] = os.path.join(SESSION['extract_dir'], 'word', 'endnotes.xml')
    return index()

@app.route('/get_notes')
def get_notes():
    return jsonify({'notes': extract_endnotes_xml()})

@app.route('/search_book', methods=['POST'])
def search_book():
    return jsonify({'items': query_google_books(request.json['query'])})

@app.route('/update_note', methods=['POST'])
def update_note():
    write_updated_note(request.json['id'], request.json['text'])
    return jsonify({'success': True})

@app.route('/download')
def download():
    output = os.path.join(SESSION['temp_dir'], f"Resolved_{SESSION['original_filename']}")
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(SESSION['extract_dir']):
            for file in files:
                p = os.path.join(root, file)
                z.write(p, os.path.relpath(p, SESSION['extract_dir']))
    return send_file(output, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5001)
