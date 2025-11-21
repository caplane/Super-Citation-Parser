import os
import zipfile
import shutil
import tempfile
import re
import json
import requests
import xml.dom.minidom as minidom
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse, unquote

app = Flask(__name__)
app.config['SECRET_KEY'] = 'resolver-v7-gov-docs'

# ==================== CONFIGURATION ====================
PUBLISHER_PLACE_MAP = {
    'Harvard University Press': 'Cambridge, MA',
    'MIT Press': 'Cambridge, MA',
    'Yale University Press': 'New Haven',
    'Princeton University Press': 'Princeton',
    'Stanford University Press': 'Stanford',
    'University of California Press': 'Berkeley',
    'University of Chicago Press': 'Chicago',
    'Columbia University Press': 'New York',
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

# ==================== BACKEND LOGIC ====================

def clean_search_term(text):
    """Strip numbers and page refs for the 'Suggested Search' field"""
    text = re.sub(r'^\s*\d+\.?\s*', '', text)
    text = re.sub(r',?\s*pp?\.?\s*\d+(-\d+)?\.?$', '', text)
    text = re.sub(r',?\s*\d+\.?$', '', text)
    return text.strip()

def fetch_web_metadata(url):
    """Fetch metadata, detecting Government (.gov) and PDF files"""
    if not url.startswith('http'):
        url = 'http://' + url
        
    try:
        # 1. Parse Domain & Check for Gov
        parsed_uri = urlparse(url)
        domain = parsed_uri.netloc
        if domain.startswith('www.'):
            domain = domain[4:]
            
        is_gov = domain.endswith('.gov')
        result_type = 'gov' if is_gov else 'web'
        
        # 2. Handle PDFs (Common for Gov Docs)
        path = parsed_uri.path.lower()
        if path.endswith('.pdf'):
            # Extract filename as candidate title
            filename = unquote(os.path.basename(parsed_uri.path))
            title_candidate = filename.replace('.pdf', '').replace('_', ' ').replace('-', ' ').title()
            
            return [{
                'type': result_type,
                'title': title_candidate,
                'authors': [],
                'publisher': 'U.S. Government' if is_gov else '',
                'year': '',
                'url': url,
                'domain': domain,
                'access_date': datetime.now().strftime("%B %d, %Y"),
                'id': 'web_pdf_result'
            }]

        # 3. Fetch HTML Page Title
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = requests.get(url, headers=headers, timeout=5)
        
        title_match = re.search(r'<title>(.*?)</title>', response.text, re.IGNORECASE | re.DOTALL)
        
        page_title = "Unknown Document"
        if title_match:
            page_title = title_match.group(1).strip()
            # Clean entities and common suffixes
            page_title = page_title.replace('&ndash;', '-').replace('&mdash;', '-').replace('&amp;', '&').replace('&#039;', "'")
            page_title = re.split(r'\s+[|\-]\s+', page_title)[0]
            
        today = datetime.now().strftime("%B %d, %Y")
        
        return [{
            'type': result_type,
            'title': page_title,
            'authors': [],
            'publisher': 'U.S. Government' if is_gov else '', 
            'year': '',
            'url': url,
            'domain': domain, 
            'access_date': today,
            'id': 'web_result'
        }]

    except Exception as e:
        print(f"URL Fetch Error: {e}")
        parsed_uri = urlparse(url)
        domain = parsed_uri.netloc.replace('www.', '')
        return [{
            'type': 'gov' if domain.endswith('.gov') else 'web',
            'title': "Web Resource (Connection Failed)",
            'authors': [],
            'publisher': '',
            'year': '',
            'url': url,
            'domain': domain,
            'access_date': datetime.now().strftime("%B %d, %Y"),
            'id': 'web_result_failed'
        }]

def query_google_books(query):
    """Router: decides whether to search Google Books or fetch URL"""
    
    # 1. Check if input is a URL
    url_pattern = re.compile(r'^(http|www\.)', re.IGNORECASE)
    if url_pattern.match(query):
        return fetch_web_metadata(query)

    # 2. Otherwise, search Google Books
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
                    'type': 'book',
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
            html_parts = []
            full_text_parts = []
            
            p = en.getElementsByTagName('w:p')[0]
            for run in p.getElementsByTagName('w:r'):
                if run.getElementsByTagName('w:endnoteRef'): continue
                
                text = "".join([t.firstChild.nodeValue for t in run.getElementsByTagName('w:t') if t.firstChild])
                if not text: continue
                
                full_text_parts.append(text)
                
                rPr = run.getElementsByTagName('w:rPr')
                is_italic = False
                if rPr and rPr[0].getElementsByTagName('w:i'):
                    is_italic = True
                
                if is_italic:
                    html_parts.append(f"<em>{text}</em>")
                else:
                    html_parts.append(text)
            
            final_html = "".join(html_parts).strip()
            clean_term = clean_search_term("".join(full_text_parts).strip())
            
            notes.append({'id': en_id, 'html': final_html, 'clean_term': clean_term})
            
    return sorted(notes, key=lambda x: int(x['id']))

def write_updated_note(note_id, html_content):
    path = SESSION['endnotes_file']
    dom = minidom.parse(str(path))
    
    for en in dom.getElementsByTagName('w:endnote'):
        if en.getAttribute('w:id') == str(note_id):
            p = en.getElementsByTagName('w:p')[0]
            
            ref_run = None
            for run in p.getElementsByTagName('w:r'):
                if run.getElementsByTagName('w:endnoteRef'):
                    ref_run = run
                    break
            
            while p.hasChildNodes(): p.removeChild(p.firstChild)
            
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
            
            parts = re.split(r'(<em>.*?</em>)', html_content)
            
            for part in parts:
                if not part: continue
                
                clean_part = part.replace('&nbsp;', ' ').replace('&amp;', '&')
                
                run = dom.createElement('w:r')
                rPr = dom.createElement('w:rPr')
                
                rFonts = dom.createElement('w:rFonts')
                rFonts.setAttribute('w:ascii', 'Times New Roman')
                rFonts.setAttribute('w:hAnsi', 'Times New Roman')
                rPr.appendChild(rFonts)
                run.appendChild(rPr)
                
                text_content = clean_part
                
                if clean_part.startswith('<em>') and clean_part.endswith('</em>'):
                    text_content = clean_part[4:-5]
                    run.getElementsByTagName('w:rPr')[0].appendChild(dom.createElement('w:i'))
                
                t = dom.createElement('w:t')
                t.setAttribute('xml:space', 'preserve')
                t.appendChild(dom.createTextNode(text_content))
                run.appendChild(t)
                p.appendChild(run)
                
    with open(path, 'w', encoding='utf-8') as f:
        f.write(dom.toxml())

# ==================== ROUTES ====================
@app.route('/')
def index():
    return render_template('index.html', 
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
    data = request.json
    write_updated_note(data['id'], data['html'])
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
