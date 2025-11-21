import os
import zipfile
import shutil
import tempfile
import re
import json
import requests
import uuid
import xml.dom.minidom as minidom
from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
from werkzeug.utils import secure_filename
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse, unquote

app = Flask(__name__)
app.config['SECRET_KEY'] = 'production-key-v17-final-link-fix'

# ==================== GLOBAL STORAGE ====================
# Storage for temporary file paths, keyed by user session ID
USER_DATA_STORE = {}

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

# Maps .gov domains to proper Agency Author Names
GOV_AGENCY_MAP = {
    'ferc.gov': 'Federal Energy Regulatory Commission',
    'epa.gov': 'Environmental Protection Agency',
    'energy.gov': 'U.S. Department of Energy',
    'doi.gov': 'U.S. Department of the Interior',
    'justice.gov': 'U.S. Department of Justice',
    'regulations.gov': 'U.S. Government', 
    'fda.gov': 'U.S. Food and Drug Administration',
    'cdc.gov': 'Centers for Disease Control and Prevention',
    'nih.gov': 'National Institutes of Health',
    'usda.gov': 'U.S. Department of Agriculture',
}

# ==================== RELATIONSHIP MANAGER ====================
class RelationshipManager:
    """Handles adding URLs to word/_rels/endnotes.xml.rels"""
    
    def __init__(self, extract_dir):
        self.rels_dir = os.path.join(extract_dir, 'word', '_rels')
        self.rels_path = os.path.join(self.rels_dir, 'endnotes.xml.rels')
        self.relationships = []
        self.next_id = 1
        self._load()

    def _load(self):
        if not os.path.exists(self.rels_dir):
            os.makedirs(self.rels_dir)
        if os.path.exists(self.rels_path):
            with open(self.rels_path, 'r', encoding='utf-8') as f:
                dom = minidom.parseString(f.read())
                for rel in dom.getElementsByTagName('Relationship'):
                    rid = rel.getAttribute('Id')
                    match = re.search(r'\d+', rid)
                    if match:
                        num_id = int(match.group())
                        self.next_id = max(self.next_id, num_id + 1)
                    self.relationships.append({
                        'Id': rid,
                        'Type': rel.getAttribute('Type'),
                        'Target': rel.getAttribute('Target'),
                        'TargetMode': rel.getAttribute('TargetMode')
                    })

    def get_or_create_hyperlink(self, url):
        """Returns the rId for a URL, creating a new Relationship if needed"""
        for rel in self.relationships:
            if rel['Target'] == url and rel['Type'].endswith('/hyperlink'):
                return rel['Id']
        new_id = f"rId{self.next_id}"
        self.next_id += 1
        self.relationships.append({
            'Id': new_id,
            'Type': "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            'Target': url,
            'TargetMode': "External"
        })
        self._save()
        return new_id

    def _save(self):
        root = minidom.Document()
        rels_elem = root.createElement('Relationships')
        rels_elem.setAttribute('xmlns', "http://schemas.openxmlformats.org/package/2006/relationships")
        root.appendChild(rels_elem)
        for rel in self.relationships:
            node = root.createElement('Relationship')
            node.setAttribute('Id', rel['Id'])
            node.setAttribute('Type', rel['Type'])
            node.setAttribute('Target', rel['Target'])
            if rel.get('TargetMode'):
                node.setAttribute('TargetMode', rel['TargetMode'])
            rels_elem.appendChild(node)
        with open(self.rels_path, 'w', encoding='utf-8') as f:
            f.write(root.toxml())

# ==================== BACKEND LOGIC ====================

def get_user_data():
    if 'user_id' not in session: return None
    return USER_DATA_STORE.get(session['user_id'])

def clean_search_term(text):
    text = re.sub(r'^\s*\d+\.?\s*', '', text)
    text = re.sub(r',?\s*pp?\.?\s*\d+(-\d+)?\.?$', '', text)
    text = re.sub(r',?\s*\d+\.?$', '', text)
    return text.strip()

def get_agency_name(domain):
    """Returns the official agency name based on the domain."""
    parts = domain.split('.')
    if len(parts) >= 2:
        root_domain = f"{parts[-2]}.{parts[-1]}"
        if root_domain in GOV_AGENCY_MAP:
            return GOV_AGENCY_MAP[root_domain]
    return "U.S. Government"

def get_heuristic_title(url):
    """Generates a title from the URL slug if scraping fails."""
    parsed_uri = urlparse(url)
    path = parsed_uri.path
    slug = [s for s in path.split('/') if s][-1] if path.split('/') else ''
    
    clean_filename = unquote(slug).replace('_', ' ').replace('-', ' ').title()
    
    if clean_filename.lower().endswith(('.pdf', '.htm', '.html', '.aspx', '.asp', '.php')):
        clean_filename = clean_filename[:-4].strip()
        
    if not clean_filename or len(clean_filename) < 5:
        domain = parsed_uri.netloc.replace('www.', '')
        return domain.split('.')[0].title() 
        
    return clean_filename

def fetch_web_metadata(url):
    if not url.startswith('http'): url = 'http://' + url
    
    try:
        parsed_uri = urlparse(url)
        domain = parsed_uri.netloc.replace('www.', '')
        is_gov = domain.endswith('.gov')
        result_type = 'gov' if is_gov else 'web'
        
        # Determine Agency Author Name (Guaranteed for .gov)
        author_name = get_agency_name(domain) if is_gov else ""
        
        last_updated = None
        
        # 1. Start with Heuristic Title as the safest option
        page_title = get_heuristic_title(url)
        
        # 2. Try to fetch real title from the server
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
        }
        
        try:
            s = requests.Session()
            response = s.get(url, headers=headers, timeout=7, allow_redirects=True)
            
            if response.status_code == 200:
                # Scrape <title>
                title_match = re.search(r'<title>(.*?)</title>', response.text, re.IGNORECASE | re.DOTALL)
                if title_match:
                    raw_title = title_match.group(1).strip()
                    if not any(block_word in raw_title for block_word in ["Just a moment", "Access Denied", "Error", "404"]):
                         page_title = re.sub(r' \| .*$', '', raw_title).strip()
                         
                # Scrape Last Modified/Published Date from headers 
                if 'Last-Modified' in response.headers:
                    try:
                        last_updated = datetime.strptime(response.headers['Last-Modified'][:25], '%a, %d %b %Y %H:%M:%S').strftime("%B %d, %Y")
                    except:
                         pass
        except:
            pass # Use heuristic title and generic date if request fails

        # Final cleanup for output
        access_date = datetime.now().strftime("%B %d, %Y")
        
        return [{
            'type': result_type, 
            'title': page_title, 
            'authors': [author_name] if author_name else [], 
            'publisher': 'U.S. Government' if is_gov else '', 
            'year': '', 
            'url': url, 
            'last_updated': last_updated,
            'access_date': access_date, 
            'id': 'web_result'
        }]
    except:
        domain = urlparse(url).netloc.replace('www.', '')
        return [{'type': 'gov' if domain.endswith('.gov') else 'web', 
            'title': "Web Resource (Fatal Error)", 
            'authors': [], 
            'publisher': '', 
            'year': '', 
            'url': url, 
            'last_updated': None,
            'access_date': datetime.now().strftime("%B %d, %Y"), 
            'id': 'web_result_failed'
        }]

def query_google_books(query):
    url_pattern = re.compile(r'^(http|www\.)', re.IGNORECASE)
    if url_pattern.match(query): return fetch_web_metadata(query)
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
    except: return []

def extract_endnotes_xml(user_data):
    """
    Reads the endnotes.xml and converts complex Word markup (italics, hyperlinks) 
    into simple HTML for the editor, preserving original link targets.
    """
    if not user_data or not user_data['endnotes_file']: return []
    with open(user_data['endnotes_file'], 'r', encoding='utf-8') as f:
        dom = minidom.parseString(f.read())
    
    notes = []
    
    # Find the relationships file path (needed to resolve existing hyperlinks)
    rels_path = os.path.join(user_data['extract_dir'], 'word', '_rels', 'endnotes.xml.rels')
    relationships = {}
    if os.path.exists(rels_path):
        rels_dom = minidom.parse(rels_path)
        for rel in rels_dom.getElementsByTagName('Relationship'):
            if rel.getAttribute('Type').endswith('/hyperlink'):
                relationships[rel.getAttribute('Id')] = rel.getAttribute('Target')

    for en in dom.getElementsByTagName('w:endnote'):
        en_id = en.getAttribute('w:id')
        if en_id and en_id not in ['-1', '0']:
            html_parts = []
            full_text_parts = []
            
            p = en.getElementsByTagName('w:p')[0]
            
            # --- Iterate through direct children of the paragraph ---
            for node in p.childNodes:
                
                if node.nodeType == node.ELEMENT_NODE:
                    
                    # Case 1: Existing Hyperlink (<w:hyperlink>)
                    if node.tagName == 'w:hyperlink':
                        r_id = node.getAttribute('r:id')
                        url = relationships.get(r_id, '#') # Get original URL target
                        
                        # Extract the text content from the runs inside the hyperlink
                        link_text = ""
                        for run in node.getElementsByTagName('w:r'):
                            text = "".join([t.firstChild.nodeValue for t in run.getElementsByTagName('w:t') if t.firstChild])
                            link_text += text
                            full_text_parts.append(text)
                            
                        # Convert to HTML anchor tag
                        html_parts.append(f'<a href="{url}">{link_text}</a>')
                        continue

                    # Case 2: Standard Run (<w:r>) (for plain text, spaces, or italics)
                    if node.tagName == 'w:r':
                        # Check for Endnote Marker and skip it
                        if node.getElementsByTagName('w:endnoteRef'): continue
                        
                        text = "".join([t.firstChild.nodeValue for t in node.getElementsByTagName('w:t') if t.firstChild])
                        if not text: continue
                        full_text_parts.append(text)
                        
                        # Check Italics
                        rPr = node.getElementsByTagName('w:rPr')
                        is_italic = False
                        if rPr and rPr[0].getElementsByTagName('w:i'): is_italic = True
                        
                        if is_italic: html_parts.append(f"<em>{text}</em>")
                        else: html_parts.append(text)
            
            final_html = "".join(html_parts).strip()
            clean_term = clean_search_term("".join(full_text_parts).strip())
            notes.append({'id': en_id, 'html': final_html, 'clean_term': clean_term})
            
    return sorted(notes, key=lambda x: int(x['id']))

def write_updated_note(user_data, note_id, html_content):
    if not user_data: return
    path = user_data['endnotes_file']
    dom = minidom.parse(str(path))
    rel_mgr = RelationshipManager(user_data['extract_dir'])
    
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
            
            # HYPERLINK PARSER: Split by <a> tags or <em> tags
            tokens = re.split(r'(<a href="[^"]+">.*?</a>|<em>.*?</em>)', html_content)
            
            for token in tokens:
                if not token: continue
                token = token.replace('&nbsp;', ' ').replace('&amp;', '&')
                
                # Case 1: Hyperlink (New or Existing)
                if token.startswith('<a href='):
                    match = re.match(r'<a href="([^"]+)">(.*?)</a>', token)
                    if match:
                        url = match.group(1)
                        text = match.group(2)
                        
                        r_id = rel_mgr.get_or_create_hyperlink(url)
                        
                        hlink = dom.createElement('w:hyperlink')
                        hlink.setAttribute('r:id', r_id)
                        
                        run = dom.createElement('w:r')
                        rPr = dom.createElement('w:rPr')
                        rStyle = dom.createElement('w:rStyle')
                        rStyle.setAttribute('w:val', 'Hyperlink')
                        rPr.appendChild(rStyle)
                        
                        color = dom.createElement('w:color')
                        color.setAttribute('w:val', '0000FF')
                        rPr.appendChild(color)
                        u = dom.createElement('w:u')
                        u.setAttribute('w:val', 'single')
                        rPr.appendChild(u)
                        
                        run.appendChild(rPr)
                        t = dom.createElement('w:t')
                        t.appendChild(dom.createTextNode(text))
                        run.appendChild(t)
                        
                        hlink.appendChild(run)
                        p.appendChild(hlink)
                        
                        # --- CRITICAL FIX: Ensure the relationships XML is saved immediately ---
                        rel_mgr._save()
                        continue

                # Case 2: Text (Italic or Plain)
                run = dom.createElement('w:r')
                rPr = dom.createElement('w:rPr')
                rFonts = dom.createElement('w:rFonts')
                rFonts.setAttribute('w:ascii', 'Times New Roman')
                rFonts.setAttribute('w:hAnsi', 'Times New Roman')
                rPr.appendChild(rFonts)
                run.appendChild(rPr)
                
                text_content = token
                if token.startswith('<em>'):
                    text_content = token[4:-5]
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
    if 'user_id' not in session: session['user_id'] = str(uuid.uuid4())
    user_data = get_user_data()
    filename = user_data['original_filename'] if user_data else None
    return render_template('index.html', filename=filename, publisher_map=PUBLISHER_PLACE_MAP)

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    if file:
        if 'user_id' not in session: session['user_id'] = str(uuid.uuid4())
        user_id = session['user_id']
        if user_id in USER_DATA_STORE:
            try: shutil.rmtree(USER_DATA_STORE[user_id]['temp_dir'])
            except: pass
        temp_dir = tempfile.mkdtemp()
        original_filename = secure_filename(file.filename)
        input_path = os.path.join(temp_dir, 'source.docx')
        file.save(input_path)
        extract_dir = os.path.join(temp_dir, 'extracted')
        with zipfile.ZipFile(input_path, 'r') as z: z.extractall(extract_dir)
        endnotes_file = os.path.join(extract_dir, 'word', 'endnotes.xml')
        USER_DATA_STORE[user_id] = {
            'temp_dir': temp_dir,
            'extract_dir': extract_dir,
            'endnotes_file': endnotes_file,
            'original_filename': original_filename
        }
    return index()

@app.route('/reset')
def reset():
    user_id = session.get('user_id')
    if user_id and user_id in USER_DATA_STORE:
        try: shutil.rmtree(USER_DATA_STORE[user_id]['temp_dir'])
        except: pass
        del USER_DATA_STORE[user_id]
    return redirect(url_for('index'))

@app.route('/get_notes')
def get_notes(): 
    user_data = get_user_data()
    if not user_data: return jsonify({'notes': []})
    return jsonify({'notes': extract_endnotes_xml(user_data)})

@app.route('/search_book', methods=['POST'])
def search_book(): return jsonify({'items': query_google_books(request.json['query'])})

@app.route('/update_note', methods=['POST'])
def update_note():
    user_data = get_user_data()
    if not user_data: return jsonify({'success': False, 'error': 'Session expired'})
    data = request.json
    write_updated_note(user_data, data['id'], data['html'])
    return jsonify({'success': True})

@app.route('/download')
def download():
    user_data = get_user_data()
    if not user_data: return "Session expired", 400
    output = os.path.join(user_data['temp_dir'], f"Resolved_{user_data['original_filename']}")
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(user_data['extract_dir']):
            for file in files:
                p = os.path.join(root, file)
                z.write(p, os.path.relpath(p, user_data['extract_dir']))
    return send_file(output, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
