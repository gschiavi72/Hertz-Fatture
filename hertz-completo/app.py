import os
import re
import json
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from werkzeug.utils import secure_filename
import pdfplumber
import xml.etree.ElementTree as ET
from xml.dom import minidom
import imaplib
import email
from email.header import decode_header

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.environ.get('UPLOAD_FOLDER', 'uploads')
app.config['OUTPUT_FOLDER'] = os.environ.get('OUTPUT_FOLDER', 'outputs')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)


class HertzProcessor:
    def __init__(self):
        self.data_file = Path("hertz_data.json")
        self.initial_data_file = Path("hertz_data_initial.json")
        self.load_data()
    
    def load_data(self):
        if self.data_file.exists():
            with open(self.data_file, 'r') as f:
                self.data = json.load(f)
        else:
            # Prova a caricare dati iniziali se esistono
            if self.initial_data_file.exists():
                print("📦 Caricamento dati iniziali da hertz_data_initial.json...")
                with open(self.initial_data_file, 'r') as f:
                    self.data = json.load(f)
                print(f"✅ Caricati: {len(self.data.get('preventivi', []))} preventivi, "
                      f"{len(self.data.get('purchase_orders', []))} PO, "
                      f"{len(self.data.get('fatture_generate', []))} fatture")
            else:
                self.data = {
                    'preventivi': [],
                    'purchase_orders': [],
                    'fatture_generate': [],
                    'config': {
                        'last_number_hg': 0,  # Per TYRES (gomme)
                        'last_number_hm': 0,  # Per altri (meccanica)
                        'year': datetime.now().year
                    },
                    'email_config': {
                        'email': 'fornitorischiavigomme@gmail.com',
                        'password': '',
                        'mittente_filtro': '',
                        'oggetto_filtro': 'PO',
                        'ultimo_controllo': None
                    }
                }
            self.save_data()
        
        # Migrazione da vecchia configurazione
        if 'last_number_hg' not in self.data['config']:
            self.data['config']['last_number_hg'] = 0
            self.data['config']['last_number_hm'] = 0
            self.save_data()
        
        if 'email_config' not in self.data:
            self.data['email_config'] = {
                'email': 'fornitorischiavigomme@gmail.com',
                'password': '',
                'mittente_filtro': '',
                'oggetto_filtro': 'PO',
                'ultimo_controllo': None,
                'data_inizio': None,
                'po_scaricati': []  # Lista dei PO già scaricati
            }
            self.save_data()
        
        # Assicura che po_scaricati esista (per upgrade)
        if 'po_scaricati' not in self.data['email_config']:
            self.data['email_config']['po_scaricati'] = []
            self.save_data()
        if 'data_inizio' not in self.data['email_config']:
            self.data['email_config']['data_inizio'] = None
            self.save_data()
    
    def save_data(self):
        with open(self.data_file, 'w') as f:
            json.dump(self.data, f, indent=2, default=str)
    
    def is_po_invoiced(self, po_number):
        return po_number in [f.get('po_number') for f in self.data['fatture_generate']]
    
    def get_stats(self):
        preventivi = self.data['preventivi']
        pos = self.data['purchase_orders']
        
        matches = []
        for prev in preventivi:
            for po in pos:
                if prev.get('pratica_hertz') == po.get('pratica_hertz'):
                    if not self.is_po_invoiced(po.get('po_number')):
                        matches.append({'preventivo': prev, 'po': po})
                    break
        
        matched_pratiche_prev = {m['preventivo']['pratica_hertz'] for m in matches}
        matched_pratiche_po = {m['po']['pratica_hertz'] for m in matches}
        
        prev_in_attesa = [p for p in preventivi if p['pratica_hertz'] not in matched_pratiche_prev]
        po_in_attesa = [p for p in pos if p['pratica_hertz'] not in matched_pratiche_po and not self.is_po_invoiced(p.get('po_number'))]
        
        return {
            'pdf_in_attesa': len(prev_in_attesa) + len(po_in_attesa),
            'lavori_pronti': len(matches),
            'in_attesa_associazione': len(prev_in_attesa) + len(po_in_attesa),
            'ordini_inviati': len(self.data['fatture_generate']),
            'preventivi': len(preventivi),
            'po': len(pos),
            'prev_in_attesa': prev_in_attesa,
            'po_in_attesa': po_in_attesa,
            'matches': matches
        }
    
    def extract_text_from_pdf(self, pdf_path):
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""
            return text
    
    def extract_tables_from_pdf(self, pdf_path):
        with pdfplumber.open(pdf_path) as pdf:
            all_tables = []
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    all_tables.extend(tables)
            return all_tables
    
    def detect_type(self, text):
        if "PREVENTIVO" in text.upper():
            return "preventivo"
        elif "PURCHASE ORDER" in text.upper():
            return "purchase_order"
        return None
    
    def parse_preventivo(self, text, filename, tables=None):
        data = {
            'id': datetime.now().strftime('%Y%m%d%H%M%S%f'),
            'type': 'preventivo',
            'filename': filename,
            'pratica_fornitore': None,
            'pratica_hertz': None,
            'targa': None,
            'telaio': None,
            'km': None,
            'veicolo': None,
            'items': [],
            'totale': 0,
            'data_caricamento': datetime.now().isoformat()
        }
        
        match = re.search(r'Pratica Fornitore:\s*(\d+)', text)
        if match: data['pratica_fornitore'] = match.group(1)
        
        match = re.search(r'Pratica Hertz:\s*(\d+)', text)
        if match: data['pratica_hertz'] = match.group(1)
        
        match = re.search(r'Targa:\s*([A-Z0-9]+)', text)
        if match: data['targa'] = match.group(1).rstrip('T')
        
        match = re.search(r'Telaio:\s*([A-Z0-9]+)', text)
        if match: data['telaio'] = match.group(1)
        
        match = re.search(r'Km:\s*(\d+)', text)
        if match: data['km'] = match.group(1)
        
        match = re.search(r'Veicolo \(Marca - Modello - Versione\):\s*([^\n]+)', text)
        if match: data['veicolo'] = match.group(1).strip()
        
        if tables:
            data['items'] = self._extract_items_from_table(tables)
        
        match = re.search(r'Smaltimento Rifiuti[^\d]*(€?[\d.,]+)', text)
        if match:
            try:
                val = float(match.group(1).replace('€', '').replace(',', '.').strip())
                if val > 0 and not any('Smaltimento' in i['description'] for i in data['items']):
                    data['items'].append({
                        'description': 'Smaltimento Rifiuti',
                        'qty': 1, 'price': val, 'discount': 0, 'total': val
                    })
            except: pass
        
        if tables:
            for table in tables:
                for row in table:
                    if not row: continue
                    row_text = ' '.join(str(c) for c in row if c)
                    
                    for tipo in ['meccanica', 'carrozzeria', 'verniciatura']:
                        if f'Manodopera {tipo}' in row_text:
                            match = re.search(r'ore\s+([\d.,]+)\s*x\s*([\d.,]+)', row_text)
                            if match:
                                try:
                                    ore = float(match.group(1).replace(',', '.'))
                                    tariffa = float(match.group(2).replace(',', '.'))
                                    totale = ore * tariffa
                                    if totale > 0 and not any(f'Manodopera {tipo}' in i['description'] for i in data['items']):
                                        data['items'].append({
                                            'description': f'Manodopera {tipo} ({ore}h x {tariffa}€/h)',
                                            'qty': 1, 'price': totale, 'discount': 0, 'total': totale
                                        })
                                except: pass
        
        data['totale'] = sum(item['total'] for item in data['items'])
        return data
    
    def _extract_items_from_table(self, tables):
        items = []
        for table in tables:
            for row in table:
                if not row or len(row) < 20: continue
                row = [cell if cell else '' for cell in row]
                
                if any(h in str(row) for h in ['C.R.', 'Voci di Danno', 'IMPONIBILE', 'Totale tempi']): continue
                
                try:
                    codice = str(row[0]).strip() if row[0] else None
                    desc = str(row[1]).strip() if row[1] else None
                    if not desc: continue
                    
                    if any(kw in desc for kw in ['Ricambi', 'Materiale', 'Smaltimento', 'Manodopera', 'TOTALI', 'Note:']): continue
                    
                    def parse_num(s):
                        if not s: return 0
                        try: return float(str(s).replace(',', '.').replace('€', '').strip())
                        except: return 0
                    
                    tempo = parse_num(row[18] if len(row) > 18 else '')
                    qty = parse_num(row[19] if len(row) > 19 else '')
                    prezzo = parse_num(row[20] if len(row) > 20 else '')
                    sconto = parse_num(row[23] if len(row) > 23 else '0')
                    totale = parse_num(row[24] if len(row) > 24 else '')
                    
                    if qty == 0 and tempo > 1: qty = tempo
                    if qty == 0: qty = 1
                    
                    if totale > 0:
                        full_desc = f"{desc} - C.R: {codice}" if codice else desc
                        items.append({
                            'description': full_desc,
                            'qty': int(qty) if qty == int(qty) else qty,
                            'price': prezzo,
                            'discount': int(sconto) if sconto else 0,
                            'total': totale,
                            'codice_ricambio': codice
                        })
                except: continue
        return items
    
    def parse_purchase_order(self, text, filename):
        data = {
            'id': datetime.now().strftime('%Y%m%d%H%M%S%f'),
            'type': 'purchase_order',
            'filename': filename,
            'po_number': None,
            'pratica_hertz': None,
            'targa': None,
            'vin': None,
            'unit_number': None,
            'model': None,
            'mileage': None,
            'date': None,
            'total': None,
            'description': None,
            'has_tyres': False,
            'data_caricamento': datetime.now().isoformat()
        }
        
        match = re.search(r'PURCHASE ORDER #.*?(\d+)', text, re.DOTALL)
        if match: data['po_number'] = match.group(1)
        
        match = re.search(r'WD:\s*(\d+)', text)
        if match: data['pratica_hertz'] = match.group(1)
        
        match = re.search(r'Plate Number:\s*([A-Z0-9]+)', text)
        if match: data['targa'] = match.group(1)
        
        match = re.search(r'Serial Number \(VIN\):\s*([A-Z0-9]+)', text)
        if match: data['vin'] = match.group(1)
        
        match = re.search(r'Unit Number:\s*(\d+)', text)
        if match: data['unit_number'] = match.group(1)
        
        match = re.search(r'Model:\s*([^\n]+)', text)
        if match: data['model'] = match.group(1).strip()
        
        match = re.search(r'Mileage:\s*(\d+)', text)
        if match: data['mileage'] = match.group(1)
        
        match = re.search(r'TOTAL\s+€\s*([\d.]+)', text)
        if match: data['total'] = float(match.group(1))
        
        # Estrai la data dal PO (formato: Date: DD/MM/YYYY o simili)
        match = re.search(r'Date:\s*(\d{1,2}[/-]\d{1,2}[/-]\d{4})', text)
        if match: 
            date_str = match.group(1)
            # Prova a parsare la data in vari formati
            for fmt in ['%d/%m/%Y', '%d-%m-%Y', '%m/%d/%Y', '%m-%d-%Y']:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    data['date'] = parsed_date.strftime('%Y-%m-%d')
                    break
                except ValueError:
                    continue
        
        # Se non trova "Date:", cerca altri pattern comuni
        if not data['date']:
            match = re.search(r'(\d{1,2}[/-]\d{1,2}[/-]\d{4})', text)
            if match:
                date_str = match.group(1)
                for fmt in ['%d/%m/%Y', '%d-%m-%Y', '%m/%d/%Y', '%m-%d-%Y']:
                    try:
                        parsed_date = datetime.strptime(date_str, fmt)
                        data['date'] = parsed_date.strftime('%Y-%m-%d')
                        break
                    except ValueError:
                        continue
        
        # Cerca TYRES nel testo del PO
        data['has_tyres'] = 'TYRES' in text.upper()
        data['description'] = text[:500]  # Salva primi 500 caratteri per riferimento
        
        return data
    
    def process_pdf(self, filepath, filename):
        text = self.extract_text_from_pdf(filepath)
        doc_type = self.detect_type(text)
        
        if doc_type == "preventivo":
            tables = self.extract_tables_from_pdf(filepath)
            doc = self.parse_preventivo(text, filename, tables)
            
            # Controlla se già fatturato
            fatturati = [f.get('pratica_hertz') for f in self.data['fatture_generate']]
            if doc['pratica_hertz'] in fatturati:
                doc['gia_fatturato'] = True
                return doc, "preventivo_fatturato"
            
            existing = [p['pratica_hertz'] for p in self.data['preventivi']]
            if doc['pratica_hertz'] not in existing:
                self.data['preventivi'].append(doc)
                self.save_data()
            return doc, "preventivo"
            
        elif doc_type == "purchase_order":
            doc = self.parse_purchase_order(text, filename)
            
            # Controlla se già fatturato
            fatturati = [f.get('po_number') for f in self.data['fatture_generate']]
            if doc['po_number'] in fatturati:
                doc['gia_fatturato'] = True
                return doc, "purchase_order_fatturato"
            
            existing = [p['po_number'] for p in self.data['purchase_orders']]
            if doc['po_number'] not in existing:
                self.data['purchase_orders'].append(doc)
                self.save_data()
            return doc, "purchase_order"
        
        return None, None
    
    def generate_xml(self, match):
        prev = match['preventivo']
        po = match['po']
        
        root = ET.Element('EasyfattDocuments')
        
        company = ET.SubElement(root, 'Company')
        ET.SubElement(company, 'Name').text = 'SCHIAVI GOMME SRL'
        ET.SubElement(company, 'Address').text = 'VIA UTA 20'
        ET.SubElement(company, 'Postcode').text = '00133'
        ET.SubElement(company, 'City').text = 'ROMA'
        ET.SubElement(company, 'Province').text = 'RM'
        ET.SubElement(company, 'FiscalCode').text = '13021431005'
        ET.SubElement(company, 'VatCode').text = '13021431005'
        ET.SubElement(company, 'Tel').text = '0622152148'
        ET.SubElement(company, 'Email').text = 'schiavigomme@gmail.com'
        
        documents = ET.SubElement(root, 'Documents')
        document = ET.SubElement(documents, 'Document')
        
        ET.SubElement(document, 'CustomerCode').text = '999999'
        ET.SubElement(document, 'CustomerName').text = 'HERTZ ITALIANA S.R.L.'
        ET.SubElement(document, 'CustomerAddress').text = 'VIA DEL CASALE CAVALLARI, 204'
        ET.SubElement(document, 'CustomerPostcode').text = '00156'
        ET.SubElement(document, 'CustomerCity').text = 'ROMA'
        ET.SubElement(document, 'CustomerProvince').text = 'RM'
        ET.SubElement(document, 'CustomerCountry').text = 'IT'
        ET.SubElement(document, 'CustomerFiscalCode').text = '00433120581'
        ET.SubElement(document, 'CustomerVatCode').text = 'IT00890931009'
        
        ET.SubElement(document, 'DocumentType').text = 'I'
        ET.SubElement(document, 'Date').text = datetime.now().strftime('%Y-%m-%d')
        
        current_year = datetime.now().year
        if self.data['config']['year'] != current_year:
            self.data['config']['year'] = current_year
            self.data['config']['last_number_hg'] = 0
            self.data['config']['last_number_hm'] = 0
        
        # Determina se è TYRES (HG) o altro (HM)
        is_tyres = po.get('has_tyres', False)
        
        if is_tyres:
            self.data['config']['last_number_hg'] += 1
            invoice_number = self.data['config']['last_number_hg']
            numbering_suffix = 'HG'
        else:
            self.data['config']['last_number_hm'] += 1
            invoice_number = self.data['config']['last_number_hm']
            numbering_suffix = 'HM'
        
        ET.SubElement(document, 'Number').text = str(invoice_number)
        ET.SubElement(document, 'Numbering').text = f"/{numbering_suffix}"
        
        total_without_tax = sum(item['total'] for item in prev['items'])
        vat_amount = total_without_tax * 0.22
        total = total_without_tax + vat_amount
        
        ET.SubElement(document, 'TotalWithoutTax').text = f"{total_without_tax:.2f}"
        ET.SubElement(document, 'VatAmount').text = f"{vat_amount:.2f}"
        ET.SubElement(document, 'Total').text = f"{total:.2f}"
        ET.SubElement(document, 'PricesIncludeVat').text = 'false'
        ET.SubElement(document, 'PaymentName').text = 'Bonifico 60 gg'
        
        # Aggiungo Commento con PO e Targa
        targa = prev.get('targa') or po.get('targa') or ''
        ET.SubElement(document, 'InternalComment').text = f"PO: {po['po_number']} - Targa: {targa}"
        
        rows = ET.SubElement(document, 'Rows')
        
        row_vehicle = ET.SubElement(rows, 'Row')
        vehicle_desc = f"""PO Number: {po['po_number']}
Plate Number: {po['targa']}
Serial Number (VIN): {po['vin']}
Unit Number: {po['unit_number']}
Model: {po['model']}
Country: IT
Type: L
Mileage: {po['mileage']}
Car/Van: V
Pratica Hertz: {prev['pratica_hertz']}"""
        ET.SubElement(row_vehicle, 'Description').text = vehicle_desc
        
        for item in prev['items']:
            row = ET.SubElement(rows, 'Row')
            if item.get('codice_ricambio'):
                ET.SubElement(row, 'Code').text = item['codice_ricambio']
            ET.SubElement(row, 'Description').text = item['description']
            ET.SubElement(row, 'Qty').text = str(item['qty'])
            ET.SubElement(row, 'Price').text = f"{item['price']:.2f}"
            if item['discount'] > 0:
                ET.SubElement(row, 'Discounts').text = f"{item['discount']:.2f}%"
            ET.SubElement(row, 'VatCode', Perc="22.0", Class="Imponibile")
            ET.SubElement(row, 'Total').text = f"{item['total']:.2f}"
        
        rough_string = ET.tostring(root, encoding='utf-8')
        reparsed = minidom.parseString(rough_string)
        xml_string = reparsed.toprettyxml(indent="  ", encoding='utf-8').decode('utf-8')
        
        # Nome file con Data, PO e targa
        targa = prev.get('targa') or po.get('targa') or 'NOTARGA'
        po_date = po.get('date', '')
        
        # Formatta la data per il nome file (YYYYMMDD o vuoto se non presente)
        date_prefix = ''
        if po_date:
            try:
                # po_date è in formato YYYY-MM-DD
                date_prefix = po_date.replace('-', '') + '_'  # Diventa YYYYMMDD_
            except:
                date_prefix = ''
        
        # Formatta il numero fattura con 3 cifre (es. 001, 010, 123)
        invoice_str = str(invoice_number).zfill(3)
        
        filename = f"Fatt_{invoice_str}_PO_{po['po_number']}_{targa}.xml"
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(xml_string)
        
        self.data['fatture_generate'].append({
            'po_number': po['po_number'],
            'targa': targa,
            'pratica_hertz': prev['pratica_hertz'],
            'filename': filename,
            'numero_fattura': invoice_number,
            'tipo': numbering_suffix,
            'totale': total,
            'data_po': po_date,  # Salva la data del PO
            'data_generazione': datetime.now().isoformat()
        })
        
        self.data['preventivi'] = [p for p in self.data['preventivi'] if p['pratica_hertz'] != prev['pratica_hertz']]
        self.data['purchase_orders'] = [p for p in self.data['purchase_orders'] if p['po_number'] != po['po_number']]
        
        self.save_data()
        return filename, total
    
    def delete_document(self, doc_type, doc_id):
        if doc_type == 'preventivo':
            self.data['preventivi'] = [p for p in self.data['preventivi'] if p['id'] != doc_id]
        elif doc_type == 'purchase_order':
            self.data['purchase_orders'] = [p for p in self.data['purchase_orders'] if p['id'] != doc_id]
        self.save_data()
    
    def clear_all(self):
        self.data['preventivi'] = []
        self.data['purchase_orders'] = []
        self.save_data()
    
    def check_email(self, data_da=None):
        """Controlla Gmail per nuovi PO via IMAP"""
        email_config = self.data.get('email_config', {})
        
        if not email_config.get('password'):
            return {'error': 'Password email non configurata. Vai in Configurazione.'}
        
        if not email_config.get('email'):
            return {'error': 'Email non configurata. Vai in Configurazione.'}
        
        # Lista PO già scaricati
        po_scaricati = email_config.get('po_scaricati', [])
        
        results = {
            'checked': 0,
            'downloaded': 0,
            'files': [],
            'errors': [],
            'skipped': [],
            'duplicati': 0
        }
        
        mail_conn = None
        try:
            import socket
            socket.setdefaulttimeout(30)  # Timeout 30 secondi
            
            mail_conn = imaplib.IMAP4_SSL('imap.gmail.com', 993)
            mail_conn.login(email_config['email'], email_config['password'])
            mail_conn.select('inbox')
            
            mittente = email_config.get('mittente_filtro', '').strip()
            oggetto = email_config.get('oggetto_filtro', 'PO')
            
            # Usa Gmail X-GM-RAW per filtro data
            if data_da:
                data_da = data_da.strip()
                if '-' in data_da:
                    data_da = data_da.replace('-', '/')
                gmail_query = f'subject:{oggetto} after:{data_da}'
                print(f"DEBUG Gmail query: {gmail_query}")
                status, messages = mail_conn.search(None, 'X-GM-RAW', f'"{gmail_query}"')
            else:
                status, messages = mail_conn.search(None, f'(SUBJECT "{oggetto}")')
            
            if status != 'OK':
                return {'error': 'Errore nella ricerca email'}
            
            email_ids = messages[0].split()
            results['checked'] = len(email_ids)
            
            for email_id in email_ids:
                try:
                    status, msg_data = mail_conn.fetch(email_id, '(RFC822)')
                    if status != 'OK': continue
                    
                    msg = email.message_from_bytes(msg_data[0][1])
                    from_header = msg.get('From', '')
                    
                    # Filtro mittente opzionale
                    if mittente and mittente.lower() not in from_header.lower():
                        results['skipped'].append(f"Mittente non corrisponde: {from_header}")
                        continue
                    
                    subject = msg.get('Subject', '')
                    if subject:
                        decoded = decode_header(subject)
                        subject = ''.join([
                            part.decode(enc or 'utf-8') if isinstance(part, bytes) else part
                            for part, enc in decoded
                        ])
                    
                    # Cerca allegati PDF
                    found_pdf = False
                    for part in msg.walk():
                        content_disposition = str(part.get('Content-Disposition', ''))
                        
                        if part.get_content_maintype() == 'multipart':
                            continue
                        
                        filename = part.get_filename()
                        
                        if not filename and 'attachment' in content_disposition:
                            filename = f"allegato_{email_id.decode()}.pdf"
                        
                        if filename:
                            decoded_filename = decode_header(filename)
                            filename = ''.join([
                                p.decode(enc or 'utf-8') if isinstance(p, bytes) else p
                                for p, enc in decoded_filename
                            ])
                            
                            if filename.lower().endswith('.pdf'):
                                found_pdf = True
                                safe_filename = secure_filename(filename)
                                filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
                                
                                payload = part.get_payload(decode=True)
                                if payload:
                                    with open(filepath, 'wb') as f:
                                        f.write(payload)
                                    
                                    try:
                                        doc, doc_type = self.process_pdf(filepath, safe_filename)
                                        if doc:
                                            po_number = doc.get('po_number')
                                            
                                            if po_number and po_number in po_scaricati:
                                                results['duplicati'] += 1
                                                results['skipped'].append(f"PO {po_number} già scaricato")
                                                os.remove(filepath)
                                                continue
                                            
                                            if po_number:
                                                po_scaricati.append(po_number)
                                            
                                            results['downloaded'] += 1
                                            results['files'].append({
                                                'filename': safe_filename,
                                                'type': doc_type,
                                                'subject': subject,
                                                'from': from_header,
                                                'po_number': po_number
                                            })
                                    except Exception as e:
                                        results['errors'].append(f"{filename}: {str(e)}")
                                else:
                                    results['errors'].append(f"{filename}: payload vuoto")
                    
                    if not found_pdf:
                        results['skipped'].append(f"Nessun PDF in: {subject}")
                    
                except Exception as e:
                    results['errors'].append(str(e))
            
            try:
                mail_conn.logout()
            except:
                pass
            
            # Salva lista PO scaricati
            self.data['email_config']['po_scaricati'] = po_scaricati
            self.data['email_config']['ultimo_controllo'] = datetime.now().isoformat()
            self.save_data()
            
            return results
            
        except imaplib.IMAP4.error as e:
            error_msg = str(e)
            if 'AUTHENTICATIONFAILED' in error_msg:
                return {'error': 'Autenticazione fallita. Verifica la password per le app Gmail.'}
            return {'error': f'Errore IMAP: {error_msg}'}
        except ConnectionRefusedError:
            return {'error': 'Connessione rifiutata. Il server Gmail non è raggiungibile.'}
        except TimeoutError:
            return {'error': 'Timeout connessione Gmail. Riprova tra qualche secondo.'}
        except Exception as e:
            return {'error': f'Errore: {str(e)}'}
        finally:
            if mail_conn:
                try:
                    mail_conn.logout()
                except:
                    pass


processor = HertzProcessor()


# ==================== ROUTES ====================

@app.route('/')
def dashboard():
    stats = processor.get_stats()
    return render_template('dashboard.html', stats=stats, page='dashboard')

@app.route('/lavori')
def lavori():
    stats = processor.get_stats()
    return render_template('lavori.html', stats=stats, page='lavori')

@app.route('/documenti')
def documenti():
    stats = processor.get_stats()
    docs = processor.data['preventivi'] + processor.data['purchase_orders']
    docs.sort(key=lambda x: x.get('data_caricamento', ''), reverse=True)
    return render_template('documenti.html', stats=stats, documents=docs, page='documenti')

@app.route('/configurazione')
def configurazione():
    stats = processor.get_stats()
    return render_template('configurazione.html', stats=stats, config=processor.data['config'], 
                          email_config=processor.data.get('email_config', {}),
                          fatture=processor.data['fatture_generate'], page='configurazione')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'Nessun file'}), 400
    
    files = request.files.getlist('file')
    results = []
    
    for file in files:
        if file.filename == '': continue
        
        if file and file.filename.lower().endswith('.pdf'):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            try:
                doc, doc_type = processor.process_pdf(filepath, filename)
                if doc:
                    result = {
                        'filename': filename,
                        'type': doc_type,
                        'targa': doc.get('targa'),
                        'pratica': doc.get('pratica_hertz'),
                        'po_number': doc.get('po_number')
                    }
                    # Segnala se già fatturato
                    if doc_type in ['preventivo_fatturato', 'purchase_order_fatturato']:
                        result['gia_fatturato'] = True
                        result['warning'] = f"Documento già fatturato!"
                    results.append(result)
                else:
                    results.append({'filename': filename, 'error': 'Tipo non riconosciuto'})
            except Exception as e:
                results.append({'filename': filename, 'error': str(e)})
    
    return jsonify({'results': results, 'stats': processor.get_stats()})

@app.route('/check-email', methods=['POST'])
def check_email():
    try:
        data = request.json or {}
        data_da = data.get('data_da')
        results = processor.check_email(data_da)
        return jsonify(results)
    except Exception as e:
        return jsonify({'error': f'Errore server: {str(e)}'}), 200

@app.route('/save-email-config', methods=['POST'])
def save_email_config():
    data = request.json
    processor.data['email_config']['email'] = data.get('email', '')
    processor.data['email_config']['password'] = data.get('password', '')
    processor.data['email_config']['mittente_filtro'] = data.get('mittente_filtro', '')
    processor.data['email_config']['oggetto_filtro'] = data.get('oggetto_filtro', 'PO')
    processor.save_data()
    return jsonify({'success': True})

@app.route('/reset-po-scaricati', methods=['POST'])
def reset_po_scaricati():
    processor.data['email_config']['po_scaricati'] = []
    processor.save_data()
    return jsonify({'success': True})

@app.route('/update-numerazione', methods=['POST'])
def update_numerazione():
    data = request.json
    try:
        new_hm = int(data.get('last_number_hm', 0))
        new_hg = int(data.get('last_number_hg', 0))
        
        if new_hm < 0 or new_hg < 0:
            return jsonify({'error': 'I numeri devono essere positivi'}), 400
        
        processor.data['config']['last_number_hm'] = new_hm
        processor.data['config']['last_number_hg'] = new_hg
        processor.save_data()
        
        return jsonify({'success': True})
    except ValueError:
        return jsonify({'error': 'Valori non validi'}), 400

@app.route('/genera/<pratica_hertz>', methods=['POST'])
def genera_fattura(pratica_hertz):
    stats = processor.get_stats()
    
    for match in stats['matches']:
        if match['preventivo']['pratica_hertz'] == pratica_hertz:
            try:
                filename, totale = processor.generate_xml(match)
                return jsonify({'success': True, 'filename': filename, 'totale': totale})
            except Exception as e:
                return jsonify({'error': str(e)}), 500
    
    return jsonify({'error': 'Match non trovato'}), 404

@app.route('/genera-tutti', methods=['POST'])
def genera_tutti():
    stats = processor.get_stats()
    generated = []
    
    # Ordina i matches per data PO (più vecchi prima)
    sorted_matches = sorted(stats['matches'], 
                           key=lambda m: m['po'].get('date') or '9999-99-99')
    
    for match in sorted_matches:
        try:
            filename, totale = processor.generate_xml(match)
            generated.append({'filename': filename, 'totale': totale})
        except: pass
    
    return jsonify({'success': True, 'generated': generated})

@app.route('/download/<filename>')
def download_file(filename):
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return "File non trovato", 404

@app.route('/delete/<doc_type>/<doc_id>', methods=['POST'])
def delete_document(doc_type, doc_id):
    processor.delete_document(doc_type, doc_id)
    return jsonify({'success': True})

@app.route('/delete-fattura/<filename>', methods=['POST'])
def delete_fattura(filename):
    try:
        # Rimuovi la fattura dalla lista
        processor.data['fatture_generate'] = [
            f for f in processor.data['fatture_generate'] 
            if f['filename'] != filename
        ]
        
        # Elimina il file XML se esiste
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.exists(filepath):
            os.remove(filepath)
        
        processor.save_data()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/delete-all-fatture', methods=['POST'])
def delete_all_fatture():
    try:
        deleted_count = len(processor.data['fatture_generate'])
        
        # Elimina tutti i file XML
        for fattura in processor.data['fatture_generate']:
            filepath = os.path.join(app.config['OUTPUT_FOLDER'], fattura['filename'])
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except:
                    pass  # Continua anche se un file non può essere eliminato
        
        # Svuota la lista fatture
        processor.data['fatture_generate'] = []
        processor.save_data()
        
        return jsonify({'success': True, 'deleted': deleted_count})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/clear-all', methods=['POST'])
def clear_all():
    processor.clear_all()
    return jsonify({'success': True})

@app.route('/api/stats')
def api_stats():
    stats = processor.get_stats()
    stats['po_scaricati'] = len(processor.data.get('email_config', {}).get('po_scaricati', []))
    return jsonify(stats)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)