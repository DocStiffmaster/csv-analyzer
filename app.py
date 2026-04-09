from flask import Flask, request, render_template_string, send_file
import pandas as pd
import re
from collections import defaultdict
from io import BytesIO

app = Flask(__name__)

HTML = """
<!doctype html>
<html>
<head>
<title>CSV Smart Analyzer</title>
<style>
body { font-family: Arial; background:#f5f7fb; text-align:center; }
.container { margin:50px auto; width:500px; background:white; padding:30px; border-radius:12px; box-shadow:0 10px 25px rgba(0,0,0,0.1); }
button { background:#4F81BD; color:white; border:none; padding:12px 20px; border-radius:8px; cursor:pointer; }
input { margin:20px 0; }
.success { color:green; margin-top:20px; }
</style>
</head>
<body>
<div class="container">
<h2>📊 CSV Analyzer</h2>
<p>Upload deine Verkaufschancen-CSV</p>
<form method=post enctype=multipart/form-data>
  <input type=file name=file required>
  <br>
  <button type=submit>Upload & Verarbeiten</button>
</form>
{% if download %}
  <p class="success">✅ Fertig verarbeitet!</p>
  <a href="/download"><button>⬇ Excel herunterladen</button></a>
{% endif %}
</div>
</body>
</html>
"""

def normalize(text):
    if not text:
        return ""
    text = str(text).lower()
    text = text.replace("ä","ae").replace("ö","oe").replace("ü","ue").replace("ß","ss")
    text = re.sub(r'\b(gmbh|mbh|ag|kg|ug|co|ltd)\b', '', text)
    text = re.sub(r'[^a-z0-9]', '', text)
    return text

from difflib import SequenceMatcher

def is_match(f1, f2):
    n1, n2 = normalize(f1), normalize(f2)
    if not n1 or not n2:
        return False
    if n1 == n2 or n1 in n2 or n2 in n1:
        return True
    return SequenceMatcher(None, n1, n2).ratio() > 0.85

def parse_contacts(text):
    results = []
    if pd.isna(text):
        return results
    for line in str(text).split('\n'):
        if ' - Firma:' in line:
            name = line.split(' - Firma')[0].strip()
            firma = re.search(r'Firma[:\s]+(.+?)(,|$)', line)
            telefon = re.search(r'(Mobile|Phone):\s*([^,]+)', line)
            email = re.search(r'Email:\s*(\S+)', line)
            results.append({
                'name': name,
                'firma': firma.group(1) if firma else '',
                'telefon': telefon.group(2) if telefon else '',
                'email': email.group(1) if email else ''
            })
    return results

def process(df):
    grouped = defaultdict(lambda: {'vc': set(), 'angebote': set(), 'proj': set()})

    for _, row in df.iterrows():
        kontakte = parse_contacts(row.get('Kontakte mit Details', ''))
        angebote_text = row.get('Angebote mit Details', '')

        for k in kontakte:
            key = (k['name'], k['firma'], k['telefon'], k['email'])

            for angebot in str(angebote_text).split('\n'):
                num = re.match(r'^(\\d+)', angebot)
                f = re.search(r'Firma:\s*(.+?)(,|$)', angebot)
                if num and f and is_match(k['firma'], f.group(1)):
                    grouped[key]['angebote'].add(num.group(1))

            grouped[key]['vc'].add(str(row.get('Verkaufschance Nummer', '')))
            grouped[key]['proj'].add(str(row.get('Projekt', '')))

    return grouped

output_file = None

@app.route('/', methods=['GET', 'POST'])
def upload():
    global output_file
    if request.method == 'POST':
        file = request.files['file']
        df = pd.read_csv(file, sep=';', encoding='latin-1')

        data = process(df)

        out = []
        for k, v in data.items():
            out.append(list(k) + [
                ', '.join(v['vc']),
                ', '.join(v['angebote']),
                ', '.join(v['proj'])
            ])

        out_df = pd.DataFrame(out, columns=[
            'Ansprechpartner','Firma','Telefon','E-Mail','Verkaufschancen','Angebote','Projekte'
        ])

        buffer = BytesIO()
        out_df.to_excel(buffer, index=False)
        buffer.seek(0)
        output_file = buffer

        return render_template_string(HTML, download=True)

    return render_template_string(HTML, download=False)

@app.route('/download')
def download():
    return send_file(output_file, as_attachment=True, download_name='Auswertung.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
