#!/usr/bin/env python3
"""
Hope Energy — Quotation Generator Server
Run: python3 server.py
Open: http://localhost:5000
"""
import os, shutil, zipfile, re, tempfile, subprocess
from flask import Flask, request, send_file, jsonify

app = Flask(__name__, static_folder='.', static_url_path='')

TEMPLATE = os.path.join(os.path.dirname(__file__), 'template_ph.docx')

# ── Helpers ──────────────────────────────────────────────

def fmt_inr(amount):
    """Format as Indian currency: 1,92,838"""
    amount = int(round(amount))
    s = str(amount)
    if len(s) <= 3:
        return s
    result = s[-3:]
    s = s[:-3]
    while len(s) > 2:
        result = s[-2:] + ',' + result
        s = s[:-2]
    if s:
        result = s + ',' + result
    return result


def generate_docx(data, output_path):
    grand_total  = int(data['grand_total'])
    project_cost = round(grand_total / 1.089)
    gst          = grand_total - project_cost
    effective    = grand_total - 108000

    with zipfile.ZipFile(TEMPLATE, 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
        all_files = z.namelist()
        file_contents = {}
        for f in all_files:
            if f != 'word/document.xml':
                file_contents[f] = z.read(f)

    # ── Simple replacements ──
    simple = {
        '{REFERENCE}':   str(data.get('reference', data.get('kilowatt','') + 'KW')),
        '{CLIENT NAME}': data['client_name'].upper(),
        '{LOCATION}':    data['location'].upper(),
        '{KILOWATT}':    str(data['kilowatt']),
        '{INVERTER}':    str(data['inverter_kva']),
        '{WATTPEAK}':    str(data['watt_peak']),
        '{NO OF PANEL}': str(data['no_of_panels']),
        '{STRUCTURE}':   data['structure'],
        '{DAY}':         str(data['day']),
        '{MONTH}':       data['month'],
        '{YEAR,2026}':   str(data['year']),
    }
    for ph, val in simple.items():
        xml = xml.replace(ph, val)

    # ── {GRAND TOTAL} — split across 3 runs ──
    grand_str = fmt_inr(grand_total)
    grand_split = (
        '<w:t>{</w:t></w:r>'
        '<w:proofErr w:type="gramEnd"/>'
        '<w:r w:rsidR="008830EE"><w:rPr><w:b/><w:spacing w:val="-4"/><w:sz w:val="28"/></w:rPr>'
        '<w:t xml:space="preserve">GRAND </w:t></w:r>'
        '<w:proofErr w:type="gramStart"/>'
        '<w:r w:rsidR="008830EE"><w:rPr><w:b/><w:spacing w:val="-4"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>TOTAL}</w:t></w:r>'
    )
    xml = xml.replace(grand_split,
        f'<w:t xml:space="preserve">{grand_str}</w:t></w:r>')

    # ── Project Cost — split runs Rs. + 1 + , + 69 + , + 881 + /- ──
    proj_old = (
        '<w:t>Rs.</w:t></w:r>'
        '<w:r w:rsidR="004D6B48"><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>1</w:t></w:r>'
        '<w:r><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>,</w:t></w:r>'
        '<w:r w:rsidR="004D6B48"><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>69</w:t></w:r>'
        '<w:r><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>,</w:t></w:r>'
        '<w:r w:rsidR="004D6B48"><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>881</w:t></w:r>'
        '<w:r><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>/-</w:t></w:r>'
    )
    xml = xml.replace(proj_old,
        f'<w:t xml:space="preserve">Rs.{fmt_inr(project_cost)}/-</w:t></w:r>')

    # ── GST — split runs Rs. + 15 + , + 119 + /- ──
    gst_old = (
        '<w:t>Rs.</w:t></w:r>'
        '<w:r w:rsidR="004D6B48"><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>15</w:t></w:r>'
        '<w:r><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>,</w:t></w:r>'
        '<w:r w:rsidR="004D6B48"><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>119</w:t></w:r>'
        '<w:r><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>/-</w:t></w:r>'
    )
    xml = xml.replace(gst_old,
        f'<w:t xml:space="preserve">Rs.{fmt_inr(gst)}/-</w:t></w:r>')

    # ── Effective Cost — Rs., + 7 + 7 + ,000/- ──
    eff_old = (
        '<w:t>Rs.,</w:t></w:r>'
        '<w:proofErr w:type="gramEnd"/>'
        '<w:r w:rsidR="004D6B48"><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>7</w:t></w:r>'
        '<w:r w:rsidR="00BD329B"><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>7</w:t></w:r>'
        '<w:r><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="28"/></w:rPr>'
        '<w:t>,000/-</w:t></w:r>'
    )
    xml = xml.replace(eff_old,
        f'<w:t xml:space="preserve">Rs.{fmt_inr(effective)}/-</w:t></w:r>')

    # ── Write DOCX ──
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        zout.writestr('word/document.xml', xml.encode('utf-8'))
        for f, content in file_contents.items():
            zout.writestr(f, content)


def docx_to_pdf(docx_path, out_dir):
    result = subprocess.run([
        'libreoffice', '--headless', '--convert-to', 'pdf',
        '--outdir', out_dir, docx_path
    ], capture_output=True, text=True, timeout=60)
    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf  = os.path.join(out_dir, base + '.pdf')
    if os.path.exists(pdf):
        return pdf
    raise RuntimeError(f"LibreOffice failed: {result.stderr}")


# ── Routes ───────────────────────────────────────────────

@app.route('/')
def index():
    return app.send_static_file('hope_energy_tool.html')


@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json()
        if not data:
            return 'Invalid JSON', 400

        # Basic validation
        required = ['client_name', 'location', 'kilowatt', 'grand_total']
        for r in required:
            if not data.get(r):
                return f'Missing field: {r}', 400

        tmpdir = tempfile.mkdtemp()
        try:
            safe_name = re.sub(r'[^A-Za-z0-9_]', '_',
                               data['client_name'].upper().replace(' ', '_'))
            kw = str(data['kilowatt'])
            docx_name = f"{safe_name}_{kw}KW.docx"
            docx_path = os.path.join(tmpdir, docx_name)

            generate_docx(data, docx_path)
            pdf_path = docx_to_pdf(docx_path, tmpdir)

            return send_file(
                pdf_path,
                mimetype='application/pdf',
                as_attachment=True,
                download_name=os.path.splitext(docx_name)[0] + '.pdf'
            )
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    except Exception as e:
        return str(e), 500


if __name__ == '__main__':
    print("=" * 50)
    print("  Hope Energy — Quotation Generator")
    print("  Server: http://localhost:5000")
    print("  Ctrl+C se band karo")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5000, debug=False)
