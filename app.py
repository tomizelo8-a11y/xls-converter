from flask import Flask, request, send_file, render_template_string
import pandas as pd
import os
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB max upload

HTML_PAGE = '''<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>XLS Converter — HD Internal Tool</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Inter', sans-serif;
    background: #0f1117;
    color: #e2e8f0;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 24px;
  }
  .card {
    background: #1a1d27;
    border: 1px solid #2d3147;
    border-radius: 16px;
    padding: 40px;
    width: 100%;
    max-width: 480px;
  }
  .badge {
    display: inline-block;
    background: #1e3a5f;
    color: #60a5fa;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    padding: 4px 10px;
    border-radius: 20px;
    margin-bottom: 16px;
  }
  h1 { font-size: 22px; font-weight: 700; color: #f1f5f9; margin-bottom: 8px; line-height: 1.3; }
  p.desc { font-size: 13px; color: #64748b; margin-bottom: 28px; line-height: 1.6; }
  .drop-zone {
    border: 2px dashed #2d3147;
    border-radius: 12px;
    padding: 36px 20px;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s;
    position: relative;
    margin-bottom: 20px;
  }
  .drop-zone:hover, .drop-zone.dragover { border-color: #3b82f6; background: #1e2235; }
  .drop-zone input[type="file"] { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; }
  .drop-icon { font-size: 32px; margin-bottom: 10px; }
  .drop-label { font-size: 14px; color: #94a3b8; }
  .drop-label span { color: #3b82f6; font-weight: 600; }
  .file-name { font-size: 12px; color: #60a5fa; margin-top: 8px; font-weight: 500; }
  button[type="submit"] {
    width: 100%;
    padding: 13px;
    background: #3b82f6;
    color: #fff;
    font-size: 14px;
    font-weight: 600;
    border: none;
    border-radius: 10px;
    cursor: pointer;
    transition: background 0.2s, opacity 0.2s;
  }
  button[type="submit"]:hover { background: #2563eb; }
  button[type="submit"]:disabled { opacity: 0.5; cursor: not-allowed; }
  .status { margin-top: 16px; font-size: 13px; text-align: center; color: #64748b; min-height: 20px; }
  .status.success { color: #34d399; }
  .status.error { color: #f87171; }
  .progress-bar {
    width: 100%;
    height: 4px;
    background: #2d3147;
    border-radius: 4px;
    margin-top: 12px;
    display: none;
    overflow: hidden;
  }
  .progress-fill {
    height: 100%;
    background: #3b82f6;
    border-radius: 4px;
    animation: progress 3s ease-in-out infinite;
  }
  @keyframes progress {
    0% { width: 0%; }
    50% { width: 70%; }
    100% { width: 95%; }
  }
  .info {
    margin-top: 24px;
    padding: 14px 16px;
    background: #12151f;
    border-radius: 10px;
    border-left: 3px solid #3b82f6;
    font-size: 12px;
    color: #64748b;
    line-height: 1.6;
  }
  .info b { color: #94a3b8; }
</style>
</head>
<body>
<div class="card">
  <div class="badge">HD Internal Tool</div>
  <h1>XLS Fulfillment Converter</h1>
  <p class="desc">Upload file .xls yang sudah diedit di WPS/Excel. File akan dikonversi balik ke format HTML-XLS agar bisa diupload ke sistem web.</p>

  <form id="uploadForm" enctype="multipart/form-data">
    <div class="drop-zone" id="dropZone">
      <input type="file" name="file" id="fileInput" accept=".xls">
      <div class="drop-icon">📂</div>
      <div class="drop-label">Drag & drop file .xls atau <span>klik untuk pilih</span></div>
      <div class="file-name" id="fileName"></div>
    </div>
    <button type="submit" id="submitBtn" disabled>Konversi & Download</button>
    <div class="progress-bar" id="progressBar"><div class="progress-fill"></div></div>
  </form>

  <div class="status" id="status"></div>

  <div class="info">
    <b>Cara pakai:</b><br>
    1. Edit file .xls seperti biasa di WPS/Excel<br>
    2. Upload file yang sudah diedit ke sini<br>
    3. Download hasilnya → upload ke web Telkom
  </div>
</div>

<script>
  const fileInput = document.getElementById('fileInput');
  const fileName = document.getElementById('fileName');
  const submitBtn = document.getElementById('submitBtn');
  const dropZone = document.getElementById('dropZone');
  const status = document.getElementById('status');
  const form = document.getElementById('uploadForm');
  const progressBar = document.getElementById('progressBar');

  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) {
      fileName.textContent = fileInput.files[0].name + ' (' + (fileInput.files[0].size/1024/1024).toFixed(1) + ' MB)';
      submitBtn.disabled = false;
      status.textContent = '';
      status.className = 'status';
    }
  });

  dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    fileInput.files = e.dataTransfer.files;
    fileInput.dispatchEvent(new Event('change'));
  });

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!fileInput.files[0]) return;

    submitBtn.disabled = true;
    submitBtn.textContent = 'Memproses...';
    status.textContent = 'Sedang mengkonversi, harap tunggu (file besar ~1-2 menit)...';
    status.className = 'status';
    progressBar.style.display = 'block';

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    try {
      const res = await fetch('/convert', {
        method: 'POST',
        body: formData
      });

      if (!res.ok) {
        const err = await res.text();
        throw new Error(err);
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileInput.files[0].name;
      a.click();
      URL.revokeObjectURL(url);

      progressBar.style.display = 'none';
      status.textContent = '✅ Berhasil! File sudah didownload.';
      status.className = 'status success';
    } catch (err) {
      progressBar.style.display = 'none';
      status.textContent = '❌ Gagal: ' + err.message;
      status.className = 'status error';
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = 'Konversi & Download';
    }
  });
</script>
</body>
</html>'''

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return 'Tidak ada file', 400

    file = request.files['file']
    if file.filename == '':
        return 'Nama file kosong', 400

    try:
        df = pd.read_html(file.stream)[0]
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)

        html_content = '''<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
   <x:ExcelWorksheet>
    <x:Name>Sheet1</x:Name>
    <x:WorksheetOptions>
     <x:DisplayGridlines/>
    </x:WorksheetOptions>
   </x:ExcelWorksheet>
  </x:ExcelWorksheets>
 </x:ExcelWorkbook>
</xml><![endif]-->
</head>
<body>
'''
        html_content += df.to_html(index=False, na_rep='', border=1)
        html_content += '\n</body>\n</html>'

        buf = io.BytesIO(html_content.encode('utf-8'))
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=file.filename,
            mimetype='application/vnd.ms-excel'
        )
    except Exception as e:
        return f'Error: {str(e)}', 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port, threaded=True)
