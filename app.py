from flask import Flask, request, send_file, render_template_string, jsonify
import pandas as pd
import os
import io
import uuid
import threading

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024

# Simpan hasil di memory sementara
results = {}

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
  h1 { font-size: 22px; font-weight: 700; color: #f1f5f9; margin-bottom: 8px; }
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
    width: 100%; padding: 13px;
    background: #3b82f6; color: #fff;
    font-size: 14px; font-weight: 600;
    border: none; border-radius: 10px;
    cursor: pointer; transition: background 0.2s;
  }
  button[type="submit"]:hover { background: #2563eb; }
  button[type="submit"]:disabled { opacity: 0.5; cursor: not-allowed; }
  .status { margin-top: 16px; font-size: 13px; text-align: center; color: #64748b; min-height: 20px; }
  .status.success { color: #34d399; }
  .status.error { color: #f87171; }
  .status.info { color: #60a5fa; }
  .progress-bar { width: 100%; height: 4px; background: #2d3147; border-radius: 4px; margin-top: 12px; display: none; overflow: hidden; }
  .progress-fill { height: 100%; background: #3b82f6; border-radius: 4px; width: 0%; transition: width 0.5s; }
  .info-box {
    margin-top: 24px; padding: 14px 16px;
    background: #12151f; border-radius: 10px;
    border-left: 3px solid #3b82f6;
    font-size: 12px; color: #64748b; line-height: 1.6;
  }
  .info-box b { color: #94a3b8; }
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
    <div class="progress-bar" id="progressBar"><div class="progress-fill" id="progressFill"></div></div>
  </form>

  <div class="status" id="status"></div>

  <div class="info-box">
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
  const progressBar = document.getElementById('progressBar');
  const progressFill = document.getElementById('progressFill');
  let originalFileName = '';

  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) {
      originalFileName = fileInput.files[0].name;
      fileName.textContent = originalFileName + ' (' + (fileInput.files[0].size/1024/1024).toFixed(1) + ' MB)';
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

  document.getElementById('uploadForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!fileInput.files[0]) return;

    submitBtn.disabled = true;
    submitBtn.textContent = 'Mengupload...';
    status.textContent = 'Mengupload file...';
    status.className = 'status info';
    progressBar.style.display = 'block';
    progressFill.style.width = '10%';

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    try {
      // Upload file dan mulai proses di background
      const uploadRes = await fetch('/upload', { method: 'POST', body: formData });
      const uploadData = await uploadRes.json();

      if (!uploadData.job_id) throw new Error('Upload gagal');

      progressFill.style.width = '30%';
      status.textContent = 'File diupload, sedang diproses...';
      submitBtn.textContent = 'Memproses...';

      // Polling status setiap 3 detik
      let attempts = 0;
      const poll = setInterval(async () => {
        attempts++;
        progressFill.style.width = Math.min(30 + attempts * 10, 90) + '%';

        const statusRes = await fetch('/status/' + uploadData.job_id);
        const statusData = await statusRes.json();

        if (statusData.status === 'done') {
          clearInterval(poll);
          progressFill.style.width = '100%';

          // Download file
          const dlRes = await fetch('/download/' + uploadData.job_id);
          const blob = await dlRes.blob();
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = originalFileName;
          a.click();
          URL.revokeObjectURL(url);

          progressBar.style.display = 'none';
          status.textContent = '✅ Berhasil! File sudah didownload.';
          status.className = 'status success';
          submitBtn.disabled = false;
          submitBtn.textContent = 'Konversi & Download';

        } else if (statusData.status === 'error') {
          clearInterval(poll);
          throw new Error(statusData.message);
        } else if (attempts > 40) {
          clearInterval(poll);
          throw new Error('Timeout — file terlalu besar atau server sibuk');
        }
      }, 3000);

    } catch (err) {
      progressBar.style.display = 'none';
      status.textContent = '❌ Gagal: ' + err.message;
      status.className = 'status error';
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

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'Tidak ada file'}), 400

    file = request.files['file']
    file_data = file.read()
    filename = file.filename
    job_id = str(uuid.uuid4())

    results[job_id] = {'status': 'processing', 'filename': filename}

    def process(job_id, file_data, filename):
        try:
            df = pd.read_html(io.BytesIO(file_data))[0]
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

            results[job_id] = {
                'status': 'done',
                'filename': filename,
                'data': html_content.encode('utf-8')
            }
        except Exception as ex:
            results[job_id] = {'status': 'error', 'message': str(ex)}

    t = threading.Thread(target=process, args=(job_id, file_data, filename))
    t.daemon = True
    t.start()

    return jsonify({'job_id': job_id})

@app.route('/status/<job_id>')
def check_status(job_id):
    job = results.get(job_id)
    if not job:
        return jsonify({'status': 'error', 'message': 'Job tidak ditemukan'})
    return jsonify({'status': job['status'], 'message': job.get('message', '')})

@app.route('/download/<job_id>')
def download(job_id):
    job = results.get(job_id)
    if not job or job['status'] != 'done':
        return 'File belum siap', 404

    buf = io.BytesIO(job['data'])
    buf.seek(0)
    filename = job['filename']

    # Hapus dari memory setelah didownload
    del results[job_id]

    return send_file(buf, as_attachment=True, download_name=filename, mimetype='application/vnd.ms-excel')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(debug=False, host='0.0.0.0', port=port, threaded=True)
