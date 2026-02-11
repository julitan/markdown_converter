"""
Document to Markdown Converter (Flask Web UI)
- PDF, DOC, DOCX 파일을 Markdown으로 변환
"""
import argparse
import tempfile
import threading
import webbrowser
from pathlib import Path

from flask import Flask, request, jsonify, send_file

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

# 변환 결과 저장 폴더 (--output-dir 인자 또는 기본값)
OUTPUT_DIR = Path(".")  # 아래 __main__에서 설정

# 변환 상태 관리
_state = {
    "running": False,
    "progress": 0,
    "status": "대기 중",
    "logs": [],
    "result_path": None,
}
_lock = threading.Lock()


def _reset_state():
    with _lock:
        _state["running"] = False
        _state["progress"] = 0
        _state["status"] = "대기 중"
        _state["logs"] = []
        _state["result_path"] = None


def _log(message: str):
    with _lock:
        _state["logs"].append(message)


def _set(status: str = None, progress: float = None, running: bool = None):
    with _lock:
        if status is not None:
            _state["status"] = status
        if progress is not None:
            _state["progress"] = progress
        if running is not None:
            _state["running"] = running


# --- HTML (인라인) ---

HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Document to Markdown</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:#f0f2f5;color:#333;min-height:100vh;display:flex;justify-content:center;padding:40px 16px}
.container{background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,.08);width:100%;max-width:680px;padding:32px}
h1{font-size:22px;margin-bottom:24px;text-align:center}
.tabs{display:flex;border-bottom:2px solid #e0e0e0;margin-bottom:20px}
.tab-btn{flex:1;padding:10px;text-align:center;cursor:pointer;border:none;background:none;font-size:15px;font-weight:600;color:#888;border-bottom:2px solid transparent;margin-bottom:-2px;transition:color .2s,border-color .2s}
.tab-btn.active{color:#2563eb;border-bottom-color:#2563eb}
.tab-panel{display:none}.tab-panel.active{display:block}
.drop-zone{border:2px dashed #d0d0d0;border-radius:10px;padding:32px 16px;text-align:center;cursor:pointer;transition:border-color .2s,background .2s;margin-bottom:14px}
.drop-zone:hover,.drop-zone.drag-over{border-color:#2563eb;background:#eff6ff}
.drop-zone-text{font-size:14px;color:#888}.drop-zone-text strong{color:#2563eb}
.drop-zone-file{font-size:14px;color:#333;font-weight:600;margin-top:8px}
.drop-zone input[type="file"]{display:none}
.field{margin-bottom:14px}
.field label{display:block;font-size:13px;font-weight:600;margin-bottom:4px;color:#555}
.field input[type="text"]{width:100%;padding:9px 12px;border:1px solid #d0d0d0;border-radius:6px;font-size:14px;outline:none;transition:border-color .2s}
.field input[type="text"]:focus{border-color:#2563eb}
.checkbox-group{display:flex;gap:20px;margin-top:4px}
.checkbox-group label{font-size:14px;font-weight:normal;display:flex;align-items:center;gap:6px;cursor:pointer}
.hint{font-size:12px;color:#999}
.convert-btn{display:block;width:100%;padding:12px;background:#2563eb;color:#fff;border:none;border-radius:8px;font-size:16px;font-weight:600;cursor:pointer;margin-top:8px;transition:background .2s}
.convert-btn:hover{background:#1d4ed8}
.convert-btn:disabled{background:#93b4f5;cursor:not-allowed}
.download-btn{display:none;width:100%;padding:10px;margin-top:8px;background:#16a34a;color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;transition:background .2s}
.download-btn:hover{background:#15803d}
.progress-section{margin-top:24px}
.progress-bar-bg{width:100%;height:10px;background:#e5e7eb;border-radius:5px;overflow:hidden}
.progress-bar-fill{height:100%;width:0%;background:#2563eb;border-radius:5px;transition:width .3s}
.status-text{font-size:13px;color:#666;margin-top:6px}
.log-section{margin-top:20px}
.log-section h3{font-size:14px;margin-bottom:6px;color:#555}
.log-box{background:#f8f9fa;border:1px solid #e0e0e0;border-radius:8px;padding:12px;height:200px;overflow-y:auto;font-family:'Consolas','D2Coding',monospace;font-size:13px;line-height:1.6;white-space:pre-wrap}
.log-success{color:#16a34a}.log-fail{color:#dc2626}.log-info{color:#555}
</style>
</head>
<body>
<div class="container">
<h1>Document to Markdown</h1>
<div class="tabs">
<button class="tab-btn active" data-tab="single">파일 변환</button>
<button class="tab-btn" data-tab="batch">폴더 일괄 변환</button>
</div>
<div id="tab-single" class="tab-panel active">
<div class="drop-zone" id="drop-zone">
<input type="file" id="file-input" accept=".pdf,.doc,.docx">
<div class="drop-zone-text">PDF / DOC / DOCX 파일을 여기에 <strong>드래그</strong>하거나 <strong>클릭</strong>하여 선택하세요</div>
<div class="drop-zone-file" id="drop-zone-file"></div>
</div>
<div class="checkbox-group">
<label><input type="checkbox" id="single-images" checked> 이미지 추출</label>
</div>
</div>
<div id="tab-batch" class="tab-panel">
<div class="field"><label>입력 폴더 경로</label><input type="text" id="input-dir" placeholder="예: C:\\PDFs"></div>
<div class="field"><label>출력 폴더 <span class="hint">(비워두면 입력과 같은 위치)</span></label><input type="text" id="batch-output" placeholder="예: C:\\Output"></div>
<div class="checkbox-group">
<label><input type="checkbox" id="batch-images" checked> 이미지 추출</label>
<label><input type="checkbox" id="recursive"> 하위 폴더 포함</label>
</div>
</div>
<button class="convert-btn" id="convert-btn">변환 시작</button>
<button class="download-btn" id="download-btn">결과 다운로드 (.md)</button>
<div class="progress-section">
<div class="progress-bar-bg"><div class="progress-bar-fill" id="progress-fill"></div></div>
<div class="status-text" id="status-text">대기 중</div>
</div>
<div class="log-section"><h3>로그</h3><div class="log-box" id="log-box"></div></div>
</div>
<script>
let currentTab='single';
document.querySelectorAll('.tab-btn').forEach(btn=>{
btn.addEventListener('click',()=>{
document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
btn.classList.add('active');
currentTab=btn.dataset.tab;
document.getElementById('tab-'+currentTab).classList.add('active');
});
});
const dropZone=document.getElementById('drop-zone');
const fileInput=document.getElementById('file-input');
const dropZoneFile=document.getElementById('drop-zone-file');
let selectedFile=null;
dropZone.addEventListener('click',()=>fileInput.click());
dropZone.addEventListener('dragover',e=>{e.preventDefault();dropZone.classList.add('drag-over')});
dropZone.addEventListener('dragleave',()=>dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop',e=>{
e.preventDefault();dropZone.classList.remove('drag-over');
const f=e.dataTransfer.files;
const fn=f[0].name.toLowerCase();
if(f.length>0&&(fn.endsWith('.pdf')||fn.endsWith('.doc')||fn.endsWith('.docx'))){selectedFile=f[0];dropZoneFile.textContent=selectedFile.name}
else{dropZoneFile.textContent='PDF, DOC, DOCX 파일만 선택 가능합니다.'}
});
fileInput.addEventListener('change',()=>{if(fileInput.files.length>0){selectedFile=fileInput.files[0];dropZoneFile.textContent=selectedFile.name}});
document.addEventListener('dragover',e=>e.preventDefault());
document.addEventListener('drop',e=>e.preventDefault());
const convertBtn=document.getElementById('convert-btn');
const downloadBtn=document.getElementById('download-btn');
const progressFill=document.getElementById('progress-fill');
const statusText=document.getElementById('status-text');
const logBox=document.getElementById('log-box');
let polling=null,lastLogCount=0;
convertBtn.addEventListener('click',async()=>{
logBox.innerHTML='';lastLogCount=0;progressFill.style.width='0%';downloadBtn.style.display='none';
if(currentTab==='single')await startUpload();else await startBatch();
});
async function startUpload(){
if(!selectedFile){appendLog('[오류] 파일을 선택해주세요 (PDF, DOC, DOCX).','fail');return}
const fd=new FormData();fd.append('file',selectedFile);fd.append('save_images',document.getElementById('single-images').checked);
try{const r=await fetch('/upload',{method:'POST',body:fd});const d=await r.json();
if(d.error){appendLog(d.error,'fail');return}convertBtn.disabled=true;startPolling(true)}
catch(e){appendLog('[오류] 서버 연결 실패','fail')}
}
async function startBatch(){
const body={mode:'batch',input_dir:document.getElementById('input-dir').value,
output_dir:document.getElementById('batch-output').value,
save_images:document.getElementById('batch-images').checked,
recursive:document.getElementById('recursive').checked};
try{const r=await fetch('/convert',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
const d=await r.json();if(d.error){appendLog(d.error,'fail');return}convertBtn.disabled=true;startPolling(false)}
catch(e){appendLog('[오류] 서버 연결 실패','fail')}
}
function startPolling(dl){
polling=setInterval(async()=>{
try{const r=await fetch('/status');const d=await r.json();
progressFill.style.width=d.progress+'%';statusText.textContent=d.status;
for(let i=lastLogCount;i<d.logs.length;i++){const m=d.logs[i];let c='info';
if(m.startsWith('[성공]'))c='success';else if(m.startsWith('[실패]')||m.startsWith('[오류]'))c='fail';
appendLog(m,c)}lastLogCount=d.logs.length;
if(!d.running){clearInterval(polling);convertBtn.disabled=false;if(dl&&d.result_path)downloadBtn.style.display='block'}}
catch(e){clearInterval(polling);convertBtn.disabled=false}
},2000);
}
downloadBtn.addEventListener('click',()=>{window.location.href='/download'});
function appendLog(m,t){const s=document.createElement('span');s.className='log-'+t;s.textContent=m+'\\n';logBox.appendChild(s);logBox.scrollTop=logBox.scrollHeight}
</script>
</body>
</html>"""


# --- 라우트 ---

@app.route("/")
def index():
    return HTML_PAGE


@app.route("/upload", methods=["POST"])
def upload_convert():
    with _lock:
        if _state["running"]:
            return jsonify({"error": "이미 변환 중입니다."}), 409

    if "file" not in request.files:
        return jsonify({"error": "파일이 없습니다."}), 400

    file = request.files["file"]
    allowed_ext = (".pdf", ".doc", ".docx")
    if not file.filename.lower().endswith(allowed_ext):
        return jsonify({"error": "PDF, DOC, DOCX 파일만 지원합니다."}), 400

    save_images = request.form.get("save_images", "true") == "true"

    # 파일을 임시 디렉토리에 저장 (변환 후 삭제)
    tmp_dir = tempfile.mkdtemp()
    file_path = Path(tmp_dir) / file.filename
    file.save(str(file_path))

    # 결과 저장 폴더 (output/ 직접)
    out_dir = OUTPUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)

    _reset_state()
    _set(running=True)
    threading.Thread(
        target=_convert_uploaded, args=(file_path, out_dir, save_images), daemon=True
    ).start()
    return jsonify({"ok": True})


@app.route("/download")
def download_result():
    with _lock:
        md_path = _state.get("result_path")
    if not md_path or not Path(md_path).exists():
        return jsonify({"error": "다운로드할 파일이 없습니다."}), 404
    return send_file(md_path, as_attachment=True)


@app.route("/convert", methods=["POST"])
def start_convert():
    with _lock:
        if _state["running"]:
            return jsonify({"error": "이미 변환 중입니다."}), 409

    data = request.json
    mode = data.get("mode")

    if mode == "batch":
        input_dir = data.get("input_dir", "").strip()
        if not input_dir:
            return jsonify({"error": "입력 폴더 경로를 입력해주세요."}), 400
        if not Path(input_dir).exists():
            return jsonify({"error": f"폴더를 찾을 수 없습니다: {input_dir}"}), 400
    else:
        return jsonify({"error": "잘못된 모드입니다."}), 400

    _reset_state()
    _set(running=True)
    threading.Thread(target=_run_convert, args=(data,), daemon=True).start()
    return jsonify({"ok": True})


@app.route("/status")
def get_status():
    with _lock:
        return jsonify({
            "running": _state["running"],
            "progress": _state["progress"],
            "status": _state["status"],
            "logs": list(_state["logs"]),
            "result_path": _state.get("result_path"),
        })


# --- 변환 워커 ---

def _run_convert(data: dict):
    try:
        _convert_batch(data)
    except Exception as e:
        _log(f"[오류] 예외 발생: {e}")
    finally:
        _set(running=False)


def _convert_uploaded(file_path: Path, output_dir: Path, save_images: bool):
    try:
        _set(status=f"변환 중: {file_path.name}", progress=0)
        _log(f"변환 시작: {file_path.name}")

        ext = file_path.suffix.lower()
        if ext == ".pdf":
            from converter import convert_pdf
            result = convert_pdf(file_path, output_dir, save_images=save_images)
        elif ext in (".doc", ".docx"):
            from converter import convert_docx
            result = convert_docx(file_path, output_dir, save_images=save_images)
        else:
            _log(f"[실패] 지원하지 않는 형식: {ext}")
            _set(status="오류", progress=0)
            return

        if result.success:
            md_path = output_dir / (file_path.stem + ".md")
            with _lock:
                _state["result_path"] = str(md_path)
            msg = f"[성공] {file_path.name} → 변환 완료"
            if result.images and save_images:
                msg += f" (이미지 {len(result.images)}개)"
            _log(msg)
        else:
            _log(f"[실패] {file_path.name}: {result.error}")

        _set(status="완료", progress=100)
    except Exception as e:
        _log(f"[오류] 예외 발생: {e}")
    finally:
        # 임시 파일 정리
        try:
            file_path.unlink(missing_ok=True)
            file_path.parent.rmdir()
        except OSError:
            pass
        _set(running=False)


def _convert_batch(data: dict):
    try:
        from converter import convert_pdf, convert_docx
    except ImportError as e:
        _log(f"[오류] 필요한 라이브러리가 설치되지 않았습니다: {e}")
        _set(status="오류", progress=0)
        return
    input_dir = Path(data["input_dir"].strip())
    output_dir_str = data.get("output_dir", "").strip()
    output_dir = Path(output_dir_str) if output_dir_str else None
    save_images = data.get("save_images", True)
    recursive = data.get("recursive", False)

    # PDF + DOC + DOCX 파일 수집
    all_files = []
    for ext in ("*.pdf", "*.doc", "*.docx"):
        if recursive:
            all_files.extend(input_dir.rglob(ext))
        else:
            all_files.extend(input_dir.glob(ext))

    if not all_files:
        _log("[알림] 변환할 파일이 없습니다 (PDF, DOC, DOCX).")
        _set(status="완료", progress=100)
        return

    total = len(all_files)
    _log(f"총 {total}개 파일 발견 (PDF/DOC/DOCX)")
    success_count = 0
    fail_count = 0

    for i, file_path in enumerate(all_files, 1):
        _set(status=f"변환 중: {file_path.name} ({i}/{total})", progress=((i - 1) / total) * 100)

        if recursive and output_dir:
            relative_path = file_path.parent.relative_to(input_dir)
            current_output = output_dir / relative_path
        else:
            current_output = output_dir

        ext = file_path.suffix.lower()
        if ext == ".pdf":
            result = convert_pdf(file_path, current_output, save_images=save_images)
        else:
            result = convert_docx(file_path, current_output, save_images=save_images)

        if result.success:
            success_count += 1
            msg = f"[성공] {file_path.name}"
            if result.images and save_images:
                msg += f" (이미지 {len(result.images)}개)"
            _log(msg)
        else:
            fail_count += 1
            _log(f"[실패] {file_path.name}: {result.error}")

    _set(status="완료", progress=100)
    _log(f"\n--- 결과: 성공 {success_count}개 / 실패 {fail_count}개 / 총 {total}개 ---")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", default=None, help="변환 결과 저장 폴더")
    args = parser.parse_args()

    if args.output_dir:
        OUTPUT_DIR = Path(args.output_dir)
    else:
        OUTPUT_DIR = Path(__file__).parent / "output"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # 모듈 레벨 변수 갱신
    globals()["OUTPUT_DIR"] = OUTPUT_DIR
    print(f"[설정] 출력 폴더: {OUTPUT_DIR.resolve()}")

    webbrowser.open("http://127.0.0.1:5000")
    app.run(debug=False, port=5000)
