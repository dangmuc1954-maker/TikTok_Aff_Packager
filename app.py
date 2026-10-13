"""
TikTok Affiliate Studio & Package Builder
Giao diện quản lý và đóng gói video TikTok Shop Affiliate bàn giao khách hàng
"""

import os
import sys
import subprocess
import json
import webbrowser
import threading
import re
from datetime import datetime

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

try:
    import tkinter as tk
    from tkinter import filedialog
    HAS_TKINTER = True
except Exception:
    tk = None
    filedialog = None
    HAS_TKINTER = False

from flask import Flask, render_template, request, jsonify, send_file, send_from_directory, redirect, url_for
from werkzeug.utils import secure_filename
from packager import TikTokAffPackager

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IS_SERVERLESS = os.environ.get("VERCEL") == "1" or os.environ.get("AWS_LAMBDA_FUNCTION_NAME") is not None
if IS_SERVERLESS:
    OUTPUT_DIR = "/tmp/output"
    SAMPLE_VIDEOS_DIR = "/tmp/sample_videos"
    TEMP_UPLOADS_DIR = "/tmp/temp_uploads"
else:
    OUTPUT_DIR = os.path.join(BASE_DIR, "output")
    SAMPLE_VIDEOS_DIR = os.path.join(BASE_DIR, "sample_videos")
    TEMP_UPLOADS_DIR = os.path.join(BASE_DIR, "temp_uploads")
CLOUDFLARED_EXE = os.path.join(BASE_DIR, "cloudflared.exe")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SAMPLE_VIDEOS_DIR, exist_ok=True)
os.makedirs(TEMP_UPLOADS_DIR, exist_ok=True)

packager = TikTokAffPackager(output_base_dir=OUTPUT_DIR)

# Quản lý Cloudflare Tunnel để tạo link online gửi khách
ACTIVE_TUNNEL_URL = None
TUNNEL_PROCESS = None
TUNNEL_LOG = os.path.join(BASE_DIR, "cloudflared.log")

def get_current_tunnel_url():
    global ACTIVE_TUNNEL_URL
    if ACTIVE_TUNNEL_URL:
        return ACTIVE_TUNNEL_URL
    if os.path.exists(TUNNEL_LOG):
        try:
            with open(TUNNEL_LOG, 'r', encoding='utf-8', errors='ignore') as f:
                log_content = f.read()
            match = re.search(r'https://[a-zA-Z0-9\-]+\.trycloudflare\.com', log_content)
            if match:
                ACTIVE_TUNNEL_URL = match.group(0)
                return ACTIVE_TUNNEL_URL
        except:
            pass
    return None

def start_tunnel_thread():
    global ACTIVE_TUNNEL_URL, TUNNEL_PROCESS
    if not os.path.exists(CLOUDFLARED_EXE):
        return
    if TUNNEL_PROCESS is not None and TUNNEL_PROCESS.poll() is None and ACTIVE_TUNNEL_URL:
        return

    # Xóa log cũ
    if os.path.exists(TUNNEL_LOG):
        try: os.remove(TUNNEL_LOG)
        except: pass

    cmd = [CLOUDFLARED_EXE, "tunnel", "--url", "http://127.0.0.1:5050", "--logfile", TUNNEL_LOG]
    try:
        TUNNEL_PROCESS = subprocess.Popen(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        )
        
        # Quét log file trong 12 giây để lấy URL
        for _ in range(24):
            threading.Event().wait(0.5)
            url = get_current_tunnel_url()
            if url:
                print(f"🌐 [CLOUDFLARE PUBLIC LINK READY]: {url}")
                break
    except Exception as e:
        print("Tunnel start error:", e)

@app.route("/api/pick-files-dialog", methods=["POST"])
def pick_files_dialog():
    """Mở hộp thoại chọn nhiều file trực tiếp trên Windows"""
    if not HAS_TKINTER:
        return jsonify({"success": False, "error": "Tính năng chọn file từ ổ đĩa trực tiếp chỉ hỗ trợ khi chạy trên máy tính cá nhân. Trên Cloud, vui lòng kéo thả hoặc upload file."})
    selected_files = []
    
    def open_dialog():
        nonlocal selected_files
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            files = filedialog.askopenfilenames(
                title="Chọn các video TikTok Shop Affiliate (Giữ Ctrl hoặc Shift để chọn nhiều file)",
                filetypes=[("Video files", "*.mp4 *.mov *.mkv *.avi *.webm"), ("All files", "*.*")]
            )
            selected_files = list(files)
            root.destroy()
        except Exception as e:
            print("Dialog error:", e)

    t = threading.Thread(target=open_dialog)
    t.start()
    t.join()

    items = []
    for idx, fpath in enumerate(selected_files, 1):
        fname = os.path.basename(fpath)
        raw_name = os.path.splitext(fname)[0]
        raw_name = re.split(r'[?&#%]', raw_name)[0]
        base_title = re.sub(r'[_\\-]+', ' ', raw_name).strip()[:50] or f"Sản Phẩm {idx}"
        size_mb = round(os.path.getsize(fpath) / (1024 * 1024), 2) if os.path.exists(fpath) else 0
        items.append({
            "stt": idx,
            "filename": fname,
            "filepath": fpath,
            "title": base_title,
            "size_mb": size_mb
        })

    return jsonify({"success": True, "files": items})

@app.route("/api/pick-single-file-dialog", methods=["POST"])
def pick_single_file_dialog():
    """Mở hộp thoại chọn đúng 1 file video (để thêm từng video hoặc đổi video của dòng)"""
    if not HAS_TKINTER:
        return jsonify({"success": False, "error": "Tính năng chọn file từ ổ đĩa trực tiếp chỉ hỗ trợ khi chạy trên máy tính cá nhân. Trên Cloud, vui lòng kéo thả hoặc upload file."})
    selected_file = ""
    
    def open_dialog():
        nonlocal selected_file
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            f = filedialog.askopenfilename(
                title="Chọn 1 video TikTok Shop Affiliate",
                filetypes=[("Video files", "*.mp4 *.mov *.mkv *.avi *.webm"), ("All files", "*.*")]
            )
            selected_file = f
            root.destroy()
        except Exception as e:
            print("Single Dialog error:", e)

    t = threading.Thread(target=open_dialog)
    t.start()
    t.join()

    if selected_file and os.path.exists(selected_file):
        fname = os.path.basename(selected_file)
        raw_name = os.path.splitext(fname)[0]
        raw_name = re.split(r'[?&#%]', raw_name)[0]
        base_title = re.sub(r'[_\\-]+', ' ', raw_name).strip()[:50] or "Sản Phẩm"
        size_mb = round(os.path.getsize(selected_file) / (1024 * 1024), 2)
        return jsonify({
            "success": True,
            "file": {
                "filename": fname,
                "filepath": selected_file,
                "title": base_title,
                "size_mb": size_mb
            }
        })

    return jsonify({"success": False, "error": "Chưa chọn file"})

@app.route("/api/preview-local-video")
def preview_local_video():
    """Phát video preview nhanh cho tác giả xem trước trên bảng quản lý"""
    fpath = request.args.get("path", "")
    if fpath and os.path.exists(fpath) and os.path.isfile(fpath):
        return send_file(fpath)
    return "File không tồn tại", 404

@app.route("/api/upload-files", methods=["POST"])
def upload_files():
    """Nhận các file video được kéo thả hoặc chọn từ trình duyệt"""
    uploaded_files = request.files.getlist("videos")
    saved_items = []
    for idx, f in enumerate(uploaded_files, 1):
        if f.filename:
            safe_name = secure_filename(f.filename) or f"video_{idx}_{int(datetime.now().timestamp())}.mp4"
            if len(safe_name) > 60:
                ext = os.path.splitext(safe_name)[1]
                safe_name = safe_name[:50] + ext
            dest_path = os.path.join(TEMP_UPLOADS_DIR, f"{int(datetime.now().timestamp())}_{safe_name}")
            f.save(dest_path)
            
            raw_name = os.path.splitext(f.filename)[0]
            raw_name = re.split(r'[?&#%]', raw_name)[0]
            base_title = re.sub(r'[_\\-]+', ' ', raw_name).strip()[:50] or f"Sản Phẩm {idx}"
            size_mb = round(os.path.getsize(dest_path) / (1024 * 1024), 2)
            saved_items.append({
                "stt": idx,
                "filename": f.filename,
                "filepath": dest_path,
                "title": base_title,
                "size_mb": size_mb
            })

    return jsonify({"success": True, "files": saved_items})

def create_sample_video(output_path, text, color="0x1e293b", duration=3):
    """Tạo 1 video mp4 dọc 9:16 mẫu bằng ffmpeg"""
    if os.path.exists(output_path) and os.path.getsize(output_path) > 1000:
        return output_path
    cmd = [
        "ffmpeg", "-y",
        "-f", "lavfi", "-i", f"color=c={color}:s=720x1280:d={duration}:r=30",
        "-vf", f"drawtext=text='{text}':fontcolor=white:fontsize=48:x=(w-text_w)/2:y=(h-text_h)/2:box=1:boxcolor=black@0.5:boxborderw=10",
        "-c:v", "libx264", "-pix_fmt", "yuv420p",
        output_path
    ]
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
    except Exception as e:
        # Fallback tạo file rỗng nếu lỗi
        with open(output_path, "wb") as f:
            f.write(b"")
    return output_path

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/create-package", methods=["POST"])
def create_package():
    data = request.json or {}
    client_name = data.get("client_name", "").strip() or "Khách Hàng Mới"
    project_title = data.get("project_title", "").strip() or "Gói Video TikTok Shop Affiliate"
    
    raw_links = data.get("links", "").strip().split("\n")
    links = [l.strip() for l in raw_links if l.strip()]
    
    raw_video_dir = data.get("video_dir", "").strip()
    
    # Tìm các video trong thư mục nếu có
    found_videos = []
    if raw_video_dir and os.path.exists(raw_video_dir):
        for f in sorted(os.listdir(raw_video_dir)):
            if f.lower().endswith(('.mp4', '.mov', '.mkv', '.avi', '.webm')):
                found_videos.append(os.path.join(raw_video_dir, f))
    
    # Nếu không có video thật, kiểm tra items gửi lên từ bảng data editor
    items_input = data.get("items", [])
    
    final_items = []
    total_count = max(len(links), len(found_videos), len(items_input), 1)

    for i in range(total_count):
        stt = i + 1
        
        # Lấy thông tin từ items_input nếu có
        item_in = items_input[i] if i < len(items_input) else {}
        
        # Link sản phẩm
        link = item_in.get("affiliate_link") or (links[i] if i < len(links) else f"https://vt.tiktok.com/ZS{stt}99abc/")
        
        # Video path
        src_video = item_in.get("src_video_path") or (found_videos[i] if i < len(found_videos) else None)
        
        # Tên sản phẩm / Tiêu đề
        product_name = item_in.get("product_name") or f"Sản Phẩm TikTok Hot Trend #{stt:02d}"
        caption = item_in.get("caption") or f"Đỉnh chóp luôn cả nhà ơi! Trải nghiệm ngay {product_name} đang siêu hot trên TikTok Shop nhé. Link ở giỏ hàng góc trái 👇 #tiktokshop #affiliate #review"
        
        final_items.append({
            "stt": stt,
            "title": f"Video #{stt:02d}",
            "product_name": product_name,
            "affiliate_link": link,
            "caption": caption,
            "hashtags": "#tiktokshop #review #affiliate #xuhuong",
            "src_video_path": src_video
        })

    result = packager.build_package(
        project_name=project_title,
        client_name=client_name,
        items=final_items
    )

    folder_name = os.path.basename(result["package_dir"])
    t_url = get_current_tunnel_url()
    public_view_url = f"{t_url}/view/{folder_name}" if t_url else f"http://127.0.0.1:5050/view/{folder_name}"

    return jsonify({
        "success": True,
        "result": result,
        "folder_name": folder_name,
        "folder_url": result["package_dir"],
        "html_url": result["html_path"],
        "zip_url": result["zip_path"],
        "view_url": public_view_url,
        "public_view_url": public_view_url,
        "public_tunnel_url": t_url,
        "download_zip_url": f"/download-zip/{folder_name}",
        "total_items": result["total_items"]
    })

def render_completed_page(meta):
    client_name = meta.get("client_name", "Quý Khách Hàng")
    project_title = meta.get("project_title", "Gói Video TikTok Shop Affiliate")
    total_videos = meta.get("total_videos", 0)
    completed_at = meta.get("completed_at") or datetime.now().strftime("%d/%m/%Y %H:%M")
    
    return f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Đã Hoàn Tất - {project_title}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Plus Jakarta Sans', sans-serif; }}
        body {{
            background: #080c14;
            color: #fff;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
            background-image: radial-gradient(at 50% 50%, rgba(16, 185, 129, 0.15) 0px, transparent 60%);
        }}
        .card {{
            background: #131b2e;
            border: 1px solid rgba(16, 185, 129, 0.4);
            border-radius: 24px;
            padding: 42px 32px;
            max-width: 500px;
            width: 100%;
            text-align: center;
            box-shadow: 0 16px 50px rgba(0, 0, 0, 0.6);
        }}
        .icon-box {{ font-size: 64px; margin-bottom: 16px; }}
        h1 {{ font-size: 24px; font-weight: 800; color: #10b981; margin-bottom: 12px; }}
        p {{ color: #94a3b8; font-size: 14px; line-height: 1.6; margin-bottom: 24px; }}
        .info-box {{
            background: rgba(0, 0, 0, 0.35);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 14px;
            padding: 16px;
            margin-bottom: 24px;
            text-align: left;
            font-size: 13px;
        }}
        .info-row {{ display: flex; justify-content: space-between; margin-bottom: 8px; }}
        .info-row:last-child {{ margin-bottom: 0; }}
        .info-label {{ color: #64748b; }}
        .info-val {{ color: #fff; font-weight: 600; }}
        .badge-done {{
            background: rgba(16, 185, 129, 0.2);
            color: #10b981;
            padding: 2px 8px;
            border-radius: 6px;
            font-weight: 700;
        }}
    </style>
</head>
<body>
    <div class="card">
        <div class="icon-box">🎉</div>
        <h1>GÓI VIDEO ĐÃ HOÀN THÀNH!</h1>
        <p>Toàn bộ video trong gói đã được đánh dấu <strong>ĐÃ ĐĂNG TIKTOK ĐỦ 100%</strong>. Đường link bàn giao này đã tự động đóng lại theo đúng quy trình.</p>
        
        <div class="info-box">
            <div class="info-row">
                <span class="info-label">Dự án:</span>
                <span class="info-val">{project_title}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Khách hàng:</span>
                <span class="info-val">{client_name}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Tiến độ:</span>
                <span class="info-val"><span class="badge-done">Đã đăng {total_videos}/{total_videos} Video (100%)</span></span>
            </div>
            <div class="info-row">
                <span class="info-label">Thời gian đóng link:</span>
                <span class="info-val" style="color: #00f2fe;">{completed_at}</span>
            </div>
        </div>

        <p style="font-size: 13px; color: #64748b; margin-bottom: 0;">
            Cảm ơn quý khách đã tin tưởng và sử dụng dịch vụ! Chúc quý khách bão đơn TikTok Shop!
        </p>
    </div>
</body>
</html>"""

def render_locked_page(meta):
    client_name = meta.get("client_name", "Quý Khách Hàng")
    project_title = meta.get("project_title", "Gói Video TikTok Shop Affiliate")
    return f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Link Tạm Khóa - {project_title}</title>
    <style>
        body {{ background: #080c14; color: #fff; font-family: sans-serif; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; padding: 20px; text-align: center; }}
        .card {{ background: #131b2e; border: 1px solid rgba(255, 0, 80, 0.3); border-radius: 20px; padding: 40px; max-width: 460px; }}
        h1 {{ color: #ff0050; font-size: 22px; margin-bottom: 12px; }}
        p {{ color: #94a3b8; font-size: 14px; line-height: 1.6; }}
    </style>
</head>
<body>
    <div class="card">
        <div style="font-size: 54px; margin-bottom: 14px;">🔒</div>
        <h1>LIÊN KẾT ĐANG TẠM KHÓA</h1>
        <p>Gói bàn giao video dành cho <strong>{client_name}</strong> đang được tạm khóa bởi người tạo.<br>Vui lòng liên hệ người tạo nếu cần mở lại link.</p>
    </div>
</body>
</html>"""

@app.before_request
def restrict_remote_admin_access():
    """Đã mở khóa: Cho phép truy cập trang quản trị trực tiếp trên Vercel và Internet"""
    pass

@app.route("/api/sync-progress", methods=["POST"])
def api_sync_progress():
    """Đồng bộ tiến độ đăng video của khách và tự động đóng link khi tick đủ 100%"""
    data = request.json or {}
    pkg_name = data.get("package_name", "")
    posted_count = int(data.get("posted_count", 0))
    total = int(data.get("total", 0))
    is_completed = bool(data.get("is_completed", False))

    pkg_dir = os.path.join(OUTPUT_DIR, pkg_name)
    if os.path.exists(pkg_dir):
        meta_path = os.path.join(pkg_dir, "metadata.json")
        meta = {}
        if os.path.exists(meta_path):
            try:
                with open(meta_path, "r", encoding="utf-8") as f:
                    meta = json.load(f)
            except: pass

        meta["project_title"] = meta.get("project_title", pkg_name)
        meta["client_name"] = meta.get("client_name", "Quý Khách Hàng")
        meta["posted_count"] = posted_count
        if total > 0:
            meta["total_videos"] = total
        else:
            meta["total_videos"] = meta.get("total_videos", 1)

        if is_completed or (posted_count >= meta["total_videos"] and meta["total_videos"] > 0):
            meta["is_completed"] = True
            if not meta.get("completed_at"):
                meta["completed_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        
        return jsonify({
            "success": True,
            "is_completed": meta.get("is_completed", False),
            "posted_count": meta.get("posted_count", 0)
        })
    return jsonify({"success": False, "error": "Package not found"})

@app.route("/view/<package_name>")
def view_package(package_name):
    """Xem trực tiếp Dashboard HTML - Tự hủy khi khách tick đủ 100%"""
    pkg_dir = os.path.join(OUTPUT_DIR, package_name)
    if not os.path.exists(pkg_dir):
        return "Gói không tồn tại", 404

    # Kiểm tra trạng thái hoàn thành từ metadata.json
    meta_path = os.path.join(pkg_dir, "metadata.json")
    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)
            
            # Nếu khách đã tick đủ 100% -> Tự động đóng link và hiển thị trang hoàn tất
            if meta.get("is_completed", False):
                return render_completed_page(meta)
            
            if not meta.get("is_active", True):
                return render_locked_page(meta), 403
        except Exception as e:
            print("Check meta error:", e)

    html_path = os.path.join(pkg_dir, "Xem_Va_Lay_Link.html")
    if os.path.exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            content = f.read()
        content = content.replace('src="videos/', f'src="/view/{package_name}/videos/')
        content = content.replace('"video_rel_path": "videos/', f'"video_rel_path": "/view/{package_name}/videos/')
        return content
    return "Không tìm thấy Dashboard", 404

@app.route("/api/reopen-link", methods=["POST"])
def api_reopen_link():
    """Mở lại link đã hoàn thành hoặc đã khóa (dành cho tác giả trên localhost)"""
    data = request.json or {}
    pkg_name = data.get("package_name", "")
    pkg_dir = os.path.join(OUTPUT_DIR, pkg_name)
    if os.path.exists(pkg_dir):
        meta_path = os.path.join(pkg_dir, "metadata.json")
        meta = {}
        if os.path.exists(meta_path):
            try:
                with open(meta_path, "r", encoding="utf-8") as f:
                    meta = json.load(f)
            except: pass
        meta["is_completed"] = False
        meta["is_active"] = True
        meta["completed_at"] = None
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        return jsonify({"success": True})
    return jsonify({"success": False, "error": "Không tìm thấy gói"})

@app.route("/api/toggle-lock-link", methods=["POST"])
def api_toggle_lock_link():
    """Khóa hoặc Mở khóa link tức thì"""
    data = request.json or {}
    pkg_name = data.get("package_name", "")
    pkg_dir = os.path.join(OUTPUT_DIR, pkg_name)
    meta_path = os.path.join(pkg_dir, "metadata.json")
    if os.path.exists(meta_path):
        with open(meta_path, "r", encoding="utf-8") as f:
            meta = json.load(f)
        meta["is_active"] = not meta.get("is_active", True)
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        return jsonify({"success": True, "is_active": meta["is_active"]})
    return jsonify({"success": False, "error": "Không tìm thấy gói"})

@app.route("/view/<package_name>/videos/<path:filename>")
def serve_package_video(package_name, filename):
    """Phát video trực tiếp trong Dashboard trên trình duyệt"""
    video_dir = os.path.join(OUTPUT_DIR, package_name, "videos")
    return send_from_directory(video_dir, filename)

@app.route("/download-zip/<package_name>")
def download_zip(package_name):
    """Tải trực tiếp file ZIP về máy từ trình duyệt"""
    clean_name = package_name.replace('.zip', '')
    zip_path = os.path.join(OUTPUT_DIR, f"{clean_name}.zip")
    if os.path.exists(zip_path):
        return send_file(zip_path, as_attachment=True, download_name=f"{clean_name}.zip")
    return "Không tìm thấy file ZIP", 404

@app.route("/api/create-demo-sample", methods=["POST"])
def create_demo_sample():
    """Tạo nhanh 1 gói mẫu 5 video demo kèm video dọc thật và link TikTok Shop"""
    sample_data = [
        {
            "product_name": "Nồi Chiên Không Dầu Điện Tử 6.5L Lock&Lock",
            "link": "https://vt.tiktok.com/ZS2x9AbCD/",
            "color": "0x0f172a",
            "caption": "Nấu ăn healthy cực nhàn với em nồi chiên siêu to khổng lồ này! Giảm ngay 40% trong giỏ hàng nha cả nhà. #noichienkhongdau #giadungthongminh #tiktokshop"
        },
        {
            "product_name": "Tai Nghe Bluetooth Chống Ồn Chủ Động ANC Pro",
            "link": "https://vt.tiktok.com/ZS2x9EfGH/",
            "color": "0x1e1b4b",
            "caption": "Bass căng, cách âm cực êm mà giá chưa tới 300k. Đeo tập gym hay học bài đều hết sảy! #tainghebluetooth #reviewcongnghe #tiktokshop"
        },
        {
            "product_name": "Máy Xay Sinh Tố Cầm Tay Sạc Pin Tiện Lợi",
            "link": "https://vt.tiktok.com/ZS2x9IjKL/",
            "color": "0x064e3b",
            "caption": "Uống sinh tố healthy mọi lúc mọi nơi chỉ trong 30 giây! Tiện lợi mang đi làm, đi chơi. #mayxaysinhto #healthy #lifestyle"
        },
        {
            "product_name": "Son Kem Lì Mịn Môi Kháng Nước Cao Cấp",
            "link": "https://vt.tiktok.com/ZS2x9MnOP/",
            "color": "0x831843",
            "caption": "Màu đỏ đất tôn da cực kỳ, ăn uống thả ga không trôi luôn các nàng ơi! Săn sale ở giỏ hàng góc trái nè. #sonmoi #lamdep #reviewmypham"
        },
        {
            "product_name": "Áo Thun Form Rộng Unisex Cotton 100%",
            "link": "https://vt.tiktok.com/ZS2x9QrST/",
            "color": "0x312e81",
            "caption": "Form rộng unisex mặc siêu tôn dáng, chất vải cotton 100% thoáng mát cực kỳ. #aothununisex #thoitrang #ootd"
        }
    ]

    items = []
    for idx, item in enumerate(sample_data, 1):
        # Tạo video mp4 mẫu dọc 9:16 thật
        vid_file = os.path.join(SAMPLE_VIDEOS_DIR, f"sample_{idx:02d}.mp4")
        create_sample_video(vid_file, f"VIDEO AI #{idx:02d}\\n{item['product_name']}", color=item["color"], duration=4)
        
        items.append({
            "stt": idx,
            "title": f"Video #{idx:02d}",
            "product_name": item["product_name"],
            "affiliate_link": item["link"],
            "caption": item["caption"],
            "hashtags": "#tiktokshop #affiliate #xuhuong #review",
            "src_video_path": vid_file
        })

    result = packager.build_package(
        project_name="Gói 5 Video TikTok Shop Demo Mẫu",
        client_name="Khách Hàng Trải Nghiệm",
        items=items
    )

    folder_name = os.path.basename(result["package_dir"])
    t_url = get_current_tunnel_url()
    public_view_url = f"{t_url}/view/{folder_name}" if t_url else f"http://127.0.0.1:5050/view/{folder_name}"

    return jsonify({
        "success": True,
        "result": result,
        "html_path": result["html_path"],
        "zip_path": result["zip_path"],
        "package_dir": result["package_dir"],
        "view_url": public_view_url,
        "public_view_url": public_view_url,
        "total_items": len(items)
    })

@app.route("/api/start-tunnel", methods=["POST"])
def api_start_tunnel():
    """Bật / Lấy link Cloudflare Public Tunnel"""
    global ACTIVE_TUNNEL_URL
    if not ACTIVE_TUNNEL_URL:
        t = threading.Thread(target=start_tunnel_thread, daemon=True)
        t.start()
        # Chờ tối đa 6 giây để lấy URL
        for _ in range(12):
            if ACTIVE_TUNNEL_URL:
                break
            threading.Event().wait(0.5)

    return jsonify({
        "success": bool(ACTIVE_TUNNEL_URL),
        "public_url": ACTIVE_TUNNEL_URL
    })

@app.route("/api/get-tunnel-status", methods=["GET"])
def api_get_tunnel_status():
    return jsonify({
        "active": bool(ACTIVE_TUNNEL_URL),
        "public_url": ACTIVE_TUNNEL_URL
    })

@app.route("/api/open-path", methods=["POST"])
def open_path():
    path = request.json.get("path", "")
    if path and os.path.exists(path):
        if os.name == 'nt':
            if os.path.isfile(path) and path.endswith('.html'):
                os.startfile(path)
            elif os.path.isdir(path):
                subprocess.Popen(f'explorer "{os.path.normpath(path)}"')
            else:
                os.startfile(path)
        return jsonify({"success": True})
    return jsonify({"success": False, "error": "Path not found"})

@app.route("/api/list-packages", methods=["GET"])
def list_packages():
    t_url = get_current_tunnel_url()
    packages = []
    if os.path.exists(OUTPUT_DIR):
        for name in sorted(os.listdir(OUTPUT_DIR), reverse=True):
            full_path = os.path.join(OUTPUT_DIR, name)
            if os.path.isdir(full_path):
                html_path = os.path.join(full_path, "Xem_Va_Lay_Link.html")
                zip_path = os.path.join(OUTPUT_DIR, f"{name}.zip")
                meta_path = os.path.join(full_path, "metadata.json")
                meta = {}
                if os.path.exists(meta_path):
                    try:
                        with open(meta_path, "r", encoding="utf-8") as f:
                            meta = json.load(f)
                    except: pass

                online_url = f"{t_url}/view/{name}" if t_url else f"http://127.0.0.1:5050/view/{name}"

                packages.append({
                    "name": name,
                    "dir": full_path,
                    "html_path": html_path if os.path.exists(html_path) else None,
                    "zip_path": zip_path if os.path.exists(zip_path) else None,
                    "view_url": online_url,
                    "public_view_url": online_url,
                    "date": datetime.fromtimestamp(os.path.getmtime(full_path)).strftime("%d/%m/%Y %H:%M"),
                    "is_completed": meta.get("is_completed", False),
                    "is_active": meta.get("is_active", True),
                    "posted_count": meta.get("posted_count", 0),
                    "total_videos": meta.get("total_videos", 0)
                })
    return jsonify({
        "packages": packages,
        "public_tunnel_url": t_url
    })

if __name__ == "__main__":
    # Tự động bật Cloudflare Tunnel ngầm khi khởi động
    if os.path.exists(CLOUDFLARED_EXE):
        t = threading.Thread(target=start_tunnel_thread, daemon=True)
        t.start()

    print(f"🚀 TikTok Aff Studio Server running at http://127.0.0.1:5050")
    app.run(host="127.0.0.1", port=5050, debug=False)
