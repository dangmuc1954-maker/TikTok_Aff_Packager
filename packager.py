import os
import re
import shutil
import zipfile
import json
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class TikTokAffPackager:
    def __init__(self, output_base_dir=None):
        if output_base_dir is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            self.output_base_dir = os.path.join(base_dir, "output")
        else:
            self.output_base_dir = output_base_dir
        os.makedirs(self.output_base_dir, exist_ok=True)

    def generate_html_dashboard(self, project_info, items):
        """Sinh ra file HTML Dashboard siêu đẹp, chạy 100% offline cho khách hàng"""
        client_name = project_info.get("client_name", "Quý Khách Hàng")
        project_title = project_info.get("project_title", "Gói Video TikTok Shop Affiliate")
        created_date = project_info.get("created_date", datetime.now().strftime("%d/%m/%Y %H:%M"))
        expiry_text = project_info.get("expiry_text", "7 ngày kể từ khi nhận")
        total_videos = len(items)
        items_json = json.dumps(items, ensure_ascii=False)

        html_content = f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{project_title} - {client_name}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        :root {{
            --bg-primary: #0b0f19;
            --bg-secondary: #131b2e;
            --bg-card: #1a233a;
            --bg-card-hover: #222e4c;
            --accent-pink: #ff0050;
            --accent-cyan: #00f2fe;
            --accent-blue: #4facfe;
            --accent-purple: #7f00ff;
            --accent-green: #10b981;
            --text-primary: #ffffff;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --border-color: rgba(255, 255, 255, 0.08);
            --border-highlight: rgba(0, 242, 254, 0.3);
            --radius-lg: 16px;
            --radius-md: 12px;
            --radius-sm: 8px;
            --shadow-sm: 0 4px 12px rgba(0, 0, 0, 0.3);
            --shadow-glow: 0 8px 24px rgba(255, 0, 80, 0.25);
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        }}

        body {{
            background: linear-gradient(135deg, #090d16 0%, #0f172a 50%, #0b0f19 100%);
            color: var(--text-primary);
            min-height: 100vh;
            padding: 24px;
            overflow-x: hidden;
        }}

        .container {{
            max-width: 1400px;
            margin: 0 auto;
        }}

        /* Header */
        header {{
            background: rgba(19, 27, 46, 0.85);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            padding: 28px 32px;
            margin-bottom: 24px;
            box-shadow: var(--shadow-sm);
            position: relative;
            overflow: hidden;
        }}

        header::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg, var(--accent-pink), var(--accent-purple), var(--accent-cyan));
        }}

        .header-top {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
        }}

        .header-title-box h1 {{
            font-size: 26px;
            font-weight: 800;
            background: linear-gradient(90deg, #ffffff, #cbd5e1);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            display: flex;
            align-items: center;
            gap: 10px;
        }}

        .header-title-box .badge-tiktok {{
            background: linear-gradient(45deg, var(--accent-pink), #e11d48);
            color: #fff;
            font-size: 11px;
            font-weight: 700;
            padding: 4px 10px;
            border-radius: 20px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .header-subtitle {{
            color: var(--text-secondary);
            font-size: 14px;
            margin-top: 6px;
        }}

        .client-info {{
            display: flex;
            gap: 20px;
            align-items: center;
        }}

        .client-tag {{
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-color);
            padding: 8px 16px;
            border-radius: var(--radius-md);
            font-size: 13px;
        }}
        .client-tag strong {{
            color: var(--accent-cyan);
        }}

        /* Stats & Progress bar */
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-top: 24px;
            padding-top: 20px;
            border-top: 1px solid var(--border-color);
        }}

        .stat-card {{
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-md);
            padding: 14px 18px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}

        .stat-label {{
            font-size: 12px;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .stat-value {{
            font-size: 22px;
            font-weight: 700;
            color: var(--text-primary);
        }}

        .progress-box {{
            grid-column: 1 / -1;
            margin-top: 8px;
        }}

        .progress-bar-bg {{
            background: rgba(255, 255, 255, 0.08);
            height: 10px;
            border-radius: 10px;
            overflow: hidden;
            position: relative;
        }}

        .progress-bar-fill {{
            background: linear-gradient(90deg, var(--accent-cyan), var(--accent-green));
            height: 100%;
            width: 0%;
            border-radius: 10px;
            transition: width 0.4s ease;
        }}

        /* Controls Bar */
        .controls-bar {{
            background: rgba(19, 27, 46, 0.7);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            padding: 16px 24px;
            margin-bottom: 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
        }}

        .search-box {{
            flex: 1;
            min-width: 260px;
            position: relative;
        }}

        .search-box input {{
            width: 100%;
            background: rgba(0, 0, 0, 0.3);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-md);
            padding: 10px 16px 10px 38px;
            color: #fff;
            font-size: 14px;
            outline: none;
            transition: all 0.2s;
        }}

        .search-box input:focus {{
            border-color: var(--accent-cyan);
            box-shadow: 0 0 10px rgba(0, 242, 254, 0.2);
        }}

        .search-icon {{
            position: absolute;
            left: 12px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            pointer-events: none;
        }}

        .filters-group {{
            display: flex;
            gap: 10px;
            align-items: center;
        }}

        .filter-btn {{
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-color);
            color: var(--text-secondary);
            padding: 8px 16px;
            border-radius: var(--radius-md);
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }}

        .filter-btn:hover, .filter-btn.active {{
            background: var(--accent-pink);
            color: #fff;
            border-color: var(--accent-pink);
            box-shadow: var(--shadow-glow);
        }}

        .view-toggle {{
            display: flex;
            background: rgba(0, 0, 0, 0.4);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-md);
            padding: 3px;
        }}

        .view-btn {{
            background: transparent;
            border: none;
            color: var(--text-muted);
            padding: 6px 12px;
            border-radius: var(--radius-sm);
            cursor: pointer;
            font-size: 13px;
            font-weight: 600;
            transition: all 0.2s;
        }}

        .view-btn.active {{
            background: rgba(255, 255, 255, 0.15);
            color: #fff;
        }}

        /* Grid View */
        .video-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
            gap: 20px;
        }}

        .video-card {{
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            overflow: hidden;
            display: flex;
            flex-direction: column;
            transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
            position: relative;
        }}

        .video-card:hover {{
            transform: translateY(-4px);
            border-color: var(--border-highlight);
            box-shadow: 0 12px 30px rgba(0, 0, 0, 0.5);
            background: var(--bg-card-hover);
        }}

        .video-card.posted {{
            border-color: rgba(16, 185, 129, 0.4);
            background: rgba(16, 185, 129, 0.05);
        }}

        .video-preview-wrapper {{
            position: relative;
            background: #000;
            aspect-ratio: 9 / 14;
            max-height: 380px;
            overflow: hidden;
            display: flex;
            align-items: center;
            justify-content: center;
        }}

        .video-preview-wrapper video {{
            width: 100%;
            height: 100%;
            object-fit: cover;
            outline: none;
        }}

        .video-stt-badge {{
            position: absolute;
            top: 12px;
            left: 12px;
            background: rgba(0, 0, 0, 0.75);
            backdrop-filter: blur(6px);
            border: 1px solid rgba(255, 255, 255, 0.2);
            color: #fff;
            font-size: 12px;
            font-weight: 800;
            padding: 4px 10px;
            border-radius: 20px;
            z-index: 2;
        }}

        .status-badge-posted {{
            position: absolute;
            top: 12px;
            right: 12px;
            background: var(--accent-green);
            color: #fff;
            font-size: 11px;
            font-weight: 700;
            padding: 4px 10px;
            border-radius: 20px;
            display: none;
            z-index: 2;
        }}

        .video-card.posted .status-badge-posted {{
            display: block;
        }}

        .video-card-body {{
            padding: 18px;
            display: flex;
            flex-direction: column;
            flex: 1;
            gap: 12px;
        }}

        .product-title {{
            font-size: 15px;
            font-weight: 700;
            color: #fff;
            line-height: 1.4;
        }}

        .product-link-box {{
            background: rgba(0, 0, 0, 0.35);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-md);
            padding: 10px 12px;
            display: flex;
            flex-direction: column;
            gap: 6px;
        }}

        .product-link-label {{
            font-size: 11px;
            color: var(--text-muted);
            text-transform: uppercase;
            font-weight: 700;
            letter-spacing: 0.5px;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .product-link-url {{
            font-size: 12px;
            color: var(--accent-cyan);
            word-break: break-all;
            text-decoration: none;
            display: -webkit-box;
            -webkit-line-clamp: 1;
            -webkit-box-orient: vertical;
            overflow: hidden;
        }}

        .caption-box {{
            background: rgba(255, 255, 255, 0.03);
            border-radius: var(--radius-sm);
            padding: 8px 10px;
            font-size: 12px;
            color: var(--text-secondary);
            line-height: 1.4;
            max-height: 60px;
            overflow-y: auto;
        }}

        .card-actions {{
            margin-top: auto;
            display: flex;
            flex-direction: column;
            gap: 8px;
            padding-top: 8px;
        }}

        .btn-action-group {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
        }}

        .btn {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            padding: 9px 14px;
            border-radius: var(--radius-md);
            font-size: 13px;
            font-weight: 700;
            cursor: pointer;
            border: none;
            transition: all 0.2s;
            text-decoration: none;
        }}

        .btn-copy-link {{
            background: linear-gradient(45deg, var(--accent-pink), #f43f5e);
            color: #fff;
            width: 100%;
        }}

        .btn-copy-link:hover {{
            box-shadow: 0 4px 16px rgba(244, 63, 94, 0.4);
            transform: translateY(-1px);
        }}

        .btn-copy-caption {{
            background: rgba(255, 255, 255, 0.08);
            color: #e2e8f0;
            border: 1px solid var(--border-color);
        }}

        .btn-copy-caption:hover {{
            background: rgba(255, 255, 255, 0.15);
            color: #fff;
        }}

        .btn-download {{
            background: rgba(0, 242, 254, 0.12);
            color: var(--accent-cyan);
            border: 1px solid rgba(0, 242, 254, 0.3);
        }}

        .btn-download:hover {{
            background: var(--accent-cyan);
            color: #000;
        }}

        .post-check-label {{
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 13px;
            color: var(--text-secondary);
            cursor: pointer;
            user-select: none;
            padding: 6px 4px;
        }}

        .post-check-label input[type="checkbox"] {{
            width: 16px;
            height: 16px;
            accent-color: var(--accent-green);
            cursor: pointer;
        }}

        /* Table View */
        .table-view-container {{
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: var(--radius-lg);
            overflow: hidden;
            display: none;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 13px;
        }}

        th {{
            background: rgba(0, 0, 0, 0.3);
            color: var(--text-secondary);
            font-weight: 700;
            padding: 14px 16px;
            border-bottom: 1px solid var(--border-color);
            text-transform: uppercase;
            font-size: 11px;
            letter-spacing: 0.5px;
        }}

        td {{
            padding: 14px 16px;
            border-bottom: 1px solid var(--border-color);
            vertical-align: middle;
        }}

        tr:hover td {{
            background: rgba(255, 255, 255, 0.02);
        }}

        .table-video-thumb {{
            width: 60px;
            height: 80px;
            object-fit: cover;
            border-radius: var(--radius-sm);
            background: #000;
        }}

        /* Toast notification */
        .toast {{
            position: fixed;
            bottom: 30px;
            right: 30px;
            background: #10b981;
            color: #fff;
            padding: 12px 24px;
            border-radius: var(--radius-md);
            font-size: 14px;
            font-weight: 700;
            box-shadow: 0 8px 24px rgba(16, 185, 129, 0.3);
            display: flex;
            align-items: center;
            gap: 10px;
            z-index: 9999;
            transform: translateY(100px);
            opacity: 0;
            transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        }}

        .toast.show {{
            transform: translateY(0);
            opacity: 1;
        }}

        /* Responsive */
        @media (max-width: 768px) {{
            body {{ padding: 12px; }}
            .header-top {{ flex-direction: column; align-items: flex-start; }}
            .controls-bar {{ flex-direction: column; align-items: stretch; }}
            .video-grid {{ grid-template-columns: 1fr; }}
        }}
    </style>
</head>
<body>
    <!-- Complete Modal -->
    <div id="complete-modal" style="display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(8, 12, 20, 0.95); backdrop-filter: blur(12px); z-index: 10000; align-items: center; justify-content: center; padding: 20px;">
        <div style="background: #131b2e; border: 2px solid #10b981; border-radius: 24px; padding: 40px; text-align: center; max-width: 480px; width: 100%; box-shadow: 0 0 50px rgba(16, 185, 129, 0.3);">
            <div style="font-size: 64px; margin-bottom: 16px;">🎉</div>
            <h2 style="font-size: 24px; color: #10b981; margin-bottom: 12px; font-weight: 800;">ĐÃ ĐĂNG HOÀN TẤT 100%!</h2>
            <p style="color: #94a3b8; font-size: 14px; line-height: 1.6; margin-bottom: 24px;">
                Chúc mừng bạn đã hoàn thành việc đăng toàn bộ video TikTok Shop Affiliate.<br>
                Đường link bàn giao sẽ tự động đóng lại sau <strong id="close-countdown" style="color: #00f2fe; font-size: 18px;">4</strong> giây.
            </p>
            <div style="background: rgba(16, 185, 129, 0.1); border-radius: 12px; padding: 14px; color: #fff; font-size: 13px; font-weight: 600;">
                ✅ Gói video đã hoàn tất xuất sắc! Chúc bạn nổ nhiều đơn hàng!
            </div>
        </div>
    </div>

    <div class="container">
        <!-- Header -->
        <header>
            <div class="header-top">
                <div class="header-title-box">
                    <h1>
                        🎬 {project_title}
                        <span class="badge-tiktok">TikTok Shop Aff</span>
                    </h1>
                    <div class="header-subtitle">Bộ sưu tập video AI chuẩn tỉ lệ 9:16 kèm link giỏ hàng sẵn sàng đăng kênh</div>
                </div>
                <div class="client-info">
                    <div class="client-tag">Khách hàng: <strong>{client_name}</strong></div>
                    <div class="client-tag">Ngày tạo: <strong>{created_date}</strong></div>
                    <div class="client-tag" style="border-color: rgba(16, 185, 129, 0.4); background: rgba(16, 185, 129, 0.08);">
                        <span>🎯 Tiến độ: </span>
                        <strong id="header-status-badge" style="color: var(--accent-green);">Đang quản lý video</strong>
                    </div>
                </div>
            </div>

            <!-- Stats -->
            <div class="stats-grid">
                <div class="stat-card">
                    <div>
                        <div class="stat-label">Tổng số video</div>
                        <div class="stat-value" id="stat-total">{total_videos}</div>
                    </div>
                    <div style="font-size: 28px;">📦</div>
                </div>
                <div class="stat-card">
                    <div>
                        <div class="stat-label">Đã đăng lên TikTok</div>
                        <div class="stat-value" id="stat-posted" style="color: var(--accent-green);">0</div>
                    </div>
                    <div style="font-size: 28px;">✅</div>
                </div>
                <div class="stat-card">
                    <div>
                        <div class="stat-label">Chưa đăng</div>
                        <div class="stat-value" id="stat-remaining" style="color: var(--accent-pink);">{total_videos}</div>
                    </div>
                    <div style="font-size: 28px;">⏳</div>
                </div>

                <div class="progress-box">
                    <div style="display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 12px; color: var(--text-secondary);">
                        <span>Tiến độ đăng kênh:</span>
                        <span id="progress-percent" style="font-weight: 700; color: var(--accent-cyan);">0%</span>
                    </div>
                    <div class="progress-bar-bg">
                        <div class="progress-bar-fill" id="progress-fill"></div>
                    </div>
                </div>
            </div>
        </header>

        <!-- Controls Bar -->
        <div class="controls-bar">
            <div class="search-box">
                <span class="search-icon">🔍</span>
                <input type="text" id="search-input" placeholder="Tìm theo tên sản phẩm, mã video, caption...">
            </div>

            <div class="filters-group">
                <button class="filter-btn active" data-filter="all">Tất cả ({total_videos})</button>
                <button class="filter-btn" data-filter="unposted">Chưa đăng</button>
                <button class="filter-btn" data-filter="posted">Đã đăng</button>
            </div>

            <div class="view-toggle">
                <button class="view-btn active" id="btn-view-grid" title="Xem dạng lưới">🔲 Thẻ</button>
                <button class="view-btn" id="btn-view-table" title="Xem dạng bảng">📋 Bảng</button>
            </div>
        </div>

        <!-- Video Grid Container -->
        <div class="video-grid" id="video-grid"></div>

        <!-- Video Table View -->
        <div class="table-view-container" id="table-view">
            <table>
                <thead>
                    <tr>
                        <th style="width: 50px;">STT</th>
                        <th style="width: 90px;">Video</th>
                        <th>Tên Sản Phẩm & Caption</th>
                        <th style="width: 280px;">Link Giỏ Hàng TikTok</th>
                        <th style="width: 140px;">Thao tác</th>
                        <th style="width: 120px;">Trạng thái</th>
                    </tr>
                </thead>
                <tbody id="table-body"></tbody>
            </table>
        </div>
    </div>

    <!-- Toast Notification -->
    <div class="toast" id="toast">
        <span>✅</span>
        <span id="toast-msg">Đã copy link giỏ hàng vào bộ nhớ tạm!</span>
    </div>

    <script>
        const items = {items_json};
        const STORAGE_KEY = 'tiktok_aff_posted_' + encodeURIComponent('{client_name}_{project_title}');
        let postedSet = new Set(JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'));

        let currentFilter = 'all';
        let currentSearch = '';

        function saveStorage() {{
            localStorage.setItem(STORAGE_KEY, JSON.stringify(Array.from(postedSet)));
            updateStats();
        }}

        function togglePosted(id) {{
            if (postedSet.has(id)) {{
                postedSet.delete(id);
            }} else {{
                postedSet.add(id);
            }}
            saveStorage();
            render();
        }}

        function showToast(msg) {{
            const toast = document.getElementById('toast');
            document.getElementById('toast-msg').innerText = msg;
            toast.classList.add('show');
            setTimeout(() => toast.classList.remove('show'), 2200);
        }}

        function copyText(text, successMsg) {{
            navigator.clipboard.writeText(text).then(() => {{
                showToast(successMsg || 'Đã copy vào bộ nhớ tạm!');
            }}).catch(() => {{
                // Fallback
                const ta = document.createElement('textarea');
                ta.value = text;
                document.body.appendChild(ta);
                ta.select();
                document.execCommand('copy');
                document.body.removeChild(ta);
                showToast(successMsg || 'Đã copy vào bộ nhớ tạm!');
            }});
        }}

        function updateStats() {{
            const total = items.length;
            const posted = postedSet.size;
            const remaining = total - posted;
            const pct = total > 0 ? Math.round((posted / total) * 100) : 0;

            document.getElementById('stat-total').innerText = total;
            document.getElementById('stat-posted').innerText = posted;
            document.getElementById('stat-remaining').innerText = remaining;
            document.getElementById('progress-percent').innerText = pct + '%';
            document.getElementById('progress-fill').style.width = pct + '%';

            const statusBadge = document.getElementById('header-status-badge');
            if (statusBadge) {{
                if (posted === total && total > 0) {{
                    statusBadge.innerText = '🎉 ĐÃ ĐĂNG ĐỦ 100% (' + posted + '/' + total + ')';
                    statusBadge.style.color = '#10b981';
                }} else {{
                    statusBadge.innerText = 'Đã đăng ' + posted + '/' + total + ' video (' + pct + '%)';
                    statusBadge.style.color = '#00f2fe';
                }}
            }}

            // Đồng bộ tiến độ về máy chủ nếu đang mở Online
            if (window.location.pathname.includes('/view/')) {{
                const parts = window.location.pathname.split('/view/');
                if (parts.length > 1) {{
                    const pkgName = parts[1].split('/')[0];
                    if (pkgName) {{
                        const isDone = (posted === total && total > 0);
                        fetch('/api/sync-progress', {{
                            method: 'POST',
                            headers: {{ 'Content-Type': 'application/json' }},
                            body: JSON.stringify({{
                                package_name: pkgName,
                                posted_count: posted,
                                total: total,
                                is_completed: isDone
                            }})
                        }}).then(res => res.json()).then(data => {{
                            if (isDone) {{
                                const modal = document.getElementById('complete-modal');
                                if (modal && modal.style.display !== 'flex') {{
                                    modal.style.display = 'flex';
                                    let cd = 4;
                                    const timer = setInterval(() => {{
                                        cd--;
                                        const el = document.getElementById('close-countdown');
                                        if (el) el.innerText = cd;
                                        if (cd <= 0) {{
                                            clearInterval(timer);
                                            window.location.reload();
                                        }}
                                    }}, 1000);
                                }}
                            }}
                        }}).catch(() => {{}});
                    }}
                }}
            }}
        }}

        function render() {{
            const gridContainer = document.getElementById('video-grid');
            const tableBody = document.getElementById('table-body');
            gridContainer.innerHTML = '';
            tableBody.innerHTML = '';

            const filtered = items.filter(item => {{
                const isPosted = postedSet.has(item.id);
                if (currentFilter === 'posted' && !isPosted) return false;
                if (currentFilter === 'unposted' && isPosted) return false;

                if (currentSearch) {{
                    const q = currentSearch.toLowerCase();
                    const matchTitle = (item.title || '').toLowerCase().includes(q);
                    const matchProduct = (item.product_name || '').toLowerCase().includes(q);
                    const matchCaption = (item.caption || '').toLowerCase().includes(q);
                    const matchStt = String(item.stt).includes(q);
                    if (!matchTitle && !matchProduct && !matchCaption && !matchStt) return false;
                }}
                return true;
            }});

            filtered.forEach(item => {{
                const isPosted = postedSet.has(item.id);

                // Grid Card
                const card = document.createElement('div');
                card.className = `video-card ${{isPosted ? 'posted' : ''}}`;
                card.innerHTML = `
                    <div class="video-preview-wrapper">
                        <span class="video-stt-badge">#${{String(item.stt).padStart(2, '0')}}</span>
                        <span class="status-badge-posted">✅ Đã đăng</span>
                        <video controls preload="metadata" playsinline>
                            <source src="${{item.video_rel_path}}" type="video/mp4">
                            Trình duyệt không hỗ trợ xem video trực tiếp.
                        </video>
                    </div>
                    <div class="video-card-body">
                        <div class="product-title">${{item.product_name || item.title}}</div>
                        
                        <div class="product-link-box">
                            <div class="product-link-label">
                                <span>🛒 Link Giỏ Hàng TikTok:</span>
                            </div>
                            <a href="${{item.affiliate_link}}" target="_blank" class="product-link-url" title="${{item.affiliate_link}}">
                                ${{item.affiliate_link}}
                            </a>
                        </div>

                        ${{item.caption ? `
                        <div class="caption-box" title="Caption & Hashtag">
                            ${{item.caption}}
                        </div>` : ''}}

                        <div class="card-actions">
                            <button class="btn btn-copy-link" onclick="copyText('${{item.affiliate_link.replace(/'/g, "\\'") }}', '📋 Đã copy link giỏ hàng #${{item.stt}}!')">
                                📋 Copy Link Giỏ Hàng
                            </button>
                            
                            <div class="btn-action-group">
                                ${{item.caption ? `
                                <button class="btn btn-copy-caption" onclick="copyText('${{item.caption.replace(/'/g, "\\'").replace(/\\n/g, ' ') }}', '📝 Đã copy Caption #${{item.stt}}!')">
                                    📝 Copy Caption
                                </button>` : ''}}
                                <a href="${{item.video_rel_path}}" download class="btn btn-download" style="${{!item.caption ? 'grid-column: 1 / -1;' : ''}}">
                                    ⬇️ Tải Video
                                </a>
                            </div>

                            <label class="post-check-label">
                                <input type="checkbox" ${{isPosted ? 'checked' : ''}} onchange="togglePosted('${{item.id}}')">
                                <span>Đánh dấu đã đăng TikTok</span>
                            </label>
                        </div>
                    </div>
                `;
                gridContainer.appendChild(card);

                // Table Row
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="font-weight: 700; color: var(--accent-cyan);">#${{String(item.stt).padStart(2, '0')}}</td>
                    <td>
                        <video class="table-video-thumb" preload="metadata" muted>
                            <source src="${{item.video_rel_path}}" type="video/mp4">
                        </video>
                    </td>
                    <td>
                        <div style="font-weight: 700; color: #fff; margin-bottom: 4px;">${{item.product_name || item.title}}</div>
                        <div style="font-size: 12px; color: var(--text-muted);">${{item.caption || 'Không có caption'}}</div>
                    </td>
                    <td>
                        <div style="display: flex; flex-direction: column; gap: 4px;">
                            <a href="${{item.affiliate_link}}" target="_blank" style="color: var(--accent-cyan); font-size: 12px; word-break: break-all;">
                                ${{item.affiliate_link}}
                            </a>
                            <button class="btn btn-copy-link" style="padding: 4px 10px; font-size: 11px; margin-top: 4px;" onclick="copyText('${{item.affiliate_link.replace(/'/g, "\\'") }}', '📋 Đã copy link giỏ hàng #${{item.stt}}!')">
                                📋 Copy Link
                            </button>
                        </div>
                    </td>
                    <td>
                        <div style="display: flex; flex-direction: column; gap: 6px;">
                            <a href="${{item.video_rel_path}}" download class="btn btn-download" style="padding: 6px 10px; font-size: 12px;">
                                ⬇️ Tải File
                            </a>
                            ${{item.caption ? `
                            <button class="btn btn-copy-caption" style="padding: 4px 8px; font-size: 11px;" onclick="copyText('${{item.caption.replace(/'/g, "\\'").replace(/\\n/g, ' ') }}', '📝 Đã copy Caption #${{item.stt}}!')">
                                📝 Copy Caption
                            </button>` : ''}}
                        </div>
                    </td>
                    <td>
                        <label class="post-check-label" style="font-size: 12px;">
                            <input type="checkbox" ${{isPosted ? 'checked' : ''}} onchange="togglePosted('${{item.id}}')">
                            <span>${{isPosted ? '<span style="color: var(--accent-green); font-weight: 700;">Đã đăng</span>' : 'Chưa đăng'}}</span>
                        </label>
                    </td>
                `;
                tableBody.appendChild(tr);
            }});

            if (filtered.length === 0) {{
                gridContainer.innerHTML = '<div style="grid-column: 1 / -1; text-align: center; padding: 60px; color: var(--text-muted); font-size: 16px;">Không tìm thấy video nào phù hợp với bộ lọc.</div>';
            }}
        }}

        // Event listeners
        document.getElementById('search-input').addEventListener('input', (e) => {{
            currentSearch = e.target.value;
            render();
        }});

        document.querySelectorAll('.filter-btn').forEach(btn => {{
            btn.addEventListener('click', () => {{
                document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                currentFilter = btn.dataset.filter;
                render();
            }});
        }});

        const btnGrid = document.getElementById('btn-view-grid');
        const btnTable = document.getElementById('btn-view-table');
        const gridView = document.getElementById('video-grid');
        const tableView = document.getElementById('table-view');

        btnGrid.addEventListener('click', () => {{
            btnGrid.classList.add('active');
            btnTable.classList.remove('active');
            gridView.style.display = 'grid';
            tableView.style.display = 'none';
        }});

        btnTable.addEventListener('click', () => {{
            btnTable.classList.add('active');
            btnGrid.classList.remove('active');
            gridView.style.display = 'none';
            tableView.style.display = 'block';
        }});

        // Init
        updateStats();
        render();
    </script>
</body>
</html>
"""
        return html_content

    def generate_excel_sheet(self, output_filepath, project_info, items):
        """Sinh ra file Excel định dạng chuẩn, có màu sắc, hyperlink mở video & link sản phẩm"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Danh Sách Video TikTok Shop"

        # Fonts & Styles
        font_title = Font(name="Segoe UI", size=16, bold=True, color="1E293B")
        font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="64748B")
        font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        font_data = Font(name="Segoe UI", size=10, color="0F172A")
        font_link = Font(name="Segoe UI", size=10, color="0284C7", underline="single")
        font_stt = Font(name="Segoe UI", size=11, bold=True, color="0F172A")

        fill_header = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        fill_accent = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")

        border_thin = Border(
            left=Side(style='thin', color='E2E8F0'),
            right=Side(style='thin', color='E2E8F0'),
            top=Side(style='thin', color='E2E8F0'),
            bottom=Side(style='thin', color='E2E8F0')
        )

        # Title Block
        ws.merge_cells("A1:G1")
        ws["A1"] = f"🎬 {project_info.get('project_title', 'BỘ VIDEO TIKTOK SHOP AFFILIATE')}"
        ws["A1"].font = font_title
        ws["A1"].alignment = Alignment(vertical="center")

        ws.merge_cells("A2:G2")
        ws["A2"] = f"Khách hàng: {project_info.get('client_name', 'Quý Khách')} | Ngày tạo: {project_info.get('created_date', datetime.now().strftime('%d/%m/%Y'))} | Tổng: {len(items)} video"
        ws["A2"].font = font_subtitle

        # Headers
        headers = ["STT", "Tên Video / File", "Tên Sản Phẩm", "Link Giỏ Hàng TikTok", "Caption Gợi Ý", "Hashtags", "Trạng Thái Đăng"]
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_thin

        ws.row_dimensions[4].height = 28

        # Data rows
        for row_idx, item in enumerate(items, 5):
            is_even = (row_idx % 2 == 0)
            row_fill = fill_zebra if is_even else None

            # STT
            c1 = ws.cell(row=row_idx, column=1, value=item['stt'])
            c1.font = font_stt
            c1.alignment = Alignment(horizontal="center", vertical="center")
            
            # Video File name + Hyperlink
            video_name = item.get('video_filename', f"Video_{item['stt']:02d}.mp4")
            c2 = ws.cell(row=row_idx, column=2, value=video_name)
            c2.font = font_link
            c2.hyperlink = item.get('video_rel_path', f"videos/{video_name}")
            c2.alignment = Alignment(vertical="center")

            # Product Name
            c3 = ws.cell(row=row_idx, column=3, value=item.get('product_name', ''))
            c3.font = font_data
            c3.alignment = Alignment(vertical="center")

            # Affiliate Link
            aff_link = item.get('affiliate_link', '')
            c4 = ws.cell(row=row_idx, column=4, value=aff_link)
            if aff_link.startswith('http'):
                c4.hyperlink = aff_link
                c4.font = font_link
            else:
                c4.font = font_data
            c4.alignment = Alignment(vertical="center")

            # Caption
            c5 = ws.cell(row=row_idx, column=5, value=item.get('caption', ''))
            c5.font = font_data
            c5.alignment = Alignment(vertical="center", wrap_text=True)

            # Hashtags
            c6 = ws.cell(row=row_idx, column=6, value=item.get('hashtags', '#tiktokshop #affiliate #review'))
            c6.font = font_data
            c6.alignment = Alignment(vertical="center")

            # Status
            c7 = ws.cell(row=row_idx, column=7, value="Chưa đăng")
            c7.font = font_data
            c7.alignment = Alignment(horizontal="center", vertical="center")

            # Apply borders and zebra fill
            for c_idx in range(1, 8):
                cell = ws.cell(row=row_idx, column=c_idx)
                cell.border = border_thin
                if row_fill:
                    cell.fill = row_fill

            ws.row_dimensions[row_idx].height = 24

        # Auto fit columns
        column_widths = {1: 8, 2: 24, 3: 28, 4: 40, 5: 35, 6: 25, 7: 16}
        for col_idx, width in column_widths.items():
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        wb.save(output_filepath)
        return output_filepath

    def build_package(self, project_name, client_name, items, source_video_dir=None):
        """
        Đóng gói toàn bộ dự án gồm:
        - Thư mục videos/
        - Xem_Va_Lay_Link.html
        - Danh_Sach_Video_Affiliate.xlsx
        - Huong_Dan_Su_Dung.txt
        - File .zip hoàn chỉnh
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_client_name = re.sub(r'[\\/*?:"<>|\s]+', '_', str(client_name)).strip('_')[:30] or "KhachHang"
        
        folder_name = f"Goi_Video_TikTokAff_{safe_client_name}_{timestamp}"
        package_dir = os.path.join(self.output_base_dir, folder_name)
        videos_dir = os.path.join(package_dir, "videos")
        
        os.makedirs(videos_dir, exist_ok=True)

        # Chuẩn hóa items và copy video vào thư mục package
        processed_items = []
        for idx, item in enumerate(items, 1):
            stt = item.get("stt", idx)
            raw_title = item.get("title", f"Video {stt}")
            product_name = item.get("product_name", raw_title)
            
            # Làm sạch và giới hạn độ dài tên file (tối đa 40 ký tự) để không bị lỗi Windows path
            safe_prod = re.sub(r'[\\/*?:"<>|]', '', str(product_name))
            safe_prod = re.sub(r'\s+', ' ', safe_prod).strip()
            if len(safe_prod) > 35:
                safe_prod = safe_prod[:35].rstrip()
            if not safe_prod:
                safe_prod = f"Video_{stt:02d}"
            
            target_filename = f"{stt:02d}_{safe_prod}.mp4"
            target_video_path = os.path.join(videos_dir, target_filename)

            # Copy video nếu có nguồn thực tế
            src_path = item.get("src_video_path")
            if src_path and os.path.exists(src_path):
                shutil.copy2(src_path, target_video_path)
            elif not os.path.exists(target_video_path):
                # Tạo 1 file placeholder video nhỏ nếu chưa có
                with open(target_video_path, "wb") as vf:
                    vf.write(b"") # Placeholder

            processed_items.append({
                "id": f"vid_{stt}",
                "stt": stt,
                "title": raw_title,
                "product_name": product_name,
                "video_filename": target_filename,
                "video_rel_path": f"videos/{target_filename}",
                "affiliate_link": item.get("affiliate_link", ""),
                "caption": item.get("caption", f"Trải nghiệm {product_name} cực xịn xò nè cả nhà ơi! Gắn giỏ hàng góc trái nhé. #affiliate #tiktokshop"),
                "hashtags": item.get("hashtags", "#tiktokshop #affiliate #review")
            })

        now_dt = datetime.now()

        project_info = {
            "project_title": project_name,
            "client_name": client_name,
            "created_date": now_dt.strftime("%d/%m/%Y %H:%M"),
            "total_videos": len(processed_items)
        }

        # Lưu metadata.json để quản lý trạng thái tự hủy khi khách tick đủ 100%
        meta_path = os.path.join(package_dir, "metadata.json")
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump({
                "project_title": project_name,
                "client_name": client_name,
                "created_at": now_dt.isoformat(),
                "total_videos": len(processed_items),
                "posted_count": 0,
                "is_completed": False,
                "completed_at": None,
                "is_active": True
            }, f, ensure_ascii=False, indent=2)

        # 1. Tạo file HTML Dashboard
        html_path = os.path.join(package_dir, "Xem_Va_Lay_Link.html")
        html_code = self.generate_html_dashboard(project_info, processed_items)
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_code)

        # 2. Tạo file Excel
        excel_path = os.path.join(package_dir, "Danh_Sach_Video_Affiliate.xlsx")
        self.generate_excel_sheet(excel_path, project_info, processed_items)

        # 3. Tạo file Hướng dẫn sử dụng
        readme_path = os.path.join(package_dir, "Huong_Dan_Su_Dung.txt")
        with open(readme_path, "w", encoding="utf-8") as f:
            f.write(f"""================================================================
🎬 HƯỚNG DẪN SỬ DỤNG GÓI VIDEO TIKTOK SHOP AFFILIATE
================================================================
Khách hàng: {client_name}
Dự án: {project_name}
Ngày bàn giao: {datetime.now().strftime("%d/%m/%Y %H:%M")}
Tổng số video: {len(processed_items)} video

CÁCH SỬ DỤNG NHANH & TIỆN LỢI NHẤT:
1. MỞ DASHBOARD BẰNG TRÌNH DUYỆT (KHUYÊN DÙNG):
   - Click đúp vào file 'Xem_Va_Lay_Link.html'
   - File sẽ mở ngay trên Chrome/Edge/Cốc Cốc mà không cần cài đặt gì.
   - Tại đây bạn có thể:
     + Xem trước từng video
     + Bấm nút 'Copy Link Giỏ Hàng' để dán thẳng vào TikTok Shop
     + Bấm 'Copy Caption'
     + Đánh dấu ô 'Đã đăng' để theo dõi tiến độ

2. DÙNG BẢNG EXCEL:
   - Mở file 'Danh_Sach_Video_Affiliate.xlsx' để xem toàn bộ danh sách, click vào tên video để mở file video trong máy.

3. THƯ MỤC VIDEO GỐC:
   - Toàn bộ file video nằm trong thư mục 'videos/' được đánh số thứ tự từ 01 đến {len(processed_items):02d}.

Chúc bạn đăng video bão đơn và nổ nhiều hoa hồng affiliate!
================================================================
""")

        # 4. Nén toàn bộ thành file ZIP
        zip_path = os.path.join(self.output_base_dir, f"{folder_name}.zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(package_dir):
                for file in files:
                    file_full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_full_path, package_dir)
                    zipf.write(file_full_path, arcname=rel_path)

        return {
            "package_dir": package_dir,
            "zip_path": zip_path,
            "html_path": html_path,
            "excel_path": excel_path,
            "total_items": len(processed_items)
        }
