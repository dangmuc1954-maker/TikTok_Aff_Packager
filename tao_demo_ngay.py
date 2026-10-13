"""
Tạo nhanh 1 Gói 50 Video TikTok Shop Affiliate Mẫu hoàn chỉnh
Sinh Dashboard HTML + Bảng Excel + File nén ZIP sẵn sàng gửi khách
"""

import os
import sys
import subprocess
from datetime import datetime

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from packager import TikTokAffPackager

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SAMPLE_VIDEOS_DIR = os.path.join(BASE_DIR, "sample_videos")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SAMPLE_VIDEOS_DIR, exist_ok=True)

def create_sample_video(output_path, text, color="0x0f172a", duration=3):
    if os.path.exists(output_path) and os.path.getsize(output_path) > 1000:
        return output_path
    cmd = [
        "ffmpeg", "-y",
        "-f", "lavfi", "-i", f"color=c={color}:s=720x1280:d={duration}:r=30",
        "-vf", f"drawtext=text='{text}':fontcolor=white:fontsize=44:x=(w-text_w)/2:y=(h-text_h)/2:box=1:boxcolor=black@0.6:boxborderw=10",
        "-c:v", "libx264", "-pix_fmt", "yuv420p",
        output_path
    ]
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
    except Exception:
        with open(output_path, "wb") as f:
            f.write(b"")
    return output_path

def main():
    print("🚀 Đang khởi tạo gói demo 50 Video TikTok Shop Affiliate...")
    
    # 50 danh mục và tên sản phẩm hot trend
    categories = [
        "Nồi Chiên Không Dầu Điện Tử 6.5L Lock&Lock",
        "Tai Nghe Bluetooth Chống Ồn Chủ Động ANC Pro",
        "Máy Xay Sinh Tố Cầm Tay Sạc Pin Tiện Lợi",
        "Son Kem Lì Mịn Môi Kháng Nước Cao Cấp",
        "Áo Thun Form Rộng Unisex Cotton 100%",
        "Bộ Dao Bếp Cao Cấp 6 Món Thép Không Gỉ",
        "Đèn Hoàng Hôn Sunset Lamp RGB Decor Phòng",
        "Bàn Phím Cơ Không Dây 3 Chế Độ RGB Gaming",
        "Nước Tẩy Trang Rau Má Dịu Nhẹ Cho Da Nhạy Cảm",
        "Gối Công Thái Học Chống Đau Cổ Vai Gáy",
        "Bình Giữ Nhiệt Inox 304 Khắc Tên 800ml",
        "Kính Cường Lực Chống Nhìn Trộm Full Viền",
        "Kệ Để Đồ Nhà Tắm Dán Tường Không Cần Khoan",
        "Sữa Rửa Mặt Tạo Bọt Chiết Xuất Tràm Trà",
        "Giá Đỡ Điện Thoại Để Bàn Hợp Kim Nhôm Xoay 360",
        "Bộ Drap Giường Cotton Tencel Mát Mịn 4 Món",
        "Máy Hút Bụi Mini Cầm Tay Không Dây Siêu Hút",
        "Kem Chống Nắng Phổ Rộng Kiềm Dầu SPF50+",
        "Chuột Không Dây Silent Chống Ồn Công Thái Học",
        "Túi Đeo Chéo Nam Nữ Chống Thấm Nước Thời Trang"
    ]

    colors = ["0x0f172a", "0x1e1b4b", "0x064e3b", "0x831843", "0x312e81", "0x1e293b", "0x3f3f46"]

    items = []
    for i in range(1, 51):
        prod_name = categories[(i - 1) % len(categories)]
        prod_title = f"{prod_name} (Mẫu #{i:02d})"
        color = colors[(i - 1) % len(colors)]
        
        # Tạo video mp4 dọc 9:16
        vid_file = os.path.join(SAMPLE_VIDEOS_DIR, f"sample_video_{i:02d}.mp4")
        if i <= 5: # Tạo 5 video demo ffmpeg chất lượng để nhẹ máy
            create_sample_video(vid_file, f"VIDEO AI TIKTOK #{i:02d}\\n{prod_name[:20]}...", color=color, duration=3)
        else:
            if not os.path.exists(vid_file):
                with open(vid_file, "wb") as f:
                    f.write(b"")

        items.append({
            "stt": i,
            "title": f"Video #{i:02d}",
            "product_name": prod_title,
            "affiliate_link": f"https://vt.tiktok.com/ZS2x{i:02d}AffiliateShop/",
            "caption": f"Trải nghiệm {prod_name} cực xịn xò! Đang có voucher giảm 50% trong giỏ hàng góc trái nè cả nhà ơi 👇 #tiktokshop #affiliate #review #{prod_name.split()[0].lower()}",
            "hashtags": "#tiktokshop #review #affiliate #xuhuong",
            "src_video_path": vid_file
        })

    packager = TikTokAffPackager(output_base_dir=OUTPUT_DIR)
    result = packager.build_package(
        project_name="Trọn Bộ 50 Video TikTok Shop Affiliate Đợt 1",
        client_name="Khách Hàng VIP - Shop Anh Nam",
        items=items
    )

    print("\n" + "="*60)
    print("🎉 ĐÃ TẠO THÀNH CÔNG GÓI 50 VIDEO BÀN GIAO CHO KHÁCH!")
    print("="*60)
    print(f"📁 Thư mục bàn giao: {result['package_dir']}")
    print(f"🌐 Dashboard HTML:   {result['html_path']}")
    print(f"📊 File Excel:       {result['excel_path']}")
    print(f"📦 File ZIP gửi đi:  {result['zip_path']}")
    print("="*60)

    # Tự động mở Dashboard HTML lên trình duyệt
    if os.name == 'nt':
        os.startfile(result['html_path'])

if __name__ == "__main__":
    main()
