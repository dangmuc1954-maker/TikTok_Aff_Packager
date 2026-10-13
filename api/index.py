import os
import sys

# Đảm bảo đường dẫn gốc được nhận diện bởi Vercel Serverless
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

from app import app

# Export cho Vercel Python runtime
app = app
