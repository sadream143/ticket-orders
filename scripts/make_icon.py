"""Creates build/app.ico from the logo drawn in app.py."""
import os
import sys

sys.path.insert(0, "src")
from app import _draw_icon  # noqa: E402

os.makedirs("build", exist_ok=True)
img = _draw_icon("logo", "#4F46E5", 256)
img.save("build/app.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("build/app.ico created")
