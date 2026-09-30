"""pytest 入口配置：确保项目根目录在 sys.path 上，
使 `import backend.thesis_builder...` 在任意工作目录下都能成立。"""

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
