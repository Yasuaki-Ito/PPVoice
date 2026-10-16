"""PPVoice の GUI の起動 (画面は ui/ パッケージ)"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.app import App  # noqa: E402,F401

if __name__ == "__main__":
    app = App()
    app.mainloop()
