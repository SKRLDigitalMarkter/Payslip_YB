import sys
import os

# Ensure project root is in sys.path so app can be imported
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from app import app

# Expose app for Vercel serverless runtime
if __name__ == "__main__":
    app.run()
