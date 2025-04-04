import os
import time
import pandas as pd

def get_base_path():
    import sys
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)
    return path

def load_excel(path, sheet=0):
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception as e:
        print(f"error loading excel file: {e}")
        return None

def save_excel(df, path, sheet_name="Sheet1"):
    try:
        df.to_excel(path, sheet_name=sheet_name, index=False)
        return True
    except Exception as e:
        print(f"error saving excel file: {e}")
        return False

def find_latest_file(folder, ext=None):
    files = []
    for f in os.listdir(folder):
        if ext and not f.endswith(ext):
            continue
        path = os.path.join(folder, f)
        if os.path.isfile(path):
            files.append((path, os.path.getmtime(path)))
    
    if not files:
        return None
    
    return sorted(files, key=lambda x: x[1], reverse=True)[0][0]

def wait_for_download(folder, timeout=30):
    start = time.time()
    while time.time() - start < timeout:
        files = [f for f in os.listdir(folder) if f.endswith('.part') or f.endswith('.crdownload')]
        if not files:
            return True
        time.sleep(0.5)
    return False