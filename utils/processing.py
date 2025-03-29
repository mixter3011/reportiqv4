import pandas as pd
from pathlib import Path
from tkinter import messagebox
from typing import List, Tuple

def ld_data(files: List[Tuple[str, int, str]], req: List[str], out: Path) -> dict:
    data = {}
    for fname in req:
        fpath = None
        
        for n, _, p in files:
            if n == fname:
                fpath = p
                break
        
        if not fpath:
            conv_p = out / fname
            if conv_p.exists():
                fpath = conv_p
        
        if fpath:
            try:
                data[fname.replace('.csv', '')] = pd.read_csv(fpath)
            except Exception as e:
                messagebox.showerror("Error", f"Could not load {fname}: {str(e)}")
                return None
        else:
            messagebox.showerror("Error", f"Missing required file: {fname}")
            return None
    
    return data

def conv(fp: str, out: Path) -> List[str]:
    try:
        xl = pd.ExcelFile(fp, engine="openpyxl")
        cvt = []
        
        for sht in xl.sheet_names:
            df = xl.parse(sht)
            csv_name = out / f"{sht}.csv"
            df.to_csv(csv_name, index=False)
            cvt.append(csv_name.name)
        
        return cvt
    except Exception as e:
        raise Exception(f"Error converting {fp} to CSV: {str(e)}")

def chk_files(files: List[Tuple[str, int, str]], req: List[str], out: Path) -> bool:
    exist = {n for n, _, _ in files}
    if out.exists():
        exist.update(f.name for f in out.glob('*.csv'))
    
    miss = set(req) - exist
    if miss:
        messagebox.showerror(
            "Missing Files",
            f"The following required files are missing:\n{', '.join(miss)}"
        )
        return False
    
    return True