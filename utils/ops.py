import threading
import time
import os
from tkinter import filedialog

def brw_files(rt, acc_types):
    ftypes = [(n, ' '.join(ext)) for n, ext in acc_types.items()]
    ftypes.append(('All files', ' '.join(sum(acc_types.values(), ()))))
    return filedialog.askopenfilenames(parent=rt, title="Select", filetypes=ftypes)

def val_file(fn: str) -> bool:
    ext = os.path.splitext(fn.lower())[1]
    exts = ('.xlsx', '.xls', '.xlsm', '.csv')
    return ext in exts

def sim_up(pb, rt):
    def upd_prog():
        for i in range(101):
            pb['value'] = i
            time.sleep(0.05)
            rt.update()

    t = threading.Thread(target=upd_prog)
    t.start()

def mk_dir(p: str) -> None:
    os.makedirs(p, exist_ok=True)

def get_sz(fp: str) -> float:
    return os.path.getsize(fp) / (1024 * 1024)

def del_f(fp: str) -> None:
    if os.path.exists(fp):
        os.remove(fp)

def lst_dir(d: str, exts: tuple = ()) -> list:
    if not os.path.exists(d):
        return []
    return [
        os.path.join(d, f)
        for f in os.listdir(d)
        if f.endswith(exts)
    ]