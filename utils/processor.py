import os
import numpy as np
import pandas as pd

class Processor:
    def __init__(self, folder):
        self.folder = folder
    
    def log(self, msg):
        print(msg)
    
    def extract_section(self, df, section, req_cols):
        start_idx = df[df.iloc[:, 0].astype(str).str.contains(section, na=False)].index
        if start_idx.empty:
            return pd.DataFrame(), 0
        
        start = start_idx[0] + 1
        if start >= len(df):
            return pd.DataFrame(), 0
        
        header = df.iloc[start].dropna().tolist()
        
        cols = {col: header.index(col) for col in req_cols if col in header}
        if len(cols) != len(req_cols):
            return pd.DataFrame(), 0
        
        df_sec = df.iloc[start + 1:].reset_index(drop=True)
        
        total_idx = df_sec[df_sec.iloc[:, 0].astype(str).str.contains(
            "TOTAL:", case=False, na=False)].index
        total = 0
        
        if not total_idx.empty:
            total_row = df_sec.loc[total_idx[0]]
            mv_idx = cols.get("Market Value")
            
            if mv_idx is not None and mv_idx < len(total_row):
                total = pd.to_numeric(total_row.iloc[mv_idx], errors='coerce')
            
            df_sec = df_sec.loc[:total_idx[0] - 1]
        
        df_sec = df_sec.loc[:, list(cols.values())]
        df_sec.columns = req_cols
        
        return df_sec.dropna(how='all'), total

    def categorize(self, df, section, asset_col=None, scheme_col=None, instr_col=None, mv_col="Market Value"):
        eq, debt, gold, cash = 0, 0, 0, 0
        
        df[mv_col] = pd.to_numeric(df[mv_col], errors='coerce').fillna(0)
        
        for _, row in df.iterrows():
            mv = row.get(mv_col, 0)
            
            if section in ["Equity", "Bond"]:
                instr = str(row.get(instr_col, "")).upper()
                
                if section == "Equity":
                    gilt_keys = ["GILT BEES", "GILT", "G-SEC", "GSEC", "LONG TERM GILT", "LTGILTBEES"]
                    
                    if any(k in instr for k in gilt_keys):
                        debt += mv
                        cash += mv
                    elif "BOND ETF" in instr:
                        debt += mv
                    elif any(k in instr for k in ["LONGTERM GILT", "5 YR BENCHMARK GSEC", "LIQUID BEES"]):
                        debt += mv
                        cash += mv
                    elif "GOLD BEES" in instr:
                        gold += mv
                    else:
                        eq += mv
                
                elif section == "Bond":
                    if any(k in instr for k in ["SGB", "GOLD", "GOLDBOND", "SOVEREIGN"]):
                        gold += mv
                        cash += mv
                    elif "GOI" in instr:
                        debt += mv
                        cash += mv
                    elif any(k in instr for k in ["TAX FREE", "NCD", "NHAI"]):
                        debt += mv
            
            elif section == "Mutual Fund":
                asset = str(row.get(asset_col, "")).upper()
                scheme = str(row.get(scheme_col, "")).upper()
                
                if asset == "BALANCED":
                    eq += mv
                elif asset == "CASH":
                    debt += mv
                    cash += mv
                elif asset == "EQUITY":
                    eq += mv
                elif asset == "DEBT":
                    if "GILT" in scheme:
                        cash += mv
                    debt += mv
        
        return eq, debt, gold, cash

    def process_file(self, path):
        try:
            self.log(f"Processing file: {os.path.basename(path)}")
            df = pd.read_excel(path, header=None, engine='openpyxl')
            
            if df.empty or df.shape[1] < 2:
                self.log(f"⚠️ Skipping empty file: {path}")
                return None
            
            df.fillna(0, inplace=True)
            
            if df.shape[1] < 13:
                raise ValueError(f"Unexpected column count in {os.path.basename(path)}")
            
            eq_df, eq_total = self.extract_section(df, "Equity:-", ["Instrument Name", "Market Value"])
            mf_df, mf_total = self.extract_section(df, "Mutual Fund:-", 
                ["Asset Type", "Scheme Name", "Market Value"])
            bond_df, bond_total = self.extract_section(df, "Bond:-", ["Instrument Name", "Market Value"])
            
            eq, debt, gold, cash = 0, 0, 0, 0
            
            if not eq_df.empty:
                e, d, g, c = self.categorize(
                    eq_df, 
                    section="Equity",
                    instr_col="Instrument Name", 
                    mv_col="Market Value"
                )
                eq += e
                debt += d
                gold += g
                cash += c
            
            if not mf_df.empty:
                e, d, g, c = self.categorize(
                    mf_df, 
                    section="Mutual Fund",
                    asset_col="Asset Type", 
                    scheme_col="Scheme Name", 
                    mv_col="Market Value"
                )
                eq += e
                debt += d
                gold += g
                cash += c
            
            if not bond_df.empty:
                e, d, g, c = self.categorize(
                    bond_df, 
                    section="Bond",
                    instr_col="Instrument Name", 
                    mv_col="Market Value"
                )
                eq += e
                debt += d
                gold += g
                cash += c
            
            eq_total = np.nan_to_num(eq_total, nan=0)
            mf_total = np.nan_to_num(mf_total, nan=0)
            bond_total = np.nan_to_num(bond_total, nan=0)
            
            port_val = eq_total + mf_total + bond_total
            
            self.log(f"Processed {path}: Value: {port_val}, Equity: {eq}, Debt: {debt}, Gold: {gold}, Cash: {cash}")
            
            return [os.path.basename(path).replace(".xlsx", ""), port_val, eq, debt, gold, cash]
        
        except Exception as e:
            self.log(f"Error processing {path}: {e}")
            return None

    def run(self):
        self.log(f"Processing holdings from: {self.folder}")
        
        data = []
        for file in os.listdir(self.folder):
            if file.endswith(".xlsx") and not file.startswith("~$"):
                path = os.path.join(self.folder, file)
                result = self.process_file(path)
                if result:
                    data.append(result)
                else:
                    self.log(f"Skipped file: {file}")
        
        if data:
            out_file = os.path.join(self.folder, "Consolidated_Holdings.xlsx")
            df = pd.DataFrame(
                data, 
                columns=["Client Code", "Portfolio Value", "Equity", "Debt", "Gold", "Cash Equivalent"]
            )
            
            df["Portfolio Value"] = df["Portfolio Value"].apply(lambda x: f"{int(x):,}")
            
            portval = df["Portfolio Value"].replace(',', '', regex=True).astype(float)
            df["Equity (%)"] = (df["Equity"] / portval) * 100
            df["Debt (%)"] = (df["Debt"] / portval) * 100
            df["Gold (%)"] = (df["Gold"] / portval) * 100
            df["Cash Equivalent (%)"] = (df["Cash Equivalent"] / portval) * 100
            
            df[["Equity (%)", "Debt (%)", "Gold (%)", "Cash Equivalent (%)"]] = df[
                ["Equity (%)", "Debt (%)", "Gold (%)", "Cash Equivalent (%)"]
            ].round(2)
            
            df.to_excel(out_file, index=False)
            
            self.log(f"Consolidated report saved: {out_file}")
            return out_file
            
        self.log("No valid holdings files found for processing.")
        return None