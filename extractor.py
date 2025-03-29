import os
import pandas as pd
import sys
import numpy as np  


class HoldingsProcessor:
    def __init__(self, folder_path):
       self.folder_path = folder_path
    
    def log(self, message):
        print(message)
    
    def extract_section(self, df, section_name, required_columns):
        section_start = df[df.iloc[:, 0].astype(str).str.contains(section_name, na=False)].index
        if section_start.empty:
            return pd.DataFrame(), 0  
        
        start_idx = section_start[0] + 1  
        
        if start_idx >= len(df):
            return pd.DataFrame(), 0  
        
        
        header_row = df.iloc[start_idx].dropna().tolist()
        
        
        column_indices = {col: header_row.index(col) for col in required_columns if col in header_row}
        
        if len(column_indices) != len(required_columns):
            return pd.DataFrame(), 0  
        
        
        df_section = df.iloc[start_idx + 1:].reset_index(drop=True)

        
        total_row_idx = df_section[df_section.iloc[:, 0].astype(str).str.contains("TOTAL:", case=False, na=False)].index
        total_value = 0
        
        if not total_row_idx.empty:
            total_row = df_section.loc[total_row_idx[0]]  
            market_value_idx = column_indices.get("Market Value")
            
            if market_value_idx is not None and market_value_idx < len(total_row):
                total_value = pd.to_numeric(total_row.iloc[market_value_idx], errors='coerce')

            df_section = df_section.loc[:total_row_idx[0] - 1]  

        
        df_section = df_section.loc[:, list(column_indices.values())]
        df_section.columns = required_columns  

        return df_section.dropna(how='all'), total_value  

    
    def process_holdings(self):
        self.log(f"Processing holdings from: {self.folder_path}")  
        
        summary_data = []
        for file_name in os.listdir(self.folder_path):
            if file_name.endswith(".xlsx") and not file_name.startswith("~$"):
                file_path = os.path.join(self.folder_path, file_name)
                result = self.process_holdings_file(file_path)
                if result:
                    summary_data.append(result)
                else:
                    self.log(f"⚠️ Skipped file (no data extracted): {file_name}")
        
        if summary_data:
            output_file = os.path.join(self.folder_path, "Consolidated_Holdings.xlsx")
            df_summary = pd.DataFrame(
                summary_data, 
                columns=["Client Code", "Portfolio Value", "Equity", "Debt", "Gold", "Cash Equivalent"]
            )

            
            df_summary["Portfolio Value"] = df_summary["Portfolio Value"].apply(lambda x: f"{int(x):,}")

            
            df_summary["Equity (%)"] = (df_summary["Equity"] / df_summary["Portfolio Value"].replace(',', '', regex=True).astype(float)) * 100
            df_summary["Debt (%)"] = (df_summary["Debt"] / df_summary["Portfolio Value"].replace(',', '', regex=True).astype(float)) * 100
            df_summary["Gold (%)"] = (df_summary["Gold"] / df_summary["Portfolio Value"].replace(',', '', regex=True).astype(float)) * 100
            df_summary["Cash Equivalent (%)"] = (df_summary["Cash Equivalent"] / df_summary["Portfolio Value"].replace(',', '', regex=True).astype(float)) * 100

            
            df_summary[["Equity (%)", "Debt (%)", "Gold (%)", "Cash Equivalent (%)"]] = df_summary[
                ["Equity (%)", "Debt (%)", "Gold (%)", "Cash Equivalent (%)"]
            ].round(2)

            
            df_summary.to_excel(output_file, index=False)

            self.log(f"Consolidated report saved: {output_file}")
            return output_file  
        self.log("⚠️ No valid holdings files found for processing.")
        return None
            
    def categorize_values(self, df, section, asset_type_col=None, scheme_name_col=None, instrument_col=None, market_value_col="Market Value"):
        equity, debt, gold, cash_equivalent = 0, 0, 0, 0
        
        df[market_value_col] = pd.to_numeric(df[market_value_col], errors='coerce').fillna(0)

        for _, row in df.iterrows():
            market_value = row.get(market_value_col, 0)
            
            
            if section in ["Equity", "Bond"]:
                instrument = str(row.get(instrument_col, "")).upper()

                if section == "Equity":
                    if any(keyword in instrument for keyword in ["GILT BEES", "GILT", "G-SEC", "GSEC", "LONG TERM GILT", "LTGILTBEES"]):
                        debt += market_value
                        cash_equivalent += market_value
                        self.log(f"Equity - {instrument}: Added {market_value} to Debt and Cash Equivalent")
                    elif "BOND ETF" in instrument:
                        debt += market_value
                        self.log(f"Equity - {instrument}: Added {market_value} to Debt")
                    elif any(keyword in instrument for keyword in ["LONGTERM GILT", "5 YR BENCHMARK GSEC", "LIQUID BEES"]):
                        debt += market_value
                        cash_equivalent += market_value
                        self.log(f"Equity - {instrument}: Added {market_value} to Debt and Cash Equivalent")
                    elif "GOLD BEES" in instrument:
                        gold += market_value
                        self.log(f"Equity - {instrument}: Added {market_value} to Gold")
                    else:
                        equity += market_value
                        self.log(f"Equity - {instrument}: Added {market_value} to Equity")

                elif section == "Bond":
                    if any(keyword in instrument for keyword in ["SGB", "GOLD", "GOLDBOND", "SOVEREIGN"]):
                        gold += market_value
                        cash_equivalent += market_value
                        self.log(f"Bond - {instrument}: Added {market_value} to Gold and Cash Equivalent")
                    elif "GOI" in instrument:
                        debt += market_value
                        cash_equivalent += market_value
                        self.log(f"Bond - {instrument}: Added {market_value} to Debt and Cash Equivalent")
                    elif "TAX FREE" in instrument:
                        debt += market_value
                        self.log(f"Bond - {instrument}: Added {market_value} to Debt")
                    elif any(keyword in instrument for keyword in ["NCD", "NHAI"]):
                        debt += market_value
                        self.log(f"Bond - {instrument}: Added {market_value} to Debt")

            
            elif section == "Mutual Fund":
                asset_type = str(row.get(asset_type_col, "")).upper()
                scheme_name = str(row.get(scheme_name_col, "")).upper()

                if asset_type == "BALANCED":
                    equity += market_value
                    self.log(f"Mutual Fund - {asset_type}: Added {market_value} to Equity")
                elif asset_type == "CASH":
                    debt += market_value
                    cash_equivalent += market_value
                    self.log(f"Mutual Fund - {asset_type}: Added {market_value} to Debt and Cash Equivalent")
                elif asset_type == "EQUITY":
                    equity += market_value
                    self.log(f"Mutual Fund - {asset_type}: Added {market_value} to Equity")
                elif asset_type == "DEBT":
                    if "GILT" in scheme_name:
                        cash_equivalent += market_value
                        self.log(f"Mutual Fund - {scheme_name}: Added {market_value} to Cash Equivalent")
                    debt += market_value
                    self.log(f"Mutual Fund - {asset_type}: Added {market_value} to Debt")
        
        return equity, debt, gold, cash_equivalent

    def process_holdings_file(self, file_path):
        try:
            self.log(f"Processing file: {os.path.basename(file_path)}")
            df = pd.read_excel(file_path, header=None, engine='openpyxl')

            
            if df.empty or df.shape[1] < 2:
                self.log(f"⚠️ Skipping empty or invalid file: {file_path}")
                return None

            
            df.fillna(0, inplace=True)

            if df.shape[1] < 13:
                raise ValueError(f"Unexpected column count in {os.path.basename(file_path)}. Expected at least 13 columns, found {df.shape[1]}")

            
            equity_df, equity_total = self.extract_section(df, "Equity:-", ["Instrument Name", "Market Value"])
            mutual_fund_df, mf_total = self.extract_section(df, "Mutual Fund:-", ["Asset Type", "Scheme Name", "Market Value"])
            bond_df, bond_total = self.extract_section(df, "Bond:-", ["Instrument Name", "Market Value"])

            
            equity, debt, gold, cash_equivalent = 0, 0, 0, 0
            
            if not equity_df.empty:
                sec_eq, sec_debt, sec_gold, sec_cash = self.categorize_values(
                    equity_df, 
                    section="Equity",
                    instrument_col="Instrument Name", 
                    market_value_col="Market Value"
                )
                equity += sec_eq
                debt += sec_debt
                gold += sec_gold
                cash_equivalent += sec_cash

            if not mutual_fund_df.empty:
                sec_eq, sec_debt, sec_gold, sec_cash = self.categorize_values(
                    mutual_fund_df, 
                    section="Mutual Fund",
                    asset_type_col="Asset Type", 
                    scheme_name_col="Scheme Name", 
                    market_value_col="Market Value"
                )
                equity += sec_eq
                debt += sec_debt
                gold += sec_gold
                cash_equivalent += sec_cash

            if not bond_df.empty:
                sec_eq, sec_debt, sec_gold, sec_cash = self.categorize_values(
                    bond_df, 
                    section="Bond",
                    instrument_col="Instrument Name", 
                    market_value_col="Market Value"
                )
                equity += sec_eq
                debt += sec_debt
                gold += sec_gold
                cash_equivalent += sec_cash

             
            equity_total = np.nan_to_num(equity_total, nan=0)
            mf_total = np.nan_to_num(mf_total, nan=0)
            bond_total = np.nan_to_num(bond_total, nan=0)

            
            portfolio_value = equity_total + mf_total + bond_total
            
            

            self.log(f"Processed {file_path}: Portfolio Value: {portfolio_value}, Equity: {equity}, Debt: {debt}, Gold: {gold}, Cash Equivalent: {cash_equivalent}")
            return [os.path.basename(file_path).replace(".xlsx", ""), portfolio_value, equity, debt, gold, cash_equivalent]
        
        except Exception as e:
            self.log(f"Error processing {file_path}: {e}")
            return None


