import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class Scraper:
    def __init__(self, dl_folder):
        self.driver = None
        self.dl_folder = dl_folder
        self.fail_list = []
        self.max_parallel = 3
        self.successful_downloads = 0
    
    def log(self, message):
        print(message)
    
    def login(self, url, user, pwd):
        self.log(f"🔍 DEBUG: URL={url}, Username={user}, Password=******")
        
        try:
            self.log("🚀 Starting Selenium WebDriver...")
            opts = webdriver.ChromeOptions()
            opts.add_argument("--start-maximized")
            opts.add_argument("--disable-popup-blocking")

            prefs = {
                "download.default_directory": self.dl_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True
            }
            opts.add_experimental_option("prefs", prefs)
            
            self.driver = webdriver.Chrome(options=opts)
            self.driver.get(url)

            wait = WebDriverWait(self.driver, 15)  

            self.log("🔵 Waiting for username field...")
            user_field = wait.until(EC.presence_of_element_located((By.ID, "Input_Username")))
            user_field.clear()
            user_field.send_keys(user)

            self.log("🔵 Waiting for password field...")
            pwd_field = wait.until(EC.presence_of_element_located((By.ID, "Input_Password")))
            pwd_field.clear()
            pwd_field.send_keys(pwd)

            self.log("🔵 Clicking login button...")
            btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-primary-dark")))
            btn.click()

            time.sleep(5)

            tabs = self.driver.window_handles
            self.log(f"📝 Open Tabs: {tabs}")

            if len(tabs) > 1:
                self.driver.switch_to.window(tabs[-1])

            self.log("✅ Login Successful")
            return True

        except Exception as e:
            self.log(f"❌ Login Failed: {e}")
            return False
    
    def process_all_clients(self, codes, update_cb=None):
        success = 0
        self.fail_list = []
        
        for i in range(0, len(codes), self.max_parallel):
            batch = codes[i:i+self.max_parallel]
            self.log(f"🚀 Processing batch: {batch}")
            
            for code in batch:
                if self.search_client(code):
                    success += 1
                    self.successful_downloads += 1
                else:
                    self.fail_list.append(code)
                
                if update_cb:
                    update_cb(success, len(codes), self.fail_list)
                    
            time.sleep(5)  
        
        return success, self.fail_list

    def search_client(self, code):
        self.log(f"🔎 Processing client: {code}")
        wait = WebDriverWait(self.driver, 15)  

        initial_tabs = self.driver.window_handles
        if len(initial_tabs) > 1:
            self.driver.switch_to.window(initial_tabs[1])

        try:
            search_box = wait.until(EC.presence_of_element_located((By.ID, "UCBanner_txtSearch")))
            search_box.clear()
            search_box.send_keys(code)
            self.log(f"⌨️ Entering client code: {code}")

            suggestions = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ui-menu-item")))
            first_suggestion = self.driver.find_element(By.CLASS_NAME, "ui-menu-item")
            first_suggestion.click()

            time.sleep(3)
            new_tabs = self.driver.window_handles
            if len(new_tabs) > len(initial_tabs):
                client_profile_tab = new_tabs[-1]
                self.driver.switch_to.window(client_profile_tab)
                self.log("🆕 Switched to the Client Profile tab.")

                capital_gain_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Capital Gain Report')]")))
                capital_gain_btn.click()

                time.sleep(3)
                dashboard_tabs = self.driver.window_handles
                if len(dashboard_tabs) > len(new_tabs):
                    client_dashboard_tab = dashboard_tabs[-1]
                    self.driver.switch_to.window(client_dashboard_tab)
                    self.log("📊 Switched to the Client Dashboard tab.")

                    self.driver.switch_to.window(client_profile_tab)
                    self.driver.close()
                    self.log("❌ Closed Client Profile tab.")

                    self.driver.switch_to.window(client_dashboard_tab)

                    result = self.dl_holdings(code)

                    self.driver.close()
                    self.driver.switch_to.window(initial_tabs[1])
                    self.log("🔄 Closed client tab and returned to search tab.")
                    return result

            self.log(f"🚨 No new Client Dashboard tab opened for {code}")
            return False

        except Exception as e:
            self.log(f"❌ Error processing {code}: {str(e)}")
            return False

    def dl_holdings(self, code):
        try:
            self.log(f"📊 Navigating to Holdings for {code}")
            wait = WebDriverWait(self.driver, 10)

            holding_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Holding')]")))
            holding_menu.click()

            time.sleep(2)  
            as_on_date_holding = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "As on date holding")))
            as_on_date_holding.click()

            time.sleep(3)  

            self.log("💾 Downloading Holdings")
            excel_button = wait.until(EC.element_to_be_clickable((By.ID, "MainContent_imgExcel")))
            excel_button.click()

            time.sleep(4)
            
            downloaded_files = sorted(
                [f for f in os.listdir(self.dl_folder) if f.endswith(".xls") or f.endswith(".xlsx")],
                key=lambda x: os.path.getmtime(os.path.join(self.dl_folder, x)),
                reverse=True
            )

            if downloaded_files:
                latest_file = os.path.join(self.dl_folder, downloaded_files[0])
                new_file_name = os.path.join(self.dl_folder, f"{code}.xlsx")
                os.rename(latest_file, new_file_name)
                self.log(f"✅ Holdings saved as: {new_file_name}")
                return True
                
            else:
                self.log(f"⚠️ No file found for {code}")
                return False

        except Exception as e:
            self.log(f"❌ Error processing {code}: {e}")
            return False
    
    def dl_mf_transactions(self, code):
        try:
            self.log(f"📊 Navigating to MF transactions for {code}")
            wait = WebDriverWait(self.driver, 10)
            reports_menu = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//span[contains(text(),'Reports')]")))
            reports_menu.click()

            time.sleep(2)
            mf_trans = wait.until(EC.element_to_be_clickable((
                By.LINK_TEXT, "MF Transaction Report")))
            mf_trans.click()

            time.sleep(3)
            try:
                from_date = wait.until(EC.presence_of_element_located((By.ID, "MainContent_txtFromDate")))
                to_date = wait.until(EC.presence_of_element_located((By.ID, "MainContent_txtToDate")))
            except:
                self.log("Date range fields not found or not needed")

            try:
                generate_btn = wait.until(EC.element_to_be_clickable((By.ID, "MainContent_btnGenerateReport")))
                generate_btn.click()
                time.sleep(3)  
            except:
                self.log("Generate report button not found or not needed")

            self.log("💾 Downloading MF transactions")
            excel_btn = wait.until(EC.element_to_be_clickable((By.ID, "MainContent_imgExcel")))
            excel_btn.click()

            time.sleep(4)
        
            files = sorted(
                [f for f in os.listdir(self.dl_folder) 
                    if f.endswith(".xls") or f.endswith(".xlsx")],
                key=lambda x: os.path.getmtime(os.path.join(self.dl_folder, x)),
                reverse=True
            )

            if files:
                latest = os.path.join(self.dl_folder, files[0])
                new_name = os.path.join(self.dl_folder, f"{code}_MFTrans.xlsx")
                os.rename(latest, new_name)
                self.log(f"✅ MF transactions saved as: {new_name}")
                return True
            
            else:
                self.log(f"⚠️ No file found for {code} MF transactions")
                return False

        except Exception as e:
            self.log(f"❌ Error downloading MF transactions for {code}: {e}")
            return False
        
    def search_client_mf_trans(self, code):
        self.log(f"🔎 Processing client MF transactions: {code}")
        wait = WebDriverWait(self.driver, 15)

        init_tabs = self.driver.window_handles
        if len(init_tabs) > 1:
            self.driver.switch_to.window(init_tabs[1])

        try:
            search = wait.until(EC.presence_of_element_located((By.ID, "UCBanner_txtSearch")))
            search.clear()
            search.send_keys(code)
            self.log(f"⌨️ Entering client code: {code}")

            sugg = self.driver.find_element(By.CLASS_NAME, "ui-menu-item")
            sugg.click()

            time.sleep(3)
            new_tabs = self.driver.window_handles
        
            if len(new_tabs) > len(init_tabs):
                prof_tab = new_tabs[-1]
                self.driver.switch_to.window(prof_tab)
                self.log("🆕 Switched to the client profile tab.")

                reports_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//a[contains(text(),'Reports')]")))
                reports_btn.click()

                time.sleep(3)
                dash_tabs = self.driver.window_handles
            
                if len(dash_tabs) > len(new_tabs):
                    dash_tab = dash_tabs[-1]
                    self.driver.switch_to.window(dash_tab)
                    self.log("📊 Switched to the client reports tab.")

                    self.driver.switch_to.window(prof_tab)
                    self.driver.close()
                    self.log("❌ Closed client profile tab.")

                    self.driver.switch_to.window(dash_tab)
                
                    result = self.dl_mf_transactions(code)
                
                    self.driver.close()
                    self.driver.switch_to.window(init_tabs[1])
                    self.log("🔄 Closed client tab and returned to search tab.")
                
                    return result

            self.log(f"🚨 No new client dashboard tab opened for {code}")
            return False

        except Exception as e:
            self.log(f"❌ Error processing MF transactions for {code}: {str(e)}")
            return False

    def process_all_clients_mf_trans(self, codes, update_cb=None):
        success = 0
        self.fail_list = []
    
        for i in range(0, len(codes), self.max_parallel):
            batch = codes[i:i+self.max_parallel]
            self.log(f"🚀 Processing MF transactions batch: {batch}")
        
            for code in batch:
                if self.search_client_mf_trans(code):
                    success += 1
                else:
                    self.fail_list.append(code)
            
                if update_cb:
                    update_cb(success, len(codes), self.fail_list)
                
            time.sleep(5)
    
        return success, self.fail_list
    
    def quit(self):
        if self.driver:
            self.driver.quit()