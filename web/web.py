import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
class Scraper:
    def __init__(self, dl_folder, mf_folder):
        self.driver = None
        self.dl_folder = dl_folder
        self.mf_folder = mf_folder
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
        retry_count = 0
        max_retries = 2

        while retry_count <= max_retries:
            try:
                initial_tabs = self.driver.window_handles
                if len(initial_tabs) > 1:
                    self.driver.switch_to.window(initial_tabs[1])
                else:
                    self.driver.switch_to.window(initial_tabs[0])

                search_box = wait.until(EC.presence_of_element_located((By.ID, "b2-Input_Search")))
                search_box.clear()
                search_box.send_keys(code)
                self.log(f"⌨️ Entering client code: {code}")

                try:
                    suggestions_container = wait.until(
                        EC.visibility_of_element_located(
                            (By.CSS_SELECTOR, ".reactive-autocomplete-list-container")
                        )
                    )
                    first_suggestion = suggestions_container.find_element(
                        By.CSS_SELECTOR, "div:first-child"  
                    )
                    first_suggestion.click()
                    self.log("✅ Selected first autocomplete suggestion")
            
                except Exception as e:
                    self.log(f"⚠️ No suggestions found for {code}: {str(e)}, retrying...")
                    retry_count += 1
                    if retry_count > max_retries:
                        return False
                    continue

                try:
                    WebDriverWait(self.driver, 10).until(EC.number_of_windows_to_be(len(initial_tabs) + 1))
                    new_tabs = self.driver.window_handles
                    client_profile_tab = new_tabs[-1]
                    self.driver.switch_to.window(client_profile_tab)
                    self.log("🆕 Switched to Client Profile tab")
                except TimeoutException:
                    self.log(f"🚨 No new tab opened for {code}")
                    retry_count += 1
                    if retry_count > max_retries:
                        return False
                    continue

                try:
                    capital_gain_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//a[contains(text(),'Capital Gain Report')]")
                    ))
                    capital_gain_btn.click()
                    self.log("📈 Clicked Capital Gain Report button")
                except Exception as e:
                    self.log(f"⚠️ Capital Gain button error: {str(e)}")
                    self.driver.close()
                    self.driver.switch_to.window(initial_tabs[1] if len(initial_tabs) > 1 else initial_tabs[0])
                    retry_count += 1
                    if retry_count > max_retries:
                        return False
                    continue

                try:
                    WebDriverWait(self.driver, 10).until(EC.number_of_windows_to_be(len(new_tabs) + 1))
                    dashboard_tabs = self.driver.window_handles
                    client_dashboard_tab = dashboard_tabs[-1]
                    self.driver.switch_to.window(client_dashboard_tab)
                    self.log("📊 Switched to Client Dashboard tab")
                except TimeoutException:
                    self.log(f"🚨 No Dashboard tab opened for {code}")
                    self.driver.close()
                    self.driver.switch_to.window(initial_tabs[1] if len(initial_tabs) > 1 else initial_tabs[0])
                    retry_count += 1
                    if retry_count > max_retries:
                        return False
                    continue

                self.driver.switch_to.window(client_profile_tab)
                self.driver.close()
                self.driver.switch_to.window(client_dashboard_tab)
        
                result = self.dl_holdings(code)  
        
                self.driver.close()
                self.driver.switch_to.window(initial_tabs[1] if len(initial_tabs) > 1 else initial_tabs[0])
                self.log("🔄 Cleanup complete")
                return result

            except Exception as e:
                self.log(f"❌ Critical error: {str(e)}")
                current_tabs = self.driver.window_handles
                # Close all except main tab
                for tab in current_tabs[1:]:
                    self.driver.switch_to.window(tab)
                    self.driver.close()
                if current_tabs:
                    self.driver.switch_to.window(current_tabs[0])
                    retry_count += 1
                if retry_count > max_retries:
                    return False
                time.sleep(2)

        return False

    def dl_holdings(self, code):
        try:
            self.log(f"📊 Navigating to Holdings for {code}")
            wait = WebDriverWait(self.driver, 15)

            try:
                holding_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Holding')]")))
                holding_menu.click()
                time.sleep(3)
        
                as_on_date_holding = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "As on date holding")))
                as_on_date_holding.click()
                time.sleep(3)
            except Exception as e:
                self.log(f"⚠️ Error navigating to Holdings menu: {str(e)}")
                return False

            try:
                self.log("💾 Downloading Holdings")
                excel_button = wait.until(EC.element_to_be_clickable((By.ID, "MainContent_imgExcel")))
                excel_button.click()
                time.sleep(5)  
            except Exception as e:
                self.log(f"⚠️ Error clicking Excel button: {str(e)}")
                return False
    
            try:
                downloaded_files = sorted(
                    [f for f in os.listdir(self.dl_folder) if f.endswith(".xls") or f.endswith(".xlsx")],
                    key=lambda x: os.path.getmtime(os.path.join(self.dl_folder, x)),
                    reverse=True
                )

                if downloaded_files:
                    latest_file = os.path.join(self.dl_folder, downloaded_files[0])
                    new_file_name = os.path.join(self.dl_folder, f"{code}.xlsx")
                
                    os.makedirs(os.path.dirname(new_file_name), exist_ok=True)
                
                    retry = 0
                    while retry < 3:
                        try:
                            os.rename(latest_file, new_file_name)
                            self.log(f"✅ Holdings saved as: {new_file_name}")
                            return True
                        except Exception as e:
                            self.log(f"⚠️ Error renaming file (attempt {retry+1}/3): {str(e)}")
                            time.sleep(2)
                            retry += 1
                
                    return False
                else:
                    self.log(f"⚠️ No file found for {code}")
                    return False
            except Exception as e:
                self.log(f"❌ Error processing downloaded file: {str(e)}")
                return False

        except Exception as e:
            self.log(f"❌ Error processing holdings for {code}: {str(e)}")
            return False
    
    def dl_mf_transactions(self, code, from_date=None, to_date=None):
        try:
            self.log(f"📊 Navigating to MF transactions for {code}")
            wait = WebDriverWait(self.driver, 15)
    
            try:
                transactions_menu = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//span[contains(text(),'Transaction')]")))
                transactions_menu.click()
                time.sleep(3)
        
                mf_section = wait.until(EC.element_to_be_clickable((
                    By.LINK_TEXT, "Mutual Fund")))
                mf_section.click()
                time.sleep(3)
                
                if from_date and to_date:
                    try:
                        self.log(f"Setting date range: {from_date} to {to_date}")
                        
                        try:
                            from_date_input = wait.until(EC.element_to_be_clickable((
                                By.XPATH, "//input[contains(@id, 'From') or contains(@id, 'from') or contains(@placeholder, 'From')]")))
                            from_date_input.clear()
                            from_date_input.send_keys(from_date)
                            self.log(f"Set From date to {from_date}")
                        except Exception as e:
                            self.log(f"⚠️ Could not set From date: {str(e)}")
                        
                        try:
                            to_date_input = wait.until(EC.element_to_be_clickable((
                                By.XPATH, "//input[contains(@id, 'To') or contains(@id, 'to') or contains(@placeholder, 'To')]")))
                            to_date_input.clear()
                            to_date_input.send_keys(to_date)
                            self.log(f"Set To date to {to_date}")
                        except Exception as e:
                            self.log(f"⚠️ Could not set To date: {str(e)}")
                        
                        try:
                            search_btn = wait.until(EC.element_to_be_clickable((
                                By.XPATH, "//button[contains(text(), 'Submit') or contains(text(), 'Search') or contains(text(), 'Go')]")))
                            search_btn.click()
                            time.sleep(3)
                        except Exception as e:
                            self.log(f"⚠️ Could not click search button: {str(e)}")
                    except Exception as e:
                        self.log(f"⚠️ Error setting date range: {str(e)}")
                
            except Exception as e:
                self.log(f"⚠️ Error navigating to MF transactions menu: {str(e)}")
                return False

            try:
                self.log("💾 Downloading MF transactions")
                try:
                    alert = self.driver.switch_to.alert
                    alert_text = alert.text
                    if "no record" in alert_text.lower():
                        self.log(f"⚠️ Alert encountered: {alert_text}")
                        alert.accept()
                        return False
                    else:
                        alert.dismiss()
                except:
                    pass
                try:
                    excel_button = wait.until(EC.element_to_be_clickable((By.ID, "MainContent_ctl02_imgExcel")))
                    excel_button.click()
                except:
                    try:
                        excel_button = wait.until(EC.element_to_be_clickable((
                            By.XPATH, "//img[contains(@id, 'Excel')]")))
                        excel_button.click()
                    except:
                        try:
                            excel_button = wait.until(EC.element_to_be_clickable((
                                By.XPATH, "//*[contains(@title, 'Excel') or contains(@alt, 'Excel')]")))
                            excel_button.click()
                        except Exception as e:
                            self.log(f"⚠️ Could not find Excel button using multiple methods: {str(e)}")
                            return False
            
                time.sleep(5)  
            except Exception as e:
                self.log(f"⚠️ Error clicking Excel button: {str(e)}")
                return False
    
            try:
                downloaded_files = sorted(
                    [f for f in os.listdir(self.dl_folder) if f.endswith(".xls") or f.endswith(".xlsx")],
                    key=lambda x: os.path.getmtime(os.path.join(self.dl_folder, x)),
                    reverse=True
                )

                if downloaded_files:
                    latest_file = os.path.join(self.dl_folder, downloaded_files[0])
                    new_file_name = os.path.join(self.mf_folder, f"{code}_MFTrans.xlsx")
                
                    os.makedirs(os.path.dirname(new_file_name), exist_ok=True)
                
                    retry = 0
                    while retry < 3:
                        try:
                            os.rename(latest_file, new_file_name)
                            self.log(f"✅ MF transactions saved as: {new_file_name}")
                            return True
                        except Exception as e:
                            self.log(f"⚠️ Error renaming file (attempt {retry+1}/3): {str(e)}")
                            time.sleep(2)
                            retry += 1
            
                    return False
                else:
                    self.log(f"⚠️ No file found for {code} MF transactions")
                    return False
            except Exception as e:
                self.log(f"❌ Error processing downloaded file: {str(e)}")
                return False

        except Exception as e:
            self.log(f"❌ Error downloading MF transactions for {code}: {str(e)}")
            return False
        
    def search_client_mf_trans(self, code, from_date=None, to_date=None):
        self.log(f"🔎 Processing client MF transactions: {code} | Dates: {from_date} to {to_date}")
        wait = WebDriverWait(self.driver, 15)
        retry_count = 0
        max_retries = 2

        while retry_count <= max_retries:
            try:
                initial_tabs = self.driver.window_handles
                if len(initial_tabs) > 1:
                    self.driver.switch_to.window(initial_tabs[1])
                else:
                    self.driver.switch_to.window(initial_tabs[0])

                search_box = wait.until(EC.presence_of_element_located((By.ID, "b2-Input_Search")))
                search_box.clear()
                search_box.send_keys(code)
                self.log(f"⌨️ Entering client code: {code}")

                try:
                    suggestions_container = wait.until(
                        EC.visibility_of_element_located(
                            (By.CSS_SELECTOR, ".reactive-autocomplete-list-container")
                        )
                    )
                    first_suggestion = suggestions_container.find_element(
                        By.CSS_SELECTOR, "div:first-child"  
                    )
                    first_suggestion.click()
                    self.log("✅ Selected first autocomplete suggestion")
            
                except Exception as e:
                    self.log(f"⚠️ No suggestions found for {code}: {str(e)}, retrying...")
                    retry_count += 1
                    if retry_count > max_retries:
                        return False
                    continue

                try:
                    wait.until(lambda d: len(d.window_handles) > len(initial_tabs))
                    new_tabs = self.driver.window_handles
                    if len(new_tabs) > len(initial_tabs):
                        client_profile_tab = new_tabs[-1]
                        self.driver.switch_to.window(client_profile_tab)
                        self.log("🆕 Switched to Client Profile tab")
                    else:
                        raise TimeoutException("No new tab opened")
                except TimeoutException:
                    self.log(f"🚨 No new tab opened for {code}")
                    retry_count += 1
                    if retry_count > max_retries:
                        return False
                    continue

                try:
                    capital_gain_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//a[contains(text(),'Capital Gain Report')]")
                    ))
                    capital_gain_btn.click()
                    self.log("📈 Clicked Capital Gain Report button")
                except Exception as e:
                    self.log(f"⚠️ Capital Gain button error: {str(e)}")
                    self.driver.close()
                    self.driver.switch_to.window(initial_tabs[1] if len(initial_tabs) > 1 else initial_tabs[0])
                    retry_count += 1
                    if retry_count > max_retries:
                        return False
                    continue

                try:
                    wait.until(lambda d: len(d.window_handles) > len(new_tabs))
                    dashboard_tabs = self.driver.window_handles
                    if len(dashboard_tabs) > len(new_tabs):
                        client_dashboard_tab = dashboard_tabs[-1]
                        self.driver.switch_to.window(client_dashboard_tab)
                        self.log("📊 Switched to Client Dashboard tab")
                    else:
                        raise TimeoutException("No dashboard tab opened")
                except TimeoutException:
                    self.log(f"🚨 No Dashboard tab opened for {code}")
                    self.driver.close()
                    self.driver.switch_to.window(initial_tabs[1] if len(initial_tabs) > 1 else initial_tabs[0])
                    retry_count += 1
                    if retry_count > max_retries:
                        return False
                    continue

                self.driver.switch_to.window(client_profile_tab)
                self.driver.close()
                self.driver.switch_to.window(client_dashboard_tab)
        
                result = self.dl_mf_transactions(code, from_date, to_date)
        
                self.driver.close()
                if len(self.driver.window_handles) > 0:
                    self.driver.switch_to.window(self.driver.window_handles[0])
                self.log("🔄 Cleanup complete")
                return result

            except Exception as e:
                self.log(f"❌ Critical error: {str(e)}")
                current_tabs = self.driver.window_handles
                for tab in current_tabs[1:]:
                    try:
                        self.driver.switch_to.window(tab)
                        self.driver.close()
                    except:
                        pass
                if current_tabs:
                    self.driver.switch_to.window(current_tabs[0])
                    retry_count += 1
                if retry_count > max_retries:
                    return False
                time.sleep(2)
    
        return False

    def process_all_clients_mf_trans(self, codes, update_cb=None, from_date=None, to_date=None):
        success = 0
        self.fail_list = []
        
        for code in codes:
            if self.search_client_mf_trans(code, from_date, to_date):
                success += 1
            else:
                self.fail_list.append(code)
            
            if update_cb:
                update_cb(success, len(codes), self.fail_list)
            
            time.sleep(1)  
        
        return success, self.fail_list
    
    def quit(self):
        if self.driver:
            self.driver.quit()