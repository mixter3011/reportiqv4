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
    
    def login(self, url, user, pwd):
        print(f"DEBUG: URL={url}, Username={user}, Password=******")
        
        try:
            print("starting selenium webdriver...")
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

            print("waiting for username field...")
            user_field = wait.until(EC.presence_of_element_located((By.ID, "Input_Username")))
            user_field.clear()
            user_field.send_keys(user)

            print("waiting for password field...")
            pwd_field = wait.until(EC.presence_of_element_located((By.ID, "Input_Password")))
            pwd_field.clear()
            pwd_field.send_keys(pwd)

            print("clicking login button...")
            btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-primary-dark")))
            btn.click()

            time.sleep(5)

            tabs = self.driver.window_handles
            print(f"open tabs: {tabs}")

            if len(tabs) > 1:
                self.driver.switch_to.window(tabs[-1])

            print("login successful")
            return True

        except Exception as e:
            print(f"login failed: {e}")
            return False
    
    def process_all_clients(self, codes, update_cb=None):
        success = 0
        self.fail_list = []
        
        for i in range(0, len(codes), self.max_parallel):
            batch = codes[i:i+self.max_parallel]
            print(f"processing batch: {batch}")
            
            for code in batch:
                if self.search_client(code):
                    success += 1
                else:
                    self.fail_list.append(code)
                
                if update_cb:
                    update_cb(success, len(codes), self.fail_list)
                    
            time.sleep(5)
        
        return success, self.fail_list

    def search_client(self, code):
        print(f"processing client: {code}")
        wait = WebDriverWait(self.driver, 10)

        init_tabs = self.driver.window_handles
        if len(init_tabs) > 1:
            self.driver.switch_to.window(init_tabs[1])

        try:
            search = wait.until(EC.presence_of_element_located((By.ID, "UCBanner_txtSearch")))
            search.clear()
            search.send_keys(code)
            print(f"entering client code: {code}")

            sugg = self.driver.find_element(By.CLASS_NAME, "ui-menu-item")
            sugg.click()

            time.sleep(3)
            new_tabs = self.driver.window_handles
            
            if len(new_tabs) > len(init_tabs):
                prof_tab = new_tabs[-1]
                self.driver.switch_to.window(prof_tab)
                print("switched to the client profile tab.")

                cg_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//a[contains(text(),'Capital Gain Report')]")))
                cg_btn.click()

                time.sleep(3)
                dash_tabs = self.driver.window_handles
                
                if len(dash_tabs) > len(new_tabs):
                    dash_tab = dash_tabs[-1]
                    self.driver.switch_to.window(dash_tab)
                    print("switched to the client dashboard tab.")

                    self.driver.switch_to.window(prof_tab)
                    self.driver.close()
                    print("closed client profile tab.")

                    self.driver.switch_to.window(dash_tab)
                    
                    result = self.dl_holdings(code)
                    
                    self.driver.close()
                    self.driver.switch_to.window(init_tabs[1])
                    print("closed client tab and returned to search tab.")
                    
                    return result

            print(f"no new client dashboard tab opened for {code}")
            return False

        except Exception as e:
            print(f"error processing {code}: {str(e)}")
            return False

    def dl_holdings(self, code):
        try:
            print(f"navigating to holdings for {code}")
            wait = WebDriverWait(self.driver, 10)

            hold_menu = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//span[contains(text(),'Holding')]")))
            hold_menu.click()

            time.sleep(2)
            as_on_date = wait.until(EC.element_to_be_clickable((
                By.LINK_TEXT, "as on date holding")))
            as_on_date.click()

            time.sleep(3)

            print("downloading holdings")
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
                new_name = os.path.join(self.dl_folder, f"{code}.xlsx")
                os.rename(latest, new_name)
                print(f"holdings saved as: {new_name}")
                return True
                
            else:
                print(f"no file found for {code}")
                return False

        except Exception as e:
            print(f"error downloading for {code}: {e}")
            return False
    
    def quit(self):
        if self.driver:
            self.driver.quit()