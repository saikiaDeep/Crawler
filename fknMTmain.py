from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import os
import json
import csv
import pandas as pd
from threading import Thread, Lock
from queue import Queue
import base64

class NHAIFormDownloader:
    def __init__(self, download_folder="nhai_pdfs", num_threads=3):
        """
        Initialize the downloader with Chrome WebDriver
        
        Args:
            download_folder: Folder where PDFs will be saved
            num_threads: Number of parallel threads to use
        """
        self.download_folder = os.path.abspath(download_folder)
        self.num_threads = num_threads
        
      
        if not os.path.exists(self.download_folder):
            os.makedirs(self.download_folder)
        
       
        self.success_count = 0
        self.failed_refs = []
        self.lock = Lock()
        

        self.print_settings = {
            "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": ""
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
            "isHeaderFooterEnabled": False,
            "isLandscapeEnabled": False
        }
        
    
        self.main_driver = None
        self.cookies = None
        
    def get_chrome_options(self):
        """Create Chrome options for each thread"""
        chrome_options = Options()
        
        prefs = {
            'printing.print_preview_sticky_settings.appState': json.dumps(self.print_settings),
            'savefile.default_directory': self.download_folder,
            'download.default_directory': self.download_folder,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'plugins.always_open_pdf_externally': True
        }
        
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_argument('--kiosk-printing')
   
        
        return chrome_options
        
    def manual_login(self, login_url):
        """
        Open login page and wait for manual login
        
        Args:
            login_url: URL of the login page
        """
        print("Opening login page...")
        chrome_options = self.get_chrome_options()
        self.main_driver = webdriver.Chrome(options=chrome_options)
        self.main_driver.get(login_url)
        
        print("\n" + "="*60)
        print("PLEASE COMPLETE THE FOLLOWING STEPS MANUALLY:")
        print("1. Enter your username")
        print("2. Enter your password")
        print("3. Solve the CAPTCHA")
        print("4. Click the login button")
        print("5. Wait until you're successfully logged in")
        print("="*60)
        
        input("\nPress ENTER after you have successfully logged in...")
        print("Login confirmed. Capturing session cookies...")
        
        # Capture cookies for other threads
        self.cookies = self.main_driver.get_cookies()
        time.sleep(2)
        
        print(f"Cookies captured. Starting {self.num_threads} worker threads...")
    
    def print_page_to_pdf(self, driver, refer_no):
        """
        Print current page to PDF using Chrome's print function
        
        Args:
            driver: WebDriver instance
            refer_no: Reference number for naming the PDF
        """
        try:
            pdf_path = os.path.join(self.download_folder, f"{refer_no}.pdf")
            
            # Execute Chrome's print command
            result = driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "landscape": False,
                "paperWidth": 8.27,  # A4 width in inches
                "paperHeight": 11.69,  # A4 height in inches
                "marginTop": 0.4,
                "marginBottom": 0.4,
                "marginLeft": 0.4,
                "marginRight": 0.4,
                "scale": 0.6,  # Reduce scale to fit content
                "preferCSSPageSize": False
            })
            
         
            with open(pdf_path, 'wb') as f:
                f.write(base64.b64decode(result['data']))
            
            return True
            
        except Exception as e:
            print(f"✗ Error saving PDF for {refer_no}: {str(e)}")
            return False
    
    def worker_thread(self, task_queue, base_url, thread_id):
        """
        Worker thread that processes reference numbers
        
        Args:
            task_queue: Queue containing reference numbers to process
            base_url: Base URL template for application forms
            thread_id: ID of this thread
        """
  
        chrome_options = self.get_chrome_options()
        driver = webdriver.Chrome(options=chrome_options)
        
        try:
           
            driver.get(base_url.split('/AIVacancy')[0])
            
          
            if self.cookies:
                for cookie in self.cookies:
                    try:
                        driver.add_cookie(cookie)
                    except Exception as e:
                        pass  # Some cookies might not be applicable
            
            # Process tasks from queue
            while True:
                try:
                    refer_no, index, total = task_queue.get(timeout=1)
                except:
                    break 
                
                try:
                    last_two = int(str(refer_no).replace("2025", "", 1))
                    url = f"{base_url}?ABID={last_two}&ReferNo={refer_no}"
                    
                    print(f"[Thread-{thread_id}] [{index}/{total}] Processing: {refer_no}")
                
                    driver.get(url)
                    time.sleep(3)  
                    
             
                    if self.print_page_to_pdf(driver, refer_no):
                        with self.lock:
                            self.success_count += 1
                        print(f"[Thread-{thread_id}] ✓ Saved: {refer_no}.pdf")
                    else:
                        with self.lock:
                            self.failed_refs.append(refer_no)
                    
                  
                    time.sleep(1)
                    
                except Exception as e:
                    print(f"[Thread-{thread_id}] ✗ Error processing {refer_no}: {str(e)}")
                    with self.lock:
                        self.failed_refs.append(refer_no)
                finally:
                    task_queue.task_done()
        
        finally:
            driver.quit()
    
    def process_reference_numbers(self, refer_numbers, base_url, abid=21):
        """
        Process all reference numbers using multiple threads
        
        Args:
            refer_numbers: List of reference numbers
            base_url: Base URL template for application forms
            abid: ABID parameter (default: 21)
        """
        print(f"\nStarting to process {len(refer_numbers)} reference numbers with {self.num_threads} threads...")
        print("="*60)
     
        task_queue = Queue()
        
 
        for i, refer_no in enumerate(refer_numbers, 1):
            task_queue.put((refer_no, i, len(refer_numbers)))

        threads = []
        for thread_id in range(self.num_threads):
            thread = Thread(
                target=self.worker_thread,
                args=(task_queue, base_url, thread_id + 1)
            )
            thread.start()
            threads.append(thread)
        
        for thread in threads:
            thread.join()
        

        print("\n" + "="*60)
        print("PROCESSING COMPLETE")
        print(f"Total processed: {len(refer_numbers)}")
        print(f"Successful: {self.success_count}")
        print(f"Failed: {len(self.failed_refs)}")
        
        if self.failed_refs:
            print(f"\nFailed reference numbers: {', '.join(map(str, self.failed_refs))}")
        
        print(f"\nPDFs saved in: {self.download_folder}")
        print("="*60)
    
    def close(self):
        """Close the main browser"""
        if self.main_driver:
            print("\nClosing main browser...")
            self.main_driver.quit()


# Main execution
if __name__ == "__main__":
    # Configuration
    LOGIN_URL = "https://vacancy.nhai.org/AIVacancy/UserManager/User/index.aspx"
    BASE_URL = "https://vacancy.nhai.org/AIVacancy/ApplicationFormView.aspx"
    

    NUM_THREADS = 10 
    

    csv_path = "reference_numbers.csv"
    df = pd.read_csv(csv_path)
    REFER_NUMBERS = df["ReferNo"].tolist()
    
    ABID = 21
    
    try:
      
        downloader = NHAIFormDownloader(
            download_folder="nhai_pdfs",
            num_threads=NUM_THREADS
        )
        
  
        downloader.manual_login(LOGIN_URL)
        
        
        downloader.process_reference_numbers(
            refer_numbers=REFER_NUMBERS,
            base_url=BASE_URL,
            abid=ABID
        )
        
    except KeyboardInterrupt:
        print("\n\nProcess interrupted by user")
    except Exception as e:
        print(f"\n\nUnexpected error: {str(e)}")
    finally:
       
        downloader.close()
        
    print("\nScript completed!")