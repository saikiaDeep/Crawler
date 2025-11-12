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

class NHAIFormDownloader:
    def __init__(self, download_folder="nhai_pdfs"):
        """
        Initialize the downloader with Chrome WebDriver
        
        Args:
            download_folder: Folder where PDFs will be saved
        """
        self.download_folder = os.path.abspath(download_folder)
        
        # Create download folder if it doesn't exist
        if not os.path.exists(self.download_folder):
            os.makedirs(self.download_folder)
        
        # Setup Chrome options
        chrome_options = Options()
        
        # Configure print settings for PDF
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
        
        # Initialize driver
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 20)
        
    def manual_login(self, login_url):
        """
        Open login page and wait for manual login
        
        Args:
            login_url: URL of the login page
        """
        print("Opening login page...")
        self.driver.get(login_url)
        
        print("\n" + "="*60)
        print("PLEASE COMPLETE THE FOLLOWING STEPS MANUALLY:")
        print("1. Enter your username")
        print("2. Enter your password")
        print("3. Solve the CAPTCHA")
        print("4. Click the login button")
        print("5. Wait until you're successfully logged in")
        print("="*60)
        
        input("\nPress ENTER after you have successfully logged in...")
        print("Login confirmed. Proceeding with automation...")
        time.sleep(2)
    
    def print_page_to_pdf(self, refer_no):
        """
        Print current page to PDF using Chrome's print function
        
        Args:
            refer_no: Reference number for naming the PDF
        """
        try:
            pdf_path = os.path.join(self.download_folder, f"{refer_no}.pdf")
            
            # Execute Chrome's print command
            result = self.driver.execute_cdp_cmd("Page.printToPDF", {
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
            
            # Save PDF
            import base64
            with open(pdf_path, 'wb') as f:
                f.write(base64.b64decode(result['data']))
            
            print(f"✓ Saved: {refer_no}.pdf")
            return True
            
        except Exception as e:
            print(f"✗ Error saving PDF for {refer_no}: {str(e)}")
            return False
    
    def process_reference_numbers(self, refer_numbers, base_url, abid=21):
        """
        Process all reference numbers and generate PDFs
        
        Args:
            refer_numbers: List of reference numbers
            base_url: Base URL template for application forms
            abid: ABID parameter (default: 21)
        """
        success_count = 0
        failed_refs = []
        
        print(f"\nStarting to process {len(refer_numbers)} reference numbers...")
        print("="*60)
        
        for i, refer_no in enumerate(refer_numbers, 1):
            try:
                last_two=int(str(refer_no).replace("2025", "", 1))
                # Construct URL
                url = f"{base_url}?ABID={last_two}&ReferNo={refer_no}"
                
                print(f"\n[{i}/{len(refer_numbers)}] Processing: {refer_no}")
                
                # Navigate to the page
                self.driver.get(url)
                
                # Wait for page to load (adjust selector based on actual page)
                time.sleep(3)  # Basic wait, can be improved with explicit waits
                
                # Try to check if page loaded successfully
                # You might need to adjust this based on the actual page structure
                try:
                    # Wait for some element that indicates page is loaded
                    # Example: self.wait.until(EC.presence_of_element_located((By.ID, "some_element")))
                    pass
                except:
                    print(f"  Warning: Page might not have loaded completely")
                
                # Print to PDF
                if self.print_page_to_pdf(refer_no):
                    success_count += 1
                else:
                    failed_refs.append(refer_no)
                
                # Small delay between requests to avoid overwhelming the server
                time.sleep(2)
                
            except Exception as e:
                print(f"✗ Error processing {refer_no}: {str(e)}")
                failed_refs.append(refer_no)
        
        # Summary
        print("\n" + "="*60)
        print("PROCESSING COMPLETE")
        print(f"Total processed: {len(refer_numbers)}")
        print(f"Successful: {success_count}")
        print(f"Failed: {len(failed_refs)}")
        
        if failed_refs:
            print(f"\nFailed reference numbers: {', '.join(failed_refs)}")
        
        print(f"\nPDFs saved in: {self.download_folder}")
        print("="*60)
    
    def close(self):
        """Close the browser"""
        print("\nClosing browser...")
        self.driver.quit()


# Main execution
if __name__ == "__main__":
    # Configuration
    LOGIN_URL = "https://vacancy.nhai.org/AIVacancy/UserManager/User/index.aspx"
    BASE_URL = "https://vacancy.nhai.org/AIVacancy/ApplicationFormView.aspx"
    
    # List of reference numbers - REPLACE THIS WITH YOUR ACTUAL LIST
   
    csv_path = "reference_numbers.csv"
    df = pd.read_csv(csv_path)

    # Convert the column to a Python list
    ref_list = df["ReferNo"].tolist()

    REFER_NUMBERS = ref_list
    # Optional: ABID parameter (default is 21)
    ABID = 21
    
    try:
        # Initialize downloader
        downloader = NHAIFormDownloader(download_folder="nhai_pdfs")
        
        # Step 1: Manual login
        downloader.manual_login(LOGIN_URL)
        
        # Step 2: Process all reference numbers
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
        # Close browser
        downloader.close()
        
    print("\nScript completed!")