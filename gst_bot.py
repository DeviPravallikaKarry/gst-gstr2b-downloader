from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

from selenium.common.exceptions import TimeoutException

import time
import os
import glob
import pandas as pd
from datetime import datetime


class GSTDownloader:

    # -----------------------------------
    # CONSTRUCTOR
    # -----------------------------------

    def __init__(self):

        self.driver = None
        self.wait = None
        self.download_folder = None

    # -----------------------------------
    # LOG FUNCTION
    # -----------------------------------

    def log(self, message):

        print(message)

    # -----------------------------------
    # READ EXCEL USERS
    # -----------------------------------

    def read_excel_users(self, excel_file):

        try:

            df = pd.read_excel(excel_file)

            users = []

            for index, row in df.iterrows():

                username = str(row["Username"]).strip()
                password = str(row["Password"]).strip()

                users.append({
                    "username": username,
                    "password": password
                })

            self.log(f"Total users loaded: {len(users)}")

            return users

        except Exception as e:

            self.log("Failed to read Excel file")

            self.log(f"Error: {e}")

            return []

    # -----------------------------------
    # SAVE REPORT
    # -----------------------------------

    def save_report(self, results):

        try:

            report_df = pd.DataFrame(results)

            timestamp = datetime.now().strftime(
                "%Y%m%d_%H%M%S"
            )

            report_folder = os.path.join(
                os.getcwd(),
                "Reports"
            )

            os.makedirs(
                report_folder,
                exist_ok=True
            )

            report_file = os.path.join(
                report_folder,
                f"GST_Report_{timestamp}.xlsx"
            )

            report_df.to_excel(
                report_file,
                index=False
            )

            self.log(f"Report saved: {report_file}")

        except Exception as e:

            self.log("Failed to save report")

            self.log(f"Error: {e}")

    # -----------------------------------
    # GET QUARTER MONTHS
    # -----------------------------------

    def get_quarter_months(self, quarter):

        quarter_mapping = {

            "Q1": [
                "April",
                "May",
                "June"
            ],

            "Q2": [
                "July",
                "August",
                "September"
            ],

            "Q3": [
                "October",
                "November",
                "December"
            ],

            "Q4": [
                "January",
                "February",
                "March"
            ]
        }

        return quarter_mapping.get(
            quarter,
            []
        )

    # -----------------------------------
    # PROCESS WHOLE QUARTER
    # -----------------------------------

    def process_quarter(

            self,
            financial_year,
            quarter

    ):

        months = self.get_quarter_months(
            quarter
        )

        quarter_labels = {

            "Q1": "Quarter 1 (Apr - Jun)",
            "Q2": "Quarter 2 (Jul - Sep)",
            "Q3": "Quarter 3 (Oct - Dec)",
            "Q4": "Quarter 4 (Jan - Mar)"
        }

        gst_quarter = quarter_labels[quarter]

        success_count = 0
        skipped_count = 0
        failed_count = 0

        # LOOP MONTHS

        for month in months:

            print("\n")
            print("-" * 50)
            print(f"PROCESSING MONTH: {month}")
            print("-" * 50)

            try:

                self.select_return_period(
                    financial_year,
                    gst_quarter,
                    month
                )

                self.search_returns()

                tile_found = self.open_gstr2b()

                # TILE FOUND

                if tile_found:

                    download_success = self.download_excel()

                    if download_success:

                        success_count += 1

                    else:

                        skipped_count += 1

                    # GO BACK TO RETURNS PAGE

                    self.return_to_returns_page()

                # TILE NOT FOUND

                else:

                    skipped_count += 1

                    print(
                        f"No GSTR2B available for {month}"
                    )



            except Exception as e:

                error_message = str(e)

                if "Failed selecting dropdown" in error_message:

                    skipped_count += 1

                    self.log(

                        f"Skipping month due to unavailable dropdown: {month}"

                    )


                else:

                    failed_count += 1

                self.log(

                    f"Error while processing {month}"

                )

                self.log(error_message)

                # RECOVER NAVIGATION

                try:

                    self.open_returns_dashboard()


                except Exception as nav_error:

                    self.log(

                        f"Navigation recovery failed: {nav_error}"

                    )

        # RETURN FINAL COUNTS

        return {

            "success": success_count,
            "skipped": skipped_count,
            "failed": failed_count
        }

    # -----------------------------------
    # WAIT FOR GST LOADER
    # -----------------------------------

    def wait_for_loader_to_disappear(self):

        try:

            WebDriverWait(self.driver, 10).until(
                EC.invisibility_of_element_located(
                    (
                        By.CLASS_NAME,
                        "dimmer-holder"
                    )
                )
            )

        except:

            pass

    def return_to_returns_page(self):

        try:

            self.driver.back()

            self.wait_for_loader_to_disappear()

            self.wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "//button[contains(text(),'Search')]"
                    )
                )
            )

            self.log(
                "Returned to Returns page"
            )

            time.sleep(2)

        except Exception as e:

            self.log(
                "Back navigation failed"
            )

            self.log(str(e))

            self.log(
                "Reopening Returns Dashboard"
            )

            self.open_returns_dashboard()

    # -----------------------------------
    # WAIT FOR DOWNLOAD COMPLETION
    # -----------------------------------

    def wait_for_download(
            self,
            files_before_download,
            timeout=120
    ):

        self.log(
            "Waiting for download completion..."
        )

        start_time = time.time()

        while True:

            current_files = set(
                glob.glob(
                    os.path.join(
                        self.download_folder,
                        "*"
                    )
                )
            )

            new_files = (
                    current_files -
                    files_before_download
            )

            completed_files = []

            for file in new_files:

                if (
                        not file.endswith(".crdownload")
                        and
                        not file.endswith(".tmp")
                ):
                    completed_files.append(file)

            if completed_files:

                latest_file = max(
                    completed_files,
                    key=os.path.getctime
                )

                # VERIFY FILE SIZE STABLE

                try:

                    size1 = os.path.getsize(
                        latest_file
                    )

                    time.sleep(2)

                    size2 = os.path.getsize(
                        latest_file
                    )

                except Exception:

                    time.sleep(1)

                    continue

                if size1 == size2:
                    self.log(
                        f"Download complete: {latest_file}"
                    )

                    return latest_file

            elapsed = (
                    time.time() - start_time
            )

            if elapsed > timeout:
                self.log(
                    "Download timeout"
                )

                return None

            time.sleep(1)

    # -----------------------------------
    # VERSIONED FOLDER CREATION
    # -----------------------------------

    def create_unique_folder(
            self,
            base_folder
    ):

        if not os.path.exists(base_folder):
            return base_folder

        counter = 1

        while True:

            new_folder = (
                f"{base_folder}_{counter}"
            )

            if not os.path.exists(new_folder):
                return new_folder

            counter += 1

    # -----------------------------------
    # SETUP DRIVER
    # -----------------------------------

    def setup_driver(self, username):

        base_folder = os.path.join(
            os.getcwd(),
            "GST_Downloads",
            username
        )

        self.download_folder = (
            self.create_unique_folder(
                base_folder
            )
        )

        os.makedirs(self.download_folder, exist_ok=True)

        self.log(f"Download folder: {self.download_folder}")

        chrome_options = webdriver.ChromeOptions()

        prefs = {
            "download.default_directory": self.download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        }

        chrome_options.add_experimental_option(
            "prefs",
            prefs
        )

        chrome_options.add_argument(
            "--disable-blink-features=AutomationControlled"
        )

        chrome_options.add_experimental_option(
            "excludeSwitches",
            ["enable-automation"]
        )

        chrome_options.add_experimental_option(
            "useAutomationExtension",
            False
        )

        self.driver = webdriver.Chrome(
            service=Service(
                ChromeDriverManager().install()
            ),
            options=chrome_options
        )

        self.driver.maximize_window()

        self.wait = WebDriverWait(self.driver, 20)

        self.log("Browser opened")

        self.driver.execute_script(
            """
            Object.defineProperty(
                navigator,
                'webdriver',
                {
                    get: () => undefined
                }
            )
            """
        )

    # -----------------------------------
    # CLICK ELEMENT
    # -----------------------------------

    def click_element(self, xpath):

        try:

            self.wait_for_loader_to_disappear()

            element = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, xpath)
                )
            )

            self.driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});",
                element
            )

            time.sleep(1)

            element.click()


        except Exception as e:

            self.log("Unable to click element")

            self.log(f"Error: {e}")

            raise

    def click_download_button(self):

        possible_xpaths = [

            # EXACT EXCEL BUTTON

            "//button[contains(.,'DOWNLOAD GSTR-2B DETAILS')]",

            # FALLBACK

            "//*[contains(text(),'DOWNLOAD GSTR-2B DETAILS')]",

            # GENERATE EXCEL

            "//*[contains(text(),'GENERATE EXCEL FILE TO DOWNLOAD')]",

            "//*[contains(text(),'GENERATE EXCEL')]"
        ]

        for xpath in possible_xpaths:

            try:

                element = self.wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, xpath)
                    )
                )

                self.driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});",
                    element
                )

                time.sleep(1)

                self.driver.execute_script(
                    "arguments[0].click();",
                    element
                )

                self.log(
                    f"Download button found: {xpath}"
                )

                return True

            except:
                continue

        return False

    # -----------------------------------
    # ENTER TEXT
    # -----------------------------------

    def enter_text(self, by_type, locator, text):

        try:

            element = self.wait.until(
                EC.presence_of_element_located(
                    (by_type, locator)
                )
            )

            element.clear()

            element.send_keys(text)

        except Exception as e:

            self.log("Unable to enter text")

            self.log(f"Error: {e}")

            raise

    def select_dropdown(
            self,
            xpath,
            visible_text,
            retries=3
    ):

        for attempt in range(retries):

            try:

                self.wait_for_loader_to_disappear()

                # WAIT FOR DROPDOWN

                dropdown = WebDriverWait(
                    self.driver,
                    15
                ).until(
                    EC.presence_of_element_located(
                        (By.XPATH, xpath)
                    )
                )

                # WAIT UNTIL ENABLED

                WebDriverWait(
                    self.driver,
                    15
                ).until(
                    lambda d: dropdown.is_enabled()
                )

                # RE-FETCH DROPDOWN
                # GST refreshes DOM often

                dropdown = self.driver.find_element(
                    By.XPATH,
                    xpath
                )

                select = Select(dropdown)

                # WAIT FOR OPTIONS

                WebDriverWait(
                    self.driver,
                    15
                ).until(

                    lambda d: any(

                        option.text.strip() == visible_text

                        for option in Select(

                            d.find_element(
                                By.XPATH,
                                xpath
                            )

                        ).options
                    )
                )

                # RE-FETCH AGAIN

                dropdown = self.driver.find_element(
                    By.XPATH,
                    xpath
                )

                select = Select(dropdown)

                select.select_by_visible_text(
                    visible_text
                )

                self.log(
                    f"Selected: {visible_text}"
                )

                time.sleep(2)

                return True

            except Exception as e:

                self.log(
                    f"Dropdown retry {attempt + 1}"
                )

                self.log(
                    f"Dropdown error: {type(e).__name__}"
                )

                self.log(str(e))

                time.sleep(3)

        raise Exception(
            f"Failed selecting dropdown: {visible_text}"
        )

    # -----------------------------------
    # LOGIN FUNCTION
    # -----------------------------------

    def login(self, username, password):

        self.driver.get(
            "https://services.gst.gov.in/services/login"
        )

        self.log("GST login page opened")

        self.enter_text(
            By.ID,
            "username",
            username
        )

        self.log("Username entered")

        self.enter_text(
            By.ID,
            "user_pass",
            password
        )

        self.log("Password entered")

        input(
            "Enter CAPTCHA manually in browser, then press Enter here..."
        )

        self.click_element(
            "//button[@type='submit']"
        )

        self.log("Login button clicked")

        time.sleep(3)

        page_source = self.driver.page_source.lower()

        if "invalid username" in page_source:
            raise Exception("Invalid username")

        if "invalid captcha" in page_source:
            raise Exception("Invalid captcha")

        if "invalid username or password" in page_source:
            raise Exception("Invalid credentials")

        try:

            self.wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "//a[contains(text(),'Services')]"
                    )
                )
            )

            self.log("Dashboard loaded successfully")


        except TimeoutException:

            raise Exception(

                "Login failed or dashboard did not load"

            )

    # -----------------------------------
    # HANDLE POPUPS FUNCTION
    # -----------------------------------

    def handle_popups(self):

        for i in range(2):

            try:

                self.wait_for_loader_to_disappear()

                popup_button = WebDriverWait(
                    self.driver,
                    5
                ).until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//*[contains(text(),'Remind me later')]"
                        )
                    )
                )

                time.sleep(1)

                popup_button.click()

                self.log(f"Popup {i + 1} handled")

            except:

                self.log(f"Popup {i + 1} not found")

    # -----------------------------------
    # OPEN RETURNS DASHBOARD
    # -----------------------------------

    def open_returns_dashboard(self):

        self.click_element(
            "//a[@class='dropdown-toggle' and contains(text(),'Services')]"
        )

        self.log("Services menu clicked")

        self.click_element(
            "//*[contains(text(),'Returns')]"
        )

        self.log("Returns option clicked")

        self.click_element(
            "//*[contains(text(),'Returns Dashboard')]"
        )

        self.log("Returns Dashboard clicked")

        self.wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//label[contains(text(),'Financial Year')]"
                )
            )
        )

        self.log("Returns dashboard fully loaded")

    # -----------------------------------
    # SELECT RETURN PERIOD
    # -----------------------------------

    def select_return_period(
            self,
            financial_year,
            quarter,
            month
    ):

        self.select_dropdown(
            "//label[contains(text(),'Financial Year')]/following::select[1]",
            financial_year
        )

        self.select_dropdown(
            "//label[contains(text(),'Quarter')]/following::select[1]",
            quarter
        )

        self.select_dropdown(
            "//label[contains(text(),'Period')]/following::select[1]",
            month
        )

    # -----------------------------------
    # SEARCH RETURNS
    # -----------------------------------

    def search_returns(self):

        self.click_element(
            "//button[contains(text(),'Search')]"
        )

        self.log("Search button clicked")

        self.wait_for_loader_to_disappear()

        self.log("Search results loaded")

    # -----------------------------------
    # OPEN GSTR2B TILE
    # -----------------------------------

    def open_gstr2b(self):

        try:

            self.wait_for_loader_to_disappear()

            self.driver.execute_script(
                "window.scrollBy(0, 700);"
            )

            time.sleep(2)

            # -----------------------------
            # PRIORITY 1 → QUARTERLY VIEW
            # -----------------------------

            quarterly_tiles = self.driver.find_elements(
                By.XPATH,
                "//*[contains(text(),'Quarterly View')]"
            )

            for tile in quarterly_tiles:

                if tile.is_displayed():
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView({block: 'center'});",
                        tile
                    )

                    time.sleep(1)

                    self.driver.execute_script(
                        "arguments[0].click();",
                        tile
                    )

                    self.log(
                        "Quarterly View tile clicked"
                    )

                    self.wait_for_loader_to_disappear()

                    return True

            # -----------------------------
            # PRIORITY 2 → GSTR2B
            # -----------------------------

            gstr2b_tiles = self.driver.find_elements(
                By.XPATH,
                "//*[contains(text(),'GSTR-2B') or contains(text(),'GSTR2B')]"
            )

            for tile in gstr2b_tiles:

                if tile.is_displayed():
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView({block: 'center'});",
                        tile
                    )

                    time.sleep(1)

                    tile.click()

                    self.log(
                        "GSTR2B tile clicked"
                    )

                    self.wait_for_loader_to_disappear()

                    return True

            self.log(
                "No GSTR2B tile found"
            )

            return False

        except Exception as e:

            self.log(
                "Unable to open GSTR2B tile"
            )

            self.log(f"Error: {e}")

            return False

    # -----------------------------------
    # DOWNLOAD EXCEL
    # -----------------------------------

    def download_excel(self):

        try:

            self.wait_for_loader_to_disappear()

            self.driver.execute_script(
                "window.scrollTo(0, document.body.scrollHeight);"
            )

            self.log("Scrolled to bottom")

            page_source = self.driver.page_source.lower()

            error_messages = [
                "being generated",
                "not generated",
                "no record found",
                "no records found",
                "please try again later"
            ]

            if any(
                    message in page_source
                    for message in error_messages
            ):
                self.log(
                    "No downloadable GSTR2B available"
                )

                return False

            files_before_download = set(
                glob.glob(
                    os.path.join(
                        self.download_folder,
                        "*"
                    )
                )
            )

            button_found = self.click_download_button()

            time.sleep(5)

            if not button_found:
                self.log(
                    "Download button not found"
                )

                return False

            self.log(
                "Excel download button clicked"
            )

            current_page = self.driver.page_source.lower()

            if "being generated" in current_page:
                self.log(
                    "Excel generation started. Skipping for now."
                )

                return False

            downloaded_file = (
                self.wait_for_download(
                    files_before_download
                )
            )

            if downloaded_file:

                self.log(
                    "File downloaded successfully"
                )

                return True

            else:

                self.log(
                    "File download failed"
                )

                return False

        except Exception as e:

            self.log("Download failed")

            self.log(f"Error: {e}")

            return False

    # -----------------------------------
    # CLOSE DRIVER
    # -----------------------------------

    def close_driver(self):

        try:

            if self.driver is not None:
                self.driver.quit()

                self.log(
                    "Browser closed"
                )

        except Exception as e:

            self.log(
                f"Driver close failed: {e}"
            )


# -----------------------------------
# MAIN FUNCTION
# -----------------------------------

if __name__ == "__main__":

    temp_bot = GSTDownloader()

    results = []

    users = temp_bot.read_excel_users(
        "sample_users.xlsx"
    )

    for user in users:

        bot = GSTDownloader()

        username = user["username"]
        password = user["password"]

        print("\n")
        print("=" * 50)
        print(f"PROCESSING USER: {username}")
        print("=" * 50)

        status = "Success"
        details = "Downloaded Successfully"

        try:

            bot.setup_driver(username)

            bot.login(
                username,
                password
            )

            bot.handle_popups()

            bot.open_returns_dashboard()

            # PROCESS ENTIRE QUARTER

            result = bot.process_quarter(
                financial_year="2025-26",
                quarter="Q1"
            )

            # FINAL STATUS

            if result["success"] > 0:

                status = "Success"

                details = (
                    f"Downloaded: {result['success']} | "
                    f"Skipped: {result['skipped']} | "
                    f"Failed: {result['failed']}"
                )

            elif result["skipped"] > 0:

                status = "Skipped"

                details = (
                    f"No GSTR2B for "
                    f"{result['skipped']} month(s)"
                )

            else:

                status = "Failed"

                details = (
                    f"Failed Months: "
                    f"{result['failed']}"
                )
                details = "GSTR2B tile not available"

                print(
                    f"No GSTR2B available for user: {username}"
                )

            bot.close_driver()

        except Exception as e:

            status = "Failed"
            details = str(e)

            print(f"ERROR FOR USER {username}")

            print(e)

            bot.close_driver()

        # SAVE RESULT

        results.append({

            "Username": username,
            "Status": status,
            "Details": details,
            "Timestamp": datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        })

    # SAVE FINAL REPORT

    temp_bot.save_report(results)