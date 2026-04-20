import os
import time
import random
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium_stealth import stealth


class MTSBase:
    LOGIN_URL = "https://mts.technoactive.in/"
    USERNAME  = "TAPL1"
    PASSWORD  = "Tapl@321"

    def __init__(self):
        self.download_dir = os.getcwd()
        self.driver = None

    # ── Driver ────────────────────────────────────────────────────────────────

    def build_driver(self):
        opts = Options()
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)
        opts.add_experimental_option("prefs", {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
        })
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()), options=opts
        )
        stealth(driver, languages=["en-US", "en"], vendor="Google Inc.",
                platform="Win32", webgl_vendor="Intel Inc.",
                renderer="Intel Iris OpenGL Engine", fix_hairline=True)
        return driver

    # ── Helpers ───────────────────────────────────────────────────────────────

    def human_type(self, el, text):
        for ch in text:
            el.send_keys(ch)
            time.sleep(random.uniform(0.08, 0.18))

    def _find_input(self, names, input_type="text"):
        self.driver.switch_to.default_content()

        def search(d):
            for name in names:
                try:
                    return d.find_element(By.NAME, name)
                except NoSuchElementException:
                    pass
            try:
                return d.find_element(By.XPATH, f"//input[@type='{input_type}'][1]")
            except NoSuchElementException:
                return None

        el = search(self.driver)
        if el:
            return el, "main"

        for i, iframe in enumerate(self.driver.find_elements(By.TAG_NAME, "iframe")):
            try:
                self.driver.switch_to.frame(iframe)
                el = search(self.driver)
                if el:
                    return el, f"iframe[{i}]"
                self.driver.switch_to.default_content()
            except Exception:
                self.driver.switch_to.default_content()

        return None, None

    def click_element(self, label, xpaths, wait_sec=4):
        print(f"[CLICK] Looking for: {label}")
        frames = [None] + list(range(len(self.driver.find_elements(By.TAG_NAME, "iframe"))))

        for frame in frames:
            self.driver.switch_to.default_content()
            if frame is not None:
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                if frame >= len(iframes):
                    continue
                try:
                    self.driver.switch_to.frame(iframes[frame])
                except Exception:
                    continue

            for xpath in xpaths:
                try:
                    el = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, xpath))
                    )
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", el)
                    time.sleep(0.3)
                    self.driver.execute_script("arguments[0].click();", el)
                    loc = "main" if frame is None else f"iframe[{frame}]"
                    print(f"[CLICK] '{label}' clicked in {loc} via: {xpath}")
                    self.driver.switch_to.default_content()
                    time.sleep(wait_sec)
                    return True
                except (TimeoutException, NoSuchElementException):
                    continue

        self.driver.switch_to.default_content()
        print(f"[CLICK] WARNING: '{label}' not found.")
        return False

    # ── Shared navigation steps ───────────────────────────────────────────────

    def login(self):
        print("[LOGIN] Opening portal...")
        self.driver.get(self.LOGIN_URL)
        time.sleep(4)

        user_field, frame_ref = self._find_input(("usr", "login", "username", "user"), "text")
        if not user_field:
            raise Exception("Cannot find username field.")

        print(f"[LOGIN] Username field found in {frame_ref}")
        user_field.clear()
        self.human_type(user_field, self.USERNAME)

        pwd_field = None
        for name in ("pwd", "password", "pass"):
            try:
                pwd_field = self.driver.find_element(By.NAME, name)
                break
            except NoSuchElementException:
                pass
        if not pwd_field:
            try:
                pwd_field = self.driver.find_element(By.XPATH, "//input[@type='password']")
            except NoSuchElementException:
                pass
        if not pwd_field:
            raise Exception("Cannot find password field.")

        pwd_field.clear()
        self.human_type(pwd_field, self.PASSWORD)

        pre_login_url = self.driver.current_url

        btn_clicked = False
        for sel in [
            (By.ID,    "sub_login"),
            (By.XPATH, "//input[@type='submit']"),
            (By.XPATH, "//button[@type='submit']"),
            (By.XPATH, "//input[@type='button']"),
            (By.XPATH, "//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'login')]"),
            (By.XPATH, "//button[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'login')]"),
            (By.XPATH, "//button"),
        ]:
            try:
                btn = self.driver.find_element(*sel)
                self.driver.execute_script("arguments[0].click();", btn)
                btn_clicked = True
                print(f"[LOGIN] Button clicked via JS ({sel[1]})")
                break
            except Exception:
                pass

        if not btn_clicked:
            print("[LOGIN] No button found — pressing Enter in password field")
            pwd_field.send_keys(Keys.RETURN)

        try:
            WebDriverWait(self.driver, 15).until(lambda d: d.current_url != pre_login_url)
            print(f"[LOGIN] Success! New URL: {self.driver.current_url}")
        except TimeoutException:
            print(f"[LOGIN] URL did not change. Still at: {self.driver.current_url}")

        self.driver.switch_to.default_content()
        time.sleep(3)

    def navigate_to_reports(self):
        self.login()

        ok = self.click_element("Planning & Schedule", [
            "//*[contains(text(),'Planning & Schedule')]",
            "//*[contains(text(),'Planning and Schedule')]",
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'planning')]",
        ])
        if not ok:
            print("Please click 'Planning & Schedule' manually. Waiting 30s...")
            time.sleep(30)

        ok = self.click_element("Reports", [
            "//*[contains(text(),'Reports')]",
            "//*[contains(text(),'REPORTS')]",
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'reports')]",
        ], wait_sec=2)
        if not ok:
            print("Please click 'Reports' manually. Waiting 20s...")
            time.sleep(20)

    def download_excel(self):
        ok = self.click_element("Excel button", [
            "//a[contains(@class,'buttons-excel')]",
            "//button[contains(@class,'buttons-excel')]",
            "//*[contains(text(),'Excel')]",
            "//*[contains(text(),'EXCEL')]",
            "//a[contains(@href,'.xls')]",
            "//button[contains(@title,'Excel')]",
        ])
        if not ok:
            print("Please click the Excel button manually. Waiting 20s...")
            time.sleep(20)

        self.click_element("Download button", [
            "//*[contains(text(),'Download')]",
            "//*[contains(text(),'DOWNLOAD')]",
            "//button[contains(@class,'download')]",
            "//a[contains(@class,'download')]",
        ], wait_sec=5)

    # ── Subclasses implement this ─────────────────────────────────────────────

    def run(self):
        raise NotImplementedError

    def execute(self):
        self.driver = self.build_driver()
        try:
            self.run()
            print(f"\n[DONE] {self.__class__.__name__} complete.")
        except Exception as e:
            print(f"[ERROR] {e}")
            self.driver.save_screenshot(os.path.join(self.download_dir, "error.png"))
        finally:
            print("Closing browser in 5s...")
            time.sleep(5)
            self.driver.quit()




