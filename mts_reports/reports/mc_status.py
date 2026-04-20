import time
from mts_reports.base import MTSBase


class MCStatusReport(MTSBase):
    def run(self):
        self.navigate_to_reports()
        ok = self.click_element("M/C Status", [
            "//*[contains(text(),'M/C Status')]",
            "//*[contains(text(),'MC Status')]",
            "//*[contains(text(),'Machine Status')]",
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'m/c status')]",
        ])
        if not ok:
            print("Please click 'M/C Status' manually. Waiting 30s...")
            time.sleep(30)
        self.download_excel()
