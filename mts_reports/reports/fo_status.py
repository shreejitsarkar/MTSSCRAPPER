import time
from mts_reports.base import MTSBase


class FOStatusReport(MTSBase):
    def run(self):
        self.navigate_to_reports()
        ok = self.click_element("FO Status", [
            "//*[contains(text(),'FO Status')]",
            "//*[contains(text(),'Fo Status')]",
            "//*[contains(text(),'FO STATUS')]",
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'fo status')]",
        ])
        if not ok:
            print("Please click 'FO Status' manually. Waiting 30s...")
            time.sleep(30)
        self.download_excel()
