import time
from mts_reports.base import MTSBase


class POStatusReport(MTSBase):
    def run(self):
        self.navigate_to_reports()
        ok = self.click_element("PO Status", [
            "//*[contains(text(),'PO Status')]",
            "//*[contains(text(),'Po Status')]",
            "//*[contains(text(),'PO STATUS')]",
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'po status')]",
        ])
        if not ok:
            print("Please click 'PO Status' manually. Waiting 30s...")
            time.sleep(30)
        self.download_excel()
