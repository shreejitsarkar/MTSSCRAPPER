import os
from mts_reports.reports.mc_status import MCStatusReport
from mts_reports.reports.fo_status import FOStatusReport
from mts_reports.reports.po_status import POStatusReport


def delete_existing_excel_files():
    cwd = os.getcwd()
    deleted = 0
    for f in os.listdir(cwd):
        if f.lower().endswith((".xls", ".xlsx")):
            os.remove(os.path.join(cwd, f))
            print(f"[CLEANUP] Deleted: {f}")
            deleted += 1
    print(f"[CLEANUP] {deleted} Excel file(s) removed from {cwd}")
def rename():
    cwd = os.getcwd()
    for f in os.listdir(cwd):
        print(f)
        
        if f.lower().startswith("po"):
            print(f)

            src = os.path.join(cwd, f)
            dst = os.path.join(cwd, "Po_Status.xlsx")  # include filename

            os.rename(src, dst)
        if f.lower().startswith("fo"):
            print(f)

            src = os.path.join(cwd, f)
            dst = os.path.join(cwd, "Fo_Status.xlsx")  # include filename

            os.rename(src, dst)
        if f.lower().startswith("mts_mc"):
            print(f)

            src = os.path.join(cwd, f)
            dst = os.path.join(cwd, "mts_mc.xlsx")  # include filename

            os.rename(src, dst)


if __name__ == "__main__":
    delete_existing_excel_files()

    reports = [MCStatusReport, FOStatusReport, POStatusReport]
    for ReportClass in reports:
        print(f"\n{'='*50}")
        print(f"  Running: {ReportClass.__name__}")
        print(f"{'='*50}")
        ReportClass().execute()
    rename()
    
