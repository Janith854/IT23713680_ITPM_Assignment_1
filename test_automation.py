from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os
import argparse
import re
from pathlib import Path
import sys
import openpyxl

def _configure_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="ignore")
    except:
        pass

def _find_column(header, name):
    normalized_name = re.sub(r"[^a-z0-9]+", "", str(name).strip().lower())
    for i, col in enumerate(header):
        if re.sub(r"[^a-z0-9]+", "", str(col).strip().lower()) == normalized_name:
            return i + 1
    return None

def run_test():
    _configure_stdout()

    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", default="IT23713680.xlsx")
    parser.add_argument("--url", default="https://www.pixelssuite.com/transliteration")
    args = parser.parse_args()

    if not os.path.exists(args.excel):
        print(f"File not found: {args.excel}")
        return

    wb = openpyxl.load_workbook(args.excel)
    ws = wb["Test cases"] if "Test cases" in wb.sheetnames else wb.active
    header = [cell.value for cell in ws[1]]

    input_col = _find_column(header, "Input")
    expected_col = _find_column(header, "Expected output")
    actual_col = _find_column(header, "Actual output")
    status_col = _find_column(header, "Status")

    with sync_playwright() as p:
        # ඉන්ටර්නෙට් හොඳ නිසා slow_mo එක 100 දක්වා අඩු කළා
        browser = p.chromium.launch(headless=False, slow_mo=30)
        page = browser.new_page()
        page.goto(args.url)
        page.wait_for_load_state("networkidle")
        
        input_box = page.get_by_placeholder("Input Your Singlish Text Here.").or_(page.locator("textarea").first)
        output_box = page.locator("textarea").nth(1)
        
        for i in range(2, ws.max_row + 1):
            input_text = ws.cell(i, input_col).value
            if not input_text:
                continue

            print(f"Processing Row {i}: {input_text}")

            input_box.click()
            input_box.fill("")
            input_box.fill(str(input_text).strip())
            
            translate_btn = page.get_by_role("button", name=re.compile("Translate", re.I))
            translate_btn.click(force=True)

            # Wait time එක තත්පර 8ක් දක්වා අඩු කළා
            page.wait_for_timeout(2000) 

            try:
                actual = output_box.input_value().strip()
                if not actual:
                    actual = output_box.inner_text().strip()
            except:
                actual = "ERROR"

            ws.cell(i, actual_col).value = actual
            expected = str(ws.cell(i, expected_col).value or "").strip()
            status = "PASS" if actual == expected else "FAIL"
            ws.cell(i, status_col).value = status
            
            print(f"  -> {status}")
            wb.save(args.excel)

        browser.close()
        print("\nSuccess! Fast mode update completed.")

if __name__ == "__main__":
    run_test()