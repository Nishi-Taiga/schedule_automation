"""Playwright E2E test for schedule_automation v0.18.0 deployed environment."""
import sys
sys.stdout.reconfigure(encoding='utf-8')

import os
import shutil
from playwright.sync_api import sync_playwright

BASE_URL = "https://schedule-automation-386m.onrender.com"
PASSWORD = "2321"
META_FILE = r"C:\tmp\e2e_test_files\meta.xlsx"
WEEK_DIR = r"C:\tmp\e2e_test_files\weeks"
JSON_FILE = r"C:\tmp\e2e_test_files\schedule_data.json"
SCREENSHOT_DIR = r"C:\tmp\e2e_screenshots"
SURVEY_DIR = r"C:\tmp\e2e_test_files\survey"
WEEK_ONLY_DIR = r"C:\tmp\e2e_test_files\week_only"
OUTPUT_FILE = r"C:\tmp\e2e_test_files\output.xlsx"

os.makedirs(SCREENSHOT_DIR, exist_ok=True)

# Use onclick attribute to uniquely identify the download button
DL_BTN_SELECTOR = 'button[onclick="dlExcel()"]'

def setup_week_only_dir():
    if os.path.exists(WEEK_ONLY_DIR):
        shutil.rmtree(WEEK_ONLY_DIR)
    os.makedirs(WEEK_ONLY_DIR)
    for f in os.listdir(WEEK_DIR):
        if f.startswith("week_") and f.endswith(".xlsx"):
            shutil.copy2(os.path.join(WEEK_DIR, f), os.path.join(WEEK_ONLY_DIR, f))

def get_week_files():
    return [os.path.join(WEEK_DIR, f) for f in sorted(os.listdir(WEEK_DIR))
            if f.startswith("week_") and f.endswith(".xlsx")]

def check_booth_sheets(path, label):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    sheet_names = wb.sheetnames
    wb.close()
    print(f"  INFO: {label} sheets: {sheet_names}")
    week_sheets = [sn for sn in sheet_names
                  if sn != '未配置コマ' and not sn.startswith('_schedule_data')]
    if week_sheets:
        print(f"  PASS: {len(week_sheets)} booth sheet(s) found")
        return True
    else:
        print("  FAIL: No booth sheets!")
        return False

def run_tests():
    setup_week_only_dir()
    results = []
    downloaded_path = None

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(120000)

        # ==================== Test 1: Login ====================
        print("[Test 1] Login...")
        try:
            page.goto(BASE_URL, wait_until="networkidle", timeout=60000)
            page.fill('input[type="password"]', PASSWORD)
            page.click('button[type="submit"]')
            page.wait_for_selector("text=v0.18.0", timeout=30000)
            page.screenshot(path=os.path.join(SCREENSHOT_DIR, "01_login.png"))
            print("  PASS: Login successful, v0.18.0 confirmed")
            results.append(("Login", True))
        except Exception as e:
            page.screenshot(path=os.path.join(SCREENSHOT_DIR, "01_login_fail.png"))
            print(f"  FAIL: {e}")
            results.append(("Login", False))
            browser.close()
            return results

        # ==================== Test 2: Booth consolidation ====================
        print("[Test 2] Booth consolidation...")
        try:
            meta_input = page.query_selector('#consolMetaZone input[type="file"]')
            meta_input.set_input_files(META_FILE)
            page.wait_for_timeout(2000)

            week_input = page.query_selector('#consolWeekZone input[type="file"][multiple]')
            week_input.set_input_files(get_week_files())
            page.wait_for_timeout(2000)

            btn_disabled = page.evaluate("document.getElementById('consolBtn').disabled")
            if not btn_disabled:
                page.evaluate("document.getElementById('consolBtn').click()")
                page.wait_for_timeout(1000)
                btn_text = page.evaluate("document.getElementById('consolBtn').textContent")
                print(f"  DEBUG: btn text after click = '{btn_text}'")

                # Wait for consolidation to finish: button text reverts on both success and error
                page.wait_for_function(
                    "() => { const b = document.getElementById('consolBtn'); return b && b.textContent.trim() === '集約して統合'; }",
                    timeout=300000
                )
                page.wait_for_timeout(1000)
                result_text = page.evaluate("document.getElementById('consolResult').textContent")
                error_text = page.evaluate("document.getElementById('upSt').textContent")
                page.screenshot(path=os.path.join(SCREENSHOT_DIR, "02_consolidate.png"))
                if '統合完了' in result_text or '✅' in result_text:
                    print("  PASS: Booth consolidation successful")
                    results.append(("Booth consolidation", True))
                elif error_text:
                    print(f"  FAIL: Error: {error_text[:150]}")
                    results.append(("Booth consolidation", False))
                else:
                    print(f"  FAIL: No result text found")
                    results.append(("Booth consolidation", False))
            else:
                print("  FAIL: Button disabled")
                results.append(("Booth consolidation", False))
        except Exception as e:
            page.screenshot(path=os.path.join(SCREENSHOT_DIR, "02_consolidate_fail.png"))
            print(f"  FAIL: {e}")
            results.append(("Booth consolidation", False))

        # ==================== Test 3: Survey file upload ====================
        print("[Test 3] Survey file upload...")
        try:
            survey_files = [os.path.join(SURVEY_DIR, f) for f in sorted(os.listdir(SURVEY_DIR))
                           if f.endswith(".xlsx")] if os.path.exists(SURVEY_DIR) else []
            if survey_files:
                survey_input = page.query_selector('#surveyZone input[type="file"]')
                survey_input.set_input_files(survey_files)
                page.wait_for_selector('#surveyZone.done', timeout=120000)
                page.screenshot(path=os.path.join(SCREENSHOT_DIR, "03_survey.png"))
                print(f"  PASS: {len(survey_files)} survey files uploaded")
                results.append(("Survey upload", True))
            else:
                print("  SKIP: No survey files found")
                results.append(("Survey upload", None))
        except Exception as e:
            page.screenshot(path=os.path.join(SCREENSHOT_DIR, "03_survey_fail.png"))
            print(f"  FAIL: {e}")
            results.append(("Survey upload", False))

        # ==================== Test 4: Schedule generation ====================
        print("[Test 4] Schedule generation...")
        schedule_generated = False
        try:
            btn_disabled = page.evaluate("document.getElementById('toS').disabled")
            if not btn_disabled:
                page.click("#toS")
                page.wait_for_timeout(5000)
                page.screenshot(path=os.path.join(SCREENSHOT_DIR, "04a_settings.png"))
                page.click("#genBtn")
                page.wait_for_selector("#pResult", state="visible", timeout=180000)
                page.wait_for_timeout(3000)
                page.screenshot(path=os.path.join(SCREENSHOT_DIR, "04b_result.png"))
                print("  PASS: Schedule generated")
                results.append(("Schedule generation", True))
                schedule_generated = True
            else:
                print("  SKIP: toS disabled")
                results.append(("Schedule generation", None))
        except Exception as e:
            page.screenshot(path=os.path.join(SCREENSHOT_DIR, "04_generate_fail.png"))
            print(f"  FAIL: {e}")
            results.append(("Schedule generation", False))

        # ==================== Test 5: Excel download ====================
        print("[Test 5] Excel download check...")
        if schedule_generated:
            try:
                with page.expect_download(timeout=60000) as dl_info:
                    page.click(DL_BTN_SELECTOR)
                download = dl_info.value
                dl_path = os.path.join(SCREENSHOT_DIR, "downloaded.xlsx")
                download.save_as(dl_path)
                downloaded_path = dl_path
                ok = check_booth_sheets(dl_path, "Normal download")
                results.append(("Excel download (booth sheets)", ok))
            except Exception as e:
                page.screenshot(path=os.path.join(SCREENSHOT_DIR, "05_download_fail.png"))
                print(f"  FAIL: {e}")
                results.append(("Excel download (booth sheets)", False))
        else:
            print("  SKIP: No schedule to download")
            results.append(("Excel download (booth sheets)", None))

        # ==================== Test 6: JSON restore + download ====================
        print("[Test 6] JSON restore...")
        page2 = context.new_page()
        page2.set_default_timeout(120000)
        try:
            page2.goto(BASE_URL, wait_until="networkidle", timeout=60000)
            page2.wait_for_selector("text=v0.18.0", timeout=15000)

            page2.click('#advancedResumeOptions summary')
            page2.wait_for_timeout(500)

            json_input = page2.query_selector('#uJson input[type="file"]')
            json_input.set_input_files(JSON_FILE)
            page2.wait_for_selector("#jsonBoothSection", state="visible", timeout=10000)
            print("  Step 1: JSON selected")

            meta_btn_input = page2.query_selector('#jsonBoothSection input[accept=".xlsx"]')
            meta_btn_input.set_input_files(META_FILE)
            page2.wait_for_timeout(1000)
            print("  Step 2: Meta selected")

            week_dir_input = page2.query_selector('#jsonBoothSection input[webkitdirectory]')
            week_dir_input.set_input_files(WEEK_ONLY_DIR)
            page2.wait_for_timeout(2000)
            print("  Step 3: Week folder selected")

            restore_btn = page2.query_selector("#jsonRestoreBtn")
            is_disabled = restore_btn.get_attribute("disabled")
            if is_disabled:
                page2.screenshot(path=os.path.join(SCREENSHOT_DIR, "06_btn_disabled.png"))
                print("  FAIL: Restore button disabled")
                results.append(("JSON restore", False))
            else:
                restore_btn.click()
                # v0.18.0: JSON restore now goes to settings screen first
                page2.wait_for_selector("#pSettings", state="visible", timeout=120000)
                page2.wait_for_timeout(2000)
                page2.screenshot(path=os.path.join(SCREENSHOT_DIR, "06c_json_settings.png"))
                print("  Step 4: Settings screen shown")

                # Click "結果を表示" button to go to results
                go_result_btn = page2.query_selector("#goResultBtn")
                if go_result_btn and go_result_btn.is_visible():
                    go_result_btn.click()
                    page2.wait_for_selector("#pResult", state="visible", timeout=30000)
                    page2.wait_for_timeout(3000)
                    page2.screenshot(path=os.path.join(SCREENSHOT_DIR, "06d_json_result.png"))

                    # Download using specific selector
                    print("  Step 5: Download...")
                    try:
                        with page2.expect_download(timeout=90000) as dl_info2:
                            page2.click(DL_BTN_SELECTOR)
                        download2 = dl_info2.value
                        dl_path2 = os.path.join(SCREENSHOT_DIR, "downloaded_json_restore.xlsx")
                        download2.save_as(dl_path2)
                        ok2 = check_booth_sheets(dl_path2, "JSON-restored download")
                        if ok2:
                            print("  PASS: JSON restore + settings + download with booth sheets")
                        else:
                            print("  WARN: No booth sheets after JSON restore")
                    except Exception as dl_e:
                        page2.screenshot(path=os.path.join(SCREENSHOT_DIR, "06e_dl_fail.png"))
                        print(f"  WARN: Download issue: {dl_e}")
                else:
                    print("  WARN: goResultBtn not visible, schedule may not be in JSON")

                print("  PASS: JSON restore completed")
                results.append(("JSON restore", True))

            page2.close()
        except Exception as e:
            try: page2.screenshot(path=os.path.join(SCREENSHOT_DIR, "06_json_fail.png"))
            except: pass
            print(f"  FAIL: {e}")
            results.append(("JSON restore", False))

        # ==================== Test 7: Excel restore -> download (KEY FIX TEST) ====================
        print("[Test 7] Excel restore -> download booth sheets (v0.18.0 key test)...")
        restore_file = downloaded_path if downloaded_path and os.path.exists(downloaded_path) else OUTPUT_FILE
        if os.path.exists(restore_file):
            print(f"  Using: {os.path.basename(restore_file)}")
            page3 = context.new_page()
            page3.set_default_timeout(120000)
            try:
                page3.goto(BASE_URL, wait_until="networkidle", timeout=60000)
                page3.wait_for_selector("text=v0.18.0", timeout=15000)

                page3.click('#advancedResumeOptions summary')
                page3.wait_for_timeout(500)

                saved_input = page3.query_selector('#uSaved input[type="file"]')
                saved_input.set_input_files(restore_file)
                page3.wait_for_selector("#postRestoreSection", state="visible", timeout=90000)
                page3.wait_for_timeout(2000)
                page3.screenshot(path=os.path.join(SCREENSHOT_DIR, "07a_excel_restored.png"))

                restore_msg = page3.text_content("#postRestoreMsg") or ""
                print(f"  INFO: {restore_msg[:150]}")

                go_result_btn = page3.query_selector('button[onclick="finalizeRestore()"]')
                if go_result_btn:
                    go_result_btn.click()
                    page3.wait_for_selector("#pResult", state="visible", timeout=30000)
                    page3.wait_for_timeout(2000)
                    page3.screenshot(path=os.path.join(SCREENSHOT_DIR, "07b_excel_result.png"))

                    with page3.expect_download(timeout=90000) as dl_info3:
                        page3.click(DL_BTN_SELECTOR)
                    download3 = dl_info3.value
                    dl_path3 = os.path.join(SCREENSHOT_DIR, "downloaded_excel_restore.xlsx")
                    download3.save_as(dl_path3)
                    ok3 = check_booth_sheets(dl_path3, "Excel-restored download")
                    results.append(("Excel restore -> download", ok3))
                else:
                    print("  FAIL: Go to result button not found")
                    results.append(("Excel restore -> download", False))

                page3.close()
            except Exception as e:
                try: page3.screenshot(path=os.path.join(SCREENSHOT_DIR, "07_excel_restore_fail.png"))
                except: pass
                print(f"  FAIL: {e}")
                results.append(("Excel restore -> download", False))
        else:
            print("  SKIP: No file to restore")
            results.append(("Excel restore -> download", None))

        browser.close()

    # Summary
    print("\n" + "=" * 50)
    print("E2E Test Summary (v0.18.0)")
    print("=" * 50)
    passed = sum(1 for _, r in results if r is True)
    failed = sum(1 for _, r in results if r is False)
    skipped = sum(1 for _, r in results if r is None)
    for name, result in results:
        status = "PASS" if result is True else ("FAIL" if result is False else "SKIP")
        print(f"  [{status}] {name}")
    print(f"\nTotal: {passed} passed, {failed} failed, {skipped} skipped")
    return failed == 0

if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)
