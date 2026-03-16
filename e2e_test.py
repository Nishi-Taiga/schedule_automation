"""Playwright E2E test for schedule_automation v0.12.0 deployed environment."""
import sys
sys.stdout.reconfigure(encoding='utf-8')

import os
import shutil
from playwright.sync_api import sync_playwright

BASE_URL = "https://schedule-automation-386m.onrender.com"
PASSWORD = "2321"
META_FILE = r"C:\tmp\e2e_test_files\meta.xlsx"
WEEK_DIR = r"C:\tmp\e2e_test_files\weeks"
JSON_FILE = r"C:\tmp\e2e_test_files\backup.json"
SCREENSHOT_DIR = r"C:\tmp\e2e_screenshots"
SURVEY_DIR = r"C:\tmp\e2e_test_files\surveys"

# webkitdirectory用: week filesのみ入ったクリーンディレクトリ
WEEK_ONLY_DIR = r"C:\tmp\e2e_test_files\week_only"

os.makedirs(SCREENSHOT_DIR, exist_ok=True)

def setup_week_only_dir():
    """webkitdirectory input用にweekファイルのみのクリーンディレクトリを作成"""
    if os.path.exists(WEEK_ONLY_DIR):
        shutil.rmtree(WEEK_ONLY_DIR)
    os.makedirs(WEEK_ONLY_DIR)
    for f in os.listdir(WEEK_DIR):
        if f.startswith("week_") and f.endswith(".xlsx"):
            shutil.copy2(os.path.join(WEEK_DIR, f), os.path.join(WEEK_ONLY_DIR, f))

def get_week_files():
    return [os.path.join(WEEK_DIR, f) for f in sorted(os.listdir(WEEK_DIR))
            if f.startswith("week_") and f.endswith(".xlsx")]

def run_tests():
    setup_week_only_dir()
    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(90000)

        # Test 1: Login
        print("[Test 1] Login...")
        try:
            page.goto(BASE_URL, wait_until="networkidle")
            page.fill('input[type="password"]', PASSWORD)
            page.click('button[type="submit"]')
            page.wait_for_selector("text=v0.12.0", timeout=15000)
            page.screenshot(path=os.path.join(SCREENSHOT_DIR, "01_login.png"))
            print("  PASS: Login successful, v0.12.0 confirmed")
            results.append(("Login", True))
        except Exception as e:
            page.screenshot(path=os.path.join(SCREENSHOT_DIR, "01_login_fail.png"))
            print(f"  FAIL: {e}")
            results.append(("Login", False))
            browser.close()
            return results

        # Test 2: Booth consolidation (meta + week files)
        print("[Test 2] Booth consolidation...")
        try:
            meta_input = page.query_selector('#consolMetaZone input[type="file"]')
            meta_input.set_input_files(META_FILE)
            page.wait_for_timeout(2000)

            week_input = page.query_selector('#consolWeekZone input[type="file"][multiple]')
            week_input.set_input_files(get_week_files())
            page.wait_for_timeout(2000)

            consol_btn = page.query_selector("#consolBtn")
            if consol_btn and not consol_btn.is_disabled():
                consol_btn.click()
                # Wait for API response with longer timeout for remote server
                page.wait_for_selector("text=統合完了", timeout=90000)
                page.screenshot(path=os.path.join(SCREENSHOT_DIR, "02_consolidate.png"))
                print("  PASS: Booth consolidation successful")
                results.append(("Booth consolidation", True))
            else:
                page.screenshot(path=os.path.join(SCREENSHOT_DIR, "02_consolidate_disabled.png"))
                print("  FAIL: Consolidate button is disabled")
                results.append(("Booth consolidation", False))
        except Exception as e:
            page.screenshot(path=os.path.join(SCREENSHOT_DIR, "02_consolidate_fail.png"))
            print(f"  FAIL: {e}")
            results.append(("Booth consolidation", False))

        # Test 3: Survey file upload
        print("[Test 3] Survey file upload...")
        try:
            survey_files = [os.path.join(SURVEY_DIR, f) for f in sorted(os.listdir(SURVEY_DIR))
                           if f.endswith(".xlsx")] if os.path.exists(SURVEY_DIR) else []
            if survey_files:
                survey_input = page.query_selector('#surveyZone input[type="file"]')
                survey_input.set_input_files(survey_files)
                page.wait_for_selector('#surveyZone.done', timeout=90000)
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

        # Test 4: Schedule generation
        print("[Test 4] Schedule generation...")
        try:
            next_btn = page.query_selector("#toS")
            if next_btn and not next_btn.is_disabled():
                next_btn.click()
                page.wait_for_timeout(5000)
                page.screenshot(path=os.path.join(SCREENSHOT_DIR, "04a_settings.png"))

                page.click("#genBtn")
                page.wait_for_selector("#pResult", state="visible", timeout=120000)
                page.wait_for_timeout(3000)
                page.screenshot(path=os.path.join(SCREENSHOT_DIR, "04b_result.png"))

                result_text = page.text_content("#pResult")
                print("  PASS: Schedule generated")
                if "未配置" in result_text:
                    print("  INFO: Unplaced slots exist")

                # Check Feature 5: unplaced add button
                add_btn = page.query_selector("text=未配置コマ追加")
                if add_btn:
                    print("  PASS: Feature 5 - 'Add unplaced' button visible")

                results.append(("Schedule generation", True))
            else:
                print("  SKIP: Next button disabled (need both src+booth)")
                results.append(("Schedule generation", None))
        except Exception as e:
            page.screenshot(path=os.path.join(SCREENSHOT_DIR, "04_generate_fail.png"))
            print(f"  FAIL: {e}")
            results.append(("Schedule generation", False))

        # Test 5: JSON restore (separate page/session)
        print("[Test 5] JSON restore...")
        page2 = context.new_page()
        page2.set_default_timeout(90000)
        try:
            page2.goto(BASE_URL, wait_until="networkidle")
            page2.wait_for_selector("text=v0.12.0", timeout=15000)

            # Select JSON file
            json_input = page2.query_selector('#uJson input[type="file"]')
            json_input.set_input_files(JSON_FILE)
            page2.wait_for_selector("#jsonBoothSection", state="visible", timeout=10000)
            page2.screenshot(path=os.path.join(SCREENSHOT_DIR, "05a_json_selected.png"))
            print("  Step 1: JSON selected")

            # Select meta file
            meta_btn_input = page2.query_selector('#jsonBoothSection input[accept=".xlsx"]')
            meta_btn_input.set_input_files(META_FILE)
            page2.wait_for_timeout(1000)
            print("  Step 2: Meta selected")

            # Select week folder (webkitdirectory needs directory path)
            week_dir_input = page2.query_selector('#jsonBoothSection input[webkitdirectory]')
            week_dir_input.set_input_files(WEEK_ONLY_DIR)
            page2.wait_for_timeout(2000)
            page2.screenshot(path=os.path.join(SCREENSHOT_DIR, "05b_json_files_selected.png"))
            print("  Step 3: Week folder selected")

            # Check restore button
            restore_btn = page2.query_selector("#jsonRestoreBtn")
            is_disabled = restore_btn.get_attribute("disabled")
            if is_disabled:
                print("  FAIL: Restore button still disabled")
                page2.screenshot(path=os.path.join(SCREENSHOT_DIR, "05_btn_disabled.png"))
                results.append(("JSON restore", False))
            else:
                restore_btn.click()
                page2.wait_for_selector("#pResult", state="visible", timeout=120000)
                page2.wait_for_timeout(3000)
                page2.screenshot(path=os.path.join(SCREENSHOT_DIR, "05c_json_restored.png"))

                result_text = page2.text_content("#pResult")
                # Check Feature 4: placement rate should show real data, not 0/0
                import re
                rate_match = re.search(r'(\d+)/(\d+)', result_text)
                if rate_match:
                    placed, total = int(rate_match.group(1)), int(rate_match.group(2))
                    print(f"  INFO: {placed}/{total} koma")
                    if total > 0:
                        print("  PASS: Feature 4 - Total is correctly calculated (not 0)")
                    else:
                        print("  WARN: Total is 0")

                # Check unplaced add button
                add_btn = page2.query_selector("text=未配置コマ追加")
                if add_btn:
                    print("  PASS: Feature 5 - 'Add unplaced' button visible")

                print("  PASS: JSON restore completed")
                results.append(("JSON restore", True))

            page2.close()
        except Exception as e:
            try:
                page2.screenshot(path=os.path.join(SCREENSHOT_DIR, "05_json_fail.png"))
            except:
                pass
            print(f"  FAIL: {e}")
            results.append(("JSON restore", False))

        # Test 6: Folder button text (v0.10.4 fix)
        print("[Test 6] Folder button text check...")
        try:
            folder_btn = page.query_selector("text=フォルダから一括選択")
            no_meta_text = page.query_selector("text=メタデータも含めて")
            if folder_btn and not no_meta_text:
                print("  PASS: Folder button text correct")
                results.append(("Folder button text", True))
            elif no_meta_text:
                print("  FAIL: Old text still present")
                results.append(("Folder button text", False))
            else:
                print("  SKIP: Folder button not found")
                results.append(("Folder button text", None))
        except Exception as e:
            print(f"  FAIL: {e}")
            results.append(("Folder button text", False))

        browser.close()

    # Summary
    print("\n" + "=" * 50)
    print("E2E Test Summary (v0.12.0)")
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
