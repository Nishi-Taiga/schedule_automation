#!/usr/bin/env python3
"""
scheduler.htmlのPlaywrightテスト
生徒・講師情報.xlsxとブース表テンプレート.xlsxを読み込んで割り当てを実行
"""

import asyncio
import os
import subprocess
import time
from pathlib import Path
from playwright.async_api import async_playwright

async def test_scheduler():
    print("=" * 80)
    print("scheduler.html テスト開始")
    print("=" * 80)

    # ファイルパスを設定
    base_dir = Path("/home/user/schedule_automation")
    html_file = base_dir / "scheduler.html"
    booth_file = base_dir / "ブース表テンプレート.xlsx"
    teacher_file = base_dir / "生徒・講師情報.xlsx"

    # ファイルの存在確認
    print("\n✓ ファイルの存在確認:")
    for file_path in [html_file, booth_file, teacher_file]:
        exists = "✓" if file_path.exists() else "✗"
        print(f"  {exists} {file_path.name}")
        if not file_path.exists():
            raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")

    # HTTPサーバーをバックグラウンドで起動
    print("\n✓ HTTPサーバーを起動中...")
    server_process = subprocess.Popen(
        ["python3", "-m", "http.server", "8888"],
        cwd=str(base_dir),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
    time.sleep(2)  # サーバー起動を待つ

    try:
        async with async_playwright() as p:
            # ブラウザを起動（ヘッドレスモード）
            print("✓ ブラウザを起動中...")
            browser = await p.chromium.launch(
                headless=True,
                args=[
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-blink-features=AutomationControlled',
                    '--disable-extensions',
                    '--disable-gpu',
                    '--single-process',
                    '--no-zygote',
                    '--js-flags=--max-old-space-size=4096',
                    '--ignore-certificate-errors'
                ]
            )
            context = await browser.new_context(
                ignore_https_errors=True
            )
            page = await context.new_page()

            # ページのメモリ制限を緩和
            await page.set_viewport_size({"width": 1920, "height": 1080})

            # コンソールログをキャプチャ
            console_messages = []
            page.on("console", lambda msg: console_messages.append(f"[{msg.type}] {msg.text}"))

            # HTMLファイルを開く
            print(f"\n✓ {html_file.name} を開いています...")
            await page.goto("http://localhost:8888/scheduler.html", timeout=30000)
            await page.wait_for_load_state("networkidle", timeout=30000)

            # タイトルを確認
            title = await page.title()
            print(f"  ページタイトル: {title}")

            # ExcelJSライブラリが読み込まれるまで待つ
            print("  ExcelJSライブラリの読み込みを待機中...")
            try:
                await page.wait_for_function("typeof ExcelJS !== 'undefined'", timeout=30000)
                print("  ✓ ExcelJSライブラリが読み込まれました")
            except Exception as e:
                print(f"  ⚠️ ExcelJSの読み込みに失敗しました: {e}")
                print("  コンソールログ:")
                for msg in console_messages:
                    print(f"    {msg}")
                raise

            # ブース表を読み込み
            print(f"\n✓ ブース表を読み込み中...")
            booth_input = await page.query_selector("#boothExcelFile")
            await booth_input.set_input_files(str(booth_file))

            # 読み込み完了を待機（ステータスメッセージを確認）
            try:
                await page.wait_for_function(
                    "document.getElementById('boothExcelStatus').textContent.includes('読み込み完了') || document.getElementById('boothExcelStatus').textContent.includes('エラー')",
                    timeout=30000
                )
                booth_status = await page.locator("#boothExcelStatus").inner_text()
                print(f"  {booth_status}")
            except Exception as e:
                print(f"  ⚠️  タイムアウト: {e}")
                # 現在のステータスを確認
                booth_status = await page.locator("#boothExcelStatus").inner_text()
                print(f"  現在のステータス: {booth_status}")
                # コンソールログを表示
                print(f"  コンソールログ:")
                for msg in console_messages[-10:]:
                    print(f"    {msg}")
                raise

            # 生徒・講師情報を読み込み
            print(f"\n✓ 生徒・講師情報を読み込み中...")
            teacher_input = await page.query_selector("#teacherExcelFile")
            await teacher_input.set_input_files(str(teacher_file))

            # 読み込み完了を待機
            await page.wait_for_function(
                "document.getElementById('teacherExcelStatus').textContent.includes('読み込み完了')",
                timeout=30000
            )
            teacher_status = await page.locator("#teacherExcelStatus").inner_text()
            print(f"  {teacher_status}")

            # 割り当てボタンをクリック
            print(f"\n✓ 講師の割り当てを実行中...")
            run_button = await page.query_selector("#runBtn")
            await run_button.click()

            # 割り当て完了を待機（最大60秒）
            try:
                await page.wait_for_function(
                    """document.getElementById('status').textContent.includes('完了') ||
                       document.getElementById('status').textContent.includes('エラー')""",
                    timeout=60000
                )
            except Exception as e:
                print(f"  ⚠️ タイムアウト: {e}")

            # 結果ステータスを取得
            status = await page.locator("#status").inner_text()
            print(f"  {status}")

            # 生徒の割り当て
            print(f"\n✓ 生徒の割り当てを実行中...")

            # 生徒選択UIが表示されるまで待機
            try:
                await page.wait_for_selector("#studentSelectionCard", state="visible", timeout=5000)
                print(f"  生徒選択UIが表示されました")
            except Exception as e:
                print(f"  ⚠️ 生徒選択UIが表示されませんでした: {e}")

            # 生徒のプルダウンから全ての生徒を取得
            student_select = await page.query_selector("#studentSelect")
            student_options = await student_select.query_selector_all("option")

            students = []
            for option in student_options:
                value = await option.get_attribute("value")
                if value:  # 空のオプションを除外
                    students.append(value)

            print(f"  生徒数: {len(students)}人")

            # 各生徒を割り当て（最大5人をテスト）
            test_student_count = min(5, len(students))
            print(f"  テスト対象: 最初の{test_student_count}人")

            for i, student_id in enumerate(students[:test_student_count], 1):
                print(f"\n  [{i}/{test_student_count}] {student_id} を割り当て中...")

                # プルダウンから生徒を選択
                await student_select.select_option(value=student_id)
                await page.wait_for_timeout(200)

                # 割り当てボタンをクリック
                assign_btn = await page.query_selector("#assignStudentBtn")
                await assign_btn.click()

                # 割り当て完了を待機（ステータスメッセージが更新されるまで）
                try:
                    await page.wait_for_function(
                        f"""document.getElementById('status').textContent.includes('{student_id}')""",
                        timeout=10000
                    )
                    result_status = await page.locator("#status").inner_text()
                    print(f"    ✓ {result_status}")
                except Exception as e:
                    print(f"    ⚠️ タイムアウト: {e}")

                await page.wait_for_timeout(500)

            # CSV出力を確認
            print(f"\n✓ 割り当て結果を確認中...")
            assignments_out = await page.locator("#assignmentsOut").input_value()
            lines = assignments_out.strip().split('\n')
            print(f"  CSV行数: {len(lines)}行")
            if len(lines) > 1:
                print(f"  ヘッダー: {lines[0]}")
                print(f"  サンプル (最初の3件):")
                for line in lines[1:4]:
                    print(f"    {line}")

            # レポートを確認
            report_out = await page.locator("#reportOut").input_value()
            if report_out:
                report_lines = report_out.strip().split('\n')
                print(f"\n✓ 検知レポート:")
                print(f"  レポート行数: {len(report_lines)}行")
                if len(report_lines) > 1:
                    print(f"  警告/エラー (最初の5件):")
                    for line in report_lines[1:6]:
                        print(f"    {line}")

            # 必要コマ数確認表を表示
            print(f"\n✓ 必要コマ数確認表を表示中...")
            toggle_required_btn = await page.query_selector("#toggleRequiredClassBtn")
            await toggle_required_btn.click()
            await page.wait_for_timeout(500)

            # テーブルの内容を取得
            required_table = await page.query_selector("#requiredClassTable")
            if required_table:
                rows = await required_table.query_selector_all("tbody tr")
                print(f"  生徒数: {len(rows) - 1}人")  # サマリー行を除く

                # 最初の5人を表示
                print(f"  サンプル (最初の5人):")
                for i, row in enumerate(rows[:5]):
                    cells = await row.query_selector_all("td")
                    if cells and len(cells) >= 5:
                        student_name = await cells[0].inner_text()
                        subject_info = await cells[1].inner_text()
                        required = await cells[2].inner_text()
                        assigned = await cells[3].inner_text()
                        rate = await cells[4].inner_text()
                        print(f"    {student_name}: {subject_info} | 必要:{required} 割当:{assigned} 達成率:{rate}")

            # エラーログを確認
            errors = [msg for msg in console_messages if 'error' in msg.lower() or '❌' in msg]
            if errors:
                print(f"\n⚠️ エラーログ ({len(errors)}件):")
                for error in errors[:10]:  # 最大10件まで表示
                    print(f"  {error}")

            # スクリーンショットを保存
            screenshot_path = base_dir / "test_result_screenshot.png"
            await page.screenshot(path=str(screenshot_path), full_page=True)
            print(f"\n✓ スクリーンショットを保存: {screenshot_path.name}")

            # ブラウザを閉じる
            await browser.close()

    finally:
        # HTTPサーバーを停止
        print("\n✓ HTTPサーバーを停止中...")
        server_process.terminate()
        server_process.wait()

    print("\n" + "=" * 80)
    print("✓ テスト完了")
    print("=" * 80)

if __name__ == "__main__":
    asyncio.run(test_scheduler())
