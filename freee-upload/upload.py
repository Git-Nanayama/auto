import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError

load_dotenv()

def run(playwright):
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    try:
        # freeeにログイン
        page.goto("https://accounts.secure.freee.co.jp/login")

        # ラベルテキストを元に入力欄を特定し、入力する (より堅牢な方法)
        page.get_by_label("メールアドレス/ログインID").fill(os.getenv("FREEE_EMAIL"), timeout=15000)
        page.get_by_test_id("password-form").fill(os.getenv("FREEE_PASSWORD"))

        # 「ログイン」ボタンをクリック
        page.get_by_role('button', name='ログイン').click()

        # ログイン後のページ遷移を待つ
        page.wait_for_load_state("networkidle")

        # 「エクセルインポート」画面に直接アクセス
        import_url = "https://secure.freee.co.jp/spreadsheets/import"
        print(f"「{import_url}」に直接アクセスします。")
        page.goto(import_url)

        # ページが遷移・更新されるのを待つ
        page.wait_for_load_state("networkidle")

        # 「振替伝票（仕訳形式）のインポート」ラジオボタンを選択
        page.check("#import-data-type-form-manual")

        print("ファイルアップロードの直前で処理を停止しました。ブラウザを確認してください。")
        input("続行するにはEnterキーを押してください...")

    except TimeoutError as e:
        print(f"タイムアウトエラーが発生しました: {e}")
        print(f"現在のページのURL: {page.url}")
        print(f"現在のページのタイトル: {page.title()}")
        page.screenshot(path="error_screenshot.png")
        print("現在の画面のスクリーンショットを 'error_screenshot.png' として保存しました。")

    finally:
        context.close()
        browser.close()

with sync_playwright() as playwright:
    run(playwright)
