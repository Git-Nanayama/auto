import os
import argparse
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# .envファイルから環境変数を読み込む
load_dotenv()

def run(playwright, file_path: Path):
    """
    freeeにログインし、指定されたファイルをアップロードする
    """
    # --- 事前チェック ---
    FREEE_EMAIL = os.getenv('FREEE_EMAIL')
    FREEE_PASSWORD = os.getenv('FREEE_PASSWORD')
    if not FREEE_EMAIL or not FREEE_PASSWORD or FREEE_EMAIL == "your_email@example.com":
        print("エラー: .envファイルが正しく設定されていません。")
        print(f"'{Path(__file__).parent / '.env'}' ファイルを開き、メールアドレスとパスワードを正しく設定してください。")
        return

    # CI環境ではヘッドレス、ローカルではヘッドフル（デバッグしやすいように）で実行
    is_ci = os.getenv('CI') == 'true'
    print(f"CI環境として実行: {is_ci}")

    launch_options = {'headless': is_ci}
    if not is_ci:
        launch_options['slow_mo'] = 500

    browser = playwright.chromium.launch(**launch_options)
    context = browser.new_context()
    page = context.new_page()

    try:
        # 1. freeeにログイン
        print("1. freeeにログインします...")
        page.goto("https://accounts.secure.freee.co.jp/login")

        page.get_by_label("メールアドレス/ログインID").fill(FREEE_EMAIL, timeout=15000)
        page.get_by_test_id("password-form").fill(FREEE_PASSWORD)
        page.get_by_role('button', name='ログイン').click()
        page.wait_for_load_state("networkidle")
        print("ログイン成功。")

        # 2. 「エクセルインポート」画面に直接アクセス
        import_url = "https://secure.freee.co.jp/spreadsheets/import"
        print(f"2. 「{import_url}」に直接アクセスします。")
        page.goto(import_url)
        page.wait_for_load_state("networkidle")

        # 3. 「振替伝票（仕訳形式）のインポート」ラジオボタンを選択
        print("3. インポート形式として「振替伝票（仕訳形式）」を選択します。")
        page.check("#import-data-type-form-manual")

        # 4. ファイルをアップロード
        print(f"4. ファイル '{file_path.name}' をアップロードします...")
        # ファイル選択のinput要素にファイルをセット
        # このセレクタはfreeeのHTML構造に依存します。'input[type="file"]'は一般的なセレクタです。
        file_input_selector = 'input[type="file"]'
        page.locator(file_input_selector).set_input_files(file_path)

        # アップロードボタンをクリック
        # ボタンのテキストが「アップロード」と仮定しています。
        upload_button = page.get_by_role("button", name="アップロード")
        upload_button.wait_for(state="visible")
        upload_button.click()
        print("アップロードボタンをクリックしました。")

        # 5. アップロード完了を待機
        print("5. アップロード処理の完了を待機します...")
        # 完了後のページ遷移や完了メッセージを待つのがより堅牢です。
        # ここではネットワークが落ち着くまで最大2分間待機します。
        page.wait_for_load_state("networkidle", timeout=120000)
        print("処理が完了しました。")

    except PlaywrightTimeoutError as e:
        print(f"タイムアウトエラーが発生しました: {e}")
        print(f"現在のページのURL: {page.url}")
        print(f"現在のページのタイトル: {page.title()}")
        screenshot_path = Path(__file__).parent / "error_screenshot.png"
        page.screenshot(path=str(screenshot_path))
        print(f"現在の画面のスクリーンショットを '{screenshot_path}' として保存しました。")
    except Exception as e:
        print(f"予期せぬエラーが発生しました: {e}")
        screenshot_path = Path(__file__).parent / "error_screenshot.png"
        page.screenshot(path=str(screenshot_path))
        print(f"現在の画面のスクリーンショットを '{screenshot_path}' として保存しました。")
    finally:
        print("\nブラウザを閉じます。")
        context.close()
        browser.close()

def main():
    """
    コマンドライン引数を解析し、アップロード処理を実行する
    """
    parser = argparse.ArgumentParser(description="freeeにファイルをアップロードするスクリプト。")
    parser.add_argument("--file", required=True, help="アップロードするCSVファイルのパス。")
    args = parser.parse_args()

    file_path = Path(args.file)
    if not file_path.is_file():
        print(f"エラー: 指定されたファイルが見つかりません: {file_path}")
        return

    with sync_playwright() as playwright:
        run(playwright, file_path)

if __name__ == "__main__":
    main()
