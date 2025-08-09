import os
import shutil
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from dotenv import load_dotenv

# .envファイルから環境変数を読み込む
load_dotenv()

# --- 設定項目 ---
# 環境変数からログイン情報を取得
TENKURA_EMAIL = os.getenv('TENKURA_EMAIL')
TENKURA_PASSWORD = os.getenv('TENKURA_PASSWORD')

# URL
LOGIN_URL = "https://app.cloud-tenkura.net/home"

# 保存先フォルダ (sales/downloaded_data)
OUTPUT_DIR = Path("sales") / "downloaded_data"

# --- メイン処理 ---
def main():
    """
    天の蔵にログインし、売上・仕入データをダウンロードして指定のフォルダに移動する
    """
    # --- 出力ディレクトリの作成 ---
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"データは '{OUTPUT_DIR.resolve()}' に保存されます。")

    # --- 事前チェック ---
    if not TENKURA_EMAIL or not TENKURA_PASSWORD or TENKURA_EMAIL == "your_email@example.com":
        print("エラー: .envファイルが正しく設定されていません。")
        print(f"'{Path(__file__).parent / '.env'}' ファイルを開き、メールアドレスとパスワードを正しく設定してください。")
        return

    with sync_playwright() as p:
        # CI環境ではヘッドレス、ローカルではヘッドフル（デバッグしやすいように）で実行
        is_ci = os.getenv('CI') == 'true'
        print(f"CI環境として実行: {is_ci}")

        launch_options = {'headless': is_ci}
        if not is_ci:
            launch_options['slow_mo'] = 500

        browser = p.chromium.launch(**launch_options)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # --- タスク1: 売上データのダウンロードと移動 ---
            print("--- タスク1: 売上データの処理を開始します ---")

            # 1-1. サイトへのログイン
            print("1-1. サイトにログインします...")
            page.goto(LOGIN_URL)

            # 実際のWebサイトの入力フィールドの属性に合わせて調整してください
            email_locator = page.locator('#username')
            password_locator = page.locator('#password')

            # 1. メールアドレス入力フィールドをクリックしてアクティブにする
            email_locator.click()
            # 2. メールアドレスを入力する
            email_locator.fill(TENKURA_EMAIL)
            print(f"デバッグ: メールアドレス入力完了 (先頭5文字: {TENKURA_EMAIL[:5]}...)")

            # 短い待機時間を追加 (パスワード入力欄が確実にアクティブになるのを待つ)
            time.sleep(0.5)

            # 3. パスワードを入力する
            password_locator.fill(TENKURA_PASSWORD)
            print(f"デバッグ: パスワード入力完了 (先頭5文字: {TENKURA_PASSWORD[:5]}...)")

            # 4. ログインボタンをクリックする
            login_button_locator = page.locator('#login_button')
            login_button_locator.wait_for(state="visible") # ボタンが表示されるまで待機
            login_button_locator.click()
            time.sleep(0.5) # 短い待機

            # ログイン成功の確認 (ページが完全に読み込まれるまで待機)
            print("ログイン後のページ読み込みを待機します...")
            page.wait_for_load_state('networkidle', timeout=60000)
            # ローディング画面が非表示になるまで待機
            page.locator('#loading-bg').wait_for(state="hidden", timeout=60000)
            print("ログイン成功。")

            # 1-2. 売上一覧画面への移動
            print("1-2. 売上一覧画面に移動します...")
            # マウスカーソルを画面の左端に移動させてメニューを展開
            page.mouse.move(0, 100) # x=0 (左端), y=100 (適当な高さ)
            time.sleep(1) # メニューが開くのを待つ
            # 「販売管理」ボタンをクリックしてメニューを展開 (メニューが開いた状態でクリック)
            page.get_by_role("button", name="販売管理").click()
            time.sleep(0.5) # サブメニューが開くのを待つ
            # 「売上」の「一覧」リンクをクリック
            page.get_by_role("listitem").filter(has_text="売上").get_by_role("link", name="一覧", exact=True).wait_for(state="visible") # リンクが表示されるまで待機
            page.get_by_role("listitem").filter(has_text="売上").get_by_role("link", name="一覧", exact=True).click()
            page.wait_for_url("https://app.cloud-tenkura.net/hanbai/uriage/denpyo_list", timeout=60000)
            print("売上一覧画面に移動しました。")

            # 1-3. 売上データのダウンロード
            print("1-3. 売上データをダウンロードします...")
            page.get_by_text("昨日").click() # 「先月」から「昨日」に変更

            with page.expect_download() as download_info:
                data_output_button = page.get_by_role("button", name="データ出力")
                data_output_button.click()
                # 「データ出力」ボタンの親要素からドロップダウンメニューを特定して待機
                data_output_button.locator('xpath=../div[contains(@class, "dropdown-menu")]').wait_for(state="visible")
                # 「売上伝票」のCSVボタンをクリック
                page.locator('div.d-inline-flex.w-100:has(label:has-text("売上伝票"))').get_by_role("button", name="CSV").click()

            download = download_info.value
            download_path = Path.home() / "Downloads" / download.suggested_filename
            download.save_as(download_path)
            print(f"ファイル '{download.suggested_filename}' のダウンロードが完了しました。")

            # 1-4. 売上ファイルの移動
            print(f"1-4. 売上ファイルを '{OUTPUT_DIR}' フォルダに移動します...")
            destination_path = OUTPUT_DIR / download.suggested_filename
            # 移動先にファイルが既に存在する場合は上書きする
            if destination_path.exists():
                destination_path.unlink()
            shutil.move(download_path, destination_path)
            print(f"ファイルを '{destination_path}' に移動しました。")
            print("--- タスク1: 完了 ---")

            time.sleep(2)

            # --- タスク2: 仕入データのダウンロードと移動 ---
            print("\n--- タスク2: 仕入データの処理を開始します ---")

            # 2-1. 仕入一覧画面への移動
            print("2-1. 仕入一覧画面に移動します...")
            # マウスカーソルを画面の左端に移動させてメニューを展開
            page.mouse.move(0, 100) # x=0 (左端), y=100 (適当な高さ)
            time.sleep(1) # メニューが開くのを待つ
            # 「購買管理」ボタンをクリックしてメニューを展開 (メニューが開いた状態でクリック)
            page.get_by_role("button", name="購買管理").click()
            time.sleep(0.5) # サブメニューが開くのを待つ
            # 「仕入」の「一覧」リンクをクリック
            page.get_by_role("listitem").filter(has_text="仕入").get_by_role("link", name="一覧", exact=True).wait_for(state="visible") # リンクが表示されるまで待機
            page.get_by_role("listitem").filter(has_text="仕入").get_by_role("link", name="一覧", exact=True).click()
            page.wait_for_url("https://app.cloud-tenkura.net/koubai/shiire/denpyo_list", timeout=60000)
            print("仕入一覧画面に移動しました。")

            # 2-2. 仕入データのダウンロード
            print("2-2. 仕入データをダウンロードします...")
            page.get_by_text("昨日").click() # 「先月」から「昨日」に変更

            with page.expect_download() as download_info:
                data_output_button = page.get_by_role("button", name="データ出力")
                data_output_button.click()
                # 「データ出力」ボタンの親要素からドロップダウンメニューを特定して待機
                data_output_button.locator('xpath=../div[contains(@class, "dropdown-menu")]').wait_for(state="visible")
                # 「仕入伝票」のCSVボタンをクリック
                page.locator('div.d-inline-flex.w-100:has(label:has-text("仕入伝票"))').get_by_role("button", name="CSV").click()

            download = download_info.value
            download_path = Path.home() / "Downloads" / download.suggested_filename
            download.save_as(download_path)
            print(f"ファイル '{download.suggested_filename}' のダウンロードが完了しました。")

            # 2-3. 仕入ファイルの移動
            print(f"2-3. 仕入ファイルを '{OUTPUT_DIR}' フォルダに移動します...")
            destination_path = OUTPUT_DIR / download.suggested_filename
            # 移動先にファイルが既に存在する場合は上書きする
            if destination_path.exists():
                destination_path.unlink()
            shutil.move(download_path, destination_path)
            print(f"ファイルを '{destination_path}' に移動しました。")
            print("--- タスク2: 完了 ---")

            print("\n売上と仕入のデータダウンロード、およびファイル移動が完了しました。")

        except PlaywrightTimeoutError:
            error_message = "エラー: タイムアウトが発生しました。ページの読み込みや要素の表示が時間内に完了しませんでした。"
            print(error_message)
            screenshot_path = Path(__file__).parent / "error.png"
            page.screenshot(path=screenshot_path)
            print(f"現在の状況を '{screenshot_path}' に保存しました。")
        except Exception as e:
            error_message = f"予期せぬエラーが発生しました: {e}"
            print(error_message)
            screenshot_path = Path(__file__).parent / "error.png"
            page.screenshot(path=screenshot_path)
            print(f"現在の状況を '{screenshot_path}' に保存しました。")
        finally:
            browser.close()
            print("\nブラウザを閉じました。")

if __name__ == "__main__":
    main()
