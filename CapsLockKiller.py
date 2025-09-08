import sys
import winreg
import ctypes

# --- 定数定義 ---
# レジストリキーのパス
KEY_PATH = r"SYSTEM\CurrentControlSet\Control\Keyboard Layout"
# 書き込む値の名前
SCANCODE_MAP_VALUE_NAME = "Scancode Map"

def is_admin():
    """管理者権限で実行されているかを確認する"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def set_caps_as_ctrl():
    """CapsLockキーを左Ctrlキーに割り当てる設定をレジストリに書き込む"""
    try:
        # レジストリのHKEY_LOCAL_MACHINEに接続
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
            # 指定したパスのキーを書き込み権限で開く
            with winreg.OpenKey(hkey, KEY_PATH, 0, winreg.KEY_WRITE) as key:
                # CapsLock(0x3A)を左Ctrl(0x1D)にマッピングするバイナリデータ
                # 構造:
                # 00 00 00 00  - ヘッダー (バージョン情報、常に0)
                # 00 00 00 00  - ヘッダー (フラグ、常に0)
                # 02 00 00 00  - マッピングエントリ数 (変更するキーの数 + 終端のNULL)
                # 1D 00 3A 00  - [変更後のキー] <- [変更前のキー] (左Ctrl <- CapsLock)
                # 00 00 00 00  - 終端のNULL
                scancode_map = bytes([
                    0x00, 0x00, 0x00, 0x00,
                    0x00, 0x00, 0x00, 0x00,
                    0x02, 0x00, 0x00, 0x00,
                    0x1D, 0x00, 0x3A, 0x00, # 左Ctrl(1D 00)にCapsLock(3A 00)を割り当て
                    0x00, 0x00, 0x00, 0x00,
                ])
                
                # Scancode Map値をバイナリデータとして書き込む
                winreg.SetValueEx(key, SCANCODE_MAP_VALUE_NAME, 0, winreg.REG_BINARY, scancode_map)
                print("成功: CapsLockが左Ctrlに割り当てられました。")
                print("設定を有効にするには、PCを再起動するか、一度サインアウトしてください。")

    except PermissionError:
        print("エラー: アクセスが拒否されました。このスクリプトは管理者権限で実行する必要があります。")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

def reset_keymap():
    """キーの割り当てをリセット（Scancode Map値を削除）する"""
    try:
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
            with winreg.OpenKey(hkey, KEY_PATH, 0, winreg.KEY_WRITE) as key:
                # Scancode Map値を削除する
                winreg.DeleteValue(key, SCANCODE_MAP_VALUE_NAME)
                print("成功: キーボードの割り当てがリセットされました。")
                print("設定を有効にするには、PCを再起動するか、一度サインアウトしてください。")
    
    except FileNotFoundError:
        print("情報: 割り当て設定は既にリセットされています（Scancode Map値が見つかりません）。")
    except PermissionError:
        print("エラー: アクセスが拒否されました。このスクリプトは管理者権限で実行する必要があります。")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

def main():
    """メイン処理"""
    if not is_admin():
        print("管理者権限が必要です。")
        print("コマンドプロンプトやPowerShellを「管理者として実行」してから、このスクリプトを実行してください。")
        input("Enterキーを押して終了します...")
        sys.exit(1)

    while True:
        print("\nCapsLockキーを左Ctrlキーに再割り当てします。")
        print("1: 有効化 (CapsLock -> 左Ctrl)")
        print("2: 無効化 (設定をリセット)")
        print("q: 終了")
        choice = input("操作を選択してください > ")

        if choice == '1':
            set_caps_as_ctrl()
            break
        elif choice == '2':
            reset_keymap()
            break
        elif choice.lower() == 'q':
            break
        else:
            print("無効な選択です。もう一度入力してください。")

if __name__ == "__main__":
    main()