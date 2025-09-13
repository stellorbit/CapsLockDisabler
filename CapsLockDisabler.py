import sys
import winreg
import ctypes

# --- 定数定義 ---
# レジストリキーのパス
KEY_PATH = r"SYSTEM\CurrentControlSet\Control\Keyboard Layout"
# 書き込む値の名前
SCANCODE_MAP_VALUE_NAME = "Scancode Map"

# --- 言語設定 ---
LANG = {
    "en": {
        "admin_required": "Administrator privileges are required.",
        "run_as_admin": "Please run this script from a command prompt or PowerShell with 'Run as administrator'.",
        "press_enter_to_exit": "Press Enter to exit...",
        "reassign_caps_lock": "\nRemap CapsLock key to Left Ctrl key.",
        "enable_option": "1: Enable (CapsLock -> Left Ctrl)",
        "disable_option": "2: Disable (Reset settings)",
        "quit_option": "q: Quit",
        "select_operation": "Select an operation > ",
        "invalid_choice": "Invalid choice. Please try again.",
        "success_assigned": "Success: CapsLock has been assigned to Left Ctrl.",
        "success_reset": "Success: Keyboard mapping has been reset.",
        "reboot_required": "To apply the settings, please restart your PC or sign out.",
        "error_permission": "Error: Access denied. This script must be run with administrator privileges.",
        "error_generic": "An error occurred: {e}",
        "info_already_reset": "Info: Mapping is already reset (Scancode Map value not found)."
    },
    "ja": {
        "admin_required": "管理者権限が必要です。",
        "run_as_admin": "コマンドプロンプトやPowerShellを「管理者として実行」してから、このスクリプトを実行してください。",
        "press_enter_to_exit": "Enterキーを押して終了します...",
        "reassign_caps_lock": "\nCapsLockキーを左Ctrlキーに再割り当てします。",
        "enable_option": "1: 有効化 (CapsLock -> 左Ctrl)",
        "disable_option": "2: 無効化 (設定をリセット)",
        "quit_option": "q: 終了",
        "select_operation": "操作を選択してください > ",
        "invalid_choice": "無効な選択です。もう一度入力してください。",
        "success_assigned": "成功: CapsLockが左Ctrlに割り当てられました。",
        "success_reset": "成功: キーボードの割り当てがリセットされました。",
        "reboot_required": "設定を有効にするには、PCを再起動するか、一度サインアウトしてください。",
        "error_permission": "エラー: アクセスが拒否されました。このスクリプトは管理者権限で実行する必要があります。",
        "error_generic": "エラーが発生しました: {e}",
        "info_already_reset": "情報: 割り当て設定は既にリセットされています（Scancode Map値が見つかりません）。"
    }
}

def get_lang():
    """言語を選択させる"""
    while True:
        print("\nSelect language")
        print("1: 日本語")
        print("2: English")
        lang_choice = input("> ").lower()
        if lang_choice == '1':
            return "ja"
        elif lang_choice == '2':
            return "en"
        else:
            print("Invalid choice. Please enter '1' or '2'.")

def is_admin():
    """管理者権限で実行されているかを確認する"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def set_caps_as_ctrl(t):
    """CapsLockキーを左Ctrlキーに割り当てる設定をレジストリに書き込む"""
    try:
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
            with winreg.OpenKey(hkey, KEY_PATH, 0, winreg.KEY_WRITE) as key:
                scancode_map = bytes([
                    0x00, 0x00, 0x00, 0x00,
                    0x00, 0x00, 0x00, 0x00,
                    0x02, 0x00, 0x00, 0x00,
                    0x1D, 0x00, 0x3A, 0x00,
                    0x00, 0x00, 0x00, 0x00,
                ])
                winreg.SetValueEx(key, SCANCODE_MAP_VALUE_NAME, 0, winreg.REG_BINARY, scancode_map)
                print(t["success_assigned"])
                print(t["reboot_required"])

    except PermissionError:
        print(t["error_permission"])
    except Exception as e:
        print(t["error_generic"].format(e=e))

def reset_keymap(t):
    """キーの割り当てをリセット（Scancode Map値を削除）する"""
    try:
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
            with winreg.OpenKey(hkey, KEY_PATH, 0, winreg.KEY_WRITE) as key:
                winreg.DeleteValue(key, SCANCODE_MAP_VALUE_NAME)
                print(t["success_reset"])
                print(t["reboot_required"])
    
    except FileNotFoundError:
        print(t["info_already_reset"])
    except PermissionError:
        print(t["error_permission"])
    except Exception as e:
        print(t["error_generic"].format(e=e))

def main():
    """メイン処理"""
    lang_choice = get_lang()
    t = LANG[lang_choice]

    if not is_admin():
        print(t["admin_required"])
        print(t["run_as_admin"])
        input(t["press_enter_to_exit"])
        sys.exit(1)

    while True:
        print(t["reassign_caps_lock"])
        print(t["enable_option"])
        print(t["disable_option"])
        print(t["quit_option"])
        choice = input(t["select_operation"])

        if choice == '1':
            set_caps_as_ctrl(t)
            break
        elif choice == '2':
            reset_keymap(t)
            break
        elif choice.lower() == 'q':
            break
        else:
            print(t["invalid_choice"])

if __name__ == "__main__":

    main()
