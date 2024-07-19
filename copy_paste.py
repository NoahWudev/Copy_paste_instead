import win32gui
import win32con
import win32com.client
import pyperclip
import time
import pyautogui


# 找到目標窗口的handle
def find_window(title_part):
    def callback(hwnd, hwnds):
        if title_part.lower() in win32gui.GetWindowText(hwnd).lower():
            hwnds.append(hwnd)
        return True
    
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds[0] if hwnds else None

# 獲取當前剪貼簿內容
def get_clipboard_content():
    return pyperclip.paste()

# 確保窗口處於活動狀態
def activate_window(hwnd):
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        # 設置窗口為前台
        win32gui.SetForegroundWindow(hwnd)
        # 開啟前台窗口
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        time.sleep(0.1)  # 確保窗口有足夠的時間成為前台窗口
    except Exception as e:
        print(f"開啟窗口失敗: {e}")

# 模擬鍵盤輸入文字
def send_text_to_window(hwnd, text):
    try:
        activate_window(hwnd)
        time.sleep(1)  # 確保窗口已經處於前台
        pyautogui.typewrite(text, interval=0.01)  # 使用 pyautogui 來模擬輸入
        print("已完成貼上")
    except Exception as e:
        print(f"貼上失敗: {e}")

# 主函數
def main():
    title_part = input("請輸入目標應用程式窗口的標題(部分匹配即可): ")
    hwnd = find_window(title_part)
    
    if hwnd:
        print(f"找到窗口句柄: {hwnd}")
        last_text = get_clipboard_content()
        print("剪貼簿監控已啟動。請複製文字。")
        
        try:
            while True:
                current_text = get_clipboard_content()
                if current_text != last_text:
                    print(f"檢測到新的剪貼簿內容，即將自動貼上: {current_text[:30]}...")
                    send_text_to_window(hwnd, current_text)
                    last_text = current_text
                time.sleep(0.5)  # 每0.5秒檢查一次剪貼板內容
        except KeyboardInterrupt:
            print("監控已停止。")
    else:
        print("未找到匹配的窗口。")

if __name__ == "__main__":
    main()
