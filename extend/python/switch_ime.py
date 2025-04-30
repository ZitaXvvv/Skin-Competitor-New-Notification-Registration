import win32api
import win32gui
from win32con import WM_INPUTLANGCHANGEREQUEST

def to_en():
    hwnd = win32gui.GetForegroundWindow()
    win32api.SendMessage(
        hwnd,
        WM_INPUTLANGCHANGEREQUEST,
        0,
        0x0409
    )

if __name__ == '__main__':
    to_en()