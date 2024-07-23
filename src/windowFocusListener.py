import win32gui
import win32con
import win32process
import ctypes
import threading
import os
import psutil
from pycaw.pycaw import AudioUtilities
from TPClient import TPClient, g_log
from tppEntry import *
from tppEntry import __version__

class WindowFocusListener:
    def __init__(self):
        self.hook = None
        self.win_event_proc = None
        self.running = False
        self.thread = None
        self.TPClient = TPClient
        self.TP_PLUGIN_STATES = TP_PLUGIN_STATES
        self.current_app_connector_id = f"pc_{TP_PLUGIN_INFO['id']}_{TP_PLUGIN_CONNECTORS['APP control']['id']}|{TP_PLUGIN_CONNECTORS['APP control']['data']['appchoice']['id']}=Current app"


    def callback(self, hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
        if event == win32con.EVENT_SYSTEM_FOREGROUND:
            window_title = win32gui.GetWindowText(hwnd)
            executable_path = self.getActiveExecutablePath()
            if executable_path:
                process_name = os.path.basename(executable_path)
                self.update_volume_info(process_name)
                g_log.info(f"Window focus changed to: {window_title} | Process: {process_name}")

    def win_event_loop(self):
        WinEventProcType = ctypes.WINFUNCTYPE(
            None, 
            ctypes.wintypes.HANDLE,
            ctypes.wintypes.DWORD,
            ctypes.wintypes.HWND,
            ctypes.wintypes.LONG,
            ctypes.wintypes.LONG,
            ctypes.wintypes.DWORD,
            ctypes.wintypes.DWORD
        )
        
        self.win_event_proc = WinEventProcType(self.callback)
        
        user32 = ctypes.windll.user32
        self.hook = user32.SetWinEventHook(
            win32con.EVENT_SYSTEM_FOREGROUND,
            win32con.EVENT_SYSTEM_FOREGROUND,
            0,
            self.win_event_proc,
            0,
            0,
            win32con.WINEVENT_OUTOFCONTEXT
        )

        if self.hook == 0:
            g_log.info("Failed to set up event hook")
            return

        msg = ctypes.wintypes.MSG()
        while self.running:
            if user32.GetMessageA(ctypes.byref(msg), 0, 0, 0) != 0:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageA(ctypes.byref(msg))

    def start(self):
        if self.running:
            g_log.info("windowFocusListener is already running")
            return

        self.running = True
        self.thread = threading.Thread(target=self.win_event_loop)
        self.thread.start()
        g_log.info("windowFocusListener started")

    def stop(self):
        if self.hook:
            self.running = False
            ctypes.windll.user32.PostThreadMessageA(self.thread.ident, win32con.WM_QUIT, 0, 0)
            self.thread.join()
            ctypes.windll.user32.UnhookWinEvent(self.hook)
            self.hook = None
            self.win_event_proc = None
            
            g_log.info("Stopped listening for focus changes.")

    def getActiveExecutablePath(self):
        hWnd = ctypes.windll.user32.GetForegroundWindow()
        if hWnd == 0:
            return None
        else:
            _, pid = win32process.GetWindowThreadProcessId(hWnd)
            return psutil.Process(pid).exe()

    def get_volume(self, process_name):
        session = self.get_session(process_name)
        if session:
            interface = session.SimpleAudioVolume
            return interface.GetMasterVolume()
        return None

    def get_session(self, process_name):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process and session.Process.name() == process_name:
                return session
        return None

    def update_volume_info(self, process_name):
        current_app_volume = self.get_volume(process_name)
        if current_app_volume is not None:
            volume_int = int(current_app_volume * 100)
            if self.current_app_connector_id in self.TPClient.shortIdTracker:
                self.TPClient.shortIdUpdate(
                    self.TPClient.shortIdTracker[self.current_app_connector_id],
                    volume_int)
            self.TPClient.stateUpdate(self.TP_PLUGIN_STATES['currentAppVolume']['id'], str(volume_int))
        else:
            if self.current_app_connector_id in self.TPClient.shortIdTracker:
                self.TPClient.shortIdUpdate(
                    self.TPClient.shortIdTracker[self.current_app_connector_id],
                    0)
            self.TPClient.stateUpdate(self.TP_PLUGIN_STATES['currentAppVolume']['id'], "0")

# if __name__ == "__main__":
#     listener = WindowFocusListener()
#     try:
#         listener.start()
#         input("Press Enter to stop...")
#     finally:
#         listener.stop()