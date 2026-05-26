import sys
import ctypes
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

import subprocess
import win32api
import win32con
import win32gui
import atexit
import time
import struct
import re
import win32com.client
import wmi as wmilib
import threading
import traceback
from PIL import Image, ImageDraw
from PyQt5.QtWidgets import QApplication, QMenu, QAction, QActionGroup
from PyQt5.QtGui import QCursor

# ============ POWER PLAN ============
BALANCED_GUID = "381b4222-f694-41f0-9685-ff5bb260df2e"
MAX_GUID      = "823a8a39-e7eb-49b0-a7d5-e39ae2d0bd15"

# ============ FAN PROFILES ============
PROFILES = {
    "performance": 0x00,
    "balanced":    0x01,
    "cool":        0x02,
    "quiet":       0x03,
}

PROFILE_NAMES = {
    0x00: "performance",
    0x01: "balanced",
    0x02: "cool",
    0x03: "quiet",
}

SIGN_BYTES = tuple(struct.pack('<I', 0x55434553))

WMAPP_NOTIFYCALLBACK = win32con.WM_APP + 1
WMAPP_EXIT           = win32con.WM_APP + 2
WATCHDOG_INTERVAL    = 10

STYLE = """
    QMenu {
        background-color: #1a1a1a;
        color: #ffffff;
        border: 1px solid #333;
        padding: 4px;
        font-size: 13px;
        font-family: Segoe UI;
    }
    QMenu::item {
        padding: 7px 24px 7px 12px;
        border-radius: 4px;
    }
    QMenu::item:selected {
        background-color: #2d2d2d;
    }
    QMenu::item:checked {
        font-weight: bold;
        color: #44ddff;
    }
    QMenu::item:disabled {
        color: #555;
    }
    QMenu::separator {
        height: 1px;
        background: #333;
        margin: 4px 8px;
    }
"""

# ============ POWER PLAN FUNCTIONS ============

def run_cmd(cmd):
    result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
    if result.returncode != 0:
        raise RuntimeError(f"Command failed: {cmd}")
    return result.stdout

def get_active_power_plan():
    """Возвращает (guid, name) текущей схемы питания"""
    try:
        output = run_cmd("powercfg /getactivescheme")
        match  = re.search(r"([0-9a-fA-F\-]{36})\s*\((.+?)\)", output)
        if match:
            return match.group(1).lower(), match.group(2).strip()
        match2 = re.search(r"([0-9a-fA-F\-]{36})", output)
        if match2:
            return match2.group(1).lower(), "Unknown"
    except Exception:
        pass
    return None, "Unknown"

def set_power_plan(guid):
    try:
        run_cmd(f"powercfg /setactive {guid}")
        return True
    except Exception:
        return False

# ============ FAN FUNCTIONS ============

def get_wmi():
    svc  = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    conn = svc.ConnectServer(".", "root\\WMI")
    c    = wmilib.WMI(namespace="root\\WMI")
    instance_name = c.hpqBIntM()[0].InstanceName
    obj  = conn.Get(f"hpqBIntM.InstanceName='{instance_name}'")
    return conn, obj, instance_name

def wmi_send(cmd, cmd_type, value):
    conn, obj, instance_name = get_wmi()
    data_in = conn.Get("hpqBDataIn").SpawnInstance_()
    data_in.Properties_("InstanceName").Value = instance_name
    data_in.Properties_("Sign").Value         = SIGN_BYTES
    data_in.Properties_("Command").Value      = cmd
    data_in.Properties_("CommandType").Value  = cmd_type
    data_in.Properties_("Size").Value         = 0x04
    data_in.Properties_("hpqBData").Value     = tuple([value, 0, 0, 0])
    method    = conn.Get("hpqBIntM").Methods_("hpqBIOSInt128")
    in_params = method.InParameters.SpawnInstance_()
    in_params.Properties_("InData").Value = data_in
    result   = obj.ExecMethod_("hpqBIOSInt128", in_params)
    return result.Properties_("OutData").Value.Properties_("rwReturnCode").Value

def set_fan_profile(profile_name):
    value = PROFILES[profile_name]
    wmi_send(0x01, 0x28, value)
    return wmi_send(0x02, 0x4c, value) == 0

def read_fan_profile():
    try:
        conn, obj, instance_name = get_wmi()
        data_in = conn.Get("hpqBDataIn").SpawnInstance_()
        data_in.Properties_("InstanceName").Value = instance_name
        data_in.Properties_("Sign").Value         = SIGN_BYTES
        data_in.Properties_("Command").Value      = 0x01
        data_in.Properties_("CommandType").Value  = 0x4c
        data_in.Properties_("Size").Value         = 0x04
        data_in.Properties_("hpqBData").Value     = tuple([0, 0, 0, 0])
        method    = conn.Get("hpqBIntM").Methods_("hpqBIOSInt128")
        in_params = method.InParameters.SpawnInstance_()
        in_params.Properties_("InData").Value = data_in
        result   = obj.ExecMethod_("hpqBIOSInt128", in_params)
        data_out = result.Properties_("OutData").Value
        return PROFILE_NAMES.get(data_out.Properties_("Data").Value[0])
    except Exception:
        return None

# ============ ICON ============

def create_icon(color, path):
    size  = 64
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw  = ImageDraw.Draw(image)
    draw.ellipse((10, 10, size - 10, size - 10), fill=color)
    image.save(path, format="ICO")

ICONS = {
    "performance": "ico_fan_perf.ico",
    "balanced":    "ico_fan_bal.ico",
    "cool":        "ico_fan_cool.ico",
    "quiet":       "ico_fan_quiet.ico",
    "unknown":     "ico_unknown.ico",
}

create_icon("#ff4444", ICONS["performance"])
create_icon("#44aaff", ICONS["balanced"])
create_icon("#44ddff", ICONS["cool"])
create_icon("#aaaaaa", ICONS["quiet"])
create_icon("#808080", ICONS["unknown"])

# ============ TRAY APP ============

class TrayApp:
    def __init__(self, qt_app):
        self.qt_app       = qt_app
        self.hwnd         = None
        self.class_name   = "HPControlTrayClass"
        self.icon_id      = 1
        self.fan_profile  = None
        self.fan_pinned   = None
        self.watchdog_on  = True
        self.power_guid   = None
        self.power_name   = "Unknown"
        self._lock        = threading.Lock()

        self.fan_profile = read_fan_profile() or "balanced"
        self.fan_pinned  = self.fan_profile
        self.power_guid, self.power_name = get_active_power_plan()

        self._wd_thread = threading.Thread(target=self._watchdog_loop, daemon=True)
        self._wd_thread.start()

        atexit.register(self.cleanup)

    def _watchdog_loop(self):
        while True:
            time.sleep(WATCHDOG_INTERVAL)
            with self._lock:
                if self.watchdog_on and self.fan_pinned:
                    try:
                        set_fan_profile(self.fan_pinned)
                    except Exception:
                        pass

    def cleanup(self):
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (self.hwnd, self.icon_id))
        except Exception:
            pass

    def toggle_power(self):
        for attempt in range(3):
            try:
                if self.power_guid == MAX_GUID.lower():
                    set_power_plan(BALANCED_GUID)
                else:
                    set_power_plan(MAX_GUID)
                time.sleep(0.3)
                self.power_guid, self.power_name = get_active_power_plan()
                self.refresh_icon()
                return
            except Exception as e:
                print(f"toggle_power attempt {attempt+1} failed: {e}")
                time.sleep(1)

    def switch_fan(self, profile_name):
        with self._lock:
            self.fan_pinned = profile_name
        threading.Thread(
            target=self._set_fan_async,
            args=(profile_name,),
            daemon=True
        ).start()

    def _set_fan_async(self, profile_name):
        try:
            set_fan_profile(profile_name)
            self.fan_profile = profile_name
        except Exception as e:
            print(f"Fan error: {e}")
        self.refresh_icon()

    def toggle_watchdog(self):
        with self._lock:
            self.watchdog_on = not self.watchdog_on
        self.refresh_icon()

    def refresh_after_sleep(self):
        time.sleep(5)
        self.power_guid, self.power_name = get_active_power_plan()
        with self._lock:
            if self.fan_pinned:
                try:
                    set_fan_profile(self.fan_pinned)
                    self.fan_profile = self.fan_pinned
                except Exception as e:
                    print(f"Fan restore after sleep failed: {e}")
        self.refresh_icon()

    def get_icon_path(self):
        return ICONS.get(self.fan_profile, ICONS["unknown"])

    def get_tooltip(self):
        fan = self.fan_profile or "?"
        wd  = "ON" if self.watchdog_on else "OFF"
        return f"Fan: {fan.capitalize()}  |  Power: {self.power_name}  |  WD: {wd}"

    def refresh_icon(self):
        try:
            icon_path = self.get_icon_path()
            tip       = self.get_tooltip()
            hicon = win32gui.LoadImage(
                0, icon_path, win32con.IMAGE_ICON, 0, 0,
                win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            )
            flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
            nid   = (self.hwnd, self.icon_id, flags, WMAPP_NOTIFYCALLBACK, hicon, tip)
            try:
                win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)
            except Exception:
                win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        except Exception as e:
            print(f"refresh_icon error: {e}")

    def show_menu(self):
        try:
            # Обновляем power plan перед показом меню
            self.power_guid, self.power_name = get_active_power_plan()

            menu = QMenu()
            menu.setStyleSheet(STYLE)

            # Статус fan
            s1 = QAction(f"🌀  Fan: {(self.fan_profile or '?').capitalize()}", menu)
            s1.setEnabled(False)
            menu.addAction(s1)

            # Статус power plan — показываем реальное название
            s2 = QAction(f"🔋  Power: {self.power_name}", menu)
            s2.setEnabled(False)
            menu.addAction(s2)

            menu.addSeparator()

            # Fan profiles
            grp = QActionGroup(menu)
            fan_items = {
                "performance": "⚡  Performance",
                "balanced":    "⚖  Balanced",
                "cool":        "❄  Cool",
                "quiet":       "🔇  Quiet",
            }
            for name, label in fan_items.items():
                a = QAction(label, menu)
                a.setCheckable(True)
                a.setChecked(name == self.fan_profile)
                a.triggered.connect(lambda _, n=name: self.switch_fan(n))
                grp.addAction(a)
                menu.addAction(a)

            menu.addSeparator()

            # Watchdog
            wd_a = QAction(f"📌  Watchdog ({WATCHDOG_INTERVAL}s)", menu)
            wd_a.setCheckable(True)
            wd_a.setChecked(self.watchdog_on)
            wd_a.triggered.connect(self.toggle_watchdog)
            menu.addAction(wd_a)

            menu.addSeparator()

            # Power plan toggle
            if self.power_guid == MAX_GUID.lower():
                power_label = "🔋  → Balanced Power"
            else:
                power_label = "🔋  → Max Performance"
            pw_a = QAction(power_label, menu)
            pw_a.triggered.connect(
                lambda: threading.Thread(target=self.toggle_power, daemon=True).start()
            )
            menu.addAction(pw_a)

            menu.addSeparator()

            quit_a = QAction("✕  Exit", menu)
            quit_a.triggered.connect(
                lambda: win32gui.PostMessage(self.hwnd, WMAPP_EXIT, 0, 0)
            )
            menu.addAction(quit_a)

            pos = QCursor.pos()
            menu.exec_(pos)

        except Exception as e:
            print(f"show_menu error: {e}")
            traceback.print_exc()

    def wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == WMAPP_NOTIFYCALLBACK:
            if lparam in (win32con.WM_LBUTTONUP, win32con.WM_RBUTTONUP):
                self.show_menu()
            return 0

        if msg == win32con.WM_POWERBROADCAST:
            if wparam == 0x8018:  # PBT_APMRESUMEAUTOMATIC
                threading.Thread(target=self.refresh_after_sleep, daemon=True).start()
            return 1

        if msg == WMAPP_EXIT:
            win32gui.DestroyWindow(hwnd)
            return 0

        if msg == win32con.WM_DESTROY:
            self.cleanup()
            win32gui.PostQuitMessage(0)
            return 0

        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def run(self):
        wc               = win32gui.WNDCLASS()
        wc.hInstance     = win32api.GetModuleHandle(None)
        wc.lpszClassName = self.class_name
        wc.lpfnWndProc   = self.wnd_proc
        class_atom       = win32gui.RegisterClass(wc)

        self.hwnd = win32gui.CreateWindow(
            class_atom, self.class_name,
            0, 0, 0, 0, 0, 0, 0,
            wc.hInstance, None
        )

        self.refresh_icon()
        win32gui.PumpMessages()


if __name__ == "__main__":
    try:
        qt_app = QApplication(sys.argv)
        qt_app.setQuitOnLastWindowClosed(False)
        app = TrayApp(qt_app)
        app.run()
    except Exception as e:
        traceback.print_exc()
        input("Press Enter to exit...")
