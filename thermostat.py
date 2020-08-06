#!/usr/bin/env python3

# Base modules needed by thermostat
import configparser
import sys
import time
import os
import argparse
import logging

# Initialize logging
log = logging.getLogger("thermostat")
log.addHandler(logging.NullHandler())

# OpenHardwareMonitorLib.dll is only 32-bit, so application must be 32-bit!
if sys.maxsize > 2**32:
    raise Exception("OpenHardwareMonitorLib is only 32-bit")

# pywin32 is used to create a message-only window to receive USB device changes
# https://github.com/mhammond/pywin32/blob/master/win32/Demos/win32gui_devicenotify.py
# https://docs.microsoft.com/en-us/windows/win32/winmsg/window-features#message-only-windows
import win32gui
import win32con
import win32api
import win32gui_struct

# For creating the system tray icon
# https://github.com/mhammond/pywin32/blob/master/win32/Demos/win32gui_taskbar.py
import win32event
import winerror

# Just for creating dialog popups
import win32ui

# To run the win32 window messagepump; should be in its own thread
import threading

# Needs 32-bit hidapi.dll and lib in same directory as binary, or os.getcwd()
# https://docs.microsoft.com/en-us/windows/win32/dlls/dynamic-link-library-search-order#standard-search-order-for-desktop-applications
import hid
import ctypes

# Meeds 32-bit OpenHardwareMonitorLib.dll in same directory as binary or sys.path
# **Must run as administrator!**
# https://stackoverflow.com/a/49909330
import clr
clr.AddReference('OpenHardwareMonitorLib')
from OpenHardwareMonitor import Hardware

class Thermostat():
    def __init__(self, *args, **kwargs):
        if len(args) % 2:
            raise Exception("Thermostat arguments must be in pairs")
        
        # Allow forward and reverse hysteresis
        fhyst = int(kwargs.pop("forward_hysteresis", 0))
        rhyst = int(kwargs.pop("reverse_hysteresis", 0))
        
        self.__tmap = {}
        
        for i in range(0, len(args), 2):
            self.__tmap[int(args[i+1])] = args[i]
            
        for k, v in kwargs.items():
            self.__tmap[int(v)] = k
        
        if len(self.__tmap.keys()) < 2:
            raise Exception("Thermostat requires at least 2 settings (hi/lo)")
        
        # Need lowest, second lowest, and highest to set thermostat ranges
        keys = sorted(self.__tmap.keys())
        tmin = keys[0]
        t1 = keys[1]
        tmax = keys[-1]
        
        # Simplifying
        fail = (None, False)
        
        # The below min, and above max ranges are special
        self.__ranges = {
            lambda t: t < t1 + fhyst:
                lambda t: (self.__tmap[tmin], True) if t < (t1 + fhyst) else fail,
            lambda t: t >= tmax + fhyst:
                lambda t: (self.__tmap[tmax], True) if t > (tmax - rhyst) else fail,
        }
        
        # Remaining hysteresis ranges filled in dynamically
        def middle_range(low, hi, target):
            forward = lambda t: t >= (low + fhyst) and t < (hi + fhyst)
            reverse = lambda t: (self.__tmap[low], True) if t > (low - rhyst) and t < (hi + fhyst) else fail
            target[forward] = reverse
        for i in range(1, len(keys)-1):
            middle_range(keys[i], keys[i+1], self.__ranges)
        
        # Set the mode to a function which always returns False
        self.__hrange = lambda t: fail
        
    def mode(self, temperature, changes=True, unchanged=None):
        newmode, hasmode = self.__hrange(temperature)
        if not hasmode:
            self.__hrange = next(self.__ranges[f] for f in self.__ranges.keys() if f(temperature))
            newmode = self.__hrange(temperature)[0]
        try:
            mode = unchanged if self.__last == newmode and changes else newmode
        except AttributeError:
            mode = newmode
        self.__last = newmode
        return mode


class VoltageSwitch(hid.Device):
    # Didn't lookup the byte order and since only looking for bytes . . .
    v12 = ctypes.create_string_buffer(b'\x01'*64, 64)
    v5 = ctypes.create_string_buffer(b'\x00'*64, 64)
    v0 = ctypes.create_string_buffer(b'\x02'*64, 64)

    def __init__(self, vid=None, pid=None, serial=None, path=None):
        if vid and pid:
            try:
                vid, pid = map(int, vid, pid)
            except ValueError:
                vid, pid = map(lambda i: int(i, 0), vid, pid)
        self.__switch = {
                            '12v': self.set12v,
                            '5v': self.set5v,
                            '0v': self.set0v,
                            None: lambda *args: None,
                        }
        super().__init__(vid, pid, serial, path)
    
    def set12v(self):
        log.debug("Setting voltage switch to 12V")
        return self.write(VoltageSwitch.v12)
    
    def set5v(self):
        log.debug("Setting voltage switch to 5V")
        return self.write(VoltageSwitch.v5)
    
    def set0v(self):
        log.debug("Setting voltage switch to 0V")
        return self.write(VoltageSwitch.v0)
    
    def __getitem__(self, key):
        return self.__switch[key]


class TemperatureSensor():
    def __init__(self, device=None, sensor=None, cpu=True, ram=False, gpu=False, mobo=False, disk=False):
        self.__handle = Hardware.Computer()
        self.__handle.CPUEnabled = cpu
        self.__handle.MainboardEnabled = mobo
        self.__handle.RAMEnabled = ram
        self.__handle.GPUEnabled = gpu
        self.__handle.HDDEnabled = disk
        self.__handle.Open()
        
        if len(self.__handle.Hardware) < 1:
            raise Exception("No hardware was detected")

        try:
            self.__hardware = next(h for h in self.__handle.Hardware if h.Name == device)
        except:
            raise Exception("Device '%s' not detected.\r\n\r\nAvailable devices:\r\n%s"
                            %(device, '\r\n'.join(sorted(f"- {h.Name}" for h in self.__handle.Hardware))))
            
        # Poll device
        self.__hardware.Update()

        try:
            self.__sensor = next(s for s in self.__hardware.Sensors if s.Index == 2 and s.Name == sensor)
        except:
            raise Exception("Sensor '%s' not found.\r\n\r\nAvailable sensors:\r\n%s"
                            %(sensor, '\r\n'.join(sorted(f"- {s.Name}" for s in self.__hardware.Sensors if s.Index == 2))))
        
        self.value = self.__sensor.Value
    
    def reading(self):
        self.__hardware.Update()
        self.value = self.__sensor.Value
        return self.value


class Win32HID(threading.Event):
    # **Everything is added to self to keep python from doing GC!!**

    # USB Device works, but HID device is more specific. Both listed anyway
    GUID_DEVINTERFACE_USB_DEVICE = "{A5DCBF10-6530-11D2-901F-00C04FB951ED}"
    GUID_DEVINTERFACE_HID = "{4D1E55B2-F16F-11CF-88CB-001111000030}"

    def __init__(self, device=hid.Device, vid=None, pid=None, path=None):
        # Initialize the base class event object
        super().__init__()
    
        # Path must be convertable to bytes, and vid/pid integers
        if path:
            try:
                self.__path = path.lower().encode()
            except AttributeError:
                self.__path = path.lower()
            try:
                with hid.Device(path=self.__path):
                    super().set()
            except hid.HIDException:
                pass
        elif vid and pid:
            try:
                v, p = map(lambda x: int(x),[vid,pid])
            except:
                v, p = map(lambda x: int(x, 0),[vid,pid])
            
            # The IDs are only used to find the full device path
            self.__ids = list(f"{x}_{y:04x}".encode() for x, y in zip(["vid","pid"],[v,p]))
            
            # Always use path for connection; check for device at startup
            self.__path = None
            for d in hid.enumerate(vid=v, pid=p):
                if self.__matchingdevice(d["path"]):
                    super().set()
                    break
        else:
            raise Exception("VID, PID or path required for HID device")
        
        self.__dev = device
    
        # Create a window class for receiving messages
        self.wc = win32gui.WNDCLASS()
        self.wc.hInstance = win32api.GetModuleHandle(None)
        self.wc.lpszClassName = "win32hidnotifier"
        self.wc.lpfnWndProc = {win32con.WM_DEVICECHANGE:self.__devicechange}
        
        # Register the window class
        winClass = win32gui.RegisterClass(self.wc)
        
        # Create a Message-Only window
        self.hwnd = win32gui.CreateWindowEx(
            0,                      #dwExStyle
			self.wc.lpszClassName,  #lpClassName
			self.wc.lpszClassName,  #lpWindowName
			0, 0, 0, 0, 0,
			win32con.HWND_MESSAGE,  #hWndParent
			0, 0, None)
        
        # Watch for all USB device notifications
        self.filter = win32gui_struct.PackDEV_BROADCAST_DEVICEINTERFACE(Win32HID.GUID_DEVINTERFACE_HID)
        self.hdev = win32gui.RegisterDeviceNotification(self.hwnd, self.filter,
                                                        win32con.DEVICE_NOTIFY_WINDOW_HANDLE)


    # Added only for ergonomics
    def attached(self):
        return self.is_set()
        
    
    def device(self, timeout=None):
        if not self.wait(timeout):
            raise Exception("Device not attached")
        return self.__dev(path=self.__path)


    def __matchingdevice(self, name):
        # Names looks like:
        # '\\\\?\\HID#VID_16C0&PID_0486&MI_01#7&2d928156&0&0000#{4d1e55b2-f16f-11cf-88cb-001111000030}'
        # https://docs.microsoft.com/en-us/windows-hardware/drivers/install/standard-usb-identifiers#multiple-interface-usb-devices
        try:
            path = name.encode().lower()
        except AttributeError:
            path = name.lower()
        
        if self.__path and self.__path == path:
            return True
        elif self.__ids:
            chunks = path.split(b'#')
            if len(chunks) > 2 and chunks[-1] == Win32HID.GUID_DEVINTERFACE_HID.encode().lower():
                ids = chunks[1].split(b'&')
                if (all(id in ids for id in self.__ids) and 
                    all(id == b'mi_00' for id in ids if id.startswith(b'mi_'))):
                    self.__path = path
                    return True
        return False


    # WM_DEVICECHANGE message handler.
    def __devicechange(self, hWnd, msg, wParam, lParam):
        info = win32gui_struct.UnpackDEV_BROADCAST(lParam)
        
        if (info.devicetype == win32con.DBT_DEVTYP_DEVICEINTERFACE and
            self.__matchingdevice(info.name)):
            if wParam == win32con.DBT_DEVICEREMOVECOMPLETE:
                log.debug("Device is being removed")
                self.clear()
            elif wParam == win32con.DBT_DEVICEARRIVAL:
                log.debug("Device is being added")
                self.set()
        return True

class TrayThermostat(threading.Thread):
    Automatic = 1023
    V12 = 1024
    V5 = 1025
    V0 = 1026
    Exit = 1027
    Connected = 1028

    def __init__(self, w32hid, sensor, thermostat):
        threading.Thread.__init__(self)
        self.__w32hid = w32hid
        self.__sensor = sensor
        self.__thermostat = thermostat
        
        # Pre-initialize a dummy event for the ticker based msg pump
        self.__event = win32event.CreateEvent(None, 0, 0, None)
    
    def __wait_msg_pump(self, timeout=2000):
        # When user does mouseover it runs QS_SENDMESSAGE 0x0040 in a loop . . .
        # Also we receive ALL mouse clicks . . .
        rc = win32event.MsgWaitForMultipleObjects(
                (self.__event,), # list of objects
                0, # wait all
                timeout,  # timeout
                win32event.QS_ALLINPUT, # type of input
                )
        if rc == win32event.WAIT_OBJECT_0+1:
            # Message waiting.
            if win32gui.PumpWaitingMessages():
                # Received WM_QUIT, so return True.
                win32gui.PostMessage(self.__w32hid.hwnd, win32con.WM_QUIT, 0, 0)
                return True
    
    def run(self):
        self.__mode = TrayThermostat.Automatic
        self.__connected = False
    
        msg_TaskbarRestart = win32gui.RegisterWindowMessage("TaskbarCreated");
        message_map = {
                msg_TaskbarRestart: self.OnRestart,
                win32con.WM_DESTROY: self.OnDestroy,
                win32con.WM_COMMAND: self.OnCommand,
                win32con.WM_USER+20 : self.OnTaskbarNotify,
        }
        # Register the Window class.
        wc = win32gui.WNDCLASS()
        hinst = wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpszClassName = "HTPCThermostat"
        wc.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW;
        wc.hCursor = win32api.LoadCursor( 0, win32con.IDC_ARROW )
        wc.hbrBackground = win32con.COLOR_WINDOW
        wc.lpfnWndProc = message_map # could also specify a wndproc.

        # Don't blow up if class already registered to make testing easier
        try:
            classAtom = win32gui.RegisterClass(wc)
        except win32gui.error as err_info:
            if err_info.winerror!=winerror.ERROR_CLASS_ALREADY_EXISTS:
                raise

        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow( wc.lpszClassName, "HTPC Thermostat Taskbar", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, hinst, None)
        win32gui.UpdateWindow(self.hwnd)
        self._DoCreateIcons()
        
        # There must be an event to wait for even if its not used. Otherwise
        # MsgWaitForMultipleObjects will just return immediately.
        last_poll = 0
        while True:
            # If we're here, then we're not connected.
            self.__connected = False
            
            # MsgWaitForMultiple objects will miss events between calls
            if win32gui.PumpWaitingMessages():
                return
        
            try:
                # Wait 0.2 seconds for the device (so as not to delay win32 Message Loop)
                with self.__w32hid.device(timeout=0.2) as vs:
                    # If we made it here, then we're connected
                    self.__connected = True
                    
                    # On the first reading always set the switch state
                    changes = False
                    
                    while True:
                        # MsgWaitForMultiple objects will miss events between calls
                        if win32gui.PumpWaitingMessages():
                            return
                        
                        # Now we wait . . .
                        if self.__wait_msg_pump():
                            return
                        
                        if (time.time() * 1000) - last_poll >= 2000:
                            last_poll = time.time() * 1000
                            
                            # Set the voltage switch based on our current mode
                            if self.__mode == TrayThermostat.Automatic:
                                # Send the sensor reading to the thermostat for the voltage switch
                                if not vs[self.__thermostat.mode(self.__sensor.reading(), changes)]():
                                    # If no command sent, check to ensure still connected
                                    if not self.__w32hid.attached():
                                        break
                                
                                # From now on, only set switch state on changes
                                changes = True
                            elif self.__mode == TrayThermostat.V12:
                                vs.set12v()
                                changes = False
                            elif self.__mode == TrayThermostat.V5:
                                vs.set5v()
                                changes = False
                            elif self.__mode == TrayThermostat.V0:
                                vs.set0v()
                                changes = False
            except:
                log.exception("Terminated HID connection")
            
            # Now we wait 300 milliseconds for GUI messages
            if self.__wait_msg_pump(timeout=300):
                return


    def _DoCreateIcons(self):
        # Try and find a custom icon
        hinst =  win32api.GetModuleHandle(None)
        iconPathName = os.path.abspath(".\\main.ico" )
        if os.path.isfile(iconPathName):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
        else:
            hicon = win32gui.ExtractIcon(hinst, sys.executable, 0)

        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, "HTPC Thermostat")
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        except win32gui.error:
            # This is common when windows is starting, and this code is hit
            # before the taskbar has been created.
            log.debug("Failed to add the taskbar icon - is explorer running?")
            # but keep running anyway - when explorer starts, we get the
            # TaskbarCreated message.


    def OnRestart(self, hwnd, msg, wparam, lparam):
        self._DoCreateIcons()


    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostMessage(self.__w32hid.hwnd, win32con.WM_QUIT, 0, 0) # Terminate win32hid
        win32gui.PostQuitMessage(0) # Terminate the app.

    def __flag_set(self, mode):
        return win32con.MF_STRING | win32con.MF_CHECKED if self.__mode == mode else win32con.MF_STRING

    def OnTaskbarNotify(self, hwnd, msg, wparam, lparam):
        if lparam==win32con.WM_LBUTTONUP:
            pass
        elif lparam==win32con.WM_LBUTTONDBLCLK:
            win32gui.DestroyWindow(self.hwnd)
        elif lparam==win32con.WM_RBUTTONUP:
            menu = win32gui.CreatePopupMenu()
            win32gui.AppendMenu( menu, win32con.MF_GRAYED | win32con.MF_STRING, TrayThermostat.Connected, "Connected" if self.__connected else "Disconnected")
            win32gui.AppendMenu( menu, self.__flag_set(TrayThermostat.Automatic), TrayThermostat.Automatic, "Automatic")
            win32gui.AppendMenu( menu, self.__flag_set(TrayThermostat.V12), TrayThermostat.V12, "12V")
            win32gui.AppendMenu( menu, self.__flag_set(TrayThermostat.V5), TrayThermostat.V5, "5V" )
            win32gui.AppendMenu( menu, self.__flag_set(TrayThermostat.V0), TrayThermostat.V0, "Off" )
            win32gui.AppendMenu( menu, win32con.MF_STRING, TrayThermostat.Exit, "Exit" )
            pos = win32gui.GetCursorPos()
            # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
            win32gui.SetForegroundWindow(self.hwnd)
            win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
            win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
        return 1

    def OnCommand(self, hwnd, msg, wparam, lparam):
        id = win32api.LOWORD(wparam)
        if id == TrayThermostat.Exit:
            log.debug("Quitting")
            win32gui.DestroyWindow(self.hwnd)
        else:
            log.debug("Setting mode - %i", id)
            self.__mode = id

def mainloop(w32hid, sensor, thermostat):
    while True:
        try:
            # Wait for the device, connect, then enter the context
            with w32hid.device() as vs:
                # On the first reading always set the switch state
                changes = False
                
                while True:
                    time.sleep(2)
                    
                    # Set the switch state by feeding sensor data into thermostat
                    vs[thermostat.mode(sensor.reading(), changes)]()
                    
                    # From now on, only set switch state on changes
                    changes = True
        except:
            log.exception("Terminated HID connection")
            time.sleep(30)


def verify_config(config):
    sections = ['thermostat', 'microcontroller', 'probe']
    types = ["cpu", "ram", "gpu", "mobo", "disk"]
    
    if not all(config.has_section(s) for s in sections):
        log.error("Config file 'thermostat.ini' not found; creating a new one.")
        config['thermostat'] = {
            "12V": "60",
            "5V": "45",
            "0V": "0",
            "reverse_hysteresis": "5"
        }
        config['microcontroller'] = {
            "vid": "0x16C0",
            "pid": "0x0486"
        }
        config['probe'] = dict({
            "device": "Intel Core i7-6600U",
            "sensor": "CPU Package"
        },**dict((t, "false" if t != "cpu" else "true") for t in types))
        with open('thermostat.ini', 'w') as configfile:
            config.write(configfile)
        win32ui.MessageBox("Config file 'thermostat.ini' not found. A new one has been created",
                           "Startup Error", win32con.MB_ICONERROR)
        sys.exit(1)


def main(argv=None):
    types = ["cpu", "ram", "gpu", "mobo", "disk"]
    
    config = configparser.ConfigParser()
    config.read('thermostat.ini')
    
    verify_config(config)
    
    parser = argparse.ArgumentParser(description='HTPC Thermostat')
    parser.add_argument('--hidden', action='store_true', help='Do not display the track icon')
    parser.add_argument('--logfile', required=False, help='Enable debug logging to file')
    args = parser.parse_args()

    try:
        sensor_types = dict((t, config.getboolean("probe", t, fallback=False)) for t in types)
        s = TemperatureSensor(**dict(config["probe"], **sensor_types))
        w = Win32HID(device=VoltageSwitch, **config['microcontroller'])
        t = Thermostat(**config['thermostat'])
    except:
        win32ui.MessageBox(sys.exc_info()[1].args[0], "Startup Error", win32con.MB_ICONERROR)
        sys.exit(1)
    
    if args.logfile:
        logging.basicConfig(
            format='%(asctime)-15s [%(levelname)s] %(threadName)s %(name)s - %(message)s',
            level=logging.DEBUG,
            filename=args.logfile)
    
    if args.hidden:
        loop = threading.Thread(target=mainloop, args=(w, s, t))
        loop.setDaemon(True)
        loop.name = "VoltageSwitch"
        loop.start()
    else:
        t = TrayThermostat(w, s, t)
        t.name = "VoltageSwitch"
        t.setDaemon(True)
        t.start()
    
    # Having trouble with thread scope; this only runs in main thread
    # or when class creating window is child of threading.Thread.
    # PumpMessages runs until PostQuitMessage() is called by someone.
    win32gui.PumpMessages()


if __name__ == '__main__':
    main()
	


# The order of this list is important since it actually represents the index
# of each sensor, ie sensor.Index == index in this list
#openhardwaremonitor_sensortypes = ['Voltage','Clock','Temperature','Load','Fan','Flow','Control','Level','Factor','Power','Data','SmallData']
