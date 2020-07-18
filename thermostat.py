#!/usr/bin/env python3

# Base modules needed by thermostat
import configparser
import sys
import time
import traceback

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

# To run the win32 window messagepump should be in its own thread
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
            raise Exception("Arugments must be in pairs")
        
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
        super().__init__(vid, pid, serial, path)
    
    def set12v(self):
        print("Setting voltage switch to 12V")
        return self.write(VoltageSwitch.v12)
    
    def set5v(self):
        print("Setting voltage switch to 5V")
        return self.write(VoltageSwitch.v5)
    
    def set0v(self):
        print("Setting voltage switch to 0V")
        return self.write(VoltageSwitch.v0)


class TemperatureSensor():
    def __init__(self, device=None, sensor=None):
        self.__handle = Hardware.Computer()
        self.__handle.CPUEnabled = True
        #self.__handle.MainboardEnabled = True
        #self.__handle.RAMEnabled = True
        #self.__handle.GPUEnabled = True
        #self.__handle.HDDEnabled = True
        self.__handle.Open()
        
        if len(self.__handle.Hardware) < 1:
            raise Exception("No hardware was detected")

        try:
            self.__hardware = next(h for h in self.__handle.Hardware if h.Name == device)
        except:
            raise Exception("Device not detected")
            
        # Poll device
        self.__hardware.Update()

        try:
            self.__sensor = next(s for s in self.__hardware.Sensors if s.Index == 2 and s.Name == sensor)
        except:
            raise Exception("Sensor not found")
        
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
            self.__ids = list(map(lambda x,y: f"{x}_{y:04x}".encode(),["vid","pid"],[v,p]))
            
            # Always use path for connection; check for device at startup
            self.__path = None
            for d in hid.enumerate(vid=v, pid=p):
                if self.__matchingdevice(d["path"]):
                    super().set()
                    break
        else:
            raise Exception("VID, PID or device path required")
        
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
        #print("Device change notification:", wParam, str(info))
        
        if (info.devicetype == win32con.DBT_DEVTYP_DEVICEINTERFACE and
            self.__matchingdevice(info.name)):
            if wParam == win32con.DBT_DEVICEREMOVECOMPLETE:
                print("Device is being removed")
                self.clear()
            elif wParam == win32con.DBT_DEVICEARRIVAL:
                print("Device is being added")
                self.set()
        return True
        

def unknown_switch(*args, **kwargs):
    raise Exception("Unknown switch state")


def nop(*args, **kwargs):
    pass


def mainloop(w32hid, sensor, thermostat):
    while True:
        try:
            # Wait for the device, connect, then enter the context
            with w32hid.device() as vs:
                # On the first reading always set the switch state
                changes = False
            
                switch = {
                    '12v': vs.set12v,
                    '5v': vs.set5v,
                    '0v': vs.set0v,
                    None: nop,
                }
                
                while True:
                    time.sleep(2)
                    
                    # Set the switch state by feeding sensor data into thermostat
                    switch.get(thermostat.mode(sensor.reading(), changes), unknown_switch)()
                    
                    # From now on, only set switch state on changes
                    changes = True
        except:
            traceback.print_exc()
            time.sleep(30)


def main(argv=None):
    sections = ['thermostat', 'microcontroller', 'probe']
    config = configparser.ConfigParser()
    config.read('thermostat.ini')
    
    if not all(config.has_section(s) for s in sections):
        print("Config file 'thermostat.ini' not found; creating a new one.")
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
        config['probe'] = {
            "device": "Intel Core i7-6600U",
            "sensor": "CPU Package"
        }
        with open('thermostat.ini', 'w') as configfile:
            config.write(configfile)
        sys.exit(1)

    t = Thermostat(**config['thermostat'])
    w = Win32HID(device=VoltageSwitch, **config['microcontroller'])
    s = TemperatureSensor(**config['probe'])
    
    loop = threading.Thread(target=mainloop, args=(w, s, t))
    loop.setDaemon(True)
    loop.name = "VoltageSwitch"
    loop.start()
    
    # Having trouble with thread scope; this only runs in main thread
    # or when class creating window is child of threading.Thread.
    # PumpMessages runs until PostQuitMessage() is called by someone.
    win32gui.PumpMessages()


if __name__ == '__main__':
    main()
	


# The order of this list is important since it actually represents the index
# of each sensor, ie sensor.Index == index in this list
#openhardwaremonitor_sensortypes = ['Voltage','Clock','Temperature','Load','Fan','Flow','Control','Level','Factor','Power','Data','SmallData']
