#!/usr/bin/env python3

# Base modules needed by thermostat
import configparser
import sys
import time

# OpenHardwareMonitorLib.dll is only 32-bit, so application must be 32-bit!
if sys.maxsize > 2**32:
    raise Exception("OpenHardwareMonitorLib is only 32-bit")

# pywin32 is used to create a message-only window to receive USB device changes
# https://github.com/mhammond/pywin32/blob/master/win32/Demos/win32gui_devicenotify.py
# https://docs.microsoft.com/en-us/windows/win32/winmsg/window-features#message-only-windows
import win32gui
import win32con
import win32file
import win32api
import win32gui_struct

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
        
        self.__temperature_map = {}
        
        for i in range(0, len(args), 2):
            self.__temperature_map[int(args[i+1])] = args[i]
            
        for k, v in kwargs.items():
            self.__temperature_map[int(v)] = k
        
        self.__min = min(self.__temperature_map.keys())
        self.__max = max(self.__temperature_map.keys())
        
        # To make temperature lookups O(c) create a mode map for all
        # temperatures between min and max
        for temp in range(self.__min, self.__max+1):
            if temp in self.__temperature_map.keys():
                mode = self.__temperature_map[temp]
                continue
            self.__temperature_map[temp] = mode
        
    def mode(self, temperature):
        if temperature <= self.__min:
            return self.__temperature_map[self.__min]
        elif temperature >= self.__max:
            return self.__temperature_map[self.__max]
        return self.__temperature_map[temperature]


class VoltageSwitch():
    v12 = ctypes.create_string_buffer(b'\x01', 64)
    v5 = ctypes.create_string_buffer(b'\x00', 64)
    v0 = ctypes.create_string_buffer(b'\x02', 64)

    def __init__(self, vid=0x16C0, pid=0x0486):
        try:
            self.__device = hid.Device(vid=int(vid), pid=int(pid))
        except ValueError:
            self.__device = hid.Device(vid=int(vid, 0), pid=int(pid, 0))
    
    def close(self):
        self.__device.close()
    
    def set12v(self):
        return self.__device.write(VoltageSwitch.v12)
    
    def set5v(self):
        return self.__device.write(VoltageSwitch.v5)
    
    def set0v(self):
        return self.__device.write(VoltageSwitch.v0)


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


class Win32HID():
    # **Everything is added to self to keep python from doing GC!!**

    # USB Device works, but HID device is more specific. Both listed anyway
    GUID_DEVINTERFACE_USB_DEVICE = "{A5DCBF10-6530-11D2-901F-00C04FB951ED}"
    GUID_DEVINTERFACE_HID = "{4D1E55B2-F16F-11CF-88CB-001111000030}"
    
    # HID Device type string
    HID_DEVICE_CLASS = '\\\\?\\HID'
    
    # Teensy multiple-interface ID (00 for HID, and 01 for serial emulation)
    # https://docs.microsoft.com/en-us/windows-hardware/drivers/install/standard-usb-identifiers#multiple-interface-usb-devices
    INTERFACE = 'MI_00'

    def __init__(self, vid=0x16C0, pid=0x0486):
        # First convert IDs into a format we can detect in device name
        try:
            self.ids = [
                f"VID_{hex(int(vid))[2:].upper()}",
                f"PID_{hex(int(pid))[2:].upper()}",
                Win32HID.INTERFACE
            ]
        except:
            self.ids = [
                f"VID_{hex(int(vid, 0))[2:].upper()}",
                f"PID_{hex(int(pid, 0))[2:].upper()}",
                Win32HID.INTERFACE
            ]
    
        # Create a window class for receiving messages
        self.wc = win32gui.WNDCLASS()
        self.wc.hInstance = win32api.GetModuleHandle(None)
        self.wc.lpszClassName = "usbnotifier"
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
        
        # Run the message loop
        #while True:
        #    try:
        #        win32gui.PumpWaitingMessages()
                # TODO should use win32event.MsgWaitForMultipleObjects
        #        time.sleep(0.25)
        #    except:
        #        win32gui.DestroyWindow(self.hwnd)
        #        win32gui.UnregisterClass(self.wc.lpszClassName, None)

                
    def messageloop(self):
        #https://stackoverflow.com/questions/51535713/pumpmessages-in-new-thread
        # PumpMessages runs until PostQuitMessage() is called by someone.
        win32gui.PumpMessages()   
                

    def __matchingdevice(self, name):
        # Name looks like:
        # '\\\\?\\HID#VID_16C0&PID_0486&MI_01#7&2d928156&0&0000#{4d1e55b2-f16f-11cf-88cb-001111000030}'
        chunks = name.split("#")
        if len(chunks) > 2 and chunk[0] == Win32HID.HID_DEVICE_CLASS:
            ids = chunks[1].split("&")
            if all(id in ids for id in self.ids):
                return True
        return False

    # WM_DEVICECHANGE message handler.
    def __devicechange(self, hWnd, msg, wParam, lParam):
        info = win32gui_struct.UnpackDEV_BROADCAST(lParam)
        print("Device change notification:", wParam, str(info))
        
        if info.devicetype == win32con.DBT_DEVTYP_DEVICEINTERFACE:
            if wParam == win32con.DBT_DEVICEREMOVECOMPLETE:
                print("Device is being removed")
            elif wParam == win32con.DBT_DEVICEARRIVAL:
                print("Device is being added")
        return True
        

def unknown_switch(*args, **kwargs):
    raise Exception("Unknown switch state")


def main(argv=None):
    config = configparser.ConfigParser()
    config.read('thermostat.ini')
    
    t = Thermostat(**config['thermostat'])
    vs = VoltageSwitch(**config['microcontroller'])
    s = TemperatureSensor(**config['probe'])
    
    mode = t.mode(s.value)
    switch = {
        '12V': vs.set12v,
        '5V': vs.set5v,
        '0V': vs.set0v
    }
    switch.get(mode, unknown_switch)(mode)
    
    # TODO Some hysteresis to prevent relay chatter if temperature is
    # oscillating at switch point.
    while True:
        time.sleep(2)
        new_mode = t.mode(s.reading())
        if mode != new_mode:
            mode = new_mode
            switch.get(mode, unknown_switch)(mode)


if __name__ == '__main__':
    main()
	


# The order of this list is important since it actually represents the index
# of each sensor, ie sensor.Index == index in this list
#openhardwaremonitor_sensortypes = ['Voltage','Clock','Temperature','Load','Fan','Flow','Control','Level','Factor','Power','Data','SmallData']
