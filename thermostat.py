#!/usr/bin/env python3

# https://stackoverflow.com/a/49909330

# Must be run as administrator!

#import os (set current directory for hidapi.dll)
#import sys (set system.path for clr.AddReference)

import configparser
import sys
import time
import ctypes

if sys.maxsize > 2**32:
    raise Exception("OpenHardwareMonitorLib is only 32-bit")

# Needs 32-bit OpenHardwareMonitorLib.dll; must run as administrator!
import clr

# Needs 32-bit hidapi.dll and lib in same directory as binary, or os.getcwd()
# https://docs.microsoft.com/en-us/windows/win32/dlls/dynamic-link-library-search-order#standard-search-order-for-desktop-applications
import hid

# Open up the OpenHardwareMonitor library with CLR
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
