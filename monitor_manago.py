import pywintypes
import screen_brightness_control as sbc
import win32api
import win32con
from PIL import Image
from pystray import MenuItem, Icon


def change_resolution():
    # get displaydevice object of the primary monitor
    primary_monitor = win32api.EnumDisplayDevices()

    # get devmodew object containing monitor information
    settings = win32api.EnumDisplaySettings(primary_monitor.DeviceName, -1)

    # check current resolution
    height = settings.PelsHeight
    width = settings.PelsWidth

    # set variables for new resolution
    # if the resolution is set to 1920x1080, change it to 1280x720
    # else (if the resolution is other than 1920x1080) change to 1920x1080 (native)
    if width == 1920 and height == 1080:
        new_width = 1280
        new_height = 720
    else:
        new_width = 1920
        new_height = 1080

    # execute of the resolution change
    # get the devmode object
    devmode = pywintypes.DEVMODEType()

    # assign variables to settings options
    devmode.PelsWidth = new_width
    devmode.PelsHeight = new_height

    # fill the options fields of devmode object
    devmode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT

    # use windows api to pass the display settings written to devmode object
    win32api.ChangeDisplaySettings(devmode, 0)


def change_refresh_rate():
    # get displaydevice object of the primary monitor
    primary_monitor = win32api.EnumDisplayDevices()

    # get devmodew object containing monitor information
    settings = win32api.EnumDisplaySettings(primary_monitor.DeviceName, -1)

    # check current refresh rate
    ref_rate = settings.DisplayFrequency

    # set variable for new refresh rate
    # if refresh rate is equal to 144hz, change it to 60hz and vice versa
    if ref_rate == 144:
        new_ref_rate = 60
    else:
        new_ref_rate = 144

    # execute of the refresh rate change
    # get the devmode object
    devmode = pywintypes.DEVMODEType()

    # assign variable to settings options
    devmode.DisplayFrequency = new_ref_rate

    # fill the options field of devmode object
    devmode.Fields = win32con.DM_DISPLAYFREQUENCY

    # use windows api to pass the display setting written to devmode object
    win32api.ChangeDisplaySettings(devmode, 0)


def rotate_screen():
    # get list of tuples containing notably the handle object and screen resolution information for each monitor
    monitors = win32api.EnumDisplayMonitors()

    # for each monitor tuple extract handle object
    for monitor in monitors:
        handle, _, _ = monitor

        # get monitor information for given handle object
        monitor_info = win32api.GetMonitorInfo(handle)

        # check if monitor is primary and its device name
        device = monitor_info['Device']
        is_primary = monitor_info['Flags'] == 1

        if is_primary:
            # get devmodew object for primary monitor
            devmodew = win32api.EnumDisplaySettings(device, win32con.ENUM_CURRENT_SETTINGS)

            # swap screen widths and heights between each other
            devmodew.PelsWidth, devmodew.PelsHeight = devmodew.PelsHeight, devmodew.PelsWidth

            # if the display orientation is equal to 0 (regular view) change the orientation value to 3 (rotate 270
            # degrees left to get vertical view - personal preference)
            if devmodew.DisplayOrientation == 0:
                devmodew.DisplayOrientation = 3
            else:
                devmodew.DisplayOrientation = 0

            # use windows api to pass the display setting written to devmode object
            win32api.ChangeDisplaySettingsEx(device, devmodew)


def set_primary(idx, monitors_list, monitor_arrangement):
    flag_for_primary = win32con.CDS_SET_PRIMARY | win32con.CDS_UPDATEREGISTRY | win32con.CDS_NORESET
    flag_for_sec = win32con.CDS_UPDATEREGISTRY | win32con.CDS_NORESET
    offset_x = - monitor_arrangement[idx][0]
    offset_y = - monitor_arrangement[idx][1]
    devices_num = len(monitors_list)

    for i in range(devices_num):
        devmode = win32api.EnumDisplaySettings(monitors_list[i].DeviceName, win32con.ENUM_CURRENT_SETTINGS)
        devmode.Position_x = monitor_arrangement[i][0] + offset_x
        devmode.Position_y = monitor_arrangement[i][1] + offset_y
        if (win32api.ChangeDisplaySettingsEx(monitors_list[i].DeviceName, devmode,
                                             flag_for_sec if i != idx else flag_for_primary) !=
                win32con.DISP_CHANGE_SUCCESSFUL):
            return False

    return win32api.ChangeDisplaySettingsEx() == win32con.DISP_CHANGE_SUCCESSFUL


def switch_primary():
    devices = []
    i = 0
    while True:
        try:
            device = win32api.EnumDisplayDevices(None, i, 0)
            if device.StateFlags & win32con.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP:
                devices.append(device)
            i += 1
        except:
            break

    for dev in devices:
        print("Name: ", dev.DeviceName)

    monitors = win32api.EnumDisplayMonitors()

    i = 0
    for monitor in monitors:
        handle, _, _ = monitor
        monitor_info = win32api.GetMonitorInfo(handle)
        is_primary = monitor_info['Flags'] == 1
        if is_primary:
            break
        i += 1

    if i:

        monitor_arrangement = {
            0: (0, 0),
            1: (-1920, 0)
        }

        set_primary(0, devices, monitor_arrangement)
    else:

        monitor_arrangement = {
            0: (1920, 0),
            1: (0, 0)
        }

        set_primary(1, devices, monitor_arrangement)


def switch_brightness():
    # # get list of tuples containing notably the handle object and screen resolution information for each monitor
    monitors = win32api.EnumDisplayMonitors()

    # get primary monitor index
    i = 0
    for monitor in monitors:
        handle, _, _ = monitor
        monitor_info = win32api.GetMonitorInfo(handle)
        if monitor_info['Flags'] == 1:
            break
        i += 1

    # get current brightness setting for primary monitor
    curr_brightness = sbc.get_brightness()[i]

    # if the brightness is set to 100 - change it to 25, otherwise change it to 100
    if curr_brightness == 100:
        sbc.set_brightness(25, display=i)
    else:
        sbc.set_brightness(100, display=i)


def on_quit():
    # change icon visibility and hide it on system tray
    icon.visible = False
    icon.stop()


# assign logo image
image = Image.open('logo.png')

# create menu
menu = (
    MenuItem('Switch resolution', change_resolution),
    MenuItem('Switch refresh rate', change_refresh_rate),
    MenuItem('Switch primary monitor', switch_primary),
    MenuItem('Switch brightness', switch_brightness),
    MenuItem('Rotate screen', rotate_screen),
    MenuItem('Quit', on_quit)
)

# load and run icon with set title, menu and image
icon = Icon("MonitorManago", image, "MonitorManago", menu)
icon.run()
