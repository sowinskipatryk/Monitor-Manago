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
    devmodew = win32api.EnumDisplaySettings(primary_monitor.DeviceName, win32con.ENUM_CURRENT_SETTINGS)

    # check current resolution
    height = devmodew.PelsHeight
    width = devmodew.PelsWidth

    # set variables for new resolution
    # if the resolution is set to 1920x1080, change it to 1280x720
    # else (if the resolution is other than 1920x1080) change to 1920x1080 (native)
    if width == 1920 and height == 1080:
        new_width = 1280
        new_height = 720
    else:
        new_width = 1920
        new_height = 1080

    # assign variables to settings options
    devmodew.PelsWidth = new_width
    devmodew.PelsHeight = new_height

    # fill the options fields of devmode object
    devmodew.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT

    # use windows api to pass the display settings written to devmode object
    win32api.ChangeDisplaySettings(devmodew, 0)


def change_refresh_rate():
    # get displaydevice object of the primary monitor
    primary_monitor = win32api.EnumDisplayDevices()

    # get devmodew object containing monitor information
    devmodew = win32api.EnumDisplaySettings(primary_monitor.DeviceName, win32con.ENUM_CURRENT_SETTINGS)

    # check current refresh rate
    ref_rate = devmodew.DisplayFrequency

    # set variable for new refresh rate
    # if refresh rate is equal to 144hz, change it to 60hz and vice versa
    if ref_rate == 144:
        new_ref_rate = 60
    else:
        new_ref_rate = 144

    # assign variable to settings options
    devmodew.DisplayFrequency = new_ref_rate

    # fill the options field of devmode object
    devmodew.Fields = win32con.DM_DISPLAYFREQUENCY

    # use windows api to pass the display settings written to devmode object
    win32api.ChangeDisplaySettings(devmodew, 0)


def rotate_screen():
    # get displaydevice object of the primary monitor
    primary_monitor = win32api.EnumDisplayDevices()

    # get devmodew object containing monitor information
    devmodew = win32api.EnumDisplaySettings(primary_monitor.DeviceName, win32con.ENUM_CURRENT_SETTINGS)

    # swap screen widths and heights between each other
    devmodew.PelsWidth, devmodew.PelsHeight = devmodew.PelsHeight, devmodew.PelsWidth

    # if the display orientation is equal to 0 (regular view) change the orientation value to 3 (rotate 270
    # degrees left to get vertical view - personal preference)
    if devmodew.DisplayOrientation == 0:
        devmodew.DisplayOrientation = 3
    else:
        devmodew.DisplayOrientation = 0

    # use windows api to pass the display settings written to devmode object
    win32api.ChangeDisplaySettingsEx(primary_monitor.DeviceName, devmodew)


def set_primary(index, monitors_list, monitor_arrangement):
    # set flags for primary and secondary monitor
    primary_flag = win32con.CDS_SET_PRIMARY | win32con.CDS_UPDATEREGISTRY | win32con.CDS_NORESET
    secondary_flag = win32con.CDS_UPDATEREGISTRY | win32con.CDS_NORESET

    # set the x and y offset to put monitors in place
    offset_x = - monitor_arrangement[index][0]
    offset_y = - monitor_arrangement[index][1]
    monitors_num = len(monitors_list)

    # iterate over monitors changing their position and state between primary and secondary
    for i in range(monitors_num):
        # get devmodew object containing monitor information
        devmodew = win32api.EnumDisplaySettings(monitors_list[i].DeviceName, win32con.ENUM_CURRENT_SETTINGS)

        # set the x and y position of monitor
        devmodew.Position_x = monitor_arrangement[i][0] + offset_x
        devmodew.Position_y = monitor_arrangement[i][1] + offset_y

        # use windows api to pass the display settings written to devmodew object
        if i != index:
            win32api.ChangeDisplaySettingsEx(monitors_list[i].DeviceName, devmodew, secondary_flag)
        else:
            win32api.ChangeDisplaySettingsEx(monitors_list[i].DeviceName, devmodew, primary_flag)

    win32api.ChangeDisplaySettingsEx()


def switch_primary():

    # create a list to store display device objects
    display_devices = []

    # iterate over display devices
    i = 0
    while True:
        try:
            device = win32api.EnumDisplayDevices(None, i, 0)
            # if the display device is attached to desktop and running, add it to the list
            if device.StateFlags & win32con.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP:
                display_devices.append(device)
            i += 1
        # break the loop when there are no more display devices
        except pywintypes.error:
            break

    # get a list of tuples containing notably the handle object and screen resolution information for each monitor
    monitors = win32api.EnumDisplayMonitors()

    # get primary monitor index by iterating over the monitors list and checking if the Flags attribute is set to 1
    i = 0
    for monitor in monitors:
        handle, _, _ = monitor
        monitor_info = win32api.GetMonitorInfo(handle)
        is_primary = monitor_info['Flags'] == 1
        if is_primary:
            break
        i += 1

    # set next monitor as primary (if last monitor is primary, set primary to first monitor -> index 0)
    if i + 1 == len(monitors):
        new_primary = 0
    else:
        new_primary = i + 1

    # change the primary monitor to secondary and vice versa and also set monitor arrangement to be able to move the
    # mouse cursor between monitors on the linked axis
    if i:

        monitor_arrangement = {
            0: (0, 0),
            1: (-1920, 0)
        }

        set_primary(new_primary, display_devices, monitor_arrangement)

    else:

        monitor_arrangement = {
            0: (1920, 0),
            1: (0, 0)
        }

        set_primary(new_primary, display_devices, monitor_arrangement)


def switch_brightness():
    # get a list of tuples containing notably the handle object and screen resolution information for each monitor
    monitors = win32api.EnumDisplayMonitors()

    # get primary monitor index by iterating over the monitors list and checking if the Flags attribute is set to 1
    i = 0
    for monitor in monitors:
        handle, _, _ = monitor
        monitor_info = win32api.GetMonitorInfo(handle)
        if monitor_info['Flags'] == 1:
            break
        i += 1

    # get current brightness settings for primary monitor
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
    MenuItem('Switch brightness', switch_brightness),
    MenuItem('Switch orientation', rotate_screen),
    MenuItem('Switch primary monitor', switch_primary),
    MenuItem('Switch refresh rate', change_refresh_rate),
    MenuItem('Switch resolution', change_resolution),
    MenuItem('Quit', on_quit)
)

# load and run icon with set title, menu and image
icon = Icon("MonitorManago", image, "MonitorManago", menu)
icon.run()
