import pywintypes
import screen_brightness_control as sbc
import win32api
import win32con
from PIL import Image
from pystray import MenuItem, Icon, Menu


def change_resolution(width, height):
    # get displaydevice object of the primary monitor
    primary_monitor = win32api.EnumDisplayDevices()

    # get devmodew object containing monitor information
    devmodew = win32api.EnumDisplaySettings(primary_monitor.DeviceName, win32con.ENUM_CURRENT_SETTINGS)

    # assign variables to settings options
    devmodew.PelsWidth = width
    devmodew.PelsHeight = height

    # fill the options fields of devmode object
    devmodew.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT

    # use windows api to pass the display settings written to devmode object
    win32api.ChangeDisplaySettings(devmodew, 0)


def change_refresh_rate(ref_rate):
    # get displaydevice object of the primary monitor
    primary_monitor = win32api.EnumDisplayDevices()

    # get devmodew object containing monitor information
    devmodew = win32api.EnumDisplaySettings(primary_monitor.DeviceName, win32con.ENUM_CURRENT_SETTINGS)

    # assign variable to settings options
    devmodew.DisplayFrequency = ref_rate

    # fill the options field of devmode object
    devmodew.Fields = win32con.DM_DISPLAYFREQUENCY

    # use windows api to pass the display settings written to devmode object
    win32api.ChangeDisplaySettings(devmodew, 0)


def rotate_screen(orient_val):
    # get displaydevice object of the primary monitor
    primary_monitor = win32api.EnumDisplayDevices()

    # get devmodew object containing monitor information
    devmodew = win32api.EnumDisplaySettings(primary_monitor.DeviceName, win32con.ENUM_CURRENT_SETTINGS)

    # get current display orientation
    curr_disp_orient = devmodew.DisplayOrientation

    # calculate the new display orientation
    new_disp_orient = curr_disp_orient + orient_val

    # lock the variable to range(4)
    if new_disp_orient > 3:
        new_disp_orient -= 4
    if new_disp_orient < 0:
        new_disp_orient += 4

    # if screen rotation by 90 degrees
    if orient_val % 2:
        # swap screen widths and heights between each other
        devmodew.PelsWidth, devmodew.PelsHeight = devmodew.PelsHeight, devmodew.PelsWidth

    # set new display orientation
    devmodew.DisplayOrientation = new_disp_orient

    primary_flag = win32con.CDS_SET_PRIMARY | win32con.CDS_UPDATEREGISTRY

    # use windows api to pass the display settings written to devmode object
    win32api.ChangeDisplaySettingsEx(primary_monitor.DeviceName, devmodew, primary_flag)


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


def change_brightness(brightness):
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

    # set brightness for primary monitor
    sbc.set_brightness(brightness, display=i)


def on_quit():
    # change icon visibility and hide it on system tray
    icon.visible = False
    icon.stop()


# assign logo image
image = Image.open('logo.png')

# create menu
menu = (
    MenuItem('Change brightness', Menu(
        MenuItem("100%", lambda: change_brightness(100)),
        MenuItem("80%", lambda: change_brightness(80)),
        MenuItem("60%", lambda: change_brightness(60)),
        MenuItem("40%", lambda: change_brightness(40)),
        MenuItem("20%", lambda: change_brightness(20)))),
    MenuItem('Change refresh rate', Menu(
        MenuItem("144Hz", lambda: change_refresh_rate(144)),
        MenuItem("120Hz", lambda: change_refresh_rate(120)),
        MenuItem("100Hz", lambda: change_refresh_rate(100)),
        MenuItem("60Hz", lambda: change_refresh_rate(60)))),
    MenuItem('Change resolution', Menu(
        MenuItem("1920x1080 (16:9)", lambda: change_resolution(1920, 1080)),
        MenuItem("1600x900 (16:9)", lambda: change_resolution(1600, 900)),
        MenuItem("1366x768 (16:9)", lambda: change_resolution(1366, 768)),
        MenuItem("1280x960 (4:3)", lambda: change_resolution(1280, 960)),
        MenuItem("1024x768 (4:3)", lambda: change_resolution(1024, 768)),
        MenuItem("800x600 (4:3)", lambda: change_resolution(800, 600)))),
    MenuItem('Rotate screen', Menu(
        MenuItem("180°", lambda: rotate_screen(2)),
        MenuItem("90° CW", lambda: rotate_screen(-1)),
        MenuItem("90° CCW", lambda: rotate_screen(1)))),
    MenuItem('Switch primary monitor', switch_primary),
    MenuItem('Quit', on_quit)
)

# load and run icon with set title, menu and image
icon = Icon("MonitorManago", image, "MonitorManago", menu)
icon.run()
