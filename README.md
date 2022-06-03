# Monitor Manago

This application is designed to allow you to change monitor resolution, refresh rate, swap primary display, change display orientation from landscape to portrait, and manipulate screen brightness from the system tray.
The script is written in Python but also Batch and Virtual Basic files are added to let the script run windowless and on system startup.
The system tray application is heavily based on the code from following site however functionality has been greatly expanded from being able to only change the resolution to much more.
As the second appendix I attach the comment that helped a lot in implementing the function responsible for changing the primary display monitor. Finally the last position is the site I used to create a simple logo for this application.

It should also be noted that the application was created for personal use so it contains some simplifications, for example changing the resolution is done between 1920x1080 and 1280x720, similarly the refresh rate is done between 60hz and 144hz, and changing the landscape to portrait orientation is done by rotating the screen 270 degrees as this is the option I use most often. It is possible to extend these options based on variables using lambda functions.

https://k0nze.dev/posts/change-resolution-python-system-tray/ <br>
https://stackoverflow.com/a/35816792/19256893 <br>
https://logomakr.com/
