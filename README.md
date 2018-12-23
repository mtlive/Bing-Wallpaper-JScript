# Bing-Wallpaper-JScript
Set daily Bing background image as your desktop background using a Windows Script written in JScript.

It also pops a message box with the explanation of the image.

## Why JScript?
Because it doesn't relay on any third-party component and the required components are already shipped with Windows.
Why not Powershell? Because Powershell is slow to load.

## Bugs:
We use `RUNDLL32.EXE USER32.DLL, UpdatePerUserSystemParameters` to refresh the desktop background, which sometimes doesn't work properly.

## If you want to silent the script
use `//B` (batch mode) switch.
