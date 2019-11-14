#Reads json file into a configuration object 
Function ParseConfig($configpath)
{
  $jsontmp = Get-Content $configpath | Out-String | ConvertFrom-Json
  $global:config = {$jsontmp}.Invoke()
}

#Modifies a given process' window to a given size and position using the Windows API
Function MoveWindow($processjson)
{
  $process = $processjson | ConvertFrom-Json
  $handle = (Get-Process -Name $process.processname -ErrorAction SilentlyContinue).MainWindowHandle

  #Select the process instance with a valid PID 
  foreach ($h in $handle)
  {
      if ($h -ne 0)
      {
          $takeHandle = $h
      }
  }

  #Skip if the process is not currently running
  if($takeHandle) {
    Write-Host "Moving window"$process.processname"-"$handle"..."
    $return = [Window]::MoveWindow($takeHandle, $process.x, $process.y, $process.width, $process.height, $True)
    Write-Host "Success?" $return "`n"
  }
  else {
    Write-Host "Process "$process.processname "not running"
  }
}

# =======
#  MAIN
# =======

#Window object type definition to interact with the Windows API
Try{
      [void][Window]
} 
Catch {
  Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Window {
  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

  [DllImport("User32.dll")]
  public extern static bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);
}
public struct RECT
{
  public int Left;        // x position of upper-left corner
  public int Top;         // y position of upper-left corner
  public int Right;       // x position of lower-right corner
  public int Bottom;      // y position of lower-right corner
}
"@
}

$global:config = New-Object System.Collections.ArrayList
#Reading file with configured size and position values
#Use `winlayout_record.ps1` to record these values
$configpath = $env:USERPROFILE + "\windowlayout.config"

#Modify windows' size and position as spec'd in the configuration file
ParseConfig($configpath)
$global:config | ForEach-Object { MoveWindow($_) }
