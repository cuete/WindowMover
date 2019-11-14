#Saves configuration object data to a json file
Function SaveConfig($configpath)
{
  $global:config | ConvertTo-Json | Out-File $configpath
}

#Queries the Windows API for a given window's size and position
#Saves it to a configuration object
Function GetWindowData($processName)
{
  $handle = (Get-Process -Name $processName).MainWindowHandle
  #Gets the process instance with a handle number 
  $handle | ForEach-Object {if ($h -ne 0) { $takeHandle = $h } }

  foreach ($h in $handle)
  {
      if ($h -ne 0)
      {
          $takeHandle = $h
      }
  }

  Write-Host "Recording data for "$processName"..."

  # This is to get the current process state
  $Rectangle = New-Object RECT
  $Return = [Window]::GetWindowRect($takeHandle,[ref]$Rectangle)
  $windowObject = New-Object -TypeName psobject
  $windowObject | Add-Member -MemberType NoteProperty -Name processname -Value $processName
  $windowObject | Add-Member -MemberType NoteProperty -Name x -Value $Rectangle.Left
  $windowObject | Add-Member -MemberType NoteProperty -Name y -Value $Rectangle.Top
  $windowObject | Add-Member -MemberType NoteProperty -Name width -Value ($Rectangle.Right - $Rectangle.Left)
  $windowObject | Add-Member -MemberType NoteProperty -Name height -Value ($Rectangle.Bottom - $Rectangle.Top)
  $json = $windowObject | ConvertTo-Json
  $global:config.Add($json)
}

# =======
#  MAIN
# =======

#Window object type definition to interact with the Windows API
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

$global:config = New-Object System.Collections.ArrayList
$configpath = $env:USERPROFILE + "\windowlayout.config"

#Creating an array with the process windows to record size and position
#To read a list of currently running processes use `Get-Process`
$processes = "Teams","Outlook","WindowsTerminal"

#Read and record window sizes and positions
$processes | ForEach-Object { GetWindowData($_) }
SaveConfig($configpath)
