if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class NativeMethods {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
    public static extern void SetThreadExecutionState(uint esFlags);
    [DllImport("Kernel32.dll")]
	public static extern IntPtr GetConsoleWindow();
	[DllImport("user32.dll")]
	public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
}
'@
$consolePtr = [NativeMethods]::GetConsoleWindow()
[NativeMethods]::ShowWindow($consolePtr, 0) > $Null
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.Stop = $false
$syncHash.Pause = $false
$syncHash.Abort = $false
$ds = New-Object System.Data.Dataset
$null = $ds.ReadXml("$PSScriptRoot\skills.xml")
if(!(Test-Path -Path "$PSScriptRoot\keybinds.json" -PathType Leaf) -or !(Test-Path -Path "$PSScriptRoot\controls.json" -PathType Leaf)) {
    [System.Windows.MessageBox]::Show('Please configure your keybinds before using this tool','Information','OK','Information')
}

. "$PSScriptRoot\ui.ps1"

Get-ChildItem -Path "$PSScriptRoot\modules" | ForEach-Object {. $($_.FullName)}

$syncHash.Window.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
$syncHash.Window.Dispose()