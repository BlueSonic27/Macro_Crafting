if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs;exit }

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()
$NativeMethods = Add-Type -name user32 -passThru -MemberDefinition '
    [DllImport("user32.dll")]
    public static extern bool PostMessageA(int hWnd, int hMsg, int wParam, int lParam);
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(IntPtr ZeroOnly, string lpWindowName);
    [DllImport("user32.dll")]
    public static extern bool BlockInput(bool fBlockIt);
    [DllImport("user32.dll")]
    public static extern IntPtr GetDC(IntPtr hwnd);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern Int32 ReleaseDC(IntPtr hwnd, IntPtr hdc);
    [DllImport("gdi32.dll", SetLastError = true)]
    public static extern uint GetPixel(IntPtr dc, int x, int y);
    [DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
    public static extern void SetThreadExecutionState(uint esFlags);
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = $NativeMethods::GetConsoleWindow()
[void]$NativeMethods::ShowWindow($consolePtr, 0)

$syncHash = [hashtable]::Synchronized(@{})
$syncHash.Stop = $false
$syncHash.Pause = $false
$AssemblyEntry = New-Object System.Management.Automation.Runspaces.SessionStateAssemblyEntry -ArgumentList System.Windows.Forms
    
$initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$initialSessionState.Assemblies.Add($AssemblyEntry)
$initialSessionState.ExecutionPolicy = 4

$Runspace = [runspacefactory]::CreateRunspace()
$Runspace.ApartmentState = "STA"
$Runspace.ThreadOptions = "ReuseThread"
$Runspace.Open()
$Runspace.SessionStateProxy.SetVariable("NativeMethods",$NativeMethods)
$Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)

$Powershell = [powershell]::Create($initialSessionState)
$Powershell.Runspace = $Runspace

$ds = New-Object System.Data.Dataset
$null = $ds.ReadXml("$PSScriptRoot\skills.xml")
if(!(Test-Path -Path "$PSScriptRoot\keybinds.json" -PathType Leaf) -or !(Test-Path -Path "$PSScriptRoot\controls.json" -PathType Leaf)) {
    [System.Windows.MessageBox]::Show('Please configure your keybinds before using this tool','Information','OK','Information')
}

. "$PSScriptRoot\ui.ps1"

Get-ChildItem -Path "$PSScriptRoot\modules" | ForEach-Object {. $($_.FullName)}

$syncHash.Window.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
$syncHash.Window.Dispose()