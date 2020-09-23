VERSION 5.00
Begin VB.UserControl API 
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   Picture         =   "API.ctx":0000
   ScaleHeight     =   1815
   ScaleWidth      =   2325
   Begin VB.Shape Shape1 
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   15
      Picture         =   "API.ctx":1042
      Top             =   15
      Width           =   435
   End
End
Attribute VB_Name = "API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'by Martin McCormick
Dim a123
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Const SND_APPLICATION = &H80
Private Const SND_ALIAS = &H10000
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000
Private Const SND_LOOP = &H8
Private Const SND_MEMORY = &H4
Private Const SND_NODEFAULT = &H2
Private Const SND_NOSTOP = &H10
Private Const SND_NOWAIT = &H2000
Private Const SND_PURGE = &H40
Private Const SND_RESOURCE = &H40004
Private Const SND_SYNC = &H0
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2
Const Internet_Autodial_Force_Unattended As Long = 2
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal _
uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As _
Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Dim retval
Private Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private XT(1) As Single, YT As Single, M As Single
Private XScreen As Single, YScreen As Single
Private i As Integer, II As Integer
Private Const MAX_DELOCATION = 250
Private Const PUPIL_DISTANCE = 30
Option Explicit
Dim timeval
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B
Private Const EWX_SHUTDOWN As Long = 1
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Const EWX_REBOOT As Long = 2
Private Const EWX_LOGOFF As Long = 0
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1
Private Const SW_SHOW = 5
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Dim dgf
Const SPI_SCREENSAVERRUNNING = 97
Private Declare Function BringWindowToTop Lib "user32.dll" (ByVal hwnd As Long) As Long
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4


Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Dim nid As NOTIFYICONDATA
Private Declare Function SwapMouseButton Lib "user32.dll" (ByVal bSwap As Long) As Long

'''''''''''''''''''''''''''''''''''''''''''''''
'Event Declarations:
Event ShutDown(ShutDown)
Event Restart(Restart)
Event LogOff(LogOff)
Event TaskBarHide(TaskBarHide)
Event TasksBarShow(TasksBarShow)
Event ScreenSaverOn(ScreenSaverOn)
Event ScreenSaverOff(ScreenSaverOff)
Event DesktopIconsHide(DesktopIconsHide)
Event ALTCTRLDELEnabled(ALT_CTRL_DEL_Enabled)
Event ALTCTRLDELDisabled(ALT_CTRL_DEL_Disabled)
Event OpenCDROM(OpenCDROM)
Event EmptRecycle(EmptRecycle)
Event MinimizeAll(MinimizeAll)
Event OpenExplore(OpenExplore)
Event FindFiles(FindFiles)
Event OpenInternetBrowser(OpenInternetBrowser)
Event InternetConnect(InternetConnect)
Event InternetDiconnect(InternetDiconnect)
Event SendEmail(SendEmail)
Event AddRemove(Add_Remove)
Event AddHardWare(Add_HardWare)
Event TimeDateSettings(Time_Date_Settings)
Event RegionalSettings(Regional_Settings)
Event DisplaySettings(Display_Settings)
Event InternetSetting(Internet_Settings)
Event KeyboardSettings(Keyboard_Settings)
Event MouseSettings(Mouse_Settings)
Event ModemSettings(Modem_Settings)
Event SystemSettings(System_Settings)
Event NetworkSettings(Network_Settings)
Event PasswordSettings(Password_Settings)
Event SoundsSettings(Sounds_Settings)
Event ShowAbout(ShowAbout)
Event CopyaFile(Copy_File)
Event DeleteaFile(Delete_File)
Event MoveaFile(Move_File)
Event FlipMouseButtons(FlipMouseButtons)
Event FormToTop(FormOnTop)
Event ShutDownDIALOG(ShutDown_DIALOG)
Event SleepMillisecs(Sleep_Millisecs)
Event CursorHidden(Cursor_Hide)
Event CursorShown(Cursor_Show)
Event WAVFilePlayed(PlayWAVFile)
Event ObjectEnabled(EnableObject)
Event ObjectDisabled(DisableObject)

Function ShutDown()
Dim lngresult
lngresult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Function
Function Restart()
Dim lngresult
lngresult = ExitWindowsEx(EWX_REBOOT, 0&)
End Function
Function LogOff()
Dim lngresult
lngresult = ExitWindowsEx(EWX_LOGOFF, 0&)
End Function
Function TaskBarHide()
Dim rtn
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Function
Function TaskBarShow()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Function
Function ScreenSaverOn()
ToggleScreenSaverActive (True)
End Function
Function ScreenSaverOff()
ToggleScreenSaverActive (False)
End Function
Public Function ToggleScreenSaverActive(Active As Boolean) _
   As Boolean
Dim lActiveFlag As Long
Dim retval As Long

lActiveFlag = IIf(Active, 1, 0)
retval = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, _
   lActiveFlag, 0, 0)
ToggleScreenSaverActive = retval > 0

End Function
Function DesktopIconsShow()
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 5
End Function
Function DesktopIconsHide()
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 0
End Function
Function ALT_CTRL_DEL_Enabled()
callme (False)
End Function
Function ALT_CTRL_DEL_Disabled()
callme (True)
End Function
Private Sub callme(huh As Boolean)
Dim gd
gd = SystemParametersInfo(97, huh, CStr(1), 0)
End Sub
Function OpenCDROM()
Dim lngReturn As Long
Dim strReturn As Long
lngReturn = mciSendString("set CDAudio door open", strReturn, 127, 0)
End Function
Function EmptRecycle()
Dim retval
retval = SHEmptyRecycleBin(UserControl.hwnd, "", SHERB_NOPROGRESSUI)
End Function
Function MinimizeAll()
Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(77, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Function
Function OpenExplore()
Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(69, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Function
Function FindFiles()
Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(70, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Function
Function OpenInternetBrowser()
ShellExecute hwnd, "open", "", vbNullString, vbNullString, conSwNormal
End Function
Function InternetConnect()
Dim lResult As Long
lResult = InternetAutodial(Internet_Autodial_Force_Unattended, 0&)
End Function
Function InternetDiconnect()
Dim lResult As Long
lResult = InternetAutodialHangup(0&)
End Function
Function SendEmail()
ShellExecute hwnd, "open", "mailto:", vbNullString, vbNullString, SW_SHOW
End Function
Function Add_Remove()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)
End Function

Function Add_HardWare()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)
End Function

Function Time_Date_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
End Function

Function Regional_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5)
End Function

Function Display_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
End Function

Function Internet_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Function

Function Keyboard_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)
End Function

Function Mouse_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0", 5)
End Function
Function Modem_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl", 5)
End Function
Function System_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
End Function
Function Network_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)
End Function
Function Password_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL password.cpl", 5)
End Function
Function Sounds_Settings()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5)
End Function
Function ShowAbout()
Dim about
about = MsgBox("CompControl.ocx was created by Martin McCormick and can be download from Http://www.planet-source-code.com any questions or comments should be sent to Slimshady_5_5_5@hotmail.com", vbOKOnly + vbInformation, "About")

End Function
Function Copy_File(FileToCopy, Destination)
retval = CopyFile(FileToCopy, Destination, 1)
End Function
Function Delete_File(file)
retval = DeleteFile(file)
End Function
Function Move_File(FileToMove, Destination)
retval = MoveFile(FileToMove, Destination)
End Function
Function FlipMouseButtons()
retval = SwapMouseButton(1)
End Function
Function FormOnTop(Form, X, Y, Width, Height)
Dim hnd
hnd = Form.hwnd
SetWindowPos hnd, conHwndTopmost, X, Y, Width, Height, conSwpNoActivate Or conSwpShowWindow
End Function
Function Sleep_Millisecs(LengthInMilliseconds)
Sleep (LengthInMilliseconds)
End Function
Function ShutDown_DIALOG()
ShutDown_DIALOG = SHShutDownDialog(0)
End Function
Function Cursor_Show()
ShowCursor (True)
End Function
Function Cursor_Hide()
ShowCursor (False)
End Function

Function Path_Exist(path)
Path_Exist = PathFileExists(path)
End Function
Function PlayWAVFile(file)
PlaySound file, ByVal 0&, SND_FILENAME Or SND_ASYNC
End Function
Function EnableObject(object)
EnableObject = EnableWindow(object.hwnd, True)
End Function
Function DisableObject(object)
DisableObject = EnableWindow(object.hwnd, False)
End Function

Private Sub UserControl_Resize()
UserControl.Width = Shape1.Width
UserControl.Height = Shape1.Height
End Sub
