VERSION 5.00
Begin VB.UserControl CpuUsageControl 
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   PropertyPages   =   "CpuUsageMonitor.ctx":0000
   ScaleHeight     =   1125
   ScaleWidth      =   5085
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   0
   End
   Begin VB.PictureBox picUsage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   0
      Width           =   990
      Begin VB.Label lblCpuUsage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   960
      ScaleHeight     =   103.279
      ScaleMode       =   0  'User
      ScaleWidth      =   100.844
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "CpuUsageControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const NO_ERROR = 0

Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const HKEY_DYN_DATA = &H80000006
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS = 0&
Private Const SPACE = 5
Private Const BAR_WIDTH = 50
Private Const HWND_TOPMOST = -1&
Private Const HWND_NOTOPMOST = -2&
Private Const SWP_NOSIZE = &H1&
Private Const SWP_NOMOVE = &H2&
Private Const SWP_NOACTIVATE = &H10&
Private Const SWP_SHOWWINDOW = &H40&
Private Const THREAD_BASE_PRIORITY_MAX = 2
Private Const HIGH_PRIORITY_CLASS = &H80

Private Type LARGE_INTEGER
    dwLow As Long
    dwHigh As Long
End Type

Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type

Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As LARGE_INTEGER
    dwSpare(0 To 75) As Long
End Type

Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As LARGE_INTEGER
    liKeSystemTime As LARGE_INTEGER
    liExpTimeZoneBias  As LARGE_INTEGER
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private liOldIdleTime As LARGE_INTEGER
Private liOldSystemTime As LARGE_INTEGER
Private hKey As Long
Private dwDataSize As Long
Private dwCpuUsage As Byte
Private dwType As Long

Private mChartPoints(0 To 99) As Long
Private mCounter As Integer

Private mBarColor As Long
Private mBackColor As Long
Private mForeColor As Long
Private mPercent As Long
Private mLastPercent As Long


Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Started()
Public Event Stopped()

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
 Call ShellAbout(UserControl.hwnd, "Cpu Usage Monitor", "Developed by Mauricio Cunha (mcunha98@terra.com.br)" & vbCrLf & "Visit http://www.mcunha98.cjb.net", UserControl.Parent.icon)
End Sub

Public Property Get SystemName() As String
 If IsWinNT Then
  SystemName = "Windows NT"
 Else
  SystemName = "Windows 9x"
 End If
End Property

Public Sub StartMonitor()
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    
    If IsWinNT Then
        Call prvStartMonitorNT
    Else
        Call prvStartMonitor
    End If
    
    RaiseEvent Started
    tmrRefresh.Enabled = True
    tmrRefresh_Timer
End Sub

Public Sub StopMonitor()
    tmrRefresh.Enabled = False
    If IsWinNT Then
    
    Else
     Call prvStopMonitor
    End If
    For mCounter = LBound(mChartPoints) To UBound(mChartPoints) - 1
     mChartPoints(mCounter) = 0
    Next
    lblCpuUsage.Caption = "0%"
    prvRefresh 0
    RaiseEvent Stopped
End Sub

Private Sub lblCpuUsage_Click()
 RaiseEvent Click
End Sub

Private Sub lblCpuUsage_DblClick()
 RaiseEvent DblClick
End Sub

Private Sub lblCpuUsage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblCpuUsage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblCpuUsage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picGraph_Click()
 RaiseEvent Click
End Sub

Private Sub picGraph_DblClick()
 RaiseEvent DblClick
End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picGraph_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picUsage_Click()
 RaiseEvent Click
End Sub

Private Sub picUsage_DblClick()
 RaiseEvent DblClick
End Sub

Private Sub picUsage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picUsage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picUsage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub tmrRefresh_Timer()
    If IsWinNT Then
     mPercent = prvQueryNT
    Else
     mPercent = prvQuery
    End If
    If mPercent = -1 Then
        tmrRefresh.Enabled = False
        lblCpuUsage.Caption = "0%"
    Else
        mLastPercent = CLng(Replace(lblCpuUsage.Caption, "%", ""))
        prvRefresh mPercent
        lblCpuUsage.Caption = CStr(mPercent) + "%"
        If mLastPercent <> mPercent Then RaiseEvent Change
    End If
End Sub

Public Property Get PercentUsed() As Integer
 PercentUsed = mPercent
End Property

Private Sub UserControl_InitProperties()
 BarColor = &H808080
 BackColor = vbBlack
 Set Font = Ambient.Font
 ForeColor = vbGreen
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 BarColor = PropBag.ReadProperty("BarColor", &H808080)
 BackColor = PropBag.ReadProperty("BackColor", vbBlack)
 Enabled = PropBag.ReadProperty("Enabled", True)
 Set Font = PropBag.ReadProperty("Font", Ambient.Font)
 ForeColor = PropBag.ReadProperty("ForeColor", vbGreen)
 Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_Show()
    picGraph.ToolTipText = UserControl.Extender.ToolTipText
    picUsage.ToolTipText = UserControl.Extender.ToolTipText
    lblCpuUsage.ToolTipText = UserControl.Extender.ToolTipText
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("BarColor", mBarColor, &H808080)
 Call PropBag.WriteProperty("BackColor", mBackColor, vbBlack)
 Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
 Call PropBag.WriteProperty("Font", lblCpuUsage.Font, Ambient.Font)
 Call PropBag.WriteProperty("ForeColor", mForeColor, vbGreen)
 Call PropBag.WriteProperty("MouseIcon", picGraph.MouseIcon, Nothing)
 Call PropBag.WriteProperty("MousePointer", picGraph.MousePointer, 0)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
 UserControl.Size picUsage.Width + picGraph.Width, picUsage.Height
 picGraph.Height = picUsage.Height
End Sub

Public Property Get BackColor() As OLE_COLOR
 BackColor = picGraph.BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
 mBackColor = NewValue
 If Enabled = True Then
  picGraph.BackColor = NewValue
  picUsage.BackColor = NewValue
  lblCpuUsage.BackColor = NewValue
 End If
 PropertyChanged "BackColor"
 prvRefresh 0
End Property

Public Property Get ForeColor() As OLE_COLOR
 ForeColor = picGraph.ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
 mForeColor = NewValue
 If Enabled = True Then
  picGraph.ForeColor = NewValue
  picUsage.ForeColor = NewValue
  lblCpuUsage.ForeColor = NewValue
 End If
 PropertyChanged "ForeColor"
 prvRefresh 0
End Property

Public Property Get BarColor() As OLE_COLOR
 BarColor = mBarColor
End Property
Public Property Let BarColor(ByVal NewValue As OLE_COLOR)
 mBarColor = NewValue
 PropertyChanged "BarColor"
 prvRefresh 0
End Property

Public Property Get Font() As StdFont
 Set Font = lblCpuUsage.Font
End Property
Public Property Set Font(ByVal NewValue As StdFont)
 Set lblCpuUsage.Font = NewValue
 PropertyChanged "Font"
 prvRefresh 0
End Property

Public Property Get Enabled() As Boolean
 Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
 UserControl.Enabled = NewValue
 If tmrRefresh.Enabled = True And NewValue = False Then tmrRefresh.Enabled = False
 PropertyChanged "Enabled"
 If Enabled = True Then
  picGraph.ForeColor = mForeColor
  picUsage.ForeColor = mForeColor
  lblCpuUsage.ForeColor = mForeColor
  picGraph.BackColor = mBackColor
  picUsage.BackColor = mBackColor
  lblCpuUsage.BackColor = mBackColor
 Else
  picGraph.ForeColor = &H80000011
  picUsage.ForeColor = &H80000011
  lblCpuUsage.ForeColor = &H80000011
  picGraph.BackColor = &H8000000F
  picUsage.BackColor = &H8000000F
  lblCpuUsage.BackColor = &H8000000F
 End If
 prvRefresh 0
End Property

Public Property Get MousePointer() As MousePointerConstants
 MousePointer = picGraph.MousePointer
End Property
Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
 picGraph.MousePointer = NewValue
 picUsage.MousePointer = NewValue
 lblCpuUsage.MousePointer = NewValue
 PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As StdPicture
 Set MouseIcon = picGraph.MouseIcon
End Property
Public Property Set MouseIcon(ByVal NewValue As StdPicture)
 Set picGraph.MouseIcon = NewValue
 Set picUsage.MouseIcon = NewValue
 Set lblCpuUsage.MouseIcon = NewValue
 PropertyChanged "MouseIcon"
End Property

Private Sub prvRefresh(lUsage As Long)
    picUsage.ScaleMode = vbPixels
    lblCpuUsage.Top = picUsage.ScaleHeight - lblCpuUsage.Height
    For mCounter = 0 To 12
      If Enabled = True Then
        picUsage.Line (SPACE, SPACE + mCounter * 3)-(SPACE + BAR_WIDTH, SPACE + mCounter * 3 + 1), IIf(lUsage >= 100 - mCounter * 10 And lUsage <> 0, lblCpuUsage.ForeColor, BarColor), BF
      Else
        picUsage.Line (SPACE, SPACE + mCounter * 3)-(SPACE + BAR_WIDTH, SPACE + mCounter * 3 + 1), &H80000011, BF
      End If
    Next mCounter
    
    prvChangePoints
    mChartPoints(UBound(mChartPoints)) = lUsage
    picGraph.Cls
    
    For mCounter = 1 To 10
      If Enabled = True Then
       picGraph.Line (0, mCounter * 14)-(100, mCounter * 14), BarColor
      Else
       picGraph.Line (0, mCounter * 14)-(100, mCounter * 14), &H80000011
      End If
    Next
    
    For mCounter = 1 To 10
      If Enabled = True Then
       picGraph.Line (mCounter * 10, 0)-(mCounter * 10, 100), BarColor
      Else
       picGraph.Line (mCounter * 10, 0)-(mCounter * 10, 100), &H80000011
      End If
    Next
    
    For mCounter = LBound(mChartPoints) To UBound(mChartPoints) - 1
      If Enabled = True Then
        picGraph.Line (mCounter, 105 - mChartPoints(mCounter))-(mCounter + 1, 105 - mChartPoints(mCounter + 1)), lblCpuUsage.ForeColor
      Else
        picGraph.Line (mCounter, 105 - mChartPoints(mCounter))-(mCounter + 1, 105 - mChartPoints(mCounter + 1)), &H8000000F
      End If
    Next
End Sub

Private Sub prvChangePoints()
    For mCounter = LBound(mChartPoints) To UBound(mChartPoints) - 1
        mChartPoints(mCounter) = mChartPoints(mCounter + 1)
    Next
End Sub

Private Function IsWinNT() As Boolean
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    GetVersionEx OSInfo
    IsWinNT = (OSInfo.dwPlatformId = 2)
End Function

Private Sub prvStartMonitor()
    If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StartStat", 0, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then
        Exit Sub
    End If
    dwDataSize = 4
    RegQueryValueEx hKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
    RegCloseKey hKey
    
    If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StatData", 0, KEY_READ, hKey) <> ERROR_SUCCESS Then
        Exit Sub
    End If
End Sub

Private Function prvQuery() As Long
    dwDataSize = 4
    RegQueryValueEx hKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
    prvQuery = CLng(dwCpuUsage)
End Function

Private Sub prvStopMonitor()
    RegCloseKey hKey
    If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StopStat", 0, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then
        Debug.Print "Error while stopping counter"
        Exit Sub
    End If
    dwDataSize = 4
    RegQueryValueEx hKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
    RegCloseKey hKey
End Sub

Private Sub prvStartMonitorNT()
    Dim SysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim SysPerfInfo As SYSTEM_PERFORMANCE_INFORMATION
    Dim Ret As Long
    Ret = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SysTimeInfo), LenB(SysTimeInfo), 0&)
    If Ret <> NO_ERROR Then
        Exit Sub
    End If
    Ret = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SysPerfInfo), LenB(SysPerfInfo), ByVal 0&)
    If Ret <> NO_ERROR Then
        Exit Sub
    End If
    liOldIdleTime = SysPerfInfo.liIdleTime
    liOldSystemTime = SysTimeInfo.liKeSystemTime
End Sub

Private Function prvQueryNT() As Long
    Dim SysBaseInfo As SYSTEM_BASIC_INFORMATION
    Dim SysPerfInfo As SYSTEM_PERFORMANCE_INFORMATION
    Dim SysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim dbIdleTime As Currency
    Dim dbSystemTime As Currency
    Dim Ret As Long
    prvQueryNT = -1
    
    Ret = NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(SysBaseInfo), LenB(SysBaseInfo), 0&)
    If Ret <> NO_ERROR Then
        Exit Function
    End If
    
    Ret = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SysTimeInfo), LenB(SysTimeInfo), 0&)
    If Ret <> NO_ERROR Then
        Exit Function
    End If
    
    Ret = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SysPerfInfo), LenB(SysPerfInfo), ByVal 0&)
    If Ret <> NO_ERROR Then
        Exit Function
    End If
    
    dbIdleTime = LI2Currency(SysPerfInfo.liIdleTime) - LI2Currency(liOldIdleTime)
    dbSystemTime = LI2Currency(SysTimeInfo.liKeSystemTime) - LI2Currency(liOldSystemTime)
    If dbSystemTime <> 0 Then dbIdleTime = dbIdleTime / dbSystemTime
    dbIdleTime = 100 - dbIdleTime * 100 / SysBaseInfo.bKeNumberProcessors + 0.5
    prvQueryNT = Int(dbIdleTime)
    liOldIdleTime = SysPerfInfo.liIdleTime
    liOldSystemTime = SysTimeInfo.liKeSystemTime
End Function

Private Function LI2Currency(liInput As LARGE_INTEGER) As Currency
    CopyMemory LI2Currency, liInput, LenB(liInput)
End Function

Public Property Get Actived() As Boolean
    Actived = tmrRefresh.Enabled
End Property
