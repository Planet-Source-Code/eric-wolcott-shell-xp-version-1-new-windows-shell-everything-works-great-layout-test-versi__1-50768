Attribute VB_Name = "modTaskListing"

'I hope this bit encourages you newbies to
'start new shells (use this to make a taskbar)

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const LB_ADDSTRING = &H180
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_ERR = (-1)

Public Const GW_OWNER = 4
Public Const GWL_EXSTYLE = (-20)

Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_TOOLWINDOW = &H80

Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Boolean
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Const DI_NORMAL = &H3

Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Integer) As Long

Public Const WM_GETICON = &H7F
Public Const GCL_HICON = (-14)
Public Const GCL_HICONSM = (-34)
Public Const WM_QUERYDRAGICON = &H37

Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

'This is used to get icons from windows >>>>
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

Public Function fEnumWindows(lst As ListBox) As Long
    With lst
      .Clear
      frmTaskbar.lstNames.Clear
      Call EnumWindows(AddressOf fEnumWindowsCallBack, .hWnd)
      fEnumWindows = .ListCount
    End With
End Function

Private Function fEnumWindowsCallBack(ByVal hWnd As Long, ByVal lParam As Long) As Long
    
    Dim lExStyle As Long, bHasNoOwner As Boolean, sAdd As String, sCaption As String
    
    If IsWindowVisible(hWnd) Then
        bHasNoOwner = (GetWindow(hWnd, GW_OWNER) = 0)
        lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
        
        If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bHasNoOwner) Or _
            ((lExStyle And WS_EX_APPWINDOW) And Not bHasNoOwner) Then
            sAdd = hWnd: sCaption = GetCaption(hWnd)
            Call SendMessage(lParam, LB_ADDSTRING, 0, ByVal sAdd)
            Call SendMessage(frmTaskbar.lstNames.hWnd, LB_ADDSTRING, 0, ByVal sCaption)
        End If
    End If

    fEnumWindowsCallBack = True
End Function

Public Function GetCaption(hWnd As Long) As String
    Dim mCaption As String, lReturn As Long
    mCaption = Space(255)
    lReturn = GetWindowText(hWnd, mCaption, 255)
    GetCaption = Left(mCaption, lReturn)
End Function

