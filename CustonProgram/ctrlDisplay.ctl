VERSION 5.00
Begin VB.UserControl Display 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   MaskColor       =   &H00FFFFFF&
   PaletteMode     =   4  'None
   ScaleHeight     =   330
   ScaleWidth      =   2400
   ToolboxBitmap   =   "ctrlDisplay.ctx":0000
   Begin VB.PictureBox picDisp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   12720
      TabIndex        =   1
      Top             =   0
      Width           =   12720
   End
   Begin VB.PictureBox picLED 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "ctrlDisplay.ctx":0312
      ScaleHeight     =   375
      ScaleWidth      =   12735
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12735
   End
   Begin VB.PictureBox picPalette 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   12735
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   12735
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private nColor As OLE_COLOR
Private nCaption As String
Private nCharWidth As Integer
Private nTextPos As Integer

Private Sub UserControl_Initialize()
    nColor = 255
    nCaption = UserControl.Name
    nTextPos = 1
End Sub
Private Sub UserControl_Resize()
    picDisp.Move 0, 0, UserControl.Width, 330
    nCharWidth = LDec((picDisp.Width / 15) / 16)
    If Not UserControl.Width = nCharWidth * 240 Then UserControl.Width = nCharWidth * 240
    UserControl.Height = 330
End Sub
Public Property Get Color() As OLE_COLOR
    Color = nColor
End Property
Public Property Let Color(aColor As OLE_COLOR)
    nColor = aColor
    PropertyChanged Color
    SetPalette
End Property
Public Property Get Caption() As String
    Caption = nCaption
End Property
Public Property Let Caption(aCaption As String)
    nCaption = aCaption
    PropertyChanged Caption
    WriteGraphicFont
End Property
Public Property Get Characters() As String
    Characters = nCharWidth
End Property
Public Property Let Characters(aCharWidth As String)
    nCharWidth = aCharWidth
    PropertyChanged Characters
    UserControl.Width = nCharWidth * 240
End Property
Public Property Get ScrollRate() As Integer
    ScrollRate = tmrScroll.Interval
End Property
Public Property Let ScrollRate(aInterval As Integer)
    If tmrScroll.Enabled = True Then If aInterval < 1 Then aInterval = 1
    tmrScroll.Interval = aInterval
    PropertyChanged ScrollRate
End Property
Public Property Get Scroll() As Boolean
    Scroll = tmrScroll.Enabled
End Property
Public Property Let Scroll(aEnabled As Boolean)
    If aEnabled = True Then If tmrScroll.Interval < 1 Then tmrScroll.Interval = 1
    tmrScroll.Enabled = aEnabled
    PropertyChanged Scroll
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    nCharWidth = PropBag.ReadProperty("Characters", 10)
    nColor = PropBag.ReadProperty("Color", 255)
    nCaption = PropBag.ReadProperty("Caption", UserControl.Name)
    tmrScroll.Interval = PropBag.ReadProperty("ScrollRate", 1)
    tmrScroll.Enabled = PropBag.ReadProperty("Scroll", 0)
    SetPalette
    WriteGraphicFont
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Characters", nCharWidth, 10)
    Call PropBag.WriteProperty("Color", nColor, 255)
    Call PropBag.WriteProperty("Caption", UserControl.Name)
    Call PropBag.WriteProperty("ScrollRate", 1)
    Call PropBag.WriteProperty("Scroll", 0)
End Sub
Private Sub tmrScroll_Timer()
    nTextPos = nTextPos + 1
    If nTextPos > Len(nCaption) + nCharWidth Then nTextPos = 1
    WriteGraphicFont
End Sub
Private Sub WriteGraphicFont()
    Dim strText As String
    Dim FrameNo As Integer
    picDisp.Cls
    If tmrScroll.Enabled = True Then
        strText = Mid(SPACE(nCharWidth) & nCaption, nTextPos, nCharWidth)
    Else
        strText = nCaption
    End If
    If Len(strText) > nCharWidth Then strText = Left(strText, nCharWidth)
    strText = strText & SPACE(nCharWidth - Len(strText))
    For SPrint = 0 To Len(strText) - 1
        FrameNo = 0
        Select Case LCase(Mid(strText, SPrint + 1, 1))
            Case " ": FrameNo = 0
            Case "0": FrameNo = 1
            Case "1": FrameNo = 2
            Case "2": FrameNo = 3
            Case "3": FrameNo = 4
            Case "4": FrameNo = 5
            Case "5": FrameNo = 6
            Case "6": FrameNo = 7
            Case "7": FrameNo = 8
            Case "8": FrameNo = 9
            Case "9": FrameNo = 10
            Case "a": FrameNo = 11
            Case "b": FrameNo = 12
            Case "c": FrameNo = 13
            Case "d": FrameNo = 14
            Case "e": FrameNo = 15
            Case "f": FrameNo = 16
            Case "g": FrameNo = 17
            Case "h": FrameNo = 18
            Case "i": FrameNo = 19
            Case "j": FrameNo = 20
            Case "k": FrameNo = 21
            Case "l": FrameNo = 22
            Case "m": FrameNo = 23
            Case "n": FrameNo = 24
            Case "o": FrameNo = 25
            Case "p": FrameNo = 26
            Case "q": FrameNo = 27
            Case "r": FrameNo = 28
            Case "s": FrameNo = 29
            Case "t": FrameNo = 30
            Case "u": FrameNo = 31
            Case "v": FrameNo = 32
            Case "w": FrameNo = 33
            Case "x": FrameNo = 34
            Case "y": FrameNo = 35
            Case "z": FrameNo = 36
            Case "-": FrameNo = 37
            Case "+": FrameNo = 38
            Case "?": FrameNo = 39
            Case "=": FrameNo = 40
            Case ":": FrameNo = 41
            Case "'": FrameNo = 42
            Case ".": FrameNo = 43
            Case "!": FrameNo = 44
            Case "$": FrameNo = 45
            Case "%": FrameNo = 46
            Case "&": FrameNo = 47
            Case "*": FrameNo = 48
            Case "(": FrameNo = 49
            Case ")": FrameNo = 50
            Case "\": FrameNo = 51
            Case "/": FrameNo = 52
        End Select
        On Error Resume Next
        picDisp.PaintPicture picPalette, 240 * SPrint, 0, , , 240 * FrameNo, 0, 240
    Next SPrint
End Sub
Private Sub SetPalette()
    Dim PalR As Integer
    Dim PalG As Integer
    Dim PalB As Integer
    GetRgb nColor, PalR, PalG, PalB
    For y = 0 To 22
        For x = 0 To 848
            Select Case picLED.Point(x * 15, y * 15)
                Case 255
                    picPalette.PSet (x * 15, y * 15), RGB(ProcByte(PalR - 155), ProcByte(PalG - 155), ProcByte(PalB - 155))
                Case 16711680
                    picPalette.PSet (x * 15, y * 15), RGB(ProcByte(PalR), ProcByte(PalG), ProcByte(PalB))
                Case Else
                    picPalette.PSet (x * 15, y * 15), 0
            End Select
        Next x
    Next y
    SavePicture picPalette.Image, App.Path & "\fontpal.tmp"
    picPalette.Cls
    picPalette.Picture = LoadPicture(App.Path & "\fontpal.tmp")
    Kill App.Path & "\fontpal.tmp"
    WriteGraphicFont
End Sub
Private Function LDec(Number As Currency) As Integer
    If InStr(1, CStr(Number), ".") > 0 Then
        LDec = CCur(Left(CStr(Number), InStr(1, CStr(Number), ".") - 1))
    Else
        LDec = Number
    End If
End Function
Private Function ProcByte(ByVal ByteNum As Integer) As Integer
    If ByteNum < 0 Then
        ByteNum = 0
    ElseIf ByteNum > 255 Then
        ByteNum = 255
    Else
        ProcByte = ByteNum
    End If
End Function
Private Sub GetRgb(ByVal Color As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
    Red = (Color And 255) And 255
    Green = Int(Color / 256) And 255
    Blue = Int(Color / 65536) And 255
End Sub
