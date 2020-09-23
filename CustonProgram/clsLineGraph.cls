VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLineGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarPictureBox As PictureBox
Private mvarDataCollection As Collection
Private mvarPicBackground As OLE_COLOR
Private mvarPicForeground As OLE_COLOR
Private mvarBorderSize As Integer
Private mvarBorderColor As Long
Private mvarGridVisible As Boolean
Private mvarGridColor As Long
Public Property Let GridColor(ByVal vData As Long)
    mvarGridColor = vData
End Property
Public Property Get GridColor() As Long
    GridColor = mvarGridColor
End Property
Public Property Let GridVisible(ByVal vData As Boolean)
    mvarGridVisible = vData
End Property
Public Property Get GridVisible() As Boolean
    GridVisible = mvarGridVisible
End Property
Public Property Let BorderColor(ByVal vData As Long)
    mvarBorderColor = vData
End Property
Public Property Get BorderColor() As Long
    BorderColor = mvarBorderColor
End Property
Public Property Let BorderSize(ByVal vData As Integer)
    mvarBorderSize = vData
End Property
Public Property Get BorderSize() As Integer
    'BorderStyle = mvarBorderSize
End Property
Public Property Let PicForeground(ByVal vData As Long)
    Let mvarPicForeground = vData
End Property
Public Property Get PicForeground() As Long
    PicForeground = mvarPicForeground
End Property
Public Property Let PicBackground(ByVal vData As Long)
    mvarPicBackground = vData
End Property
Public Property Get PicBackground() As Long
    PicBackground = mvarPicBackground
End Property
Public Property Set DataCollection(ByVal vData As Collection)
    Set mvarDataCollection = vData
End Property
Public Property Get DataCollection() As Collection
    Set DataCollection = mvarDataCollection
End Property
Public Property Set PictureBox(ByVal vData As PictureBox)
    Set mvarPictureBox = vData
End Property
Public Property Get PictureBox() As PictureBox
    Set PictureBox = mvarPictureBox
End Property
Public Sub Draw()
    Dim BDR As Integer, X As Integer
    Dim NewX As Double, NewY As Double
    Dim OldX As Double, OldY As Double
    Dim GridHeight As Double, GridWidth As Double
    
    On Error GoTo NoPicBox      ' In case the PicBox isn't set yet
    
    If mvarPictureBox.AutoRedraw = False Then mvarPictureBox.AutoRedraw = True
    mvarPictureBox.Cls
    
    BDR = mvarPictureBox.BorderStyle
    If mvarPictureBox.ScaleMode <> 3 Then mvarPictureBox.ScaleMode = 3
    If mvarPictureBox.BackColor <> mvarPicBackground Then mvarPictureBox.BackColor = mvarPicBackground
    If mvarBorderSize > 0 Then
        For X = 0 To mvarBorderSize
            mvarPictureBox.Line (X, X)-(mvarPictureBox.ScaleWidth - (BDR + X), mvarPictureBox.ScaleHeight - (BDR + X)), mvarBorderColor, B
        Next X
    End If
    
    ' Display Grid On Screen
    If mvarGridVisible = True Then
        For X = 1 To 20
            mvarPictureBox.Line (mvarBorderSize, mvarBorderSize)-((((mvarPictureBox.ScaleWidth - (mvarBorderSize * 2)) / 20) * X), (mvarPictureBox.ScaleHeight - (mvarBorderSize * 2))), mvarGridColor, B
        Next X
        For X = 1 To 10
            mvarPictureBox.Line (mvarBorderSize, mvarBorderSize)-((mvarPictureBox.ScaleWidth - (mvarBorderSize * 2)), (((mvarPictureBox.ScaleHeight - (mvarBorderSize * 2)) / 10) * X)), mvarGridColor, B
        Next X
    End If
    
    If mvarDataCollection.Count > 0 Then
        GridHeight = ((mvarPictureBox.ScaleHeight - (mvarBorderSize * 2)) / 100) + 0    ' 0-100%
        GridWidth = ((mvarPictureBox.ScaleWidth - (mvarBorderSize * 2)) / 100) + 0      ' 1-100 Items
        Do
            If mvarDataCollection.Count > 100 Then _
                mvarDataCollection.Remove 1
        Loop While mvarDataCollection.Count > 100
        
        OldX = mvarBorderSize + 2
        OldY = ((mvarPictureBox.ScaleHeight - (mvarBorderSize * 2)) - (mvarDataCollection(1) * GridHeight))
        For X = 1 To 100
            NewX = (mvarPictureBox.ScaleWidth - (mvarBorderSize * 2)) - ((100 - (X - 1)) * GridWidth)
            NewY = ((mvarPictureBox.ScaleHeight - (mvarBorderSize * 2)) - (mvarDataCollection(X) * GridHeight))
            NewX = NewX + 2
            If NewX < mvarBorderSize Then NewX = mvarBorderSize
            If NewY < mvarBorderSize Then NewY = mvarBorderSize
            
            mvarPictureBox.Line (OldX, OldY)-(NewX, NewY), mvarPicForeground
            OldX = NewX: OldY = NewY
            If OldX < mvarBorderSize Then OldX = mvarBorderSize
            If OldY < mvarBorderSize Then OldY = mvarBorderSize
        Next X
    End If
NoPicBox:
End Sub
