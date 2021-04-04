VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl ColorPicker 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ScaleHeight     =   2190
   ScaleWidth      =   2190
   ToolboxBitmap   =   "ColorPicker.ctx":0000
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   370
      Left            =   45
      ScaleHeight     =   375
      ScaleWidth      =   2085
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   45
      Width           =   2085
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   45
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1915
      Width           =   200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   0
      Picture         =   "ColorPicker.ctx":0312
      ScaleHeight     =   1350
      ScaleWidth      =   2160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   470
      Width           =   2160
      Begin VB.Shape shpValue 
         BorderColor     =   &H00C56A31&
         FillColor       =   &H00FF8080&
         Height          =   270
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00C56A31&
         FillColor       =   &H00FF8080&
         Height          =   270
         Left            =   1890
         Top             =   1080
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1800
      Top             =   1935
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblColor 
      Caption         =   "&HFFFFFF"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1918
      UseMnemonic     =   0   'False
      Width           =   1365
   End
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarColor As OLE_COLOR
Private mvarAutoColor As OLE_COLOR

Public Event pOK()
Public Event pCancel()

Public Property Get AutoColor() As OLE_COLOR
    AutoColor = mvarAutoColor
End Property

Public Property Let AutoColor(vData As OLE_COLOR)
    mvarAutoColor = vData
    PropertyChanged "AutoColor"
End Property
    
Public Property Get Tag() As String
    Tag = UserControl.Tag
End Property

Public Property Let Tag(vData As String)
    UserControl.Tag = vData
    PropertyChanged "Tag"
End Property
    
Public Property Get Hwnd() As Long
    Hwnd = UserControl.Hwnd
End Property

Public Property Get COLOR() As OLE_COLOR
    COLOR = mvarColor
End Property

Public Property Let COLOR(vData As OLE_COLOR)
    mvarColor = vData
    Dim lRow As Long, lCol As Long
    shpValue.Visible = True
    Select Case CStr(Hex(vData))
    Case "0"
        lblColor = "黑色"
        lRow = 0
        lCol = 0
    Case "3399"
        lblColor = "褐色"
        lRow = 0
        lCol = 1
    Case "3333"
        lblColor = "橄榄色"
        lRow = 0
        lCol = 2
    Case "3300"
        lblColor = "深绿"
        lRow = 0
        lCol = 3
    Case "663300"
        lblColor = "深青"
        lRow = 0
        lCol = 4
    Case "800000"
        lblColor = "深蓝"
        lRow = 0
        lCol = 5
    Case "993333"
        lblColor = "靛蓝"
        lRow = 0
        lCol = 6
    Case "333333"
        lblColor = "灰色-80%"
        lRow = 0
        lCol = 7
    Case "80"
        lblColor = "深红"
        lRow = 1
        lCol = 0
    Case "66FF"
        lblColor = "橙色"
        lRow = 1
        lCol = 1
    Case "8080"
        lblColor = "深黄"
        lRow = 1
        lCol = 2
    Case "8000"
        lblColor = "绿色"
        lRow = 1
        lCol = 3
    Case "808000"
        lblColor = "青色"
        lRow = 1
        lCol = 4
    Case "FF0000"
        lblColor = "蓝色"
        lRow = 1
        lCol = 5
    Case "996666"
        lblColor = "蓝-灰"
        lRow = 1
        lCol = 6
    Case "808080"
        lblColor = "灰色-50%"
        lRow = 1
        lCol = 7
    Case "FF"
        lblColor = "红色"
        lRow = 2
        lCol = 0
    Case "99FF"
        lblColor = "浅橙色"
        lRow = 2
        lCol = 1
    Case "CC99"
        lblColor = "酸橙色"
        lRow = 2
        lCol = 2
    Case "669933"
        lblColor = "海绿"
        lRow = 2
        lCol = 3
    Case "CCCC33"
        lblColor = "水绿色"
        lRow = 2
        lCol = 4
    Case "FF6633"
        lblColor = "浅蓝"
        lRow = 2
        lCol = 5
    Case "800080"
        lblColor = "紫罗兰"
        lRow = 2
        lCol = 6
    Case "999999"
        lblColor = "灰色-40%"
        lRow = 2
        lCol = 7
    Case "FF00FF"
        lblColor = "粉红"
        lRow = 3
        lCol = 0
    Case "CCFF"
        lblColor = "金色"
        lRow = 3
        lCol = 1
    Case "FFFF"
        lblColor = "黄色"
        lRow = 3
        lCol = 2
    Case "FF00"
        lblColor = "鲜绿"
        lRow = 3
        lCol = 3
    Case "FFFF00"
        lblColor = "青绿"
        lRow = 3
        lCol = 4
    Case "FFCC00"
        lblColor = "天蓝"
        lRow = 3
        lCol = 5
    Case "663399"
        lblColor = "梅红"
        lRow = 3
        lCol = 6
    Case "C0C0C0"
        lblColor = "灰色-25%"
        lRow = 3
        lCol = 7
    Case "CC99FF"
        lblColor = "玫瑰红"
        lRow = 4
        lCol = 0
    Case "99CCFF"
        lblColor = "茶色"
        lRow = 4
        lCol = 1
    Case "99FFFF"
        lblColor = "浅黄"
        lRow = 4
        lCol = 2
    Case "CCFFCC"
        lblColor = "浅绿"
        lRow = 4
        lCol = 3
    Case "FFFFCC"
        lblColor = "浅青绿"
        lRow = 4
        lCol = 4
    Case "FFCC99"
        lblColor = "淡蓝"
        lRow = 4
        lCol = 5
    Case "FF99CC"
        lblColor = "淡紫"
        lRow = 4
        lCol = 6
    Case "FFFFFF"
        lblColor = "白色"
        lRow = 4
        lCol = 7
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
        shpValue.Visible = False
    End Select
    shpValue.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    If vData = tomAutoColor Or vData = -1 Then
    
    Else
        picColor.BackColor = vData
    End If
    If picColor.Visible Then picColor.SetFocus
    If COLOR = tomAutoColor Then
        DrawButton 2
    Else
        DrawButton 0
    End If
    
    PropertyChanged "Color"
End Property

Private Sub picColor_Click()
'    SendKeys "{ESCAPE}"
'    DoEvents
'    dlgThis.Color = IIf(mvarColor = tomAutoColor, vbBlack, mvarColor)
'    dlgThis.CancelError = True
'    On Error GoTo LL
'    dlgThis.ShowColor
'    mvarColor = dlgThis.Color
'    RaiseEvent pOK
'    Exit Sub
'LL:
'    RaiseEvent pCancel
End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    lblColor.Caption = "更多颜色..."
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        RaiseEvent pCancel
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And X < Picture1.ScaleWidth And Y > 0 And Y < Picture1.ScaleHeight Then
        SetCapture Picture1.Hwnd
        shpBorder.Visible = True
    Else
        ReleaseCapture
        COLOR = mvarColor
        shpBorder.Visible = False
    End If

    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    
    shpBorder.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    
    If Picture1.Point(lX, lY) = -1 Then Exit Sub
    picColor.BackColor = Picture1.Point(lX, lY)
    Select Case CStr(Hex(picColor.BackColor))
    Case "0"
        lblColor = "黑色"
    Case "3399"
        lblColor = "褐色"
    Case "3333"
        lblColor = "橄榄色"
    Case "3300"
        lblColor = "深绿"
    Case "663300"
        lblColor = "深青"
    Case "800000"
        lblColor = "深蓝"
    Case "993333"
        lblColor = "靛蓝"
    Case "333333"
        lblColor = "灰色-80%"
    Case "80"
        lblColor = "深红"
    Case "66FF"
        lblColor = "橙色"
    Case "8080"
        lblColor = "深黄"
    Case "8000"
        lblColor = "绿色"
    Case "808000"
        lblColor = "青色"
    Case "FF0000"
        lblColor = "蓝色"
    Case "996666"
        lblColor = "蓝-灰"
    Case "808080"
        lblColor = "灰色-50%"
    Case "FF"
        lblColor = "红色"
    Case "99FF"
        lblColor = "浅橙色"
    Case "CC99"
        lblColor = "酸橙色"
    Case "669933"
        lblColor = "海绿"
    Case "CCCC33"
        lblColor = "水绿色"
    Case "FF6633"
        lblColor = "浅蓝"
    Case "800080"
        lblColor = "紫罗兰"
    Case "999999"
        lblColor = "灰色-40%"
    Case "FF00FF"
        lblColor = "粉红"
    Case "CCFF"
        lblColor = "金色"
    Case "FFFF"
        lblColor = "黄色"
    Case "FF00"
        lblColor = "鲜绿"
    Case "FFFF00"
        lblColor = "青绿"
    Case "FFCC00"
        lblColor = "天蓝"
    Case "663399"
        lblColor = "梅红"
    Case "C0C0C0"
        lblColor = "灰色-25%"
    Case "CC99FF"
        lblColor = "玫瑰红"
    Case "99CCFF"
        lblColor = "茶色"
    Case "99FFFF"
        lblColor = "浅黄"
    Case "CCFFCC"
        lblColor = "浅绿"
    Case "FFFFCC"
        lblColor = "浅青绿"
    Case "FFCC99"
        lblColor = "淡蓝"
    Case "FF99CC"
        lblColor = "淡紫"
    Case "FFFFFF"
        lblColor = "白色"
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    
    COLOR = Picture1.Point(lX, lY)
    RaiseEvent pOK
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton 3
    Picture2.Tag = "Down"
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 0 And X <= Picture2.ScaleWidth And Y >= 0 And Y <= Picture2.ScaleHeight Then
        SetCapture Picture2.Hwnd         '导致ToolTipText不起作用了！
        '鼠标移入！！！
        If Picture2.Tag = "Down" Then
            DrawButton 3
        Else
            DrawButton 1
        End If
    Else
        If Picture2.Tag <> "" Then
            DrawButton 3
        Else
            '鼠标移出！！！                 '导致ToolTipText不起作用了！
            ReleaseCapture
            If COLOR = tomAutoColor Then
                DrawButton 2
            Else
                DrawButton 0
            End If
        End If
    End If
End Sub

Private Sub DrawButton(lDrawStyle As Long)
    '0:普通 &H8000000F    1:移动  &HEED2C1   2:选中 &HE8E6E1    3:按下  &HE2B598          边框:&HC56A31
    On Error Resume Next
    If lDrawStyle = 2 Then lDrawStyle = 0
    Cls
    Select Case lDrawStyle
    Case 0  '普通
        Picture2.BackColor = &H8000000F
    Case 1  '移动
        Picture2.BackColor = &HEED2C1
        Picture2.Line (0, 0)-(Picture2.ScaleWidth - Screen.TwipsPerPixelX, Picture2.ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
    Case 2  '选中
        shpBorder.Visible = False
        shpValue.Visible = False
        Picture2.BackColor = &HE8E6E1
        Picture2.Line (0, 0)-(Picture2.ScaleWidth - Screen.TwipsPerPixelX, Picture2.ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
    Case 3  '按下
        Picture2.BackColor = &HE2B598
        Picture2.Line (0, 0)-(Picture2.ScaleWidth - Screen.TwipsPerPixelX, Picture2.ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
    End Select
    
    Picture2.Line (90, 90)-(290, 290), AutoColor, BF
    Picture2.Line (90, 90)-(290, 290), RGB(133, 133, 133), B
    Picture2.CurrentX = 900
    Picture2.CurrentY = 90
    Picture2.Print "自动"
    Refresh
    Err.Clear
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture2.Tag = ""
    mvarColor = tomAutoColor
    DrawButton 3
    RaiseEvent pOK
End Sub

Private Sub UserControl_Initialize()
    COLOR = vbWhite
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        RaiseEvent pCancel
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    COLOR = PropBag.ReadProperty("Color", vbWhite)
    AutoColor = PropBag.ReadProperty("AutoColor", vbBlack)
    If mvarColor = tomAutoColor Then
        DrawButton 2
    Else
        DrawButton 0
    End If
End Sub

Private Sub UserControl_Resize()
    Width = 2190
    Height = 2190
End Sub

Private Sub UserControl_Show()
'    If mvarColor = tomAutoColor Then
'        DrawButton 2
'    Else
'        DrawButton 0
'    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Color", COLOR, vbWhite
    PropBag.WriteProperty "AutoColor", AutoColor, vbBlack
    
    PropertyChanged "Color"
    PropertyChanged "AutoColor"
End Sub
