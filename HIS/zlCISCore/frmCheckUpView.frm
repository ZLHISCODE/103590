VERSION 5.00
Begin VB.Form frmCheckUpView 
   Caption         =   "图像查看"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   Icon            =   "frmCheckUpView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox PicMain 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4005
      Left            =   15
      MouseIcon       =   "frmCheckUpView.frx":058A
      MousePointer    =   99  'Custom
      ScaleHeight     =   4005
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   45
      Width           =   4665
   End
   Begin VB.Menu mnuPopMenu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopMenuPrint 
         Caption         =   "输出到打印机(&P)"
      End
   End
End
Attribute VB_Name = "frmCheckUpView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnDown As Boolean
Private mX As Single, mY As Single
Private mblnLoaded As Boolean


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Public Sub SetPicture(ObjPic As StdPicture)
'为窗体设置图片
    Set Me.picMain.Picture = ObjPic
    If Me.picMain.Width < Screen.Width Then
        Me.Width = Me.picMain.Width
    Else
        Me.Width = Screen.Width
    End If
    If Me.picMain.Height < Screen.Height Then
        Me.Height = Me.picMain.Height
    Else
        Me.Height = Screen.Height
    End If
End Sub

Private Sub Form_Load()
    Me.picMain.Top = 0
    Me.picMain.Left = 0
    mblnLoaded = False
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.WindowState = 0 Then
        PicMain_MouseDown 1, 0, 0, 0
        PicMain_MouseMove 1, 0, 0, 0
        If (Me.ScaleWidth - Me.picMain.Width) / 2 < 0 Then
            Me.picMain.Left = 0
        Else
            Me.picMain.Left = (Me.ScaleWidth - Me.picMain.Width) / 2
        End If
        If (Me.ScaleHeight - Me.picMain.Height) / 2 < 0 Then
            Me.picMain.Top = 0
        Else
            Me.picMain.Top = (Me.ScaleHeight - Me.picMain.Height) / 2
        End If
        If mblnLoaded = False Then
            Me.Left = (Screen.Width - Me.Width) / 2
            Me.Top = (Screen.Height - Me.Height) / 2
            mblnLoaded = True
        End If
    Else
        PicMain_MouseDown 1, 0, 0, 0
        PicMain_MouseMove 1, 0, 0, 0
        If (Me.ScaleWidth - Me.picMain.Width) / 2 < 0 Then
            Me.picMain.Left = 0
        Else
            Me.picMain.Left = (Me.ScaleWidth - Me.picMain.Width) / 2
        End If
        If (Me.ScaleHeight - Me.picMain.Height) / 2 < 0 Then
            Me.picMain.Top = 0
        Else
            Me.picMain.Top = (Me.ScaleHeight - Me.picMain.Height) / 2
        End If
    End If
End Sub

Private Sub mnuPopMenuPrint_Click()
Dim lngLeft As Long
Dim lngRight  As Long
Dim lngTop As Long
Dim lngBottom  As Long
Dim lngWidth As Long
Dim lngHeight  As Long
Dim m As Long
Dim lngStdPicHeight As Long
Dim lngStdPicWidth As Long
Dim dblPic比例  As Double

If Not picMain.Image Is Nothing Then
    '得到纸张的边界与宽高
    lngLeft = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "左边距", OFFSET_LEFT) * 56.7 + Screen.TwipsPerPixelX * 2
    lngRight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "右边距", OFFSET_RIGHT) * 56.7 - Screen.TwipsPerPixelX * 2
    lngTop = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "上边距", OFFSET_TOP) * 56.7 + Screen.TwipsPerPixelY * 2
    lngBottom = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "下边距", OFFSET_BOTTOM) * 56.7 - Screen.TwipsPerPixelY * 2
    lngWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "宽度", Printer.Width)
    lngHeight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "高度", Printer.Height)
    
    m = lngHeight - lngTop - lngBottom - Screen.TwipsPerPixelY * 6
    '得到图片的高
    lngStdPicHeight = picMain.ScaleHeight
    '得到图片的宽
    lngStdPicWidth = picMain.ScaleWidth
    '得到宽与高的比
    dblPic比例 = picMain.ScaleWidth / picMain.ScaleHeight
    
    '求出最大图片高
    If lngStdPicHeight > m Then
        lngStdPicHeight = m
        '再得到宽
        lngStdPicWidth = lngStdPicHeight * dblPic比例
    End If
    If lngStdPicWidth > lngWidth - lngLeft - lngRight - Screen.TwipsPerPixelX * 3 Then
        lngStdPicWidth = lngWidth - lngLeft - lngRight - Screen.TwipsPerPixelX * 3
        lngStdPicHeight = lngStdPicWidth / dblPic比例
    End If
    Printer.PaintPicture picMain.Image, lngLeft, lngTop, lngStdPicWidth, lngStdPicHeight, 0, 0, picMain.ScaleWidth, picMain.ScaleHeight
    Printer.EndDoc
End If
End Sub

Private Sub PicMain_DblClick()
    Unload Me
End Sub

Private Sub PicMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    mblnDown = True
    mX = X
    mY = Y
End If
End Sub

Private Sub PicMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And mblnDown Then
    If (Me.ScaleWidth - Me.picMain.Width) < 0 Then
        If picMain.Left + X - mX > 0 Then
            picMain.Left = 0
        ElseIf picMain.Left + X - mX < ScaleWidth - picMain.Width Then
            picMain.Left = ScaleWidth - picMain.Width
        Else
            picMain.Left = picMain.Left + X - mX
        End If
    Else
        If (Me.ScaleWidth - Me.picMain.Width) / 2 < 0 Then
            Me.picMain.Left = 0
        Else
            Me.picMain.Left = (Me.ScaleWidth - Me.picMain.Width) / 2
        End If
    End If
    If (Me.ScaleHeight - Me.picMain.Height) < 0 Then
        If picMain.Top + Y - mY > 0 Then
            picMain.Top = 0
        ElseIf picMain.Top + Y - mY < ScaleHeight - picMain.Height Then
            picMain.Top = ScaleHeight - picMain.Height
        Else
            picMain.Top = picMain.Top + Y - mY
        End If
    Else
        If (Me.ScaleHeight - Me.picMain.Height) / 2 < 0 Then
            Me.picMain.Top = 0
        Else
            Me.picMain.Top = (Me.ScaleHeight - Me.picMain.Height) / 2
        End If
    End If
End If
End Sub

Private Sub PicMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDown = False
    If Button = 2 Then
        PopupMenu mnuPopMenu
    End If
End Sub
