VERSION 5.00
Begin VB.Form frmPassPara 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   4200
      Top             =   120
   End
   Begin VB.Frame fraPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4455
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   0
         Left            =   3120
         ScaleHeight     =   420
         ScaleWidth      =   1095
         TabIndex        =   10
         Top             =   720
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "全选"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   11
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   1
         Left            =   3120
         ScaleHeight     =   420
         ScaleWidth      =   1095
         TabIndex        =   8
         Top             =   1320
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "全清"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.CheckBox chkPara 
         BackColor       =   &H80000005&
         Caption         =   "允许下达药品医嘱弹出要点提示"
         Height          =   375
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.ListBox lstType 
         Appearance      =   0  'Flat
         Columns         =   1
         Height          =   2760
         ItemData        =   "frmPassPara.frx":0000
         Left            =   480
         List            =   "frmPassPara.frx":000A
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "要点提示项目设置"
         ForeColor       =   &H00C000C0&
         Height          =   180
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   1440
      End
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00D48A00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.PictureBox picClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   1200
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   0
         Width           =   500
         Begin VB.Label lblClose 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "×"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   300
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.Label lblFrmName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "个性化"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   675
      End
   End
   Begin VB.Line linScope 
      Index           =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8040
   End
   Begin VB.Line linScope 
      Index           =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8040
   End
   Begin VB.Line linScope 
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8040
   End
   Begin VB.Line linScope 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8040
   End
End
Attribute VB_Name = "frmPassPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytShowStyle As Byte
Private mbytOpen        As Byte        '是否加载

Public Function IsOpen() As Boolean
'判断窗体是否加载
    IsOpen = mbytOpen = 1
End Function

Private Sub chkPara_Click()
    lstType.Enabled = chkPara.Value
    Call SetPicBtnEnable
End Sub

Private Sub Form_Load()
    Dim arrTemp As Variant
    Dim i As Long
    
    gstrParaTip = zlDatabase.GetPara(299, glngSys)
    chkPara.Value = Val(Mid(gstrParaTip, 1, 1))
    lstType.Clear
    arrTemp = Split(conSTR_Key_Tip, ",")
    For i = LBound(arrTemp) To UBound(arrTemp)
        lstType.AddItem arrTemp(i)
        If Mid(gstrParaTip, i + 2, 1) = "1" Then lstType.Selected(i) = True
    Next
    lstType.Enabled = chkPara.Value
    SetPicBtnEnable
    picTop.BackColor = conCOLOR_TITLE_BAR
    chkPara.BackColor = fraPara.BackColor
    Me.Width = 4740: Me.Height = 4410
    Call ShowStyle
    mbytOpen = 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picTop.Move 15, 15, Me.ScaleWidth - 30, 495
    fraPara.Move 15, picTop.Top + picTop.Height, Me.ScaleWidth - 30, Me.ScaleHeight - picTop.Top - picTop.Height - 30
    'Left
    With linScope(0)
        .X1 = 0: .X2 = 0: .Y1 = 0: .Y2 = Me.ScaleHeight
        .BorderColor = conCOLOR_TITLE_BAR
        '&H00808080&
        '&H80000010& '按钮阴影
    End With
    'bottom
    With linScope(1)
        .X1 = 0: .X2 = Me.ScaleWidth: .Y1 = Me.ScaleHeight - 15: .Y2 = Me.ScaleHeight - 15
        .BorderColor = conCOLOR_TITLE_BAR
    End With
    'right
    With linScope(2)
        .X1 = Me.ScaleWidth - 15: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = Me.ScaleHeight - 15
        .BorderColor = conCOLOR_TITLE_BAR
    End With
    'Top
    With linScope(3)
        .X1 = 0: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = 0
        .BorderColor = conCOLOR_TITLE_BAR
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytOpen = 0
End Sub

Private Sub lblBtn_Click(Index As Integer)
    SetLstSelected lstType, Index = 0
End Sub

Private Sub lblClose_Click()
    Dim arrTemp As Variant
    Dim i As Long
    gstrParaTip = IIf(chkPara.Value, "1", "0")
    arrTemp = Split(conSTR_Key_Tip, ",")
    For i = 0 To lstType.ListCount - 1
        gstrParaTip = gstrParaTip & IIf(lstType.Selected(i), "1", "0")
    Next
    Call zlDatabase.SetPara(299, gstrParaTip, glngSys)
    Unload Me
End Sub

Private Sub lstType_ItemCheck(Item As Integer)
    Dim strTemp As String
    
    If Item = 0 Then
       Debug.Print ""
    End If
End Sub

Private Sub picBtn_Click(Index As Integer)
    lblBtn_Click Index
End Sub

Private Sub picClosed_Click()
    lblClose_Click
End Sub

Private Sub picClosed_Resize()
    On Error Resume Next
    lblClose.Move picClosed.ScaleWidth / 2 + lblClose.Width / 2, (picClosed.ScaleHeight - lblClose.Height) / 2
End Sub

Private Sub picTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
       Call ReleaseCapture
       Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub picTop_Resize()
    On Error Resume Next
    lblFrmName.Move 120, picTop.ScaleHeight / 2 - lblFrmName.Height / 2
    picClosed.Move picTop.ScaleWidth - picTop.Height, 0, picTop.Height, picTop.Height
End Sub

Private Sub tmrTime_Timer()
    Dim lngRet As Long
    Dim udtPoint As POINTAPI
    Dim udtRectClose As RECT
    
    Call GetWindowRect(picClosed.hwnd, udtRectClose)
    lngRet = GetCursorPos(udtPoint)
    If PtInRect(udtRectClose, udtPoint.X, udtPoint.Y) Then
        picClosed.BackColor = "&H" & Hex(RGB(212, 64, 39))  '红色
    Else
        picClosed.BackColor = picTop.BackColor
    End If
End Sub

Private Sub ShowStyle()
'功能:根据主窗体位置,决定界面显示位置
    Dim objPoint As RECT
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngSplit As Long
    lngSplit = 60
    Call GetWindowRect(gobjFrm.hwnd, objPoint)
    lngTop = objPoint.Top * Screen.TwipsPerPixelY
    lngLeft = objPoint.Left * Screen.TwipsPerPixelX
    If lngTop + gobjFrm.Height + Me.Height < Screen.Height And lngLeft + Me.Width < Screen.Width Then
        mbytShowStyle = 0  '左下显示
        Me.Top = lngTop + gobjFrm.Height + lngSplit
        Me.Left = lngLeft
    ElseIf lngTop + gobjFrm.Height + Me.Height < Screen.Height And lngLeft + Me.Width > Screen.Width Then
        mbytShowStyle = 1  '右下显示
        Me.Top = lngTop + gobjFrm.Height + lngSplit
        Me.Left = lngLeft - Me.Width + gobjFrm.Width
    ElseIf lngTop - Me.Height > 0 And lngLeft + Me.Width < Screen.Width Then
        mbytShowStyle = 2  '左上显示
        Me.Top = lngTop - Me.Height - lngSplit
        Me.Left = lngLeft
    ElseIf lngTop - Me.Height > 0 And lngLeft + Me.Width > Screen.Width Then
        mbytShowStyle = 3  '右上显示
        Me.Top = lngTop - Me.Height - lngSplit
        Me.Left = lngLeft - Me.Width + gobjFrm.Width
    ElseIf lngTop + Me.Height + gobjFrm.Height > Screen.Height Then
        lngTop = Screen.Height - Me.Height - gobjFrm.Height
        gobjFrm.Top = lngTop '悬浮窗体高度不够向上移动
        If lngLeft + Me.Width < Screen.Width Then
            mbytShowStyle = 0  '左下显示
            Me.Top = lngTop + gobjFrm.Height + lngSplit
            Me.Left = lngLeft
        Else
            mbytShowStyle = 1  '右下显示
            Me.Top = lngTop + gobjFrm.Height + lngSplit
            Me.Left = lngLeft - Me.Width + gobjFrm.Width
        End If
    End If
End Sub

Private Sub SetLstSelected(ByRef lst As ListBox, ByVal blnSel As Boolean)
'功能：全选或全消ListBox项目，保持位置不变
    Dim i As Long, Y As Long
    
    With lstType
        Y = .ListIndex
        For i = 0 To .ListCount - 1
            .Selected(i) = blnSel    '将触发lst_ItemCheck事件
        Next
        .ListIndex = Y
    End With
End Sub

Private Sub SetPicBtnEnable()
    If chkPara.Value Then
        picBtn(0).Enabled = True
        picBtn(1).Enabled = True
        picBtn(0).BackColor = &HD48A00
        picBtn(1).BackColor = &HD48A00
    Else
        picBtn(0).Enabled = False
        picBtn(1).Enabled = False
        picBtn(0).BackColor = "&H" & Hex(RGB(144, 158, 149))
        picBtn(1).BackColor = "&H" & Hex(RGB(144, 158, 149))
    End If
End Sub
