VERSION 5.00
Begin VB.Form frmPatiFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定位设置"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4380
      TabIndex        =   10
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4380
      TabIndex        =   9
      Top             =   270
      Width           =   1100
   End
   Begin VB.Frame fraBdr 
      Height          =   2280
      Left            =   120
      TabIndex        =   11
      Top             =   15
      Width           =   4155
      Begin VB.OptionButton optHead 
         Caption         =   "所有"
         Height          =   195
         Left            =   2460
         TabIndex        =   7
         Top             =   1950
         Width           =   660
      End
      Begin VB.OptionButton optCur 
         Caption         =   "向下"
         Height          =   195
         Left            =   3180
         TabIndex        =   8
         Top             =   1950
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.TextBox txt床号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1095
         Width           =   1110
      End
      Begin VB.TextBox txt身份证 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         MaxLength       =   18
         TabIndex        =   6
         Top             =   1515
         Width           =   3030
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   900
         TabIndex        =   4
         Top             =   1095
         Width           =   1110
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2820
         MaxLength       =   18
         TabIndex        =   3
         Top             =   675
         Width           =   1110
      End
      Begin VB.TextBox txt门诊号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         MaxLength       =   18
         TabIndex        =   2
         Top             =   675
         Width           =   1110
      End
      Begin VB.TextBox txt就诊卡 
         BackColor       =   &H00EBFFFF&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2820
         TabIndex        =   1
         Top             =   255
         Width           =   1110
      End
      Begin VB.TextBox txt病人ID 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         TabIndex        =   0
         Top             =   255
         Width           =   1110
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位号"
         Height          =   180
         Left            =   2250
         TabIndex        =   18
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label lbl身份证 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证"
         Height          =   180
         Left            =   300
         TabIndex        =   17
         Top             =   1575
         Width           =   540
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   480
         TabIndex        =   16
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   2250
         TabIndex        =   15
         Top             =   735
         Width           =   540
      End
      Begin VB.Label lbl门诊号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Left            =   300
         TabIndex        =   14
         Top             =   735
         Width           =   540
      End
      Begin VB.Label lbl就诊卡 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡"
         Height          =   180
         Left            =   2250
         TabIndex        =   13
         Top             =   315
         Width           =   540
      End
      Begin VB.Label lbl病人ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         Height          =   180
         Left            =   300
         TabIndex        =   12
         Top             =   315
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmPatiFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mbytType As Byte

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub



Private Sub cmdOK_Click()
    If txt病人ID.Text = "" And txt就诊卡.Text = "" And txt门诊号.Text = "" And txt住院号.Text = "" And txt姓名.Text = "" And txt床号.Text = "" And txt身份证.Text = "" Then
        MsgBox "请至少输入一个定位条件！", vbInformation, gstrSysName
        txt病人ID.SetFocus: Exit Sub
    End If
    gblnOK = True
    Hide
End Sub

Private Sub Form_Activate()
    txt病人ID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    If glngSys Like "8??" Then
        lbl门诊号.Visible = False
        lbl住院号.Visible = False
        txt门诊号.Visible = False
        txt住院号.Visible = False
        lbl床号.Visible = False
        txt床号.Visible = False
        
        lbl病人ID.Caption = "客户ID"
        lbl就诊卡.Caption = "会员卡"
        
        lbl姓名.Top = lbl姓名.Top - 420
        txt姓名.Top = txt姓名.Top - 420
        lbl身份证.Top = lbl身份证.Top - 420
        txt身份证.Top = txt身份证.Top - 420
        
        optHead.Top = optHead.Top - 420
        optCur.Top = optCur.Top - 420
        fraBdr.Height = fraBdr.Height - 420
        Me.Height = Me.Height - 420
    End If

    txt就诊卡.Enabled = gblnShowCard
    If Not txt就诊卡.Enabled Then txt就诊卡.BackColor = Me.BackColor
    
    Select Case mbytType
        Case 0 '所有病人
        Case 1 '在院病人
            txt门诊号.Enabled = False
            txt门诊号.BackColor = Me.BackColor
        Case 2 '出院病人
            txt门诊号.Enabled = False
            txt门诊号.BackColor = Me.BackColor
            txt床号.Enabled = False
            txt床号.BackColor = Me.BackColor
        Case 3 '门诊病人
            txt住院号.Enabled = False
            txt住院号.BackColor = Me.BackColor
            txt床号.Enabled = False
            txt床号.BackColor = Me.BackColor
    End Select
End Sub

Private Sub txt病人ID_GotFocus()
    zlControl.TxtSelAll txt病人ID
End Sub

Private Sub txt病人ID_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt床号_KeyPress(KeyAscii As Integer)
    If InStr("' " & Chr(8), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt就诊卡_GotFocus()
    zlControl.TxtSelAll txt就诊卡
End Sub

Private Sub txt就诊卡_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt门诊号_GotFocus()
    zlControl.TxtSelAll txt门诊号
End Sub

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt身份证_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then KeyAscii = 0
End Sub
'问题29712 by lesfeng 2010-05-11
Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("[]:：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
End Sub

Private Sub txt床号_GotFocus()
    zlControl.TxtSelAll txt床号
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub txt身份证_GotFocus()
    zlControl.TxtSelAll txt身份证
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
