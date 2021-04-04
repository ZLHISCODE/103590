VERSION 5.00
Begin VB.Form frmHosRegFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定位设置"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   120
      TabIndex        =   9
      Top             =   30
      Width           =   4365
      Begin VB.TextBox txt床号 
         Height          =   300
         Left            =   915
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1080
         Width           =   1110
      End
      Begin VB.TextBox txt病人ID 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         MaxLength       =   10
         TabIndex        =   0
         Top             =   255
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
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   915
         MaxLength       =   10
         TabIndex        =   2
         Top             =   667
         Width           =   1110
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   3
         Top             =   667
         Width           =   1110
      End
      Begin VB.OptionButton optCur 
         Caption         =   "向下"
         Height          =   180
         Left            =   3180
         TabIndex        =   6
         Top             =   1140
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton optHead 
         Caption         =   "所有"
         Height          =   180
         Left            =   2460
         TabIndex        =   5
         Top             =   1140
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         Height          =   180
         Left            =   300
         TabIndex        =   13
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡"
         Height          =   180
         Left            =   2220
         TabIndex        =   12
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   300
         TabIndex        =   11
         Top             =   727
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   2400
         TabIndex        =   10
         Top             =   720
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2100
      TabIndex        =   7
      Top             =   1695
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3315
      TabIndex        =   8
      Top             =   1695
      Width           =   1100
   End
End
Attribute VB_Name = "frmHosRegFind"
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
    If txt病人ID.Text = "" And txt就诊卡.Text = "" And txt住院号.Text = "" And txt姓名.Text = "" And txt床号.Text = "" Then
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
   Dim strSql As String, rsTemp As ADODB.Recordset
   Dim blnPassShowCard As Boolean
   
    On Error GoTo errHandle
    strSql = "Select 卡号密文 From 医疗卡类别 where 名称='就诊卡' and 是否固定=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTemp.EOF Then
        blnPassShowCard = Nvl(rsTemp!卡号密文) = ""
    End If
    txt就诊卡.Enabled = blnPassShowCard
    If Not blnPassShowCard Then txt就诊卡.BackColor = Me.BackColor
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt病人ID_GotFocus()
    zlControl.TxtSelAll txt病人ID
End Sub

Private Sub txt病人ID_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt床号_GotFocus()
    zlControl.TxtSelAll txt床号
End Sub

Private Sub txt床号_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("[]:：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
'    If InStr("' " & Chr(8), Chr(KeyAscii)) <> 0 Then KeyAscii = 0
    '问题29712 by lesfeng 2010-05-11
End Sub

Private Sub txt就诊卡_GotFocus()
    zlControl.TxtSelAll txt就诊卡
End Sub
Private Sub txt就诊卡_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub
'问题29741 by lesfeng 2010-05-12
Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("[]:：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
