VERSION 5.00
Begin VB.Form frmDeptGo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定位设置"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraPati 
      Height          =   1440
      Left            =   285
      TabIndex        =   14
      Top             =   0
      Width           =   4515
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1245
         MaxLength       =   18
         TabIndex        =   4
         Top             =   615
         Width           =   1275
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   6
         Top             =   990
         Width           =   3075
      End
      Begin VB.TextBox txt床号 
         Height          =   300
         Left            =   1245
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号(&2)"
         Height          =   180
         Left            =   375
         TabIndex        =   3
         Top             =   675
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&3)"
         Height          =   180
         Left            =   555
         TabIndex        =   5
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号(&1)"
         Height          =   180
         Left            =   555
         TabIndex        =   1
         Top             =   300
         Width           =   630
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4965
      TabIndex        =   13
      Top             =   1500
      Width           =   4965
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3690
         TabIndex        =   10
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2370
         TabIndex        =   9
         Top             =   135
         Width           =   1100
      End
   End
   Begin VB.Frame fraBill 
      Height          =   735
      Left            =   285
      TabIndex        =   11
      Top             =   30
      Width           =   4395
      Begin VB.OptionButton optCur 
         Caption         =   "向下"
         Height          =   195
         Left            =   3330
         TabIndex        =   8
         Top             =   315
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton optHead 
         Caption         =   "所有"
         Height          =   195
         Left            =   2610
         TabIndex        =   7
         Top             =   315
         Width           =   660
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   0
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   525
         TabIndex        =   12
         Top             =   315
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmDeptGo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub



Private Sub cmdOK_Click()
    If fraBill.Visible Then
        If txtNO.Text = "" Then
            MsgBox "请设定定位条件！", vbInformation, gstrSysName
            txtNO.SetFocus: Exit Sub
        End If
    Else
        If txt住院号.Text = "" And txt姓名.Text = "" And txt床号.Text = "" Then
            MsgBox "请至少设定一个条件！", vbInformation, gstrSysName
            txt床号.SetFocus: Exit Sub
        End If
    End If
    '问题:30532
    If InStr(1, txtNO.Text, "[") > 0 Then
        MsgBox "单据号中含用非法字符[]", vbInformation, gstrSysName
        txtNO.SetFocus: Exit Sub
    End If
    If InStr(1, txtNO.Text, "]") > 0 Then
        MsgBox "单据号中含用非法字符[]", vbInformation, gstrSysName
        txtNO.SetFocus: Exit Sub
    End If
    If InStr(1, txt姓名.Text, "[") > 0 Then
        MsgBox "姓名中含用非法字符[]", vbInformation, gstrSysName
        txt姓名.SetFocus: Exit Sub
    End If
    If InStr(1, txt姓名.Text, "]") > 0 Then
        MsgBox "姓名中含用非法字符[]", vbInformation, gstrSysName
        txt姓名.SetFocus: Exit Sub
    End If
    
    gblnOK = True
    Hide
End Sub

Private Sub Form_Activate()
    If fraPati.Visible Then
        txt床号.SetFocus
    Else
        txtNO.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
    If InStr(1, "[]", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    gblnOK = False
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
End Sub

Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> 13 Then
        If Not (txtNO.Text = "" Or txtNO.SelLength = Len(txtNO.Text) Or txtNO.SelStart = 0) And _
            InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

Private Sub txtNO_LostFocus()
    If txtNO.Text <> "" Then txtNO.Text = GetFullNO(txtNO.Text, 14)
End Sub

Private Sub txt床号_GotFocus()
    zlControl.TxtSelAll txt床号
End Sub


Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
