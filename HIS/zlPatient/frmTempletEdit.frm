VERSION 5.00
Begin VB.Form frmTempletEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "审批项目模板编辑"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3585
   Icon            =   "frmTempletEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2280
      TabIndex        =   5
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   1100
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   720
      MaxLength       =   3
      TabIndex        =   3
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   720
      MaxLength       =   20
      TabIndex        =   2
      Top             =   420
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "名称"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "编码"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmTempletEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr名称 As String

Private Sub cmdCancel_Click()
    mstr名称 = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not IsNumeric(txt编码.Text) Then
        MsgBox "编码必须输入数字!", vbInformation, gstrSQL
        txt编码.SetFocus
        Exit Sub
    End If

    If zlCommFun.StrIsValid(txt名称.Text, 20) = False Then Exit Sub
     '问题30020 by lesfeng 2010-06-01 解决名称为空情况
    If Trim(txt名称.Text) = "" Then
        MsgBox "名称必须输入!", vbInformation + vbOKOnly, gstrSysName
        If txt名称.Enabled Then txt名称.SetFocus
        Exit Sub
    End If
          
    If InStr(1, txt名称.Text, "'") > 0 Then
        MsgBox "名称存在非法字符!", vbInformation + vbOKOnly, gstrSysName
        If txt名称.Enabled Then txt名称.SetFocus
        Exit Sub
    End If
    
    If Exist编码(txt编码.Text, txt名称.Text) = True Then
        MsgBox "该编码或名称重复,不能新增!", vbInformation, gstrSysName
        txt编码.SetFocus
        Exit Sub
    End If
    
    mstr名称 = txt编码.Text & "," & txt名称.Text
    Unload Me
End Sub

Private Sub txt编码_GotFocus()
    Call zlControl.TxtSelAll(txt编码)
    zlCommFun.OpenIme False
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt名称.SetFocus
    Else
        If InStr("0123456789" & vbKeyBack, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt名称_GotFocus()
    Call zlControl.TxtSelAll(txt名称)
    zlCommFun.OpenIme True
End Sub

Public Function EditTemplet(frmMain As Object) As String
    txt编码.Text = Get编码
    frmTempletEdit.Show 1, frmMain
    EditTemplet = mstr名称
End Function

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdOk.SetFocus
    Else
        If InStr("'?/<>~!@#$%^&*()_+|-=\,.", Chr(KeyAscii)) <> 0 Then KeyAscii = 0
    End If
End Sub

Private Function Exist编码(lng编码 As Long, str名称 As String) As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    'by lesfeng 2010-03-08 性能优化 select *
    strSQL = "select 编码,名称,项目ID from 审批项目模板 where 编码=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng编码)
    If Not rsTemp.EOF Then Exist编码 = True: Exit Function
    strSQL = "select 编码,名称,项目ID from 审批项目模板 where 名称=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, str名称)
    If Not rsTemp.EOF Then Exist编码 = True: Exit Function
    Exit Function
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get编码() As Long
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "select (nvl(max(编码),0)+1) 编码 from 审批项目模板"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    Get编码 = rsTemp!编码
    Exit Function
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Function

