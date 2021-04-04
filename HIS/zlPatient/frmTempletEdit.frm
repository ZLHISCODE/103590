VERSION 5.00
Begin VB.Form frmTempletEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀģ��༭"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2280
      TabIndex        =   5
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   720
      MaxLength       =   3
      TabIndex        =   3
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   720
      MaxLength       =   20
      TabIndex        =   2
      Top             =   420
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
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
Private mstr���� As String

Private Sub cmdCancel_Click()
    mstr���� = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not IsNumeric(txt����.Text) Then
        MsgBox "���������������!", vbInformation, gstrSQL
        txt����.SetFocus
        Exit Sub
    End If

    If zlCommFun.StrIsValid(txt����.Text, 20) = False Then Exit Sub
     '����30020 by lesfeng 2010-06-01 �������Ϊ�����
    If Trim(txt����.Text) = "" Then
        MsgBox "���Ʊ�������!", vbInformation + vbOKOnly, gstrSysName
        If txt����.Enabled Then txt����.SetFocus
        Exit Sub
    End If
          
    If InStr(1, txt����.Text, "'") > 0 Then
        MsgBox "���ƴ��ڷǷ��ַ�!", vbInformation + vbOKOnly, gstrSysName
        If txt����.Enabled Then txt����.SetFocus
        Exit Sub
    End If
    
    If Exist����(txt����.Text, txt����.Text) = True Then
        MsgBox "�ñ���������ظ�,��������!", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    mstr���� = txt����.Text & "," & txt����.Text
    Unload Me
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
    zlCommFun.OpenIme False
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt����.SetFocus
    Else
        If InStr("0123456789" & vbKeyBack, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
    zlCommFun.OpenIme True
End Sub

Public Function EditTemplet(frmMain As Object) As String
    txt����.Text = Get����
    frmTempletEdit.Show 1, frmMain
    EditTemplet = mstr����
End Function

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdOk.SetFocus
    Else
        If InStr("'?/<>~!@#$%^&*()_+|-=\,.", Chr(KeyAscii)) <> 0 Then KeyAscii = 0
    End If
End Sub

Private Function Exist����(lng���� As Long, str���� As String) As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    'by lesfeng 2010-03-08 �����Ż� select *
    strSQL = "select ����,����,��ĿID from ������Ŀģ�� where ����=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng����)
    If Not rsTemp.EOF Then Exist���� = True: Exit Function
    strSQL = "select ����,����,��ĿID from ������Ŀģ�� where ����=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    If Not rsTemp.EOF Then Exist���� = True: Exit Function
    Exit Function
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get����() As Long
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "select (nvl(max(����),0)+1) ���� from ������Ŀģ��"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    Get���� = rsTemp!����
    Exit Function
errHandle:
    If errCenter() = 1 Then Resume
    Call SaveErrLog
End Function

