VERSION 5.00
Begin VB.Form frmEInvoiceInsureSet 
   Caption         =   "֧������������"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   Icon            =   "frmEInvoiceInsureSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5385
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   960
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1920
      Width           =   2955
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   960
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1390
      Width           =   2955
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   860
      Width           =   2955
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   330
      Width           =   2955
   End
   Begin VB.Frame fra 
      Height          =   3400
      Left            =   4000
      TabIndex        =   6
      Top             =   -120
      Width           =   15
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4140
      TabIndex        =   5
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4140
      TabIndex        =   4
      Top             =   345
      Width           =   1100
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   10
      Top             =   1950
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   9
      Top             =   1425
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "֧������"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   8
      Top             =   885
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "frmEInvoiceInsureSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TXT_Idex
    Idex_�������� = 0
    Idex_֧������ = 1
    Idex_������� = 2
    Idex_�������� = 3
End Enum
Private mByteMode As Byte      '�޸ģ�1��������0
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If Save֧�������� = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Function Save֧��������() As Boolean
    Dim strSQL As String

    On Error GoTo errHandle
    '�������޸��շ���������
    strSQL = "Zl_֧��������_Update("
    '��������_In In Number,
    strSQL = strSQL & mByteMode & ","
    '���մ���id_In In ֧��������.���մ���id%Type,
    strSQL = strSQL & Val(txtEdit(Idex_֧������).Tag) & ","
    '�������_In   In ֧��������.�������%Type := Null,
    strSQL = strSQL & "'" & txtEdit(Idex_�������).Text & "',"
    '��������_In   In ֧��������.��������%Type := Null
    strSQL = strSQL & "'" & txtEdit(Idex_��������).Text & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "�շ���������")
    
    Save֧�������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = Idex_������� Then
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    End If
End Sub

Public Sub ShowMe(ByVal frmMain As Object, ByVal bytMode As Byte, ByVal str�������� As String, ByVal lng���մ���ID As Long, _
                                ByVal str֧������ As String, Optional ByVal str������� As String, Optional ByVal str�������� As String, _
                                Optional blnRefresh As Boolean)
    mblnOK = False
    txtEdit(Idex_��������).Text = str��������
    txtEdit(Idex_֧������).Tag = lng���մ���ID
    txtEdit(Idex_֧������).Text = str֧������
    txtEdit(Idex_�������).Text = str�������
    txtEdit(Idex_��������).Text = str��������
    mByteMode = bytMode
    Me.Show 1, frmMain
    blnRefresh = mblnOK
End Sub

Private Function IsValid() As Boolean
    On Error GoTo errHandle

    If Len(txtEdit(Idex_�������).Text) = 0 Then
        MsgBox "������벻��Ϊ�ա�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_�������)
        Exit Function
    End If

    If Not IsNumeric(txtEdit(Idex_�������).Text) Or InStr(txtEdit(Idex_�������).Text, ",") > 0 Or InStr(txtEdit(Idex_�������).Text, ".") > 0 Then
        MsgBox "����Ӧ��������ɡ�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_�������)
        Exit Function
    End If
    
    If Len(txtEdit(Idex_��������).Text) = 0 Then
        MsgBox "�������Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_��������)
        Exit Function
    End If
    
    IsValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



