VERSION 5.00
Begin VB.Form frmEInvoiceChannelSet 
   Caption         =   "�շ�������������"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   Icon            =   "frmEInvoiceChannelSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5385
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   5
      Top             =   340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   3400
      Left            =   3840
      TabIndex        =   3
      Top             =   -120
      Width           =   15
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   330
      Width           =   2475
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   1
      Top             =   915
      Width           =   2475
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1455
      Width           =   2475
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���������"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   360
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���㷽ʽ"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   945
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   1485
      Width           =   720
   End
End
Attribute VB_Name = "frmEInvoiceChannelSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TXT_Idex
    Idex_��������� = 0
    Idex_���㷽ʽ = 1
    Idex_�������� = 2
End Enum
Private mByteMode As Byte      '�޸ģ�1��������0
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If Save�շ��������� = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Function Save�շ���������() As Boolean
    Dim strSQL As String

    On Error GoTo errHandle
    '�������޸��շ���������
    strSQL = "Zl_�շ���������_Update("
    '��������_In In Number,
    strSQL = strSQL & mByteMode & ","
    '���㷽ʽ_In In �շ���������.���㷽ʽ%Type,
    strSQL = strSQL & "'" & txtEdit(Idex_���㷽ʽ).Text & "',"
    '�����id_In In �շ���������.�����id%Type,
    strSQL = strSQL & ZVal(txtEdit(Idex_���������).Tag) & ","
    '��������_In In �շ���������.��������%Type
    strSQL = strSQL & "'" & txtEdit(Idex_��������).Text & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "�շ���������")
    
    Save�շ��������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
End Sub

Public Sub ShowMe(ByVal frmMain As Object, ByVal bytMode As Byte, ByVal lng�����ID As Long, ByVal str��������� As String, _
                                ByVal str���㷽ʽ As String, Optional ByVal str�������� As String, Optional blnRefresh As Boolean)
    mblnOK = False
    txtEdit(Idex_���㷽ʽ).Text = str���㷽ʽ
    txtEdit(Idex_���������).Tag = lng�����ID
    txtEdit(Idex_���������).Text = IIf(str��������� = "", "��", str���������)
    txtEdit(Idex_��������).Text = str��������
    mByteMode = bytMode
    Me.Show 1, frmMain
    blnRefresh = mblnOK
End Sub

Private Function IsValid() As Boolean
    On Error GoTo errHandle

    If Len(txtEdit(Idex_��������).Text) = 0 Then
        MsgBox "�������벻��Ϊ�ա�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_��������)
        Exit Function
    End If

    If Not IsNumeric(txtEdit(Idex_��������).Text) Or InStr(txtEdit(Idex_��������).Text, ",") > 0 Or InStr(txtEdit(Idex_��������).Text, ".") > 0 Then
        MsgBox "����Ӧ��������ɡ�", vbExclamation, gstrSysName
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

