VERSION 5.00
Begin VB.Form frmAuditLogClear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��־����"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3090
      TabIndex        =   4
      Top             =   1005
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1755
      TabIndex        =   3
      Top             =   1005
      Width           =   1100
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   165
      ScaleHeight     =   435
      ScaleWidth      =   2730
      TabIndex        =   0
      Top             =   210
      Width           =   2730
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   2
         Text            =   "90"
         ToolTipText     =   "������Ҫ����90�����־���ݣ�"
         Top             =   75
         Width           =   570
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "����_____��֮ǰ����־"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   1
         Top             =   90
         Width           =   2520
      End
   End
End
Attribute VB_Name = "frmAuditLogClear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean

Public Function ShowMe() As Boolean
    Me.Show vbModal, frmMDIMain
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strRemarks As String

    On Error GoTo errH
    mblnOk = False
    If Val(txtDate.Text) < 90 Then
        MsgBox "������Ҫ����90������ݣ���������д��", vbInformation, gstrSysName
        txtDate.Text = 90
        txtDate.SetFocus
        Exit Sub
    End If
    If MsgBox("ȷ��Ҫ����" & Val(txtDate.Text) & "����֮ǰ����־ȫ�������", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then Exit Sub
    '��֤��ݲ��������˵��
    If Not CheckAuditStatus("0314", "��־����", strRemarks) Then Exit Sub
    '��������ʱ�䣬ִ����־�������
    Call ExecuteProcedure("Zl_Zlauditlog_Delete(" & Val(txtDate.Text) & ")", "������������Ҫ�����䶯��־")
    MsgBox "������ɣ�", vbInformation, gstrSysName
    '������Ҫ������־
    Call SaveAuditLog(3, "��־����", "�ɹ�����" & Val(txtDate.Text) & "��֮ǰ����Ҫ������־", strRemarks)
    mblnOk = True
    Unload Me
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    'ֻ����������
    If Not (InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
