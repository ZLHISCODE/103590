VERSION 5.00
Begin VB.Form frmInterfaceUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ZLInterface�û�����"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   Icon            =   "frmInterfaceUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "ZLInterface"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   1230
   End
   Begin VB.TextBox txtComfirmPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtNewPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   640
      Width           =   2295
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   120
      Picture         =   "frmInterfaceUser.frx":6633E
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "�û�"
      Height          =   180
      Left            =   1080
      TabIndex        =   0
      Top             =   180
      Width           =   360
   End
   Begin VB.Label lblNewPwd 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1080
      TabIndex        =   2
      Top             =   700
      Width           =   360
   End
   Begin VB.Label lblComfirmPwd 
      AutoSize        =   -1  'True
      Caption         =   "������֤"
      Height          =   180
      Left            =   720
      TabIndex        =   4
      Top             =   1220
      Width           =   720
   End
End
Attribute VB_Name = "frmInterfaceUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrOperate     As String '�޸����룬�����û�
Private mblnOK          As Boolean
'===========================================================================
'==�����ӿ�
'===========================================================================
Public Function ShowMe(ByVal strOperate As String) As Boolean
    mstrOperate = strOperate
    mblnOK = False
    Me.Show vbModal, frmMDIMain
    ShowMe = mblnOK
End Function

'===========================================================================
'==�¼�
'===========================================================================
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strError  As String
    
    On Error GoTo errH
    If Trim(txtNewPWD.Text) = "" Then
        MsgBox "���������룡", vbInformation, gstrSysName
        txtNewPWD.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtNewPWD.Text)) < 8 Then
        MsgBox "����������8λ�����룡", vbInformation, gstrSysName
        txtNewPWD.SetFocus
        Exit Sub
    End If
    
    If Trim(txtComfirmPwd.Text) = "" Then
        MsgBox "������������֤��", vbInformation, gstrSysName
        txtComfirmPwd.SetFocus
        Exit Sub
    End If
    If txtNewPWD.Text <> txtComfirmPwd.Text Then
        MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
        txtComfirmPwd.SetFocus
        Exit Sub
    End If
    If Not RepairGeneralAccount(gcnOracle, "ZLINTERFACE", Trim(txtNewPWD.Text), strError) Then
        MsgBox mstrOperate & "ʧ�ܡ���Ϣ��" & strError, vbInformation, gstrSysName
    End If
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    MsgBox mstrOperate & "ʧ�ܡ���Ϣ��" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub Form_Activate()
    txtNewPWD.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0: PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    If mstrOperate = "�����û�" Then
        Me.Caption = "ZLInterface�û�����"
    Else
        Me.Caption = "ZLInterface�û���������"
    End If
    
    HookDefend txtNewPWD.hwnd
End Sub

