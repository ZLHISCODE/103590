VERSION 5.00
Begin VB.Form frmSetParameter 
   Caption         =   "��������"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4410
   Icon            =   "frmSetParameter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4410
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3225
      TabIndex        =   1
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CheckBox chkFreeInput 
      Caption         =   "�������޸Ĳ���ʱ����¼�����"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "frmSetParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    '�������
    Call zlDatabase.SetPara("����¼�����", IIF(chkFreeInput.Value = 1, 1, 0), glngSys, 1001)
    Unload Me
End Sub

Private Sub Form_Load()
    chkFreeInput.Value = IIF(Val(zlDatabase.GetPara("����¼�����", glngSys, 1001, "0")) = 1, 1, 0)
End Sub

Public Sub ShowMe(ByVal frmParent As Form)
    Me.Show vbModal, frmParent
End Sub

