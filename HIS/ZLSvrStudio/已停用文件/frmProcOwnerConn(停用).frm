VERSION 5.00
Begin VB.Form frmProcOwnerConn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������������"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3765
   Icon            =   "frmProcOwnerConn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtPwd 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   825
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   420
      Width           =   2790
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1380
      TabIndex        =   1
      Top             =   825
      Width           =   1100
   End
   Begin VB.CommandButton cmdJump 
      Caption         =   "����(&J)"
      Height          =   350
      Left            =   2565
      TabIndex        =   0
      Top             =   825
      Width           =   1100
   End
   Begin VB.Label Label1 
      Caption         =   "���룺"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   495
      Width           =   570
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "����������������"
      Height          =   180
      Left            =   195
      TabIndex        =   3
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmProcOwnerConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjMain As Object
Private mstrPwd As String
Private mblnValid As Boolean
Private mblnOK As Boolean
Private mstrOwner As String
Public Function ShowDialog(ByVal objMain As Object, ByVal strOwner As String, ByRef strPwd As String) As Boolean
    Set mobjMain = objMain
    mstrOwner = strOwner
    Me.Show 1, mobjMain
    strPwd = mstrPwd
    ShowDialog = mblnValid
End Function

Private Sub cmdJump_Click()
    If MsgBox("�����˲��轫�޷���֤������ȷ��ͬʱ�޷����棬ȷ��������", vbInformation + vbOKCancel, "�������") = vbOK Then
        mblnOK = True
        mblnValid = False
        Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    mblnOK = True
    mstrPwd = txtPWD.Text
    Unload Me
End Sub

Private Sub Form_Load()
    mblnOK = False
    lblTitle.Caption = "������" & mstrOwner & "������"
    mblnValid = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not mblnOK Then
        Cancel = 1
    End If
End Sub
