VERSION 5.00
Begin VB.Form frmLisLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����������¼"
   ClientHeight    =   2325
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4650
   Icon            =   "frmLisLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt�û� 
      Height          =   300
      Left            =   1425
      TabIndex        =   1
      Top             =   600
      Width           =   2490
   End
   Begin VB.TextBox TXT���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1425
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1035
      Width           =   2490
   End
   Begin VB.CommandButton CMDȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2070
      TabIndex        =   6
      Top             =   1890
      Width           =   1100
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3225
      TabIndex        =   7
      Top             =   1890
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   330
      TabIndex        =   8
      Top             =   4065
      Width           =   5520
   End
   Begin VB.ComboBox cmb���ݿ� 
      Height          =   300
      Left            =   1425
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1455
      Width           =   2490
   End
   Begin VB.Label Lbl�û��� 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   795
      TabIndex        =   0
      Top             =   660
      Width           =   540
   End
   Begin VB.Label Lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��  ��"
      Height          =   180
      Left            =   795
      TabIndex        =   2
      Top             =   1095
      Width           =   540
   End
   Begin VB.Label Lbl������ 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   795
      TabIndex        =   4
      Top             =   1515
      Width           =   540
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Left            =   120
      Picture         =   "frmLisLogin.frx":000C
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "��¼��������Ժ�����������"
      Height          =   210
      Left            =   915
      TabIndex        =   9
      Top             =   150
      Width           =   3060
   End
End
Attribute VB_Name = "frmLisLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mstrUser As String
Private mstrPsw As String
Private mstrSvr As String
Private mblnOK As Boolean

Private mblnStartUp As Boolean

Public Function ShowLogin(ByVal frmMain As Object, ByRef strUser As String, ByRef strPsw As String, ByRef strSvr As String) As Boolean
    
    mblnStartUp = True
    mblnOK = False
    mstrUser = strUser
    mstrPsw = strPsw
    mstrSvr = strSvr
    
    txt�û�.Text = mstrUser
    cmb���ݿ�.Text = mstrSvr
            
    Me.Show 1, frmMain
    
    strUser = mstrUser
    strPsw = mstrPsw
    strSvr = mstrSvr
    ShowLogin = mblnOK
    
End Function

Private Sub CMD����_Click()
    Unload Me
End Sub

Private Sub CMDȷ��_Click()
    
    mstrUser = txt�û�.Text
    mstrPsw = TXT����.Text
    mstrSvr = cmb���ݿ�.Text
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
        
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    If mstrUser <> "" Then TXT����.SetFocus
    
End Sub

