VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5715
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CMDȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4470
      TabIndex        =   2
      Top             =   195
      Width           =   1100
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4470
      TabIndex        =   1
      Top             =   720
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "ZLHIS��Ʒ"
      Height          =   1710
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   4245
      Begin VB.TextBox txt�û� 
         Height          =   300
         Left            =   855
         TabIndex        =   5
         Text            =   "ZLHIS"
         Top             =   315
         Width           =   2850
      End
      Begin VB.TextBox TXT���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   855
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "aqa"
         Top             =   750
         Width           =   2850
      End
      Begin VB.ComboBox cmb���ݿ� 
         Height          =   300
         Left            =   855
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "2.133.ORCL"
         Top             =   1170
         Width           =   2850
      End
      Begin VB.Label Lbl�û��� 
         AutoSize        =   -1  'True
         Caption         =   "�û���"
         Height          =   180
         Left            =   135
         TabIndex        =   8
         Top             =   375
         Width           =   540
      End
      Begin VB.Label Lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   135
         TabIndex        =   7
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Lbl������ 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   135
         TabIndex        =   6
         Top             =   1230
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean

Public Function ShowLogin() As Boolean
    mblnOK = False
    
    Me.Show 1
    ShowLogin = mblnOK
End Function

Private Sub CMD����_Click()
    Unload Me
End Sub

Private Sub CMDȷ��_Click()
    Dim blnTransPassword As Boolean
    
    blnTransPassword = Not (UCase(txt�û�.Text) = "SYS" Or UCase(txt�û�.Text) = "SYSTEM")
    Set gcnOracle = gobjRegister.GetConnection(cmb���ݿ�.Text, txt�û�.Text, TXT����.Text, blnTransPassword)
    If gcnOracle.State = adStateClosed Then
        TXT����.Text = ""
        If TXT����.Enabled Then TXT����.SetFocus
        Exit Sub
    Else
        gstrDbUser = UCase(txt�û�.Text)
    End If

    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    txt�û�.Text = GetSetting(App.ProductName, "��¼", "�û�", "")
    cmb���ݿ�.Text = GetSetting(App.ProductName, "��¼", "������", "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting(App.ProductName, "��¼", "�û�", txt�û�.Text)
    Call SaveSetting(App.ProductName, "��¼", "������", cmb���ݿ�.Text)
    
End Sub
