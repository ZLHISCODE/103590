VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frm���ڷ�Χ_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ѯ��Χ����"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   Icon            =   "frm���ڷ�Χ_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2790
      TabIndex        =   6
      Top             =   1590
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1590
      TabIndex        =   5
      Top             =   1590
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   1305
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   3705
      Begin MSMask.MaskEdBox txt��ʼ���� 
         Height          =   300
         Left            =   1440
         TabIndex        =   2
         Top             =   330
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   19
         Format          =   "yyyy-MM-dd HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt�������� 
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   19
         Format          =   "yyyy-MM-dd HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   3
         Top             =   780
         Width           =   990
      End
      Begin VB.Label lbl��ʼ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   390
         Width           =   990
      End
   End
End
Attribute VB_Name = "frm���ڷ�Χ_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnReturn As Boolean
Private mstr��ʼ���� As String
Private mstr�������� As String

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    blnReturn = True
    mstr��ʼ���� = txt��ʼ����.Text
    mstr�������� = txt��������.Text
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    blnReturn = False
    Me.txt��ʼ����.Text = Format(DateAdd("m", -1, zlDatabase.Currentdate()), "yyyy-MM-dd") & " 00:00:00"
    Me.txt��������.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd") & " 23:59:59"
End Sub

Public Function Show_ME(str��ʼ���� As String, str�������� As String) As Boolean
    Me.Show 1
    str��ʼ���� = mstr��ʼ����
    str�������� = mstr��������
    Show_ME = blnReturn
End Function
