VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm����תԺ_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   Icon            =   "frm����תԺ_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2400
      TabIndex        =   8
      Top             =   1860
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1170
      TabIndex        =   7
      Top             =   1860
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "����(&S)"
      Height          =   1605
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3435
      Begin VB.ComboBox cbo��˱�־ 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtp��ʼ���� 
         Height          =   300
         Left            =   1410
         TabIndex        =   2
         Top             =   300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   64552963
         CurrentDate     =   38063
      End
      Begin MSComCtl2.DTPicker Dtp�������� 
         Height          =   300
         Left            =   1410
         TabIndex        =   4
         Top             =   690
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   64552963
         CurrentDate     =   38063
      End
      Begin VB.Label lbl��˱�־ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��˱�־(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   1140
         Width           =   990
      End
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   750
         Width           =   990
      End
      Begin VB.Label lbl��ʼ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
   End
End
Attribute VB_Name = "frm����תԺ_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strStart As String
Public strEnd As String
Public strState As String
Private blnOK As Boolean

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    strStart = Format(Me.dtp��ʼ����.Value, "yyyy-MM-dd")
    strEnd = Format(Me.Dtp��������.Value, "yyyy-MM-dd")
    strState = Me.cbo��˱�־.ItemData(Me.cbo��˱�־.ListIndex)
    If strState = -1 Then strState = "all"
    
    blnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Me.dtp��ʼ����.Value = Format(DateAdd("d", -10, zldatabase.Currentdate()), "yyyy��MM��DD��")
    Me.Dtp��������.Value = Format(zldatabase.Currentdate(), "yyyy��MM��DD��")
    With cbo��˱�־
        .Clear
        .AddItem "δ���"
        .ItemData(.NewIndex) = 0
        .AddItem "���ͨ��"
        .ItemData(.NewIndex) = 1
        .AddItem "���δͨ��"
        .ItemData(.NewIndex) = 2
        .AddItem "ȫ��תԺ����"
        .ItemData(.NewIndex) = -1
        .ListIndex = 0
    End With
End Sub

Public Function ShowME(str��ʼ���� As String, str�������� As String, str��˱�־ As String) As Boolean
    blnOK = False
    
    Me.Show 1
    
    str��ʼ���� = strStart
    str�������� = strEnd
    str��˱�־ = strState
    ShowME = blnOK
End Function
