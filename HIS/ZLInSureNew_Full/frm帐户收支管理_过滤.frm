VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�ʻ���֧����_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frm�ʻ���֧����_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4410
      TabIndex        =   19
      Top             =   2490
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3180
      TabIndex        =   18
      Top             =   2490
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   210
      TabIndex        =   20
      Top             =   2490
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "����(&F)"
      Height          =   2295
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   5445
      Begin VB.CheckBox chk��ʾ�����嵥 
         Caption         =   "��ʾ�����嵥(&D)"
         Height          =   225
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   3600
         MaxLength       =   16
         TabIndex        =   16
         Top             =   1500
         Width           =   1605
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1500
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   300
         Left            =   1050
         TabIndex        =   2
         Top             =   330
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   62521347
         CurrentDate     =   37914
      End
      Begin VB.TextBox txt���_���� 
         Height          =   300
         Left            =   3600
         MaxLength       =   16
         TabIndex        =   12
         Top             =   1110
         Width           =   1605
      End
      Begin VB.TextBox txt���_��ʼ 
         Height          =   300
         Left            =   1050
         MaxLength       =   16
         TabIndex        =   10
         Top             =   1110
         Width           =   1605
      End
      Begin VB.TextBox txt����_���� 
         Height          =   300
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   8
         Top             =   720
         Width           =   1605
      End
      Begin VB.TextBox txt����_��ʼ 
         Height          =   300
         Left            =   1050
         MaxLength       =   20
         TabIndex        =   6
         Top             =   720
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   300
         Left            =   3600
         TabIndex        =   4
         Top             =   330
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   62521347
         CurrentDate     =   37914
      End
      Begin VB.Label lbl������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2760
         TabIndex        =   15
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label lblҽ������ 
         AutoSize        =   -1  'True
         Caption         =   "����(&R)"
         Height          =   180
         Left            =   330
         TabIndex        =   13
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3030
         TabIndex        =   3
         Top             =   390
         Width           =   180
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   3030
         TabIndex        =   11
         Top             =   1170
         Width           =   180
      End
      Begin VB.Label lbl��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���(&M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   9
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3030
         TabIndex        =   7
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lbl����_��ʼ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   5
         Top             =   780
         Width           =   630
      End
   End
End
Attribute VB_Name = "frm�ʻ���֧����_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private mstrFind As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrFind = ""
    
    '��ϲ��Ҵ�
    mstrFind = mstrFind & " And Trunc(B.ʱ��) Between To_Date('" & Format(dtp��ʼʱ��.Value, "yyyy-MM-dd") & "','yyyy-MM-dd')" & _
            " And To_Date('" & Format(dtp����ʱ��.Value, "yyyy-MM-dd") & "','yyyy-MM-dd')"
    If Trim(txt����_��ʼ.Text) <> "" Then mstrFind = mstrFind & " And A.����>='" & UCase(Trim(txt����_��ʼ.Text)) & "'"
    If Trim(txt����_����.Text) <> "" Then mstrFind = mstrFind & " And A.����<='" & UCase(Trim(txt����_����.Text)) & "'"
    If Trim(txt���_��ʼ.Text) <> "" Then mstrFind = mstrFind & " And B.���>='" & Val(txt���_��ʼ.Text) & "'"
    If Trim(txt���_����.Text) <> "" Then mstrFind = mstrFind & " And B.����<='" & Val(txt���_����.Text) & "'"
    If cbo����.ListIndex <> 0 Then mstrFind = mstrFind & " And A.����=" & cbo����.ItemData(cbo����.ListIndex)
    If Trim(txt������.Text) <> "" Then mstrFind = mstrFind & " And B.������='" & Trim(txt������.Text) & "'"
    If chk��ʾ�����嵥.Value = 0 Then mstrFind = mstrFind & " And B.����=1"
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    Me.dtp��ʼʱ��.Value = Format(DateAdd("m", -1, zlDataBase.Currentdate()), "yyyy��MM��dd��")
    Me.dtp����ʱ��.Value = Format(zlDataBase.Currentdate(), "yyyy��MM��dd��")
    txt������ = gstrUserName
    
    gstrSQL = "Select ����,��� ID From ��������Ŀ¼ Where ����=" & mint����
    Call OpenRecordset(rsTemp, Me.Caption)
    cbo����.Clear
    cbo����.AddItem "����ҽ������"
    cbo����.ItemData(cbo����.NewIndex) = 0
    Call zlControl.CboAddData(Me.cbo����, rsTemp, False)
    Me.cbo����.ListIndex = 0
End Sub

Public Function ShowME(ByVal frmParent As Object, ByVal int���� As Integer) As String
    mstrFind = ""
    
    mint���� = int����
    Me.Show 1, frmParent
    ShowME = mstrFind
End Function
