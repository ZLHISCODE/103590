VERSION 5.00
Begin VB.Form frmLabMBSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5700
      TabIndex        =   16
      Top             =   2610
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4110
      TabIndex        =   15
      Top             =   2610
      Width           =   1100
   End
   Begin VB.Frame fra������� 
      Caption         =   "�������"
      Height          =   2325
      Left            =   2820
      TabIndex        =   11
      Top             =   120
      Width           =   4365
      Begin VB.TextBox txt���Զ��� 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   698
         Width           =   1185
      End
      Begin VB.CheckBox chk���Զ��� 
         Caption         =   "���Զ���С��              ʱ���趨ֵ����"
         Height          =   180
         Left            =   180
         TabIndex        =   13
         Top             =   750
         Width           =   4065
      End
      Begin VB.CheckBox chk�հ׶��� 
         Caption         =   "�Ƿ��ȥ�հ׶���"
         Height          =   345
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Width           =   2445
      End
   End
   Begin VB.Frame fraͨѶ���� 
      Caption         =   "ͨѶ����"
      Height          =   2355
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2625
      Begin VB.ComboBox cboͨѶ�� 
         Height          =   300
         ItemData        =   "frmLabMBSetup.frx":0000
         Left            =   1095
         List            =   "frmLabMBSetup.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1290
      End
      Begin VB.ComboBox cbo������ 
         Height          =   300
         ItemData        =   "frmLabMBSetup.frx":0004
         Left            =   1095
         List            =   "frmLabMBSetup.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   645
         Width           =   1290
      End
      Begin VB.ComboBox cbo����λ 
         Height          =   300
         ItemData        =   "frmLabMBSetup.frx":0008
         Left            =   1095
         List            =   "frmLabMBSetup.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1050
         Width           =   1290
      End
      Begin VB.ComboBox cboֹͣλ 
         Height          =   300
         ItemData        =   "frmLabMBSetup.frx":000C
         Left            =   1095
         List            =   "frmLabMBSetup.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1455
         Width           =   1290
      End
      Begin VB.ComboBox cboУ��λ 
         Height          =   300
         ItemData        =   "frmLabMBSetup.frx":0010
         Left            =   1095
         List            =   "frmLabMBSetup.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1875
         Width           =   1290
      End
      Begin VB.Label lblͨѶ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͨѶ��(&1)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   810
      End
      Begin VB.Label lbl������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������(&2)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   705
         Width           =   810
      End
      Begin VB.Label lbl����λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����λ(&3)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label lblֹͣλ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ֹͣλ(&4)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   1515
         Width           =   810
      End
      Begin VB.Label lblУ��λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "У��λ(&5)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   1935
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmLabMBSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub chk���Զ���_Click()
    Me.txt���Զ���.Enabled = Me.chk���Զ���.Value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    zlDatabase.SetPara "frmLabMB_ͨѶ��", Me.cboͨѶ��.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_������", Me.cbo������.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_����λ", Me.cbo����λ.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_ֹͣλ", Me.cboֹͣλ.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_У��λ", Me.cboУ��λ.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_���հ׶���", Me.chk�հ׶���.Value, 100, 1208
    zlDatabase.SetPara "frmLabMB_���Զ���", Me.chk���Զ���.Value & "," & Me.txt���Զ���.Text, 100, 1208
    Unload Me
End Sub

Private Sub Form_Load()
    Dim aryTemp() As String
    Dim lngCount As Long
    
    '�����̶�����װ��
    For lngCount = 1 To 50: Me.cboͨѶ��.AddItem "COM" & lngCount: Next
    Me.cboͨѶ��.ListIndex = 0

    aryTemp = Split("110|300|600|1200|2400|4800|9600|14400|19200|28800|38400|56000|128000|256000", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo������.AddItem aryTemp(lngCount): Next
    Me.cbo������.ListIndex = 0

    aryTemp = Split("4|5|6|7|8", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo����λ.AddItem aryTemp(lngCount): Next
    Me.cbo����λ.ListIndex = 0

    aryTemp = Split("1|1.5|2", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp):
        Me.cboֹͣλ.AddItem aryTemp(lngCount):
    Next
    Me.cboֹͣλ.ListIndex = 0

    aryTemp = Split("E-ż��|M-���|N-ȱʡ|None|O-����|S-�ո�", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cboУ��λ.AddItem aryTemp(lngCount): Next
    Me.cboУ��λ.ListIndex = 0
    
    On Error Resume Next
    
    Me.cboͨѶ�� = zlDatabase.GetPara("frmLabMB_ͨѶ��", 100, 1208, "")
    Me.cbo������ = zlDatabase.GetPara("frmLabMB_������", 100, 1208, "")
    Me.cbo����λ = zlDatabase.GetPara("frmLabMB_����λ", 100, 1208, "")
    Me.cboֹͣλ = zlDatabase.GetPara("frmLabMB_ֹͣλ", 100, 1208, "")
    Me.cboУ��λ = zlDatabase.GetPara("frmLabMB_У��λ", 100, 1208, "")
    Me.chk�հ׶��� = zlDatabase.GetPara("frmLabMB_���հ׶���", 100, 1208, "0")
    Me.chk���Զ���.Value = Mid(zlDatabase.GetPara("frmLabMB_���Զ���", 100, 1208, "0,"), 1, 1)
    Me.txt���Զ���.Text = Mid(zlDatabase.GetPara("frmLabMB_���Զ���", 100, 1208, "0,"), 3)
End Sub
