VERSION 5.00
Begin VB.Form frm���������޸� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ѡ���������"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frm���������޸�.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txt�е����� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1080
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   7
      Top             =   1590
      Width           =   4425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1770
      TabIndex        =   4
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3015
      TabIndex        =   5
      Top             =   1800
      Width           =   1100
   End
   Begin VB.ComboBox cboҽ����� 
      Height          =   300
      Left            =   1830
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   690
      Width           =   2025
   End
   Begin VB.Label lbl�е����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�е�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1020
      TabIndex        =   2
      Top             =   1140
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frm���������޸�.frx":000C
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Ϊ�ò���ѡ��������ͣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1020
      TabIndex        =   6
      Top             =   330
      Width           =   2160
   End
   Begin VB.Label lblҽ����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1020
      TabIndex        =   0
      Top             =   750
      Width           =   720
   End
End
Attribute VB_Name = "frm���������޸�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long

Private Sub cboҽ�����_Click()
    Me.txt�е�����.Enabled = False
    If cboҽ�����.ItemData(cboҽ�����.ListIndex) = 22 Then
        '��ͨ�¹�
        Me.txt�е�����.Enabled = True
        Me.txt�е�����.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex) = 22 Then
        If Val(txt�е�����.Text) < 0 Then
            MsgBox "�е���������С���㣡", vbInformation, gstrSysName
            txt�е�����.SetFocus
            Exit Sub
        End If
        If Val(txt�е�����.Text) > 100 Then
            MsgBox "�е��������ܴ���һ�٣�", vbInformation, gstrSysName
            txt�е�����.SetFocus
            Exit Sub
        End If
    End If
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_��Ϫũҽ & ",'ҵ������','''" & Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҵ������")
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Public Sub ShowME(ByVal lng����ID As Long)
    mlng����ID = lng����ID
    Me.Show 1
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    With cboҽ�����
        .AddItem "��ͨסԺ"
        .ItemData(.NewIndex) = 21
        .AddItem "��ͨ�¹�"
        .ItemData(.NewIndex) = 22
        .AddItem "�󲡾���"
        .ItemData(.NewIndex) = 23
        .AddItem "�Ѳ�"
        .ItemData(.NewIndex) = 24
        .AddItem "����"
        .ItemData(.NewIndex) = 25
        .ListIndex = 0
    End With
    
    '��ȡ�ò��˵�ǰ�Ľ�������
    gstrSQL = "Select Nvl(ҵ������,21) AS �������� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˵�ǰ��������", TYPE_��Ϫũҽ, mlng����ID)
    If rsTemp.RecordCount <> 0 Then
        cboҽ�����.ListIndex = (rsTemp!�������� - 21)
    End If
End Sub
