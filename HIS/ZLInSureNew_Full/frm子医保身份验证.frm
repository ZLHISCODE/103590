VERSION 5.00
Begin VB.Form frm��ҽ�������֤ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ҽ�������֤"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "frm��ҽ�������֤.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2490
      TabIndex        =   12
      Top             =   2610
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1230
      TabIndex        =   11
      Top             =   2610
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   10
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox txt�ʻ���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1140
      TabIndex        =   9
      Top             =   1860
      Width           =   2415
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1140
      TabIndex        =   7
      Top             =   1470
      Width           =   2415
   End
   Begin VB.TextBox txtҽ���� 
      Height          =   300
      Left            =   1140
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox cbo�Ż���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   690
      Width           =   2415
   End
   Begin VB.ComboBox cbo��ҽ�� 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   2415
   End
   Begin VB.Label lbl�ʻ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ʻ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   6
      Top             =   1530
      Width           =   360
   End
   Begin VB.Label lblҽ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   540
      TabIndex        =   4
      Top             =   1140
      Width           =   540
   End
   Begin VB.Label lbl�Ż���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ż����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   750
      Width           =   720
   End
   Begin VB.Label lbl��ҽ�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ҽ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   540
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "frm��ҽ�������֤"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim marrData
Dim mstrReg As String
Dim mstr��� As String
Private mstr���� As String
Dim i As Integer, j As Integer
Dim rsTemp As New ADODB.Recordset
Private mstrReturn As String             '�������|�Ż����|ҽ����|���|ͣ��

Public Function ShowME(ByVal STR���� As String) As String
    mstr���� = STR����
    mstrReturn = ""
    Me.Show 1
    ShowME = mstrReturn
End Function

Private Sub cbo��ҽ��_Click()
    For i = 0 To j
        If Split(marrData(i), ";")(0) = cbo��ҽ��.ItemData(cbo��ҽ��.ListIndex) Then
            Call zlControl.CboLocate(cbo�Ż����, Split(marrData(i), ";")(1), False)
            Exit For
        End If
    Next
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    If Trim(txt����.Text) = "" Then
        MsgBox "������ҽ���Ű��س�ȷ��������ݣ�", vbInformation, gstrSysName
        txtҽ����.SetFocus
        Exit Sub
    End If
    If Me.txt����.Text <> mstr���� Then
        MsgBox "�����ӿڷ��صĲ���������ͬ�����飡", vbInformation, gstrSysName
        txtҽ����.SetFocus
        Exit Sub
    End If
    
    mstrReturn = Me.cbo��ҽ��.ItemData(Me.cbo��ҽ��.ListIndex) & "|" & Me.cbo�Ż����.Text & "|" & Me.txtҽ����.Text & "|" & txt�ʻ����.Text & "|" & Val(txt�ʻ����.Tag)
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    With Me.cbo�Ż����
        .Clear
        .AddItem "��ͨ"
        .AddItem "����"
        .AddItem "����"
        .AddItem "�����"
        .ListIndex = 0
    End With
    
    Me.cbo��ҽ��.Clear
    mstrReg = GetSetting("ZLSOFT", "����ȫ��", "����ҽ���ӿ�", "")
    marrData = Split(mstrReg, ",")
    j = UBound(marrData)
    For i = 0 To j
        mstr��� = mstr��� & "," & Split(marrData(i), ";")(0)
    Next
    mstr��� = Mid(mstr���, 2)
    
    gstrSQL = " Select ���,���� From ������� Where ��� IN ([1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҽ��", mstr���)
    
    For i = 0 To j
        rsTemp.Filter = "���=" & Split(marrData(i), ";")(0)
        Me.cbo��ҽ��.AddItem rsTemp!��� & "-" & rsTemp!����
        Me.cbo��ҽ��.ItemData(Me.cbo��ҽ��.NewIndex) = rsTemp!���
    Next
    Me.cbo��ҽ��.ListIndex = 0
End Sub

Private Sub txtҽ����_GotFocus()
    Call zlControl.TxtSelAll(txtҽ����)
End Sub

Private Sub txtҽ����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intinsure As Integer
    Dim intOrder As Integer
    Dim str���׺� As String
    Dim str���1 As String
    Dim str���2 As String
    Dim str���� As String
    On Error GoTo errHand
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    intinsure = Me.cbo��ҽ��.ItemData(Me.cbo��ҽ��.ListIndex)
    If Not CreateObject_Insure(intinsure, intOrder) Then Exit Sub
    If Not gobjInsure_Obj(intOrder).InitInsure(gcnOracle, intinsure) Then Exit Sub
    
    '�����֤
    str���׺� = "01"
    str���1 = Trim(txtҽ����.Text)
    If gobjInsure_Obj(intOrder).CallAPI(str���׺�, str���1, str���2, str����) Then
        '0֤��|1����|2�Ա�|3��������|4���֤��|5���|6���|7��ͥסַ|8�ʱ�|9��ί��
        txt����.Text = Split(str����, "|")(1)
        txt�ʻ����.Text = Val(Split(str����, "|")(5))
        txt�ʻ����.Tag = Val(Split(str����, "|")(10))
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
