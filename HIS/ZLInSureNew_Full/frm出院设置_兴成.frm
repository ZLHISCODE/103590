VERSION 5.00
Begin VB.Form frm��Ժ����_�˳� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ�����˳�Ժ����"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frm��Ժ����_�˳�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2550
      Width           =   3660
   End
   Begin VB.TextBox TxtEdit 
      Height          =   300
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2160
      Width           =   3660
   End
   Begin VB.ComboBox cbo��Ժ��� 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   945
      Width           =   3660
   End
   Begin VB.ComboBox cboסԺ��� 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1365
      Width           =   3660
   End
   Begin VB.ComboBox cbo��Ժ��� 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1770
      Width           =   3660
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4095
      TabIndex        =   14
      Top             =   3195
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2910
      TabIndex        =   13
      Top             =   3195
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   720
      Width           =   7665
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -1935
      TabIndex        =   11
      Top             =   3000
      Width           =   7665
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "���ҽԺ����"
      Height          =   180
      Index           =   3
      Left            =   255
      TabIndex        =   9
      Top             =   2610
      Width           =   1080
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "���ҽԺ����"
      Height          =   180
      Index           =   2
      Left            =   255
      TabIndex        =   7
      Top             =   2220
      Width           =   1080
   End
   Begin VB.Label lblInfor 
      Caption         =   "��Ժ���"
      Height          =   210
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   990
      Width           =   735
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "סԺ���"
      Height          =   180
      Index           =   0
      Left            =   615
      TabIndex        =   3
      Top             =   1425
      Width           =   720
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "��Ժ���"
      Height          =   180
      Index           =   1
      Left            =   615
      TabIndex        =   5
      Top             =   1830
      Width           =   720
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frm��Ժ����_�˳�.frx":0E42
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "    ���ó�Ժ���˵���Ժ���סԺ��𼰳�Ժ���,����Ժ���Ϊת���ʱ�������������ҽԺ���Ƽ�����."
      Height          =   390
      Left            =   885
      TabIndex        =   0
      Top             =   240
      Width           =   4500
   End
End
Attribute VB_Name = "frm��Ժ����_�˳�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mblnOK As Boolean



Private Function IsValid() As Boolean
    '�жϴ���
    IsValid = False
    If zlCommFun.StrIsValid(txtEdit.Text, txtEdit.MaxLength) = False Then
        zlControl.TxtSelAll txtEdit
        txtEdit.SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Sub cbo��Ժ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub



Private Sub cbo����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub



Private Sub cbo��Ժ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboסԺ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If IsValid = False Then Exit Sub
    '���̲���:
    '       סԺ���_IN,��Ժ���_IN,���ҽԺ����_IN,���ҽԺ����_IN
    
    gstrSQL = "ZL_ҽ�����˳�Ժ�Ǽ�_UPDATE("
    gstrSQL = gstrSQL & mlng����ID & ","
    gstrSQL = gstrSQL & "'" & cboסԺ���.ItemData(cboסԺ���.ListIndex) & "',"
    gstrSQL = gstrSQL & "'" & cbo��Ժ���.ItemData(cbo��Ժ���.ListIndex) & "',"
    If cbo��Ժ���.ItemData(cbo��Ժ���.ListIndex) = 3 Then
        gstrSQL = gstrSQL & "'" & txtEdit.Text & "',"
        gstrSQL = gstrSQL & "'" & cbo��Ժ���.ItemData(cbo��Ժ���.ListIndex) & "')"
    Else
        gstrSQL = gstrSQL & "NULL,"
        gstrSQL = gstrSQL & "NULL)"
    End If
    ExecuteProcedure_�˳� Me.Caption
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�˳ɺ˹�ҵ & ",'��Ա���','" & cbo��Ժ���.ItemData(cbo��Ժ���.ListIndex) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ���")
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()

    Call InitData
End Sub

Private Sub InitData()
    Dim rsTemp As New ADODB.Recordset
    Dim str��Ժ��� As String
    Dim i As Long
    Me.cbo��Ժ���.Clear
    Me.cboסԺ���.Clear
    Me.cbo��Ժ���.Clear
    With Me.cbo��Ժ���
        .AddItem "1-������Ժ"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
        .ListIndex = .NewIndex
        .AddItem "2-����ת��"
        .ItemData(.NewIndex) = 2
        .AddItem "3-����ת��"
        .ItemData(.NewIndex) = 3
        .AddItem "4-�����Բ����ص�һ��סԺ"
        .ItemData(.NewIndex) = 4
    End With
    With Me.cboסԺ���
        .AddItem "0-����סԺ"
        .ItemData(.NewIndex) = 0
        .ListIndex = .NewIndex
        .AddItem "1-��������"
        .ItemData(.NewIndex) = 1
    End With
    With Me.cbo��Ժ���
        .AddItem "1-������Ժ"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "2-ת������"
        .ItemData(.NewIndex) = 2
        .AddItem "3-ת������"
        .ItemData(.NewIndex) = 3
    End With
    
    '��Ժ���ȷ��
    gstrSQL = "Select ����id,��Ա��� From �����ʻ� where ����id=" & mlng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.EOF = False Then
        For i = 0 To Me.cbo��Ժ���.ListCount - 1
            If Me.cbo��Ժ���.ItemData(i) = Val(Nvl(rsTemp!��Ա���)) Then
                Me.cbo��Ժ���.ListIndex = i: Exit For
            End If
            
        Next
    End If
        
    '��Ժ���ȷ��
    gstrSQL = "Select AF17,AF18,AF19 as סԺ���,AF20 as ��Ժ��� From ҽ�����˸�����Ϣ where ����id=" & mlng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With Me.cboסԺ���
        For i = 0 To .ListCount - 1
            If .ItemData(i) = Val(Nvl(rsTemp!סԺ���)) Then
                .ListIndex = i: Exit For
            End If
        Next
    End With
    With Me.cbo��Ժ���
        For i = 0 To .ListCount - 1
            If .ItemData(i) = Val(Nvl(rsTemp!��Ժ���)) Then
                .ListIndex = i: Exit For
            End If
        Next
    End With
End Sub


Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m�ı�ʽ
End Sub
Private Sub cbo��Ժ���_Click()
     
    If cbo��Ժ���.ItemData(cbo��Ժ���.ListIndex) <> 3 Then
        Me.cbo����.Clear
        With Me.cbo����
            .AddItem "11: ת����һ��ҽԺ"
            .ItemData(.NewIndex) = 11
            .ListIndex = .NewIndex
            .AddItem "12: ת���ڶ���ҽԺ"
            .ItemData(.NewIndex) = 12
            .AddItem "13: ת��������ҽԺ"
            .ItemData(.NewIndex) = 13
            .AddItem "14: ת����һ��ҽԺ"
            .ItemData(.NewIndex) = 14
            .AddItem "15: ת�������ҽԺ"
            .ItemData(.NewIndex) = 15
            .AddItem "16: ת��������ҽԺ"
            .ItemData(.NewIndex) = 16
            .AddItem "17: תʡ��һ��ҽԺ"
            .ItemData(.NewIndex) = 17
            .AddItem "18: תʡ�����ҽԺ"
            .ItemData(.NewIndex) = 18
            .AddItem "19: תʡ������ҽԺ"
            .ItemData(.NewIndex) = 19
        End With
        Exit Sub
    Else
        Me.cbo����.Clear
'        With Me.cbo����
'            .AddItem "01: һ��ҽԺ"
'            .ItemData(.NewIndex) = 1
'            .AddItem "02: ����ҽԺ"
'            .ItemData(.NewIndex) = 2
'            .AddItem "03: ����ҽԺ"
'            .ItemData(.NewIndex) = 3
'        End With
    End If
End Sub
Public Function ShowCard(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    Me.Show vbModal
    ShowCard = mblnOK
End Function
