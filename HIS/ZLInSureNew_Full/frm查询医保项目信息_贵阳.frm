VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm��ѯҽ����Ŀ��Ϣ_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ѯҽ����Ŀ��Ϣ"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frm��ѯҽ����Ŀ��Ϣ_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd��λ 
      Caption         =   "��λ(&L)"
      Height          =   350
      Left            =   2730
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "�ڲ�ѯ�Ľ�����϶�λĳ����Ŀ"
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmd��ѯ 
      Caption         =   "��ѯ(&R)"
      Height          =   350
      Left            =   7140
      TabIndex        =   5
      ToolTipText     =   "��������ȡָ��֧�������������Ŀ��ĳ����Ŀ�ı�����Ϣ"
      Top             =   120
      Width           =   1100
   End
   Begin VB.ComboBox cbo֧����� 
      Height          =   300
      Left            =   5070
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   150
      Width           =   1905
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1020
      TabIndex        =   1
      Top             =   150
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   4695
      Left            =   90
      TabIndex        =   6
      Top             =   600
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorSel    =   12285290
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.Label lbl֧����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "֧�����(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3990
      TabIndex        =   3
      Top             =   210
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   630
   End
End
Attribute VB_Name = "frm��ѯҽ����Ŀ��Ϣ_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Enum Columns
    ����
    ����
    ����޼�
    �Ը�����
    ������־
    ���˱�־
    ���ⱨ����־    '01-��ͨ��Ŀ��02-������ȫ�Ը���Ŀ(����ҽ�Ʋ�����Χ)��03-ҽ���չ���Ա������Ŀ��04-����ֱ��֧����Ŀ
    ���ɽ������    '01-��ͨ��Ŀ��02-���ɽ�����շ�Χ����Ŀ��03-ҽ���չ���Ա������Ŀ�� 04-����ֱ��֧����Ŀ��05-���ɽ�������Է���Ŀ
    ����
End Enum

Private Sub cmd��ѯ_Click()
    Dim arrData
    Dim str��� As String
    Dim rsData As New ADODB.Recordset
    '��ѯָ����Ŀ��������Ŀ����֧��������Ϣ
    On Error GoTo errHand
    
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Fields.Append "CLASSCODE", adVarChar, 6    '�������
        .Fields.Append "CODE", adVarChar, 20        '����
        .Fields.Append "NAME", adVarChar, 300       '����
        .Fields.Append "PY", adVarChar, 150         'ƴ������
        .Fields.Append "MEMO", adVarChar, 500       '��ע
        .Open
    End With
    Call InitBill(True)
    
    str��� = cbo֧�����.ItemData(cbo֧�����.ListIndex)
    If Not ҽ����Ŀ_����(rsData, str���) Then Exit Sub
    
    'װ������
    mshDetail.Redraw = False
    With rsData
        Do While Not .EOF
            arrData = Split(!Memo, "|")
            mshDetail.TextMatrix(.AbsolutePosition, ����) = !CODE
            mshDetail.TextMatrix(.AbsolutePosition, ����) = !Name
            mshDetail.TextMatrix(.AbsolutePosition, ����޼�) = arrData(0)
            mshDetail.TextMatrix(.AbsolutePosition, �Ը�����) = arrData(1)
            mshDetail.TextMatrix(.AbsolutePosition, ������־) = arrData(2)
            mshDetail.TextMatrix(.AbsolutePosition, ���˱�־) = arrData(3)
            mshDetail.TextMatrix(.AbsolutePosition, ���ⱨ����־) = Switch(arrData(4) = "01", "��ͨ��Ŀ", arrData(4) = "02", "������ȫ�Ը���Ŀ(����ҽ�Ʋ�����Χ)", arrData(4) = "03", "ҽ���չ���Ա������Ŀ", arrData(4) = "04", "����ֱ��֧����Ŀ")
            mshDetail.TextMatrix(.AbsolutePosition, ���ɽ������) = Switch(arrData(5) = "01", "��ͨ��Ŀ", arrData(5) = "02", "���ɽ�����շ�Χ����Ŀ", arrData(5) = "03", "ҽ���չ���Ա������Ŀ", arrData(5) = "04", "����ֱ��֧����Ŀ", arrData(5) = "05", "���ɽ�������Է���Ŀ")
            .MoveNext
        Loop
    End With
    mshDetail.Redraw = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mshDetail.Redraw = True
End Sub

Private Sub cmd��λ_Click()
    Dim intDO As Integer, intMAX As Integer
    
    If txt����.Text = "" Then Exit Sub
    intMAX = mshDetail.Rows - 1
    For intDO = 1 To intMAX
        If txt����.Text = (mshDetail.TextMatrix(intDO, ����)) Then
            With mshDetail
                .TopRow = intDO
                .Row = intDO: .RowSel = intDO
                .COL = 0: .ColSel = .Cols - 1
            End With
            Exit Sub
        End If
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With cbo֧�����
        .Clear
        .AddItem "��ͨסԺ"
        .ItemData(.NewIndex) = 12
        .AddItem "����סԺ"
        .ItemData(.NewIndex) = 22
        .AddItem "����סԺ"
        .ItemData(.NewIndex) = 42
        .AddItem "����סԺ"
        .ItemData(.NewIndex) = 32
        .AddItem "��ͨ����"
        .ItemData(.NewIndex) = 11
        .AddItem "��������"
        .ItemData(.NewIndex) = 21
        .AddItem "��������"
        .ItemData(.NewIndex) = 41
        .AddItem "��������"
        .ItemData(.NewIndex) = 31
        .ListIndex = 0
    End With
End Sub

Private Sub InitBill(Optional ByVal blnInit As Boolean)
    With mshDetail
        If blnInit Then
            .Clear
            .Rows = 2: .Cols = ����
            .TextMatrix(0, ����) = "����"
            .TextMatrix(0, ����) = "����"
            .TextMatrix(0, ����޼�) = "����޼�"
            .TextMatrix(0, �Ը�����) = "�Ը�����"
            .TextMatrix(0, ������־) = "������־"
            .TextMatrix(0, ���˱�־) = "���˱�־"
            .TextMatrix(0, ���ⱨ����־) = "���ⱨ����־"
            .TextMatrix(0, ���ɽ������) = "���ɽ������"
        End If
        .ColWidth(����) = 1000
        .ColWidth(����) = 1500
        .ColWidth(����޼�) = 1000
        .ColWidth(�Ը�����) = 1000
        .ColWidth(������־) = 500
        .ColWidth(���˱�־) = 500
        .ColWidth(���ⱨ����־) = 1500
        .ColWidth(���ɽ������) = 1500
    End With
End Sub
