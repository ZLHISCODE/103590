VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm����������ݺ˶�_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������ݺ˶�_����"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10485
   Icon            =   "frm����������ݺ˶�_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk���� 
      Caption         =   "����ѯ����������"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      Top             =   5400
      Value           =   1  'Checked
      Width           =   2835
   End
   Begin VB.CheckBox chk��ʷ���� 
      Caption         =   "��ѯ��������µ���ʷ����"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1470
      TabIndex        =   2
      Top             =   5700
      Width           =   2835
   End
   Begin VB.CommandButton cmd��ѯ 
      Caption         =   "��ѯ(&R)"
      Height          =   350
      Left            =   210
      TabIndex        =   3
      Top             =   5490
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&O)"
      Height          =   350
      Left            =   9180
      TabIndex        =   4
      Top             =   5310
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   9180
      TabIndex        =   5
      Top             =   5730
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   13275520
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm����������ݺ˶�_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
Private mint���� As Integer

Public Sub ShowME(ByVal int���� As Integer, ByVal intinsure As Integer)
    mintInsure = intinsure
    mint���� = int����
    Me.Show 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmd��ѯ_Click()
    Call LoadData
End Sub

Private Sub cmd����_Click()
    Dim blnOK As Boolean
    Dim intRow As Integer, intRows As Integer
    
    If MsgBox("��ȷ��Ҫ����ѡ������ݷ���������Ͻ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    intRows = mshDetail.Rows - 1
    For intRow = 1 To intRows
        If Val(mshDetail.TextMatrix(intRow, 0)) <> 0 And mshDetail.TextMatrix(intRow, mshDetail.Cols - 1) = "��" Then
            If mint���� = 1 Then
                If blnOK = False Then blnOK = �������(intRow)
            Else
                'Call סԺ����(intRow)
            End If
        End If
    Next
    
    If blnOK Then Call LoadData
End Sub

Private Sub Form_Load()
    Call LoadData
End Sub

Private Sub LoadData()
    Dim strStart As String
    Dim strStation As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '��ȡ��ǰ����Ա�ĵ�����쳣��¼
    
    mshDetail.Rows = 2
    mshDetail.Cols = 2
    mshDetail.Clear
    
    strStation = AnalyseComputer
    strStation = " And ����վ='" & strStation & "'"
    If chk����.Value = 0 Then strStation = ""
    If chk��ʷ����.Value = 0 Then
        strStart = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    Else
        strStart = Format(DateAdd("m", -2, zlDatabase.Currentdate), "yyyy-MM-dd 00:00:00")
    End If
    
    gstrSQL = "" & _
              "        (Select ����ID From ������־_���� " & _
              "         Where Nvl(�ѳ���,0)=0 And ����=" & mint���� & _
                        strStation & " And ����ʱ�� >= to_date('" & strStart & "','yyyy-MM-dd hh24:mi:ss')" & _
              "         MINUS" & _
              "         Select ��¼ID From ���ս����¼" & _
              "         Where ����=" & mint���� & _
                        strStation & " And ����ʱ�� >= to_date('" & strStart & "','yyyy-MM-dd hh24:mi:ss')) B"
    gstrSQL = " Select A.����ID,A.����ID,C.ҽ����,D.����,A.����˳���,A.������," & _
              "        A.֧�����,DECODE(A.֧�����,'11','��ͨ','����') AS ֧���������,A.����Ա,A.����վ,A.����ʱ�� AS ����ʱ��,'��' AS ��־" & _
              " From ������־_���� A," & gstrSQL & ",�����ʻ� C,������Ϣ D" & _
              " Where A.����ID=B.����ID And A.����ID=C.����ID And A.����ID=D.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡδ������쳣����")
    If rsTemp.RecordCount = 0 Then
        MsgBox "δ�����쳣���ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set mshDetail.DataSource = rsTemp
    mshDetail.ColWidth(10) = 2000
    mshDetail.ColWidth(11) = 600
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mshDetail_DblClick()
    Call mshDetail_KeyDown(vbKeySpace, 0)
End Sub

Private Sub mshDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If Val(mshDetail.TextMatrix(mshDetail.Row, 0)) = 0 Then Exit Sub
        If mshDetail.TextMatrix(mshDetail.Row, mshDetail.Cols - 1) = "" Then
            mshDetail.TextMatrix(mshDetail.Row, mshDetail.Cols - 1) = "��"
        Else
            mshDetail.TextMatrix(mshDetail.Row, mshDetail.Cols - 1) = ""
        End If
    End If
End Sub

Private Function �������(ByVal intRow As Integer) As Boolean
    Dim bln���� As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '����������ݳ������Լ��������ݵĸ���
    
    gstrSQL = " Select 1 From ���ս����¼ Where ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ����HIS�Ľ�������", CLng(Val(mshDetail.TextMatrix(intRow, 0))))
    If rsTemp.RecordCount <> 0 Then
        MsgBox "�ü�¼�����������¼,���������ϣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", mshDetail.TextMatrix(intRow, 4))    ' ����˳���
    Call InsertChild(mdomInput.documentElement, "BALANCEID", mshDetail.TextMatrix(intRow, 5))    ' ������
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", mshDetail.TextMatrix(intRow, 6))   ' ֧�����
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)    ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))   ' ��������
    
    '���ýӿ�
    bln���� = IS����(Val(mshDetail.TextMatrix(intRow, 1)))
    If CommServer("RETBALANCE", IIf(bln����, 1, 0)) = False Then Exit Function
    
    '����
    gcnOracle.Execute "ZL_������־_����_����(" & Val(mshDetail.TextMatrix(intRow, 0)) & ")", , adCmdStoredProc
    ������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


