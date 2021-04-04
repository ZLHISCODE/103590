VERSION 5.00
Begin VB.Form frmPriorityCause 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ԭ��"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmPriorityCause.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdDel 
      Height          =   300
      Left            =   4250
      Picture         =   "frmPriorityCause.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "ɾ����ǰѡ��ĳ���ԭ��"
      Top             =   2750
      Width           =   300
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   300
      Left            =   3950
      Picture         =   "frmPriorityCause.frx":699C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "����ǰԭ����Ϊ����ԭ��"
      Top             =   2750
      Width           =   300
   End
   Begin VB.ComboBox cboPriorityCause 
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   3855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3240
      Width           =   1075
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3240
      Width           =   1075
   End
   Begin VB.ListBox lstJQueueList 
      Height          =   2040
      Left            =   75
      TabIndex        =   0
      Top             =   315
      Width           =   4480
   End
   Begin VB.Label lblCause 
      Caption         =   "����ԭ��"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   2430
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Caption         =   "����      ��������        ״̬"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   3405
   End
End
Attribute VB_Name = "frmPriorityCause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrQueueName As String        '��������
Private mstrCruuentWorkID As String      'ҵ��ID
Private mfrmParent As Form             '������
Private mstrTempQueueName As String    '��ǰѡ�����ݵĶ�������
Private mstrSelectedName As String     '��ǰѡ��Ļ�������
Private mintCurNextIndex As String     '���������ݵ���һ�����ݵ�Index
Private mstrMaxCode As String          '��ȡ����ԭ���������
Private mstrArrQueueNum() As String

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const CB_LIMITTEXT = &H141

Private Enum mCol
    �������� = 0: Id: ����ID: �Ŷӱ��: �ŶӺ���:  �Ŷ����: ��������: ����: �������: ���������: ����ID: ����: ҽ������: �Ŷ�״̬: �Ŷ�ʱ��: ����ҽ��: ҵ������: ҵ��ID: ����ʱ��
End Enum




Private Sub cmdDel_Click()
'����: ɾ������ԭ��
    Dim strSql As String
    Dim strDelCode As String
    Dim strMaxCode As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    If cboPriorityCause.ListIndex = -1 Then
        cboPriorityCause.Text = ""
        cboPriorityCause.SetFocus
        Exit Sub
    End If
    
    strDelCode = cboPriorityCause.ItemData(cboPriorityCause.ListIndex)
    strSql = "zl_�Ŷ�����ԭ��_delete('" & Format(strDelCode, "00000") & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "����ԭ��")
    
    Call cboPriorityCause.RemoveItem(cboPriorityCause.ListIndex)
    
    strMaxCode = "0"
    
    If CLng(strDelCode) = CLng(mstrMaxCode) Then '��ȡɾ������֮������code
        If cboPriorityCause.ListCount > 0 Then
            For i = 0 To cboPriorityCause.ListCount - 1
                If cboPriorityCause.ItemData(i) > strMaxCode Then
                    strMaxCode = cboPriorityCause.ItemData(i)
                End If
            Next
        End If
        
        mstrMaxCode = strMaxCode
    End If
    
    If cboPriorityCause.ListCount <= 0 Then mstrMaxCode = ""
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdAdd_Click()
'����: ��������ԭ��
    Dim i As Integer
    Dim strSql As String
    
    On Error GoTo errHandle
    
    If Trim(cboPriorityCause.Text) = "" Then
        MsgBox "��������Ҫ���������ݣ�", vbOKOnly Or vbInformation, Me.Caption
        cboPriorityCause.SetFocus
        Exit Sub
    End If
    
    For i = 0 To cboPriorityCause.ListCount - 1
        If UCase(Trim(cboPriorityCause.List(i))) = UCase(Trim(cboPriorityCause.Text)) Then
            MsgBox "�������Ѿ�������ԭ���У�", vbOKOnly Or vbInformation, Me.Caption
            cboPriorityCause.SetFocus
            Exit Sub
        End If
    Next
    
    strSql = "zl_�Ŷ�����ԭ��_insert('" & Trim(cboPriorityCause.Text) & "','" & zlCommFun.zlGetSymbol(cboPriorityCause.Text) & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "����ԭ��")
    
    Call cboPriorityCause.AddItem(Trim(cboPriorityCause.Text))
    cboPriorityCause.ItemData(cboPriorityCause.ListCount - 1) = CLng(IIf(mstrMaxCode = "", 0, mstrMaxCode) + 1)
    mstrMaxCode = IIf(mstrMaxCode = "", 0, mstrMaxCode) + 1
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub LoadPriorityCause()
'����: ��������ԭ��
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    cboPriorityCause.Clear
    
    strSql = "select ����,����,ʹ��Ƶ�� from �Ŷ�����ԭ�� order by ʹ��Ƶ�� desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "����ԭ��")
    
    If rsTemp.RecordCount <= 0 Then Exit Sub

    rsTemp.MoveFirst
    mstrMaxCode = Nvl(rsTemp!����)
    
    Do While Not rsTemp.EOF
        cboPriorityCause.AddItem Nvl(rsTemp!����)
        cboPriorityCause.ItemData(cboPriorityCause.ListCount - 1) = Nvl(rsTemp!����)
        If CLng(Nvl(rsTemp!����)) > CLng(mstrMaxCode) Then mstrMaxCode = CLng(Nvl(rsTemp!����))
        rsTemp.MoveNext
    Loop
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errHandle
    
    '���ش���
    Me.Hide

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub ShowPriorityCause(frmParent As Form, ByVal strCurrentCaption As String, ByVal strCurrentWorkID As String, _
                            ByVal strTempQueueName As String, ByVal strSelectedName As String)
    
    '���洫�����
    mstrQueueName = strCurrentCaption
    mstrCruuentWorkID = strCurrentWorkID
    mstrTempQueueName = strTempQueueName
    mstrSelectedName = strSelectedName
    
    Set mfrmParent = frmParent
    
    '����List�ؼ�����
    Call LoadListData
    
    '�򿪴���
    Me.Show 1, frmParent
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    Dim i As Integer
    Dim strSql As String
    Dim intJQueueID As Long        '����ID
    Dim intNeedJQueueID As Long    '�����ID
    
    '�ж��Ƿ�ѡ��������
    If lstJQueueList.SelCount < 1 Then
        MsgBox "��ѡ�񱻲�ӵĲ��ˡ�", vbOKOnly Or vbInformation, Me.Caption
        Exit Sub
    End If
    
    '�ж��Ƿ�ѡ������ͬ������
    If mstrArrQueueNum(lstJQueueList.ListIndex) = mstrSelectedName Then
        MsgBox "���ܲ��Լ��Ķӡ�", vbOKOnly Or vbInformation, Me.Caption
        Exit Sub
    End If
    
    '�ж��Ƿ�ѡ���˵�ǰ���ݵ���һ�����ݣ���Ϊ��ӵ���һ��������ʵ����λ��û�䡣
    If lstJQueueList.ListIndex = mintCurNextIndex Then
        MsgBox "��ӵ��ò���ǰ������λ�ò��䣬������ѡ��", vbOKOnly Or vbInformation, Me.Caption
        Exit Sub
    End If
    
    '�ж�¼��ԭ���Ƿ�Ϊ��
    If Trim(cboPriorityCause.Text) = "" Then
         MsgBox "����ԭ��Ϊ�ա�", vbOKOnly Or vbInformation, Me.Caption
         Exit Sub
    End If
    
    '�õ������ �� ���ӵ� �Ŷ�ID
    intJQueueID = Val(Mid(mstrArrQueueNum(lstJQueueList.ListIndex), InStr(mstrArrQueueNum(lstJQueueList.ListIndex), ",") + 1, 100))
    intNeedJQueueID = Val(Mid(mstrSelectedName, InStr(mstrSelectedName, ",") + 1, 100))
    
    'ִ�в�� �����޸��Լ�ԭ��д��
    strSql = "ZL_�ŶӽкŶ���_����('" & mstrQueueName & "'," & mstrCruuentWorkID & ",'" & Trim(cboPriorityCause.Text) & "'," & intNeedJQueueID & "," & intJQueueID & ")"
    zlDatabase.ExecuteProcedure strSql, "����ԭ��"
    
    For i = 0 To cboPriorityCause.ListCount - 1
        If UCase(Trim(cboPriorityCause.List(i))) = UCase(Trim(cboPriorityCause.Text)) Then
            '����ʹ��Ƶ��
            strSql = "zl_�Ŷ�����ԭ��_Update('" & Format(cboPriorityCause.ItemData(i), "00000") & "')"
            Call zlDatabase.ExecuteProcedure(strSql, "����ԭ��")
            
            Me.Hide
            Exit Sub
        End If
    Next
    '������ԭ��д�����ݿ�
    strSql = "zl_�Ŷ�����ԭ��_insert('" & Trim(cboPriorityCause.Text) & "','" & zlCommFun.zlGetSymbol(cboPriorityCause.Text) & "',1)"
    Call zlDatabase.ExecuteProcedure(strSql, "����ԭ��")
    
    Call cboPriorityCause.AddItem(Trim(cboPriorityCause.Text))
    cboPriorityCause.ItemData(cboPriorityCause.ListCount - 1) = CLng(IIf(mstrMaxCode = "", 0, mstrMaxCode) + 1)
    mstrMaxCode = IIf(mstrMaxCode = "", 0, mstrMaxCode) + 1
    '��ɺ����ش���
    Me.Hide
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub LoadListData()
'����ListBox����
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With mfrmParent.rptQueueList
        ReDim mstrArrQueueNum(.Rows.Count)
        
        For i = 0 To .Rows.Count - 1
            If .Rows(i).GroupRow <> True Then
                'ͨ���Ա��ж�ֻ���ؿؼ��е�ǰѡ�ж��е�����
                If .Rows(i).Record(mCol.��������).value = mstrTempQueueName Then
                
                    '����Ӧ�����ݴ�������
                    mstrArrQueueNum(j) = .Rows(i).Record(mCol.�ŶӺ���).value & .Rows(i).Record(mCol.��������).value & "," & .Rows(i).Record(mCol.Id).value
                
                    '��ListBox ��ֵ
                    lstJQueueList.List(j) = "  " & .Rows(i).Record(mCol.�ŶӺ���).value & "��   " & .Rows(i).Record(mCol.��������).value & IIf(mstrArrQueueNum(j) = mstrSelectedName, "   ������", "")
                    
                    If mstrArrQueueNum(j) = mstrSelectedName Then mintCurNextIndex = j + 1
                    
                    j = j + 1
                End If
            End If
        Next i
    End With
    
    'Ĭ��ѡ�е�һ��
    If lstJQueueList.ListCount > 0 Then lstJQueueList.ListIndex = 0
    
End Sub

Private Sub Form_Load()
    '��������ԭ��
    Call LoadPriorityCause
    '����ԭ�����󳤶�����Ϊ64λ
    SendMessage cboPriorityCause.hwnd, CB_LIMITTEXT, 64, 0&
End Sub
