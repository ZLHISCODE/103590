VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSystemParaDep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ұ��"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5565
   Icon            =   "frmSystemParaDep.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   300
      Left            =   180
      TabIndex        =   3
      Top             =   4140
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   4200
      TabIndex        =   2
      Top             =   4140
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   300
      Left            =   2910
      TabIndex        =   1
      Top             =   4140
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit BillҩƷ���� 
      Height          =   3885
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   6853
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
End
Attribute VB_Name = "frmSystemParaDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mIntItem As Integer                 '����ļ�¼��
Dim mIntSequence As Integer             '�������
Dim mCboIndex As Integer                '�б�ؼ�Index
Private Const mstrChar As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

Private Sub BillҩƷ����_cboClick(ListIndex As Long)
    mCboIndex = ListIndex
End Sub

Private Sub BillҩƷ����_cboKeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        With Me.BillҩƷ����
            .RowData(.Row) = .ItemData(mCboIndex)
        End With
    End If
End Sub

Private Sub BillҩƷ����_EditChange(curText As String)
    If Len(curText) > 1 Then
        BillҩƷ����.Text = Mid(curText, 1, 1)
    End If
    BillҩƷ����.Text = UCase(BillҩƷ����.Text)
    BillҩƷ����.SelStart = 1
    BillҩƷ����.SelLength = Len(BillҩƷ����.Text)
End Sub

Private Sub BillҩƷ����_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode = 13 Then
        Me.CmdOK.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '��������Ƿ�Ϸ�
    If IsValid = True Then Exit Sub
    Call SaveҩƷ����
    Unload Me
End Sub

Private Sub Form_Load()
    InitAll
End Sub
Public Function ShowMe(objfrm As Object, IntSequence As Integer, DepStr As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''
    '����               �ṩ���ϼ��������
    '����
    'IntSequence        ���ID
    'DepStr             ���ҺͿ��ұ���ִ�
    '����               ���ҺͿ��ұ���ִ�
    '''''''''''''''''''''''''''''''''''''''''
    mIntSequence = IntSequence
    mDepStr = DepStr
    Me.Show vbModal, objfrm
End Function
Sub InitAll()
    Call InitBill
    Call loadҩƷ����
End Sub

Sub InitBill()
    With BillҩƷ����
        .Cols = 2 '����һ��������
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "���"
        .ColWidth(0) = 2000
        .ColWidth(1) = 1000
        .ColData(0) = 3
        .ColData(1) = 4
        .PrimaryCol = 0
        .Active = True
        .TxtCheck = True
        .TextMask = mstrChar
        .PrimaryCol = 0
    End With
End Sub
Sub loadҩƷ����()
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    gstrSQL = "select distinct A.ID,A.����,A.���� " & _
                   " from  ��������˵�� b,���ű� a " & _
                   " where B.�������� in ('��ҩ��','��ҩ��','��ҩ��','�Ƽ���','��ҩ��','��ҩ��','��ҩ��') " & _
                   " and  b.����ID=a.ID  order by ����"
    On Error GoTo errH
    zldatabase.OpenRecordset rsTmp, gstrSQL, Me.Caption
    i = 0
    With BillҩƷ����
        Do Until rsTmp.EOF
            .AddItem rsTmp("����") & "-" & rsTmp("����")
            .ItemData(i) = rsTmp("id")
            rsTmp.MoveNext
            i = i + 1
        Loop

        gstrSQL = "select A.���, B.id, B.����, B.���� from ���Һ���� a ,���ű� b where a.����id = b.id and ��Ŀ��� = [1] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mIntSequence)
        
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, 0) = rsTmp("����") & "-" & rsTmp("����")
            .TextMatrix(.Rows - 1, 1) = rsTmp("���")
            .RowData(.Rows - 1) = rsTmp("id")
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If .Rows > 2 Then
            .Rows = .Rows - 1
        End If
    End With
    Exit Sub
errH:
    If ERRCENTER() = 1 Then Resume
End Sub
Private Function IsValid() As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����           �����ұ�������Ƿ���ȷ
    '����           =True��ʾ������ =False��ʾ����ͨ��ȥʱ
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strDept As String
    Dim strNumber As String
    With Me.BillҩƷ����
        For i = 1 To .Rows - 1
            If Len(Trim(.TextMatrix(i, 0))) > 0 And Len(Trim(.TextMatrix(i, 1))) > 0 Then
                If InStr(1, strDept & ",", "," & .TextMatrix(i, 0) & ",") > 0 Then
                    MsgBox "��" & i & "�г��ֿ����ظ�!", vbInformation, gstrSysName
                    .Row = i
                    .Col = 0
                    .TxtSetFocus
                    IsValid = True
                    Exit Function
                End If
                strDept = strDept & "," & .TextMatrix(i, 0)
                
                If InStr(1, strDept & ",", "," & .TextMatrix(i, 1) & ",") > 0 Then
                    If InStr(1, strDept & ",", "," & .TextMatrix(i, 1) & ",") > 0 Then
                        If MsgBox("��" & i & "�г��ֿ��ұ���ظ�!", vbYesNo + vbInformation, gstrSysName) = vbNo Then
                            .Row = i
                            .Col = 1
                            .TxtSetFocus
                            IsValid = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    End With
End Function
Sub SaveҩƷ����()
    
    On Error GoTo errH
    '��ɾ����ǰ���ٱ���
    gstrSQL = "ZL_���Һ����_DELETE(" & mIntSequence & ")"
    zldatabase.ExecuteProcedure gstrSQL, Me.Caption
    With Me.BillҩƷ����
        For i = 1 To .Rows - 1
            If Len(Trim(.TextMatrix(i, 0))) > 0 And Len(Trim(.TextMatrix(i, 1))) > 0 Then
                gstrSQL = "ZL_���Һ����_INSERT(" & mIntSequence & "," & .RowData(i) & ",'" & .TextMatrix(i, 1) & "')"
                zldatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With
    Exit Sub
errH:
    If ERRCENTER() = 1 Then Resume
End Sub

