VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSendBillPrint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ƶ��ݴ�ӡ"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmSendBillPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Ԥ��(&V)"
      Height          =   350
      Left            =   3540
      TabIndex        =   4
      ToolTipText     =   "Ԥ����ǰ����"
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   2445
      TabIndex        =   3
      ToolTipText     =   "���õ�ǰ����"
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "��ӡ����ѡ��ĵ���"
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   350
      Left            =   5895
      TabIndex        =   2
      Top             =   4665
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwBill 
      Height          =   3795
      Left            =   75
      TabIndex        =   0
      Top             =   750
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   6694
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���ݺ�"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���Ƶ���"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "˵��"
         Object.Width           =   6350
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "frmSendBillPrint.frx":058A
      Top             =   165
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSendBillPrint.frx":0E54
      Height          =   525
      Left            =   930
      TabIndex        =   5
      Top             =   120
      Width           =   6090
   End
End
Attribute VB_Name = "frmSendBillPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mstrBillPrint As String '��ǰ��ӡ�����Ƶ��ݣ������š�NO����¼����

Private mlng���ͺ� As Long
Private mint���� As Integer
Private mstrǰ��IDs As String
Private mint��ӡ��ʽ As Integer
Private mblnItem As Boolean
Private mint���뵥��ӡģʽ As Integer  '1-����ʱ��ӡ��2-�¿�ʱ��ӡ

Public Sub ShowMe(ByVal lng���ͺ� As Long, ByVal int���� As Integer, frmParent As Object, Optional ByVal strǰ��IDs As String)
'������lng���ͺ�=���η��͵ķ��ͺ�
'      int����=1-����,2-סԺ(���ݳ���,���ǵ��ó���)
'      strǰ��IDsҽ��վ���ڵ�ǰ����ִ�е�����ҽ��
    mlng���ͺ� = lng���ͺ�
    mint���� = int����
    mstrǰ��IDs = strǰ��IDs
    
    On Error Resume Next
    Me.Show 1, frmParent
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
'���ܣ������Ƶ��ݶ�Ӧ���Զ��屨�����Ԥ��
    If lvwBill.SelectedItem Is Nothing Then Exit Sub
    With lvwBill.SelectedItem
        '��Ѫҽ����ӡ���뵥������غ������м��
        If InStr(1, ",ZL1_INSIDE_1254_17_1,ZL1_INSIDE_1254_17_2,", "," & .Tag & ",") <> 0 Then
            If BloodApplyPrintCheck(Val(.ListSubItems(2).Tag), mint����, IIF(.Tag = "ZL1_INSIDE_1254_17_1", 1, 2), 0) = False Then Exit Sub
        End If
        mstrBillPrint = .Tag & "," & .Text & "," & .ListSubItems(1).Tag
        Call mobjReport.ReportOpen(gcnOracle, glngSys, .Tag, Me, "NO=" & .Text, "����=" & Val(.ListSubItems(1).Tag), "ҽ��ID=" & Val(.ListSubItems(2).Tag), 1)
        mstrBillPrint = ""
    End With
End Sub

Private Sub cmdPrint_Click()
'���ܣ���ѡ������Ƶ��ݽ��д�ӡ
    Dim i As Long, j As Long
    Dim blnALL As Boolean
    
    If lvwBill.SelectedItem Is Nothing Then Exit Sub
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Checked Then j = j + 1
    Next
    If j = 0 Then
        MsgBox "����ѡ����Ҫ��ӡ�����Ƶ��ݡ�", vbInformation, gstrSysName
        Exit Sub
    ElseIf j = lvwBill.ListItems.Count Then
        blnALL = True
    End If
    
    '��Ѫҽ����ӡ���뵥������غ������м��
    For i = 1 To lvwBill.ListItems.Count
        With lvwBill.ListItems(i)
            If .Checked Then
                If InStr(1, ",ZL1_INSIDE_1254_17_1,ZL1_INSIDE_1254_17_2,", "," & .Tag & ",") <> 0 Then
                    If BloodApplyPrintCheck(Val(.ListSubItems(2).Tag), mint����, IIF(.Tag = "ZL1_INSIDE_1254_17_1", 1, 2), 1) = False Then
                        .Checked = False
                        If blnALL = True Then blnALL = False
                    End If
                End If
            End If
        End With
    Next
    
    cmdPrint.Enabled = False
    Screen.MousePointer = 11
    For i = 1 To lvwBill.ListItems.Count
        With lvwBill.ListItems(i)
            If .Checked Then
                .Selected = True: .EnsureVisible: Me.Refresh
                
                mstrBillPrint = .Tag & "," & .Text & "," & .ListSubItems(1).Tag
                Call mobjReport.ReportOpen(gcnOracle, glngSys, .Tag, Me, "NO=" & .Text, "����=" & Val(.ListSubItems(1).Tag), "ҽ��ID=" & Val(.ListSubItems(2).Tag), "PrintEmpty=0", 2)
                mstrBillPrint = ""
                
                '�Ѵ�ӡ������ɫ��ʶ
                .Checked = False: .ForeColor = vbBlue
                For j = 1 To .ListSubItems.Count
                    .ListSubItems(j).ForeColor = vbBlue
                Next
            End If
        End With
    Next
    Screen.MousePointer = 0
    cmdPrint.Enabled = True
    
    '�ֹ���ӡʱ��ȫ����ӡ��Ϻ��Զ��˳�
    If mint��ӡ��ʽ = 1 And blnALL Then
        Unload Me: Exit Sub
    ElseIf Visible Then
        cmdCancel.SetFocus
    End If
End Sub

Private Sub cmdSetup_Click()
'���ܣ������Ƶ��ݶ�Ӧ���Զ��屨���������
    If lvwBill.SelectedItem Is Nothing Then Exit Sub
    Call mobjReport.ReportPrintSet(gcnOracle, glngSys, lvwBill.SelectedItem.Tag, Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    mblnItem = False
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To lvwBill.ListItems.Count
            lvwBill.ListItems(i).Checked = True
        Next
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        For i = 1 To lvwBill.ListItems.Count
            lvwBill.ListItems(i).Checked = False
        Next
    End If
End Sub

Private Sub Form_Load()
    '���Ƶ��ݴ�ӡ��ʽ:0-����ӡ,1-�ֹ���ӡ,2-�Զ���ӡ
    If mstrǰ��IDs = "" Then
        If mint���� = 1 Then
            mint��ӡ��ʽ = Val(zlDatabase.GetPara("���﷢�͵��ݴ�ӡ", glngSys, p����ҽ���´�))
        Else
            mint��ӡ��ʽ = Val(zlDatabase.GetPara("סԺ���͵��ݴ�ӡ", glngSys, pסԺҽ������))
        End If
    Else
        mint��ӡ��ʽ = 1
    End If
    If mint��ӡ��ʽ = 0 Then Unload Me: Exit Sub
    mint���뵥��ӡģʽ = Val(zlDatabase.GetPara("��Ѫ���뵥��ӡģʽ", glngSys, pסԺҽ������, "1"))
    
    Call RestoreListViewState(lvwBill, App.ProductName & "\" & Me.Name, "")
    If Not LoadBill Then Unload Me: Exit Sub
    If lvwBill.ListItems.Count = 0 Then Unload Me: Exit Sub
    mblnItem = False
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    
    '�Զ���ӡ���˳�
    If mint��ӡ��ʽ = 2 Then
        Call cmdPrint_Click
        Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set mobjReport = Nothing   '�Զ������Ա㱨�����еĻ������ظ�ʹ��
    Call SaveListViewState(lvwBill, App.ProductName & "\" & Me.Name, "")
End Sub

Private Sub lvwBill_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwBill, ColumnHeader.Index)
End Sub

Private Function LoadBill() As Boolean
'���ܣ���ȡ���η��Ϳ��Դ�ӡ�����Ƶ����б�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim objItem As ListItem
    Dim strTmp As String
    
    lvwBill.ListItems.Clear
    
    On Error GoTo errH
    
    '�����סԺ�����ų����뵥�´����Ѫҽ����ͨ�����ⷽʽ����
    If mint���� = 2 And mint���뵥��ӡģʽ = 1 Then
        If gblnѪ��ϵͳ = True Then
            strTmp = " And (NVL(b.�������,0)=0 Or b.������� <>'K')" & _
                " Union All " & _
                " Select 0, No, ��¼����, '-17', Decode(���, 1, '��Ѫ���뵥', 'ȡѪ֪ͨ��') ����, '�Բ��˽�����Ѫ���Ƶ����뵥��', ҽ��id, ���" & vbNewLine & _
                " From (Select b.No, b.��¼����, b.ҽ��id, Decode(C.��������, '8', Nvl(C.ִ�з���, 0), 0)+1 ���" & vbNewLine & _
                "       From ������ĿĿ¼ c, ����ҽ����¼ d, ����ҽ����¼ a, ����ҽ������ b" & vbNewLine & _
                "       Where Instr(',8,9,', ',' || c.�������� || ',') > 0 And c.Id = d.������Ŀid And d.������� = 'E' And d.���id = a.Id And" & vbNewLine & _
                "             a.Id = b.ҽ��id And a.ҽ����Ч = 1 And b.���ͺ� = [1] And Nvl(a.�������, 0) <> 0 And a.������� = 'K' And a.ҽ��״̬ = 8)"
        Else
            strTmp = " And (NVL(b.�������,0)=0 Or b.������� <>'K')" & _
                " Union All " & _
                " Select 0,B.NO,B.��¼����,'-17','��Ѫ���뵥','�Բ��˽�����Ѫ���Ƶ����뵥��',B.ҽ��ID,0 From ����ҽ����¼ A,����ҽ������ B Where A.ID=B.ҽ��ID And A.ҽ����Ч=1 And B.���ͺ�=[1] And NVL(A.�������,0)<>0 And A.������� = 'K' And A.ҽ��״̬=8 "
        End If
    End If
    
    '�����������Ƶ���,���ݵ��ݱ�ŵ��ñ���(�൱��֪ͨ��)
    strSql = "Select Distinct D.ID,A.NO,A.��¼����,D.���,D.����,D.˵��,0 AS ҽ��ID,0 ���" & _
        " From ����ҽ������ A,����ҽ����¼ B,��������Ӧ�� C,�����ļ��б� D" & _
        " Where A.���ͺ�=[1] And A.ҽ��ID=B.ID" & _
        " And B.������ĿID=C.������ĿID And C.Ӧ�ó���=[2] and (not D.˵�� like '%<�¿�ʱ��ӡ>%' Or NVL(D.��ʽ,0)<>1)" & _
        " And C.�����ļ�ID=D.ID And D.����=7" & _
        strTmp & _
        " Order by NO"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng���ͺ�, mint����)
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID & "_" & rsTmp!NO & "_" & rsTmp!��¼����, rsTmp!NO)
        objItem.SubItems(1) = Nvl(rsTmp!����)
        objItem.SubItems(2) = Nvl(rsTmp!˵��)
        '���С��0��ʾʹ�ò����̶�����
        If Val(rsTmp!��� & "") < 0 And Val(rsTmp!ID & "") = 0 Then
            objItem.Tag = "ZL1_INSIDE_1254_" & Abs(Val(rsTmp!��� & "")) & IIF(Val(rsTmp!��� & "") = 0, "", "_" & Val(rsTmp!��� & "")) '��Ӧ���Զ��屨����
        Else
            objItem.Tag = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" '��Ӧ���Զ��屨����
        End If
        objItem.ListSubItems(1).Tag = rsTmp!��¼����
        objItem.ListSubItems(2).Tag = rsTmp!ҽ��ID
        objItem.Checked = True
        rsTmp.MoveNext
    Next
    LoadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lvwBill_DblClick()
    If mblnItem Then Call lvwBill_KeyPress(13)
End Sub

Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If mblnItem Then
        Item.Selected = True
        Item.EnsureVisible
    End If
End Sub

Private Sub lvwBill_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdSetup_Click
    End If
End Sub

Private Sub lvwBill_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    Dim strSql As String
    
    '���뵥�ݴ�ӡ֮��Ĵ���
    If mstrBillPrint <> "" Then
        If Split(mstrBillPrint, ",")(0) = ReportNum Then
            strSql = "Zl_���Ƶ��ݴ�ӡ_Insert('" & Split(mstrBillPrint, ",")(1) & "'," & Val(Split(mstrBillPrint, ",")(2)) & ",1,'" & UserInfo.���� & "')"
        End If
    End If
    
    On Error GoTo errH
    If strSql <> "" Then
        zlDatabase.ExecuteProcedure strSql, Me.Name
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
