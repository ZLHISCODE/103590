VERSION 5.00
Begin VB.Form frmTechnicQueueCfg 
   BorderStyle     =   0  'None
   Caption         =   "�Ŷӽк�����"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "frmTechnicQueueCfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkUseQueue 
      Caption         =   "�����Ŷӽк�"
      Height          =   180
      Left            =   240
      TabIndex        =   22
      ToolTipText     =   "�����ŶӽкŹ��ܣ�������Ӱ��ɼ�վ��Ӱ��ҽ��վ��"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame framGroup 
      Caption         =   "��������"
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7755
      Begin VB.CheckBox chkSelectRoom 
         Caption         =   "����ʱ����Ĭ��ִ�м�"
         Height          =   210
         Left            =   4080
         TabIndex        =   17
         Top             =   4970
         Width           =   2220
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "��������(&A)"
         Height          =   375
         Left            =   45
         Picture         =   "frmTechnicQueueCfg.frx":000C
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4860
         Width           =   1170
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ������(&D)"
         Height          =   375
         Left            =   1260
         Picture         =   "frmTechnicQueueCfg.frx":0156
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4860
         Width           =   1170
      End
      Begin VB.CommandButton cmdStudyAcc 
         Caption         =   "������Ŀ(&R)"
         Height          =   375
         Left            =   6360
         Picture         =   "frmTechnicQueueCfg.frx":02A0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4860
         Width           =   1260
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "�޸ķ���(&M)"
         Height          =   375
         Left            =   2460
         Picture         =   "frmTechnicQueueCfg.frx":03EA
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4860
         Width           =   1170
      End
      Begin zl9PACSWork.ucFlexGrid ufgGroupCfg 
         Height          =   4560
         Left            =   90
         TabIndex        =   14
         Top             =   255
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   8043
         DefaultCols     =   ""
         ColNames        =   "|ID,hide,key|����,w1400,read|����ǰ׺,w1500,read|"
         KeyName         =   "ID"
         DisCellColor    =   16777215
         IsRowNumber     =   0   'False
         HeadCheckValue  =   1
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         AllowExtCol     =   0   'False
         IsShowPopupMenu =   0   'False
         HeadFontCharset =   134
         HeadFontWeight  =   400
         HeadColor       =   0
         DataFontCharset =   134
         DataFontWeight  =   400
         DataColor       =   -2147483640
         RowHeightMin    =   260
         ExtendLastCol   =   -1  'True
      End
      Begin zl9PACSWork.ucFlexGrid ufgStudyProCfg 
         Height          =   2550
         Left            =   3345
         TabIndex        =   15
         Top             =   2265
         Width           =   4320
         _ExtentX        =   7408
         _ExtentY        =   4498
         DefaultCols     =   ""
         ColNames        =   "|���������Ŀ>����,w2100,read|��Ŀ����>����,w1100,read|"
         KeyName         =   "��"
         DisCellColor    =   16777215
         HeadCheckValue  =   1
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         IsShowPopupMenu =   0   'False
         HeadFontCharset =   134
         HeadFontWeight  =   400
         HeadColor       =   0
         DataFontCharset =   134
         DataFontWeight  =   400
         DataColor       =   -2147483640
         RowHeightMin    =   260
         ExtendLastCol   =   -1  'True
      End
      Begin zl9PACSWork.ucFlexGrid ufgRoomCfg 
         Height          =   1965
         Left            =   3345
         TabIndex        =   16
         Top             =   255
         Width           =   4320
         _ExtentX        =   7408
         _ExtentY        =   3466
         DefaultCols     =   ""
         ColNames        =   "|ID,hide|ִ�м�,w1400,read|����ǰ׺,w1400,read|"
         KeyName         =   "ID"
         DisCellColor    =   16777215
         IsRowNumber     =   0   'False
         HeadCheckValue  =   1
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         AllowExtCol     =   0   'False
         IsShowPopupMenu =   0   'False
         HeadFontCharset =   134
         HeadFontWeight  =   400
         HeadColor       =   0
         DataFontCharset =   134
         DataFontWeight  =   400
         DataColor       =   -2147483640
         RowHeightMin    =   260
         ExtendLastCol   =   -1  'True
      End
   End
   Begin VB.Frame framConfig 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   5355
      Width           =   7815
      Begin VB.CheckBox chkAutoInQueue 
         Caption         =   "�������Զ��Ŷ�"
         Height          =   180
         Left            =   3840
         TabIndex        =   21
         Top             =   1130
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkUseQueueMsg 
         Caption         =   "�����Ŷ���Ϣ����"
         Height          =   180
         Left            =   5880
         TabIndex        =   20
         Top             =   1130
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.ComboBox cbxPrintQueueNoWay 
         Height          =   300
         ItemData        =   "frmTechnicQueueCfg.frx":0534
         Left            =   1635
         List            =   "frmTechnicQueueCfg.frx":0541
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1080
         Width           =   1740
      End
      Begin VB.Frame Frame1 
         Caption         =   "δָ�����ִ�м���Ŷӷ�ʽ"
         Height          =   810
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   2745
         Begin VB.OptionButton optNumberRule 
            Caption         =   "���������Ŷ�"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   8
            ToolTipText     =   "���ڷ�����ִ�м�ļ�飬�ŶӺ��뽫��ִ�м��������ɣ���δ����ִ�еļ�飬�ŶӺ��뽫�������������ɡ�"
            Top             =   240
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.OptionButton optNumberRule 
            Caption         =   "���������Ŷ�"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   7
            ToolTipText     =   "���ڷ�����ִ�м�ļ�飬�ŶӺ��뽫��ִ�м��������ɣ���δ����ִ�еļ�飬�ŶӺ��뽫���ݼ�����������������ɡ�"
            Top             =   480
            Width           =   1665
         End
      End
      Begin VB.CheckBox chkSynStudyList 
         Caption         =   "ͬ����λ����б�"
         Height          =   180
         Left            =   2880
         TabIndex        =   5
         ToolTipText     =   "����Ŷ��б������б����ݺ�ͬ����λ������б�"
         Top             =   330
         Width           =   1815
      End
      Begin VB.TextBox txtQueueReport 
         Height          =   315
         Left            =   1635
         TabIndex        =   4
         Top             =   690
         Width           =   2940
      End
      Begin VB.TextBox txtValidDays 
         Height          =   315
         Left            =   1635
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "1"
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "�źŵ���ӡ��ʽ��"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1115
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "�źŵ������ţ�"
         Height          =   225
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "�ŶӴ��ʱ��Ӧ���Զ��屨���š�"
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "������Ч������       ��"
         Height          =   210
         Left            =   420
         TabIndex        =   1
         Top             =   330
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmTechnicQueueCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptId As Long
Private mblnRefreshed  As Boolean '�жϸý����Ƿ��Ѿ�ˢ��


Private Sub LoadGroupInf()
'����ҽ��������Ϣ
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select Id, ����,����ǰ׺ from Ӱ��ִ�з��� where ����ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������Ϣ", mlngDeptId)
    
    Call ufgGroupCfg.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "���� asc"
    
    Set ufgGroupCfg.AdoData = rsData
    Call ufgGroupCfg.BindData
End Sub

Private Sub LoadTechniRoom(ByVal lngGroupId As Long)
'�������������ҽ��ִ�з���
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select ִ�м�, ����ǰ׺ from ҽ��ִ�з��� where ����Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ��ִ�з���", lngGroupId)
    
    Call ufgRoomCfg.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "ִ�м� asc"
    
    Set ufgRoomCfg.AdoData = rsData
    Call ufgRoomCfg.BindData
End Sub

Private Sub LoadStudyProAssociation(ByVal lngGroupId As Long)
'��������Ŀ����
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select ����,���� from ������ĿĿ¼ a, Ӱ�������� b where a.id=b.������ĿId and b.����Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯӰ�������������Ŀ", lngGroupId)
    
    Call ufgStudyProCfg.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "����"
    
    Set ufgStudyProCfg.AdoData = rsData
    Call ufgStudyProCfg.BindData
End Sub

Private Sub chkUseQueue_Click()
On Error GoTo ErrHandle
    optNumberRule(0).Enabled = chkUseQueue.value
    optNumberRule(1).Enabled = chkUseQueue.value
        
    'ufgGroupCfg.Enabled = chkUseQueue.value
    'ufgRoomCfg.Enabled = chkUseQueue.value
    'ufgStudyProCfg.Enabled = chkUseQueue.value
    
    'cmdAdd.Enabled = chkUseQueue.value
    'cmdDel.Enabled = chkUseQueue.value
    'cmdModify.Enabled = chkUseQueue.value
    'cmdStudyAcc.Enabled = chkUseQueue.value
    chkSynStudyList.Enabled = chkUseQueue.value
    
    txtValidDays.Enabled = chkUseQueue.value
    txtQueueReport.Enabled = chkUseQueue.value
    cbxPrintQueueNoWay.Enabled = chkUseQueue.value
    chkAutoInQueue.Enabled = chkUseQueue.value
    chkUseQueueMsg.Enabled = chkUseQueue.value
    
    Label1.Enabled = chkUseQueue.value
    Label2.Enabled = chkUseQueue.value
    Frame1.Enabled = chkUseQueue.value
    
    'framGroup.Enabled = chkUseQueue.value
    
    mblnRefreshed = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdAdd_Click()
'����������Ϣ
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim strGroupName As String
    Dim strPrefix As String
    Dim objFrmAdd As frmTechnicGroup
    Dim lngRow As Long
    
    '���÷�����Ӵ���
    Set objFrmAdd = New frmTechnicGroup
    If objFrmAdd.ShowGroupCfg(Me, mlngDeptId, lngGroupId, strGroupName, strPrefix) Then
        lngRow = ufgGroupCfg.NewRow
    
        ufgGroupCfg.Text(lngRow, "ID") = lngGroupId
        ufgGroupCfg.Text(lngRow, "����") = strGroupName
        ufgGroupCfg.Text(lngRow, "����ǰ׺") = strPrefix
        
        '�������ִ�м�
        Call LoadTechniRoom(lngGroupId)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdDel_Click()
On Error GoTo ErrHandle
    Dim strSQL As String
    Dim lngGroupId As Long
    Dim lngMsgResult As Long
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBoxD Me, "��ѡ����Ҫɾ���ķ������ݡ�", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    lngMsgResult = MsgBoxD(Me, "�Ƿ�ȷ��ɾ���÷�������? ɾ������齫���ɻָ���", vbYesNo, "��ʾ")
    If lngMsgResult = vbNo Then Exit Sub
    
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    
    strSQL = "zl_Ӱ��ִ�з���_Del(" & lngGroupId & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "ɾ��ִ�з���")
    
    Call ufgRoomCfg.ClearListData
    Call ufgGroupCfg.DelCurRow(False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdModify_Click()
'�޸ķ�����Ϣ
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim strGroupName As String
    Dim strPrefix As String
    Dim objFrmUpdate As frmTechnicGroup
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBoxD Me, "��ѡ����Ҫ�޸ĵķ������ݡ�", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    strGroupName = ufgGroupCfg.Text(ufgGroupCfg.SelectionRow, "����")
    strPrefix = ufgGroupCfg.Text(ufgGroupCfg.SelectionRow, "����ǰ׺")
    
    '���÷�����´���
    Set objFrmUpdate = New frmTechnicGroup
    If objFrmUpdate.ShowGroupCfg(Me, mlngDeptId, lngGroupId, strGroupName, strPrefix) Then
        ufgGroupCfg.CurText("����") = strGroupName
        ufgGroupCfg.CurText("����ǰ׺") = strPrefix
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdStudyAcc_Click()
'Ӱ������Ŀ��������
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim objStudyAssocia As frmTechnicStudy
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBoxD Me, "��ѡ����Ҫ���й����ķ������ݡ�", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    
    Set objStudyAssocia = New frmTechnicStudy
    If objStudyAssocia.ShowStudyAssociation(mlngDeptId, lngGroupId, Me) Then
        Call ufgStudyProCfg.ClearListData
        Call LoadStudyProAssociation(lngGroupId)
    End If
    
Exit Sub
ErrHandle:
If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
''Debug Code
'    InitDebugObject 1290, Me, "zlhis", "HIS"
'    mlngDeptID = 63
'
'    LoadGroupInf
''Debug End
End Sub


Private Sub Form_Resize()
    framGroup.Left = (Me.ScaleWidth - framGroup.Width) / 2
    framConfig.Left = framGroup.Left
    chkUseQueue.Left = framConfig.Left + 120
End Sub

Private Sub optNumberRule_Click(Index As Integer)
On Error GoTo ErrHandle
    mblnRefreshed = True
    
    chkSelectRoom.Enabled = IIf(Index = 1, True, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgGroupCfg_OnSelChange()
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    lngGroupId = Val(ufgGroupCfg.CurKeyValue)
    
    '����ҽ��ִ�з���
    Call LoadTechniRoom(lngGroupId)
    
    '�����������Ŀ����
    Call LoadStudyProAssociation(lngGroupId)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgRoomCfg_OnDblClick()
'˫��ִ�м�ʱ�����з����޸Ĵ���
On Error GoTo ErrHandle
    Call cmdModify_Click
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyProCfg_OnDblClick()
'˫��Ӱ������Ŀʱ�����й������ô���
On Error GoTo ErrHandle
    Call cmdStudyAcc_Click
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub zlRefresh(lngDeptID As Long)
'ˢ�����ò���
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngIndex As Long

    On Error GoTo err

    mblnRefreshed = False
    mlngDeptId = lngDeptID

    lngIndex = Val(GetDeptPara(mlngDeptId, "�Ŷӽкű������", 0))
    txtValidDays.Text = GetDeptPara(mlngDeptId, "�Ŷ����ݱ�������", 1)
    txtQueueReport.Text = GetDeptPara(mlngDeptId, "�Ŷӵ�������", "")
    chkSynStudyList.value = Val(GetDeptPara(mlngDeptId, "ͬ����λ����б�", 0))
    chkSelectRoom.value = Val(GetDeptPara(mlngDeptId, "����ʱ����Ĭ��ִ�м�", 0))
    chkUseQueueMsg.value = Val(GetDeptPara(mlngDeptId, "�����Ŷ���Ϣ����", 1))
    chkAutoInQueue.value = Val(GetDeptPara(mlngDeptId, "�������Զ��Ŷ�", 1))
    
    '0-����ӡ��1-�Զ���ӡ��2-��ʾ��ӡ
    cbxPrintQueueNoWay.ListIndex = Val(GetDeptPara(mlngDeptId, "�Ŷӵ���ӡ��ʽ", 0))
    
    chkUseQueue.value = Val(GetDeptPara(mlngDeptId, "�����Ŷӽк�", 0))
    
    Call LoadGroupInf

    optNumberRule(lngIndex).value = True

    Call chkUseQueue_Click

    mblnRefreshed = True

    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub
 
Public Sub zlSave()
'�������ò���
    If mblnRefreshed = False Then Exit Sub
    If mlngDeptId < 0 Then Exit Sub

    SetDeptPara mlngDeptId, "�����Ŷӽк�", chkUseQueue.value
    SetDeptPara mlngDeptId, "�Ŷӽкű������", IIf(optNumberRule(0).value, 0, 1)
    SetDeptPara mlngDeptId, "�Ŷ����ݱ�������", Val(txtValidDays.Text)
    SetDeptPara mlngDeptId, "�Ŷӵ�������", txtQueueReport.Text
    SetDeptPara mlngDeptId, "ͬ����λ����б�", chkSynStudyList.value
    SetDeptPara mlngDeptId, "����ʱ����Ĭ��ִ�м�", chkSelectRoom.value
    SetDeptPara mlngDeptId, "�Ŷӵ���ӡ��ʽ", cbxPrintQueueNoWay.ListIndex
    SetDeptPara mlngDeptId, "�����Ŷ���Ϣ����", chkUseQueueMsg.value
    SetDeptPara mlngDeptId, "�������Զ��Ŷ�", chkAutoInQueue.value
End Sub
