VERSION 5.00
Begin VB.Form frmTechnicQueueCfg 
   BorderStyle     =   0  'None
   Caption         =   "�Ŷӽк�����"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "frmTechnicQueueCfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkUseQueue 
      Caption         =   "�����Ŷӽк�"
      Height          =   180
      Left            =   195
      TabIndex        =   0
      ToolTipText     =   "�����ŶӽкŹ��ܣ�������Ӱ��ɼ�վ��Ӱ��ҽ��վ��"
      Top             =   165
      Width           =   1455
   End
   Begin VB.Frame framConfig 
      Height          =   6090
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   7815
      Begin VB.Frame framGroup 
         Caption         =   "��������"
         Height          =   5175
         Left            =   90
         TabIndex        =   4
         Top             =   810
         Width           =   7635
         Begin VB.CommandButton cmdModify 
            Caption         =   "�޸ķ���(&M)"
            Height          =   375
            Left            =   2460
            Picture         =   "frmTechnicQueueCfg.frx":000C
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   4710
            Width           =   1170
         End
         Begin VB.CommandButton cmdStudyAcc 
            Caption         =   "������Ŀ(&R)"
            Height          =   375
            Left            =   6270
            Picture         =   "frmTechnicQueueCfg.frx":0156
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   4710
            Width           =   1260
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "ɾ������(&D)"
            Height          =   375
            Left            =   1260
            Picture         =   "frmTechnicQueueCfg.frx":02A0
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   4710
            Width           =   1170
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "��������(&A)"
            Height          =   375
            Left            =   45
            Picture         =   "frmTechnicQueueCfg.frx":03EA
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   4710
            Width           =   1170
         End
         Begin zl9PACSWork.ucFlexGrid ufgGroupCfg 
            Height          =   4395
            Left            =   90
            TabIndex        =   8
            Top             =   285
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   7752
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
            HeadFontSize    =   10.5
            HeadFontCharset =   134
            HeadFontWeight  =   400
            HeadColor       =   0
            DataFontSize    =   10.5
            DataFontCharset =   134
            DataFontWeight  =   400
            DataColor       =   -2147483640
            RowHeightMin    =   340
            ExtendLastCol   =   -1  'True
         End
         Begin zl9PACSWork.ucFlexGrid ufgStudyProCfg 
            Height          =   2385
            Left            =   3345
            TabIndex        =   9
            Top             =   2295
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   4207
            DefaultCols     =   ""
            ColNames        =   "|���������Ŀ>����,w2100,read|��Ŀ����>����,w1100,read|"
            KeyName         =   "��"
            DisCellColor    =   16777215
            HeadCheckValue  =   1
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            IsShowPopupMenu =   0   'False
            HeadFontSize    =   10.5
            HeadFontCharset =   134
            HeadFontWeight  =   400
            HeadColor       =   0
            DataFontSize    =   10.5
            DataFontCharset =   134
            DataFontWeight  =   400
            DataColor       =   -2147483640
            RowHeightMin    =   340
            ExtendLastCol   =   -1  'True
         End
         Begin zl9PACSWork.ucFlexGrid ufgRoomCfg 
            Height          =   1965
            Left            =   3345
            TabIndex        =   10
            Top             =   285
            Width           =   4200
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
            HeadFontSize    =   10.5
            HeadFontCharset =   134
            HeadFontWeight  =   400
            HeadColor       =   0
            DataFontSize    =   10.5
            DataFontCharset =   134
            DataFontWeight  =   400
            DataColor       =   -2147483640
            RowHeightMin    =   340
            ExtendLastCol   =   -1  'True
         End
      End
      Begin VB.OptionButton optNumberRule 
         Caption         =   "�������������ŶӺ�(����ִ�м�İ�ִ�м������źţ����򰴷��������ź�)"
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   3
         Top             =   555
         Width           =   6720
      End
      Begin VB.OptionButton optNumberRule 
         Caption         =   "��Ĭ�Ϲ�������ŶӺ�(����ִ�м�İ�ִ�м������źţ����򰴿��������ź�)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   6780
      End
   End
End
Attribute VB_Name = "frmTechnicQueueCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptID As Long
Private mblnRefreshed  As Boolean '�жϸý����Ƿ��Ѿ�ˢ��


Private Sub LoadGroupInf()
'����ҽ��������Ϣ
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select Id, ����,����ǰ׺ from Ӱ��ִ�з��� where ����ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ������Ϣ", mlngDeptID)
    
    Call ufgGroupCfg.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "���� asc"
    
    Set ufgGroupCfg.AdoData = rsData
    Call ufgGroupCfg.BindData
End Sub

Private Sub LoadTechniRoom(ByVal lngGroupId As Long)
'�������������ҽ��ִ�з���
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ִ�м�, ����ǰ׺ from ҽ��ִ�з��� where ����Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯҽ��ִ�з���", lngGroupId)
    
    Call ufgRoomCfg.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "ִ�м� asc"
    
    Set ufgRoomCfg.AdoData = rsData
    Call ufgRoomCfg.BindData
End Sub

Private Sub LoadStudyProAssociation(ByVal lngGroupId As Long)
'��������Ŀ����
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ����,���� from ������ĿĿ¼ a, Ӱ������Ŀ b where a.id=b.������ĿId and b.����Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯӰ�������������Ŀ", lngGroupId)
    
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
        
    ufgGroupCfg.Enabled = chkUseQueue.value
    ufgRoomCfg.Enabled = chkUseQueue.value
    ufgStudyProCfg.Enabled = chkUseQueue.value
    
    cmdAdd.Enabled = chkUseQueue.value
    cmdDel.Enabled = chkUseQueue.value
    cmdModify.Enabled = chkUseQueue.value
    cmdStudyAcc.Enabled = chkUseQueue.value
    
    framGroup.Enabled = chkUseQueue.value
    
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
    If objFrmAdd.ShowGroupCfg(Me, mlngDeptID, lngGroupId, strGroupName, strPrefix) Then
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
    Dim strSql As String
    Dim lngGroupId As Long
    Dim lngMsgResult As Long
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBoxD Me, "��ѡ����Ҫɾ���ķ������ݡ�", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    lngMsgResult = MsgBoxD(Me, "�Ƿ�ȷ��ɾ���÷�������? ɾ������齫���ɻָ���", vbYesNo, "��ʾ")
    If lngMsgResult = vbNo Then Exit Sub
    
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    
    strSql = "zl_Ӱ��ִ�з���_Del(" & lngGroupId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "ɾ��ִ�з���")
    
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
    If objFrmUpdate.ShowGroupCfg(Me, mlngDeptID, lngGroupId, strGroupName, strPrefix) Then
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
    If objStudyAssocia.ShowStudyAssociation(lngGroupId, Me) Then
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


Private Sub optNumberRule_Click(Index As Integer)
On Error GoTo ErrHandle
    mblnRefreshed = True
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
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngIndex As Long

    On Error GoTo err

    mblnRefreshed = False
    mlngDeptID = lngDeptID


    chkUseQueue.value = Val(GetDeptPara(mlngDeptID, "�����Ŷӽк�", 0))
    lngIndex = Val(GetDeptPara(mlngDeptID, "�Ŷӽкű������", 0))

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
    If mlngDeptID < 0 Then Exit Sub

    SetDeptPara mlngDeptID, "�����Ŷӽк�", chkUseQueue.value
    SetDeptPara mlngDeptID, "�Ŷӽкű������", IIf(optNumberRule(0).value, 0, 1)
End Sub
