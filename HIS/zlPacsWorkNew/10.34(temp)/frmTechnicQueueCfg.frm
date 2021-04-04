VERSION 5.00
Begin VB.Form frmTechnicQueueCfg 
   BorderStyle     =   0  'None
   Caption         =   "排队叫号设置"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "frmTechnicQueueCfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkUseQueue 
      Caption         =   "启用排队叫号"
      Height          =   180
      Left            =   195
      TabIndex        =   0
      ToolTipText     =   "激活排队叫号功能，仅限于影像采集站和影像医技站。"
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
         Caption         =   "分组设置"
         Height          =   5175
         Left            =   90
         TabIndex        =   4
         Top             =   810
         Width           =   7635
         Begin VB.CommandButton cmdModify 
            Caption         =   "修改分组(&M)"
            Height          =   375
            Left            =   2460
            Picture         =   "frmTechnicQueueCfg.frx":000C
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   4710
            Width           =   1170
         End
         Begin VB.CommandButton cmdStudyAcc 
            Caption         =   "关联项目(&R)"
            Height          =   375
            Left            =   6270
            Picture         =   "frmTechnicQueueCfg.frx":0156
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   4710
            Width           =   1260
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "删除分组(&D)"
            Height          =   375
            Left            =   1260
            Picture         =   "frmTechnicQueueCfg.frx":02A0
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   4710
            Width           =   1170
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "新增分组(&A)"
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
            ColNames        =   "|ID,hide,key|组名,w1400,read|分组前缀,w1500,read|"
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
            ColNames        =   "|关联检查项目>名称,w2100,read|项目编码>编码,w1100,read|"
            KeyName         =   "≡"
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
            ColNames        =   "|ID,hide|执行间,w1400,read|号码前缀,w1400,read|"
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
         Caption         =   "按分组规则产生排队号(分配执行间的按执行间连续排号，否则按分组连续排号)"
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   3
         Top             =   555
         Width           =   6720
      End
      Begin VB.OptionButton optNumberRule 
         Caption         =   "按默认规则产生排队号(分配执行间的按执行间连续排号，否则按科室连续排号)"
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
Private mblnRefreshed  As Boolean '判断该界面是否已经刷新


Private Sub LoadGroupInf()
'载入医技分组信息
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select Id, 组名,分组前缀 from 影像执行分组 where 科室ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询分组信息", mlngDeptID)
    
    Call ufgGroupCfg.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "组名 asc"
    
    Set ufgGroupCfg.AdoData = rsData
    Call ufgGroupCfg.BindData
End Sub

Private Sub LoadTechniRoom(ByVal lngGroupId As Long)
'载入分组所含的医技执行房间
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select 执行间, 号码前缀 from 医技执行房间 where 分组Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询医技执行房间", lngGroupId)
    
    Call ufgRoomCfg.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "执行间 asc"
    
    Set ufgRoomCfg.AdoData = rsData
    Call ufgRoomCfg.BindData
End Sub

Private Sub LoadStudyProAssociation(ByVal lngGroupId As Long)
'载入检查项目关联
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select 名称,编码 from 诊疗项目目录 a, 影像检查项目 b where a.id=b.诊疗项目Id and b.分组Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询影像分组关联检查项目", lngGroupId)
    
    Call ufgStudyProCfg.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "名称"
    
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
'新增分组信息
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim strGroupName As String
    Dim strPrefix As String
    Dim objFrmAdd As frmTechnicGroup
    Dim lngRow As Long
    
    '调用分组添加窗口
    Set objFrmAdd = New frmTechnicGroup
    If objFrmAdd.ShowGroupCfg(Me, mlngDeptID, lngGroupId, strGroupName, strPrefix) Then
        lngRow = ufgGroupCfg.NewRow
    
        ufgGroupCfg.Text(lngRow, "ID") = lngGroupId
        ufgGroupCfg.Text(lngRow, "组名") = strGroupName
        ufgGroupCfg.Text(lngRow, "分组前缀") = strPrefix
        
        '载入分组执行间
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
        MsgBoxD Me, "请选择需要删除的分组数据。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    lngMsgResult = MsgBoxD(Me, "是否确认删除该分组数据? 删除后分组将不可恢复。", vbYesNo, "提示")
    If lngMsgResult = vbNo Then Exit Sub
    
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    
    strSql = "zl_影像执行分组_Del(" & lngGroupId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "删除执行分组")
    
    Call ufgRoomCfg.ClearListData
    Call ufgGroupCfg.DelCurRow(False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdModify_Click()
'修改分组信息
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim strGroupName As String
    Dim strPrefix As String
    Dim objFrmUpdate As frmTechnicGroup
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBoxD Me, "请选择需要修改的分组数据。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    strGroupName = ufgGroupCfg.Text(ufgGroupCfg.SelectionRow, "组名")
    strPrefix = ufgGroupCfg.Text(ufgGroupCfg.SelectionRow, "分组前缀")
    
    '调用分组更新窗口
    Set objFrmUpdate = New frmTechnicGroup
    If objFrmUpdate.ShowGroupCfg(Me, mlngDeptID, lngGroupId, strGroupName, strPrefix) Then
        ufgGroupCfg.CurText("组名") = strGroupName
        ufgGroupCfg.CurText("分组前缀") = strPrefix
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdStudyAcc_Click()
'影像检查项目关联设置
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim objStudyAssocia As frmTechnicStudy
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBoxD Me, "请选择需要进行关联的分组数据。", vbOKOnly, "提示"
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
    
    '载入医技执行房间
    Call LoadTechniRoom(lngGroupId)
    
    '载入分组检查项目关联
    Call LoadStudyProAssociation(lngGroupId)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgRoomCfg_OnDblClick()
'双击执行间时，进行分组修改处理
On Error GoTo ErrHandle
    Call cmdModify_Click
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyProCfg_OnDblClick()
'双击影像检查项目时，进行关联配置处理
On Error GoTo ErrHandle
    Call cmdStudyAcc_Click
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub zlRefresh(lngDeptID As Long)
'刷新配置参数
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngIndex As Long

    On Error GoTo err

    mblnRefreshed = False
    mlngDeptID = lngDeptID


    chkUseQueue.value = Val(GetDeptPara(mlngDeptID, "启动排队叫号", 0))
    lngIndex = Val(GetDeptPara(mlngDeptID, "排队叫号编码规则", 0))

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
'保存配置参数
    If mblnRefreshed = False Then Exit Sub
    If mlngDeptID < 0 Then Exit Sub

    SetDeptPara mlngDeptID, "启动排队叫号", chkUseQueue.value
    SetDeptPara mlngDeptID, "排队叫号编码规则", IIf(optNumberRule(0).value, 0, 1)
End Sub
