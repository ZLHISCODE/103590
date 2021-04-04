VERSION 5.00
Begin VB.Form frmTechnicQueueCfg 
   BorderStyle     =   0  'None
   Caption         =   "排队叫号设置"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "frmTechnicQueueCfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkUseQueue 
      Caption         =   "启用排队叫号"
      Height          =   180
      Left            =   240
      TabIndex        =   22
      ToolTipText     =   "激活排队叫号功能，仅限于影像采集站和影像医技站。"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame framGroup 
      Caption         =   "分组设置"
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7755
      Begin VB.CheckBox chkSelectRoom 
         Caption         =   "报到时分配默认执行间"
         Height          =   210
         Left            =   4080
         TabIndex        =   17
         Top             =   4970
         Width           =   2220
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增分组(&A)"
         Height          =   375
         Left            =   45
         Picture         =   "frmTechnicQueueCfg.frx":000C
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4860
         Width           =   1170
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除分组(&D)"
         Height          =   375
         Left            =   1260
         Picture         =   "frmTechnicQueueCfg.frx":0156
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4860
         Width           =   1170
      End
      Begin VB.CommandButton cmdStudyAcc 
         Caption         =   "关联项目(&R)"
         Height          =   375
         Left            =   6360
         Picture         =   "frmTechnicQueueCfg.frx":02A0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4860
         Width           =   1260
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "修改分组(&M)"
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
         ColNames        =   "|ID,hide,key|组名,w1400,read|分组前缀,w1500,read|"
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
         ColNames        =   "|关联检查项目>名称,w2100,read|项目编码>编码,w1100,read|"
         KeyName         =   "≡"
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
         ColNames        =   "|ID,hide|执行间,w1400,read|号码前缀,w1400,read|"
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
         Caption         =   "报到后自动排队"
         Height          =   180
         Left            =   3840
         TabIndex        =   21
         Top             =   1130
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkUseQueueMsg 
         Caption         =   "启用排队消息处理"
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
         Caption         =   "未指定检查执行间的排队方式"
         Height          =   810
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   2745
         Begin VB.OptionButton optNumberRule 
            Caption         =   "按检查科室排队"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   8
            ToolTipText     =   "对于分配了执行间的检查，排队号码将按执行间连续生成，对未分配执行的检查，排队号码将按科室连续生成。"
            Top             =   240
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.OptionButton optNumberRule 
            Caption         =   "按检查分组排队"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   7
            ToolTipText     =   "对于分配了执行间的检查，排队号码将按执行间连续生成，对未分配执行的检查，排队号码将根据检查所属分组连续生成。"
            Top             =   480
            Width           =   1665
         End
      End
      Begin VB.CheckBox chkSynStudyList 
         Caption         =   "同步定位检查列表"
         Height          =   180
         Left            =   2880
         TabIndex        =   5
         ToolTipText     =   "点击排队列表或呼叫列表数据后，同步定位到检查列表"
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
         Caption         =   "排号单打印方式："
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1115
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "排号单报表编号："
         Height          =   225
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "排队打号时对应的自定义报表编号。"
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "数据有效天数：       天"
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
Private mblnRefreshed  As Boolean '判断该界面是否已经刷新


Private Sub LoadGroupInf()
'载入医技分组信息
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select Id, 组名,分组前缀 from 影像执行分组 where 科室ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询分组信息", mlngDeptId)
    
    Call ufgGroupCfg.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "组名 asc"
    
    Set ufgGroupCfg.AdoData = rsData
    Call ufgGroupCfg.BindData
End Sub

Private Sub LoadTechniRoom(ByVal lngGroupId As Long)
'载入分组所含的医技执行房间
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select 执行间, 号码前缀 from 医技执行房间 where 分组Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询医技执行房间", lngGroupId)
    
    Call ufgRoomCfg.ClearListData
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Sort = "执行间 asc"
    
    Set ufgRoomCfg.AdoData = rsData
    Call ufgRoomCfg.BindData
End Sub

Private Sub LoadStudyProAssociation(ByVal lngGroupId As Long)
'载入检查项目关联
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select 名称,编码 from 诊疗项目目录 a, 影像分组关联 b where a.id=b.诊疗项目Id and b.分组Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询影像分组关联检查项目", lngGroupId)
    
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
'新增分组信息
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim strGroupName As String
    Dim strPrefix As String
    Dim objFrmAdd As frmTechnicGroup
    Dim lngRow As Long
    
    '调用分组添加窗口
    Set objFrmAdd = New frmTechnicGroup
    If objFrmAdd.ShowGroupCfg(Me, mlngDeptId, lngGroupId, strGroupName, strPrefix) Then
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
    Dim strSQL As String
    Dim lngGroupId As Long
    Dim lngMsgResult As Long
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBoxD Me, "请选择需要删除的分组数据。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    lngMsgResult = MsgBoxD(Me, "是否确认删除该分组数据? 删除后分组将不可恢复。", vbYesNo, "提示")
    If lngMsgResult = vbNo Then Exit Sub
    
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    
    strSQL = "zl_影像执行分组_Del(" & lngGroupId & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "删除执行分组")
    
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
    If objFrmUpdate.ShowGroupCfg(Me, mlngDeptId, lngGroupId, strGroupName, strPrefix) Then
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
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngIndex As Long

    On Error GoTo err

    mblnRefreshed = False
    mlngDeptId = lngDeptID

    lngIndex = Val(GetDeptPara(mlngDeptId, "排队叫号编码规则", 0))
    txtValidDays.Text = GetDeptPara(mlngDeptId, "排队数据保存天数", 1)
    txtQueueReport.Text = GetDeptPara(mlngDeptId, "排队单报表编号", "")
    chkSynStudyList.value = Val(GetDeptPara(mlngDeptId, "同步定位检查列表", 0))
    chkSelectRoom.value = Val(GetDeptPara(mlngDeptId, "报到时分配默认执行间", 0))
    chkUseQueueMsg.value = Val(GetDeptPara(mlngDeptId, "启用排队消息处理", 1))
    chkAutoInQueue.value = Val(GetDeptPara(mlngDeptId, "报到后自动排队", 1))
    
    '0-不打印，1-自动打印，2-提示打印
    cbxPrintQueueNoWay.ListIndex = Val(GetDeptPara(mlngDeptId, "排队单打印方式", 0))
    
    chkUseQueue.value = Val(GetDeptPara(mlngDeptId, "启动排队叫号", 0))
    
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
    If mlngDeptId < 0 Then Exit Sub

    SetDeptPara mlngDeptId, "启动排队叫号", chkUseQueue.value
    SetDeptPara mlngDeptId, "排队叫号编码规则", IIf(optNumberRule(0).value, 0, 1)
    SetDeptPara mlngDeptId, "排队数据保存天数", Val(txtValidDays.Text)
    SetDeptPara mlngDeptId, "排队单报表编号", txtQueueReport.Text
    SetDeptPara mlngDeptId, "同步定位检查列表", chkSynStudyList.value
    SetDeptPara mlngDeptId, "报到时分配默认执行间", chkSelectRoom.value
    SetDeptPara mlngDeptId, "排队单打印方式", cbxPrintQueueNoWay.ListIndex
    SetDeptPara mlngDeptId, "启用排队消息处理", chkUseQueueMsg.value
    SetDeptPara mlngDeptId, "报到后自动排队", chkAutoInQueue.value
End Sub
