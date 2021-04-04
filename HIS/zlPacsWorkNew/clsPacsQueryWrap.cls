VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacsQueryWrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLngOwnerHwnd As Long
Private mstrLoadMsg As String   '方案加载提示内容
Private mObjOwner As Object
Private mobjPacsMainPicList As PictureBox
Private mblnLoadFilterMenu As Boolean '是否正在创建快速过滤菜单
Private mobjPublicPatient As Object

Private WithEvents mobjFilterCmdBar As CommandBars
Attribute mobjFilterCmdBar.VB_VarHelpID = -1
Private WithEvents mobjList As VSFlexGrid
Attribute mobjList.VB_VarHelpID = -1
Private WithEvents mobjFindPati As PatiIdentify
Attribute mobjFindPati.VB_VarHelpID = -1

Private mobjSquareCard As Object
Private mobjFind As Object
Private mobjIconPanel As Object
Private mobjtxtHistory As TextBox
Private mobjPicHistory As PictureBox
Private mobjImg16 As ImageList
Private mobjImg24 As ImageList
Private mobjFindCmd As CommandButton
Private mobjTab As TabControl
Private mobjGetDataContainer As Object
Private mobjcmdQuery As Object
Private mobjtxtAppend As RichTextBox
Private mobjPicTmp As PictureBox
Private mobjlabPatiInfo As Label

'|---------------------------------------------|
'                   功能跟随容器
Private mobjPicFollow As PictureBox
'                功能跟随图标1
Private mobjImgFollow1 As Image
Private mobjImgFollow2 As Image
Private mobjImgFollow3 As Image
Private mobjImgFollow4 As Image

Private WithEvents mobjTim As Timer
Attribute mobjTim.VB_VarHelpID = -1
Private WithEvents mobjTimHistory  As Timer
Attribute mobjTimHistory.VB_VarHelpID = -1

'主界面控件从上到下 mobjFilterCmdBar，mobjFindPati+mobjFindCmd，mobjList，mobjIconPanel，mobjTab
Private Const C_LAYOUT_BASEHEIGHTOFDETAILINFO As Long = 1200 ' 详细信息基准高度1200
Private Const C_LAYOUT_LISTLEFT As Long = 30 ' 列表左侧设计空出宽度 30
Private Const C_ICON_FIND As Long = 4
Private Const C_ICON_LOCATE As Long = 3
Private Const C_ICON_MENUCHOOSE单选 As Long = 90003
Private Const C_ICON_MENUNOCHOOSE单选 As Long = 90002
Private Const C_ICON_MENUCHOOSE As Long = 90001
Private Const C_ICON_MENUNOCHOOSE As Long = 90000
Private Const C_ICON_STR_报到 As String = "PACS报到"
Private Const C_ICON_STR_书写报告 As String = "PACS书写"
Private Const C_ICON_STR_观片 As String = "观片"
Private Const C_ICON_STR_完成 As String = "PACS完成"
Private Const C_ICON_STR_查看病人信息 As String = "PACS查看病人信息"
Private Const M_STR_MODULE_QUERY As String = "查询"
Private Const M_STR_MODULE_QUERYMORE As String = "查询更多菜单"
Private Const C_BLN_行关联刷新完成 As Boolean = True
Private Const CONST_EXT_PAR_SYNCNEWADVICE As String = "EXTPAR_同步新增医嘱ID"
Private Const C_LNG_使用医疗卡查找定位但是未找到对应病人ID As Long = -1
Private Const CONST_EXT_FILTERKEY As String = "关键标识"
Private Const CONST_EXT_FILTERVALUE As String = "执行间"
Private Const CONST_EXT_FILTERKEY_SAVEWAY As String = "快速过滤项目保存方式"
Private Const CONST_EXT_FILTERKEY_SAVEWAYVALUE As String = "本机"
Private Const CONST_EXT_FILTERKEY_SAVEWAYDEFALT As String = "NULL"
            
'某些功能必须参数变量
Private mblnAssignment As Boolean 'pati控件避免额外进入不必要的事件
Private mblExitSelChange As Boolean '是否需要避免selchang流程:DataSource操作中，避免进入selchange 流程
Private mrtReportType As ReportType
Private mblnCanPrint As Boolean
Private mblnSelRowTransparent As Boolean

'公共信息/系统常用变量参数
Private mcnOracle As ADODB.Connection
Private mlngModule As Long              '模块号
Private mintQueryTime As Integer       '查询间隔时间-根据流程参数中的参数获取
Private mstrDepartId As String
Private mstrSelQueueRooms As String

'当前查询相关
Private mColConfig() As clsQueryPar.TColConfig

Private mrsData As ADODB.Recordset  '数据库查询出的记录集（不能做任何修改）
Private mrsDataShow As ADODB.Recordset 'mrsData经过一些转换后的记录集
Private mstrFindName As String   'pati控件参数名(用于pati控件的查询)
Private mPatiID As Long '用于定位查找功能的病人ID
Private mTqueryType As TqueryType  '查询类型，用于查询条件赋值（LSQ待优化整理）
Private mintShowType As Integer '显示类型0：pacsMain    1：测试
Private mlngSortCol As Long               '当前进行排序的列
Private mintSortOrder As Integer         '当前进行排序的方式

Private WithEvents mobjSqlParse As clsSqlParse  '用于快速过滤项目来源为关联项目选中情况
Attribute mobjSqlParse.VB_VarHelpID = -1
Private WithEvents mobjSqlParseFilter As clsSqlParse  '用快速过滤项目来源为 系统参数
Attribute mobjSqlParseFilter.VB_VarHelpID = -1

''方案信息
Private mstrListKeyCol As String '列表关键列  比如"医嘱ID"

Private mTColSort As TColSort   '排序信息

Private mTPatiIdentifyInfo As TPatiIdentifyInfo

Private mdtStartDate As Date
Private mdtEndDate As Date

Private mPicDictionary As Scripting.Dictionary    '图标缓存
Private mQuickFIlterSaveDict As Scripting.Dictionary    '快速过滤菜单缓存
Private mTQuickFilterState As TQuickFilterState   '用于快速过滤处理（选中状态缓存）
Private mlngSchemeNo As Long '当前使用的方案号

Private mSqlScheme As clsSqlScheme '当前方案
Private mstrSchemeCfg As TSchemeCfg '方案配置参数（查询条件、快速过滤、列头等）

Private WithEvents mObjQuery As clsPacsQuery
Attribute mObjQuery.VB_VarHelpID = -1

Private mstrTemp_新增医嘱ID As String

''当前列信息
Private mobjCurStudyInfo As clsStudyInfo  '用于操作的检查信息

Private mlngAdviceId As Long '当前医嘱ID(选中医嘱ID),
'本参数作用：保存当前选中的检查，若发生查询或者快速过滤，如果当前列表中存在本医嘱ID的检查，则定位到该检查，否则定位到列表第一个数据

''''''''''''''''''''窗体自身控件、界面展示相关变量
Private mlngMove As Long '界面布局位置相关

'''''''''''''事件
Public Event OnLocateBackColor() '需要重新设置颜色
Public Event OnSelChange() 'vsflist onselchange
Public Event OnQueueRoomChanged() '执行间快速过滤项目改变
Public Event OnSwipeCard() '需要触发主界面双击事件
Public Event OnClearFace() '主界面初始化检查
Public Event OnDoStateImage(ByVal lngRow As Long) '需要处理状态图
Public Event OnFunsFollow(ByVal strFun As String) '功能跟随，用于触发对应的功能
Public Event OnListRowSelClear() '列表选中项目重置（例如经过快速过滤后导致列表不显示数据，此时需要同步更新一些状态）
Public Event OnColStatistics(ByVal strStatisticsInfo As String)   '进行列统计
Public Event OnRefreshSelectTab(ByVal lngAdviveID As Long)  '保持原来的功能
Public Event OnSelectScheme(ByVal strName As String)
Public Event OnChangeData(ByVal blnRefreshModul As Boolean, ByVal blnIsSelChange As Boolean)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''''''''''''''
'主要或者复杂sub/fun说明
''刷新列表行数据  UpdateRow
'
'
'
'快速过滤：GetFilterFromQuickFilter  DoQuickFilterMenu

Private Type TPatiIdentifyInfo
    blIsFinding As Boolean '正在进行查找功能
    blShowPatiIdentify As Boolean '是否显示刷卡控件
    
    strFindItems As String '查找项目串
    strLocateItems As String '定位项目串
    strFindItem As String '当前查找项目
    strLocateItem As String '当前定位项目
    strDefault As String '默认查找项
End Type

'行排序信息
Private Type TColSort
    LngSchemeNo As Long '当前方案号
    dictSortInfo As Dictionary
    dictNumberSort As Dictionary
End Type

'行颜色信息
Private Type TRowColorInfo
    LngSchemeNo As Long '当前方案号
    intRowColorIndex As Long '涉及行颜色的列
    blHaveRowColor As Boolean
End Type

'闪烁超时信息
Private Type TFlickerInfo
    LngSchemeNo As Long '当前方案号
    strName As String '闪烁字段名 如： "检查过程"
    strInfo As String '详细信息 如"已登记,申请时间,30|已报到,采样时间,40|"
End Type

'列统计信息
Private Type TColTotalInfo
    strName As String
    intCount As Integer
End Type
                
Private Type TSchemeCfg
    strFilterCfg As String  '快速过滤配置 数据库保存
    strListCfg As String  '列表配置 数据库保存
    strHistoryCfg As String '历史检查配置
    
End Type

Private Type TQuickFilterCmdItem
    intItemIndex As Integer '序号（1,2,3,4...,99）
    blChoose As Boolean '是否选中
    strName As String '项目名称
    strFilterValue As String '项目过滤内容 "a,b,c,d"
End Type
    
Private Type TQuickFilterCmdState
    cmdItem() As TQuickFilterCmdItem '子项目信息
    
    intItemCount As Integer '子项目数量（这一个快速过滤菜单包含的可选项目）
    intFilterWay As Integer '参与过滤方式  0 不参与   1 普通    2 脚本
    lngID As Long '主菜单ID
    
    lngRelationIndex() As Long '关联菜单序号
    intRelaitonCount As Integer '关联过滤数量
    
    strName As String '自身名称(字段名称)
    strChooseMenuName As String '子菜单名称选中情况
    strCustomScript As String '菜单过滤脚本
    strMenuSql As String '菜单项目生成 SQL
    strRelationValueForVBSFilter As String '菜单选中情况  按过滤值记录 直接参与过滤的菜单项目"A,B,C,1213,324,CT"
    strRelationChooseMenu As String '菜单选中情况 按名字记录
    
    blSingleChoose As Boolean '菜单是否单选
    blIsRelationFilter As Boolean '是否关联过滤后者项
    blHaveItem As Boolean '菜单是否存在子项目（下拉选项）
    strKey As String '关键标识，默认为空
End Type

Private Type TQuickFilterState
    intQuickFilterMenuCount As Integer '过滤项目数量
    TCmdState() As TQuickFilterCmdState '子项目信息
End Type

Private Enum EFilterType
    不过滤 = 0
    普通 = 1
    脚本 = 2
    未知 = 3
End Enum

Private Enum EUpDateRowType
    新增 = 1
    更新 = 2
    强制显示 = 3
    后台刷新 = 4
End Enum

Private Enum TqueryType
    更新一行 = 0
    过滤 = 1
    刷新 = 2
    查找 = 3
    数据检索 = 4
End Enum

'默认定位
Private mblnDefaultLocate As Boolean

Property Get DefaultLocate() As Boolean
    DefaultLocate = mblnDefaultLocate
End Property

Property Let DefaultLocate(ByVal value As Boolean)
    mblnDefaultLocate = value
End Property


Property Get StartDate() As Date
    StartDate = Empty
    If Not mObjQuery Is Nothing Then StartDate = mObjQuery.GetFilterValue("系统.开始日期")
    If StartDate = Empty Then StartDate = mdtStartDate
End Property

Property Let StartDate(ByVal value As Date)
    mdtStartDate = value
End Property

Property Get EndDate() As Date
    EndDate = Empty
    If Not mObjQuery Is Nothing Then EndDate = mObjQuery.GetFilterValue("系统.结束日期")
    If EndDate = Empty Then EndDate = mdtEndDate
End Property

Property Let EndDate(ByVal value As Date)
    mdtEndDate = value
End Property

Property Get PatiColor() As Boolean
    PatiColor = False
    If mSqlScheme Is Nothing Then Exit Property
    
    PatiColor = mSqlScheme.PatiColor <> ""
End Property

Property Get SelQueueRooms() As String
On Error GoTo errH
    Dim i As Integer
    
    '135083 在用户新装并且从未使用过的情况下，出现13583问题，原因是mstrSelQueueRooms为空并且无法进入Call DoQueueSelRoom过程
    
    If mstrSelQueueRooms = "" Then
        For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
            If mTQuickFilterState.TCmdState(i).strKey = CONST_EXT_FILTERVALUE Then
            
                '146706 增加这个判断，否则如果没有设置执行间点击排队叫号模块会无限循环
                If mTQuickFilterState.TCmdState(i).intItemCount > 0 Then
                    Call DoQueueSelRoom(i)
                    Exit For
                End If
            End If
        Next
    End If
    SelQueueRooms = mstrSelQueueRooms
    Exit Sub
errH:
    SelQueueRooms = ""
End Property

Property Let AdviceId(value As Long)
    mlngAdviceId = value
End Property

Property Let HistoryCfg(value As String)
    mstrSchemeCfg.strHistoryCfg = value
End Property

Property Get HistoryCfg() As String
    HistoryCfg = mstrSchemeCfg.strHistoryCfg
End Property

Property Get StudyInfo() As clsStudyInfo
    Set StudyInfo = mobjCurStudyInfo
End Property

Property Get blShowPatiIdentify() As Boolean
    blShowPatiIdentify = mTPatiIdentifyInfo.blShowPatiIdentify
End Property

Property Get blExitSelChange() As Long
    blExitSelChange = mblExitSelChange
End Property

Property Get SchemeNo() As Long
    SchemeNo = mlngSchemeNo
End Property

Property Get objSqlParse() As clsSqlParse
    Set objSqlParse = mobjSqlParse
End Property

Property Get SqlScheme() As clsSqlScheme
    Set SqlScheme = mSqlScheme
End Property

Property Get SelectionRow() As Long
    SelectionRow = mobjList.RowSel
End Property

Property Get CurPacsQuery() As clsPacsQuery
    Set CurPacsQuery = mObjQuery
End Property

Property Get ShowingRowCount() As Long
'取得当前显示行的数量
    Dim i As Long
    Dim lngCount As Long
    
    lngCount = 0
    For i = 1 To mobjList.Rows - 1
        If Not mobjList.RowHidden(i) Then lngCount = lngCount + 1
    Next i
    
    ShowingRowCount = lngCount
End Property

Property Get Text(ByVal lngRow As Long, ByVal strColName As String) As String
    Dim iCol As Long
    
    Text = ""
    If lngRow >= mobjList.Rows Then Exit Function
    
    iCol = mobjList.ColIndex(strColName)
    
    If iCol <= -1 Then Exit Function
    
    Text = mobjList.TextMatrix(lngRow, iCol)
End Property

Public Sub Init(ByVal lngDepartId As Long, ByVal lngUserID As Long, ByVal lngModule As Long, ByVal rtReportType As ReportType, ByVal blnCanPrint As Boolean, _
                objSquareCard As Object, objOwner As Object, ByRef objclsQueryPar As clsQueryPar)
                           
'参数说明
'lngDepartId 科室ID
'lngUserID 用户ID
'lngModule 模块号
'rtReportType 报告类型
'objOwner frmPacsWorkNew
'其余objOwner 控件  直接使用objOwner赋值
    Dim strPath As String

    mstrListKeyCol = "医嘱ID"
    mlngModule = lngModule
    mintShowType = 0
    mstrDepartId = lngDepartId
    mrtReportType = rtReportType
    mblnCanPrint = blnCanPrint
    
    Set mPicDictionary = New Dictionary
    Set mObjQuery = New clsPacsQuery
    Set mObjOwner = objOwner
    
    Set mobjSquareCard = objSquareCard

    Set mobjFindCmd = objclsQueryPar.cmdFind
    Set mobjImg16 = objclsQueryPar.ImageList16
    Set mobjImg24 = objclsQueryPar.ImageList24
    Set mobjImgFollow1 = objclsQueryPar.img1
    Set mobjImgFollow2 = objclsQueryPar.img2
    
    Set mobjImgFollow3 = objclsQueryPar.img3
    Set mobjImgFollow4 = objclsQueryPar.img4
    Set mobjFilterCmdBar = objclsQueryPar.objFilterCmdBar
    Set mobjFindPati = objclsQueryPar.objPatiIdentify
    Set mobjGetDataContainer = objclsQueryPar.PicContainer
    Set mobjcmdQuery = objclsQueryPar.cmdDo
    
    Set mobjPicFollow = objclsQueryPar.picFollow
    Set mobjPacsMainPicList = objclsQueryPar.picList
    Set mobjtxtAppend = objclsQueryPar.rtpAppend
    Set mobjTab = objclsQueryPar.TabCtl
    Set mobjTim = objclsQueryPar.TimerFunc
    Set mobjTimHistory = objclsQueryPar.TimerHisFunc
    Set mobjPicHistory = objclsQueryPar.PicHisFollow
    Set mobjPicTmp = objclsQueryPar.picTemp
    Set mobjlabPatiInfo = objclsQueryPar.labPatiInfo
    
    Set mobjList = objclsQueryPar.vsfList
    With mobjList
        .AllowUserResizing = flexResizeColumns '用于改遍列宽
        .ExplorerBar = 7 '用于列头拖动和排序
        .SelectionMode = flexSelectionByRow '用于选择整行
        .AllowSelection = False '用于选择整行
        .ScrollTrack = True '滚动条随时更新
        .FixedCols = 1
        .BackColorSel = &HFEE0E2      '&HFECFD2
        .ForeColorSel = RGB(0, 0, 0) '选中行字体黑色
        .ExtendLastCol = True
    End With
    
    mLngOwnerHwnd = objOwner.hwnd
     
    strPath = FormatFilePath(SysRootPath() & "\Apply\QueryScheme\")
    
    If Dir(strPath, vbDirectory) = "" Then
        MkDir (strPath)   '创建文件夹
    End If
    
    Call InitPatiComponent
       
    Call WriteLog("准备初始化自自定义方案", "clsPacsQueryWrap", "Init")
    mObjQuery.Init gcnOracle, UserInfo.ID, strPath
    mObjQuery.LoadQueryScheme glngSys, mlngModule, lngDepartId, lngUserID
    Call WriteLog("初始化自自定义方案完成", "clsPacsQueryWrap", "Init")
End Sub

Public Function RefreshCustomQueryMenu(objQueryMenu As Object, ByRef intValue As Integer, ByRef tabScheme As TabControl, ByVal blnShowLeftTab As Boolean) As String
'自定义查询方案加载到查询菜单,菜单号码预留 conMenu_Manage_CustomQuery * 1000+1 到 conMenu_Manage_CustomQuery * 1000+99
'intvalue 查询方案状态  0 未初始化  ，1 正常  ，2 没有任何有效方案   3：没有已经启用的方案
'blnShowLeftTab 是否显示 左侧tab快速标签
On Error GoTo errhandle
    Dim i As Integer
    Dim intIndex As Integer
    Dim objCurQueryMenu As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim lngID As Long
    Dim cbrMenuBar As CommandBarPopup
    Dim strReturn As String
    
    strReturn = ""
    intValue = 0
    RefreshCustomQueryMenu = ""
    If objQueryMenu Is Nothing Then Exit Function
    If mObjQuery Is Nothing Then Exit Function
    
    lngID = Val(conMenu_Manage_CustomQuery) * 100
    
    Set objCurQueryMenu = objQueryMenu
    
    Call objCurQueryMenu.CommandBar.Controls.DeleteAll

    With objCurQueryMenu.CommandBar.Controls
    
        If mObjQuery.SchemeCount = 0 Then
            intValue = 2
        End If

        Set cbrMenuBar = .Add(xtpControlPopup, Val(conMenu_Manage_CustomQuery) * 100 + 500, "更多")
        cbrMenuBar.ToolTipText = "显示其他已经启用的方案"
        cbrMenuBar.Category = M_STR_MODULE_QUERYMORE
        cbrMenuBar.Parameter = "查询更多"
        cbrMenuBar.Visible = False
        cbrMenuBar.Enabled = False
        cbrMenuBar.BeginGroup = True

        For i = mObjQuery.SchemeCount To 1 Step -1
            lngID = lngID + 1
            If mObjQuery.SchemeInfo(i).IsDefault Or mObjQuery.SchemeInfo(i).IsOften Then
                Set cbrControl = .Add(xtpControlButton, lngID, mObjQuery.SchemeInfo(i).Name, 1)
                intValue = 1
                
                If strReturn <> "" Then strReturn = strReturn & "|"
                strReturn = strReturn & mObjQuery.SchemeInfo(i).Name & "%" & mObjQuery.SchemeInfo(i).SchemeId
            Else
                Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, lngID, mObjQuery.SchemeInfo(i).Name)
                cbrMenuBar.Visible = True
                cbrMenuBar.Enabled = True
            End If
        
            cbrControl.ToolTipText = mObjQuery.SchemeInfo(i).Memo
            cbrControl.Category = M_STR_MODULE_QUERY
            cbrControl.Parameter = mObjQuery.SchemeInfo(i).SchemeId
            cbrControl.Visible = True
            cbrControl.Enabled = True
        Next i

        If intValue = 0 Then
            intValue = 3
        End If
    End With
    
    Call initTabScheme(tabScheme, strReturn, blnShowLeftTab)
    
    Exit Function
errhandle:
    MsgBox "读取查询方案异常-" & err.Description, vbExclamation, gstrSysName
End Function

Public Function DoBackgroundRefresh(ByVal strDepartsId As String) As Boolean
'后台刷新数据
'刷新数据时，只需读取当天申请的，没有报告人和审核人以及处于未审核之前过程的数据
'如果有新数据需要刷新 返回true  否则  false

On Error GoTo errhandle

    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim RsDataShow As ADODB.Recordset
    
    Dim dtBegin As Date
    Dim dtEnd As Date
    
    Dim strAdviceIds As String
    Dim lngCol As Long
    Dim lngRow As Long
    Dim lngSelAdviceID
    
    Dim i As Long
    Dim strColName As String
    
    Dim blFind As Boolean
    Dim intCount As Integer  '需要新显示的检查数量
    Dim blnOnlyRefreshSel As Boolean '是否因为时间范围不包括今天只需要刷新当前选中
    Dim dtNow As Date
    
    blnOnlyRefreshSel = False
    intCount = 0
    DoBackgroundRefresh = False
    lngCol = mobjList.ColIndex("医嘱ID")
    
    dtNow = zlDatabase.Currentdate
    
    dtBegin = mObjQuery.GetFilterValue("系统.开始日期")
    dtEnd = mObjQuery.GetFilterValue("系统.结束日期")
    
    If dtNow <= dtBegin Or dtNow >= dtEnd Then
        blnOnlyRefreshSel = True
    Else
        dtBegin = CDate(Format(dtNow, "yyyy-mm-dd 00:00:00"))
        dtEnd = CDate(Format(dtNow, "yyyy-mm-dd 23:59:59"))
    End If
    
    If Not blnOnlyRefreshSel Then
        strSQL = "select a.ID as ID From 病人医嘱记录 a, 病人医嘱发送 b, 影像检查记录 c " & _
                    " where a.Id = b.医嘱ID and a.相关Id is Null and b.医嘱id=C.医嘱id(+) " & _
                    " and a.开嘱时间 between [1] and [2] " & _
                    " and Instr([3], a.执行科室id)>0 " & _
                    " and (c.报告人 is Null  or c.复核人 is null) " & _
                    " and nvl(b.执行过程,0) < 5"
    
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "后台检查数据刷新", dtBegin, dtEnd, strDepartsId)
        
    End If
    
    '首先添加当前选中行的医嘱ID
    'LSQTODO 是否存在strAdviceIds="0"的情况
    If mobjList.Row > 0 Then
        lngSelAdviceID = Val(mobjList.TextMatrix(mobjList.Row, lngCol))
        strAdviceIds = lngSelAdviceID
    End If

    If Not blnOnlyRefreshSel Then
        If rsData.RecordCount > 0 Then
            While Not rsData.EOF
                blFind = False
                
                '主查询记录集数据没有找到时，且列表中数据没有找到时，才从数据库中查找加载
                Call mrsData.Find("医嘱ID=" & rsData!ID, , , 1)
                If mrsData.AbsolutePosition > adPosUnknown Then blFind = True
                
                If mobjList.FindRow(rsData!ID, 1, lngCol) < 0 And Not blFind Then
                    If Len(strAdviceIds) > 0 Then strAdviceIds = strAdviceIds & ","
                    strAdviceIds = strAdviceIds & rsData!ID
                    intCount = intCount + 1
                End If
                
                '138712 strAdviceIds是构成后面 f_str2list 的入参 超过4000后报错，而且实际情况一般不会超过4000  超过4000也没必要全部显示。
                If Len(strAdviceIds) > 3900 Then
                    DoBackgroundRefresh = True
                    GoTo NextStep
                End If
                
                rsData.MoveNext
            Wend
            DoBackgroundRefresh = True
        End If
    End If
    
    If strAdviceIds = "" Then Exit Function
NextStep:

    mstrTemp_新增医嘱ID = strAdviceIds
    strSQL = "select column_value as 医嘱ID from table(f_str2list([" & CONST_EXT_PAR_SYNCNEWADVICE & "])) "
    
    Set rsData = mObjQuery.ExecuteWithLink(strSQL, True, 1)
    
    
    '此时的 rsData：当日，已登记  未在记录集 和 列表中出现的 所有医嘱ID+当前选中的医嘱ID 最后一个需要判断是否在列表中存在
    Call UpdateRow(rsData, lngSelAdviceID, SyncDataType.rsShow)
    On Error Resume Next
    For i = intCount To 1 Step -1
        mobjList.TextMatrix(i, 0) = "*"
    Next
    
    Exit Function
errhandle:
    MsgBox "DoBackgroundRefresh,错误信息:" & err.Description, vbExclamation, gstrSysName
End Function

Public Function ExecuteWithLink(ByVal strSQL As String) As Boolean
'收藏功能
On Error GoTo errH
    Dim RowIndex As Long
    Dim lngH As Long
    
    Set mrsData = mObjQuery.ExecuteWithLink(strSQL)
    
    If Not mrsData Is Nothing Then
        If mrsData.RecordCount > 0 Then
            Set mrsDataShow = GetFilterFromQuickFilter(mrsData)
            Set mrsDataShow = mObjQuery.DataConvert(mrsDataShow, mlngSchemeNo)
            If Not mrsDataShow Is Nothing Then
                mblExitSelChange = True
                Set mobjList.DataSource = mrsDataShow
                mblExitSelChange = False
                '列统计
                Call ColStatistics
            End If
        End If
    Else
        Call ColStatistics(True)
    End If
    
    Call SetSelectRow(False, True, False, 0)
    Call LoadListHeadCfg
    
    If mobjList.TopRow = mobjList.BottomRow Then Exit Function
    
    lngH = mObjOwner.TextHeight("字") + 120
    For RowIndex = mobjList.TopRow To mobjList.BottomRow
        mobjList.RowHeight(RowIndex) = lngH
        Call RefreshRowRelation(RowIndex)
    Next
    
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "提取收藏数据异常" & vbCrLf & err.Description
End Function

Public Function ExecuteMenu(ByVal lngNo As Long, Optional ByVal blnValidTime As Boolean = False, Optional ByVal dtStartTime As Date = Empty, Optional ByVal dtEndTime As Date = Empty) As String
On Error GoTo errH
'本过程大多数处理有先后顺序，改动需谨慎
'顺序 ： 保存参数，更新方案号，加载参数 ，设置列表，界面布局刷新
'blById 是否直接通过方案号执行
    Dim i As Long
    Dim intDays As Integer
    Dim dtStart As Date
    Dim intQueryTime As Integer
    
    ExecuteMenu = ""
    Call SaveLocalPara
    
    If mlngSchemeNo > 0 And mintShowType = 0 Then Call SaveShemeCustomCfg(mlngSchemeNo)
    
    '变量初始化
    mTPatiIdentifyInfo.strFindItems = ""
    mTPatiIdentifyInfo.strLocateItems = ""
    mTPatiIdentifyInfo.blShowPatiIdentify = False
    mstrSelQueueRooms = ""
    
    ReDim mColConfig(0)
    
    mlngSchemeNo = lngNo
    
    Call GetLocalPara
    
    Call mObjQuery.ChangeCurScheme(mlngSchemeNo)
    Set mSqlScheme = mObjQuery.GetSqlScheme(mlngSchemeNo)
    
    
    mObjQuery.FontSize = 9
    Call mObjQuery.EmbedFindWindow(mobjGetDataContainer, mobjcmdQuery)
    mObjQuery.FontSize = mObjOwner.FontSize
    Call mObjQuery.EmbedSize(mobjGetDataContainer)
    
    Call LoadShemeCustomCfg(mlngSchemeNo)
    
    mstrLoadMsg = ""
    If Not RefreshQueryWindow(mlngSchemeNo) Then
        ExecuteMenu = mstrLoadMsg
        Exit Function
    End If
    
    Call RefreshTabBottom(mSqlScheme, mobjTab)

    mobjList.BackColorSel = mSqlScheme.BackColor
    mobjList.ForeColorSel = mSqlScheme.ForeColor
    mblnSelRowTransparent = mSqlScheme.SelRowTransparent
   
    '切换方案后，恢复默认录入条件
    Call mObjQuery.EmbedConditionRestore(blnValidTime, dtStartTime, dtEndTime)
    Call mObjQuery.RefreshCfgFontSize(mObjOwner.FontSize)
    
    Call ExecuteQuery(C_QUERY_首次刷新)
    mobjFindCmd.Enabled = IsPatiFilterCanFind(mobjFindPati.GetCurCard.接口序号 > 0, mobjFindPati.GetCurCard.名称)
    If Not mblnDefaultLocate Then mblnDefaultLocate = Not mobjFindCmd.Enabled
    RaiseEvent OnLocateBackColor
    
    Exit Function
errH:
    MsgBox "执行查询异常" & err.Description, vbOKOnly, gstrSysName
End Function

Private Sub SaveLocalPara()
'设置本地参数
    Dim i As Integer
     
    If mQuickFIlterSaveDict.Count > 0 Then
        For i = 0 To mQuickFIlterSaveDict.Count - 1
            Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & "快速过滤项目", mQuickFIlterSaveDict.Keys(i), mQuickFIlterSaveDict.Items(i))
        Next
    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & mlngSchemeNo & "\", "列表检查信息高度设置", mlngMove)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\", "附加信息选中标签", GetTabSelectName(True))
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & mlngSchemeNo & "\", "排序列", mlngSortCol)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & mlngSchemeNo & "\", "排序方向", mintSortOrder)
    
    Call SaveLocalPara_PatiIdentify
End Sub

Private Sub SaveLocalPara_PatiIdentify()
'设置Pati控件相关参数
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & mlngSchemeNo & "\", "查找项目", mTPatiIdentifyInfo.strFindItem)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & mlngSchemeNo & "\", "定位项目", mTPatiIdentifyInfo.strLocateItem)
End Sub

Private Sub SaveShemeCustomCfg(ByVal LngSchemeNo As Long)
'保存个性化配置
On Error GoTo errH
    Dim strSQL As String
    Dim objCboControl As CommandBarComboBox
    strSQL = "Zl_影像查询_个性化配置(" & UserInfo.ID & "," & LngSchemeNo & ",'" & "" & "%" & mstrSchemeCfg.strFilterCfg & "','" & mstrSchemeCfg.strListCfg & mstrSchemeCfg.strHistoryCfg & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "保存个性化配置")
    
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[保存个性化配置]" & vbCrLf & err.Description
End Sub

Private Sub GetLocalPara()
'获取本地参数（注册表参数）
    mlngSortCol = Val(GetSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & mlngSchemeNo & "\", "排序列", 0))
    mintSortOrder = Val(GetSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & mlngSchemeNo & "\", "排序方向", -1))
    mlngMove = Val(GetSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & mlngSchemeNo & "\", "列表检查信息高度设置", 0))
    mTPatiIdentifyInfo.strFindItem = GetSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & mlngSchemeNo & "\", "查找项目")
    mTPatiIdentifyInfo.strLocateItem = GetSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & mlngSchemeNo & "\", "定位项目")
    
End Sub

Private Sub LoadShemeCustomCfg(ByVal LngSchemeNo As Long)
'根据用户ID/方案ID加载 个性化配置,检查配置是否符合要求，不符合则对应配置自动设置为空
On Error GoTo errH
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strTmp As String
    Dim intTMP As Integer
    Dim strName As String
    Dim strValue As String
    
    strSQL = "select 过滤配置,列表配置 from 影像查询特性  where 用户ID=[1] and 查询方案ID =[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "加载查询个性化配置", UserInfo.ID, LngSchemeNo)

    Set mQuickFIlterSaveDict = New Dictionary
    
    For i = 1 To mSqlScheme.FilterCfgCount
        strName = mSqlScheme.FilterCfg(i).Name
        If mSqlScheme.GetFilterExtValue(strName, CONST_EXT_FILTERKEY_SAVEWAY) = CONST_EXT_FILTERKEY_SAVEWAYVALUE Then
            Call mQuickFIlterSaveDict.Add(strName, "")
            strValue = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & "clsPacsQueryWrap" & "\" & "快速过滤项目", strName, CONST_EXT_FILTERKEY_SAVEWAYDEFALT)
            mQuickFIlterSaveDict.Item(strName) = strValue
        End If
    Next
    '初始化查询参数
    
    On Error Resume Next
    
    If rsTemp.RecordCount = 1 Then
        mstrSchemeCfg.strFilterCfg = Split(rsTemp!过滤配置, "%")(1)
        
        strTmp = rsTemp!列表配置
        intTMP = InStr(strTmp, "[历史]")
        If intTMP > 0 Then
            mstrSchemeCfg.strListCfg = Split(strTmp, "[历史]")(0)
            mstrSchemeCfg.strHistoryCfg = Mid(strTmp, intTMP)
        Else
            mstrSchemeCfg.strListCfg = rsTemp!列表配置
            mstrSchemeCfg.strHistoryCfg = ""
        End If
        
        '兼容性处理，由35.100 升级到35.110 之后的用户
        '列表配置字段格式发生变化，所以若格式是处理前的格式，直接清空，需要重新设置列头个性化配置
        
        If mstrSchemeCfg.strListCfg <> "" Then
            strTmp = Split(mstrSchemeCfg.strListCfg, "|")(0)
            If UBound(Split(strTmp, ",")) <> 3 Then mstrSchemeCfg.strListCfg = ""
        
        End If
    
    Else
        mstrSchemeCfg.strFilterCfg = ""
        mstrSchemeCfg.strListCfg = ""
    End If

    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[加载个性化配置]" & vbCrLf & err.Description
End Sub

Private Function RefreshQueryWindow(ByVal LngSchemeNo As Long) As Boolean
'根据方案改变快速过滤界面,加载快速过滤菜单，根据个性化参数加载菜单选中项
On Error GoTo errH
    RefreshQueryWindow = False
    
    Call InitPatiIdentify
    If Not DoQuickFilterMenu Then
        MsgBox "创建快速过滤菜单出现错误，请检查快速过滤配置", vbOKOnly, gstrSysName
    End If
    Call InitFunctionFollow
    
    RefreshQueryWindow = True
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[加载主窗体]" & vbCrLf & err.Description
End Function


Private Function InitCardType(ByVal strCardNames As String) As String
'按指定格式初始化卡类型
On Error GoTo errH
    Dim i As Integer
    Dim aryKindInfo() As String
    Dim strKinds As String
    
    aryKindInfo = Split(strCardNames, ";")
    
    strKinds = ""
    For i = 0 To UBound(aryKindInfo) - 1
        If strKinds <> "" Then strKinds = strKinds & ";"
        strKinds = strKinds & aryKindInfo(i) & "|" & aryKindInfo(i) & "|-1"
    Next i
    
    InitCardType = strKinds & ";"
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[初始化定位数据]" & vbCrLf & err.Description
End Function

Private Sub RefreshCbrQuickFilterALL()
'刷新所有动态快速过滤选中情况，也就是初始化
On Error GoTo errH
    Dim i As Integer, j As Integer, k As Integer
    Dim lngID As Long

    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        If mTQuickFilterState.TCmdState(i).blIsRelationFilter Then

            Call RefreshCbrQuickFilter(i, True)
            
        End If
    Next

    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[刷新全部快速过滤菜单]" & vbCrLf & err.Description
End Sub

Private Sub RefreshCbrQuickFilter(ByVal lngIndex As Long, ByVal blInit As Boolean)
'更新自定义及快速过滤菜单，参数： 菜单ID，触发菜单选中项"123,456,789"这种形式
'根据触发菜单信息查找自定义菜单信息，删除原来的自定义菜单子项，然后重新生成。
'blInit 是否初始化
'lngIndex表示需要改变的菜单
On Error GoTo errH
    Dim lngRelationID As Long '自定义菜单ID
    Dim i As Long, j As Long
    Dim ObjPopMenu As CommandBarPopup
    Dim cbrControl As CommandBarControl

    Dim rsTemp As Recordset
    Dim intMenuCount As Integer
    Dim strSQL As String, strTmp As String
    Dim strItemName As String, strMeneShow As String, strMenuValue As String

    '''清除已有菜单子项开始
    Call LockWindowUpdate(mLngOwnerHwnd)

    lngRelationID = mTQuickFilterState.TCmdState(lngIndex).lngID
    Set ObjPopMenu = mobjFilterCmdBar.FindControl(, lngRelationID)

    ObjPopMenu.CommandBar.Controls.DeleteAll
    '''清除已有菜单子项结束

    Call LockWindowUpdate(0)

    '创建菜单
    strSQL = mTQuickFilterState.TCmdState(lngIndex).strMenuSql

    If mobjSqlParseFilter Is Nothing Then Set mobjSqlParse = New clsSqlParse
    Call mobjSqlParseFilter.Init(strSQL)
    strSQL = mobjSqlParseFilter.GetQuerySql

    Set rsTemp = ExecuteCore(strSQL, "快速过滤获取项目", mobjSqlParseFilter.ParValues)

    ''''判断是否子菜单显示数据是过滤数据,跟 rsTemp.Fields.Count 的值有关
    ''是，距离 影像类别-部位分组     对于部位分组来说，显示的是分组，参与过滤的是部位。rsTemp.Fields.Count 应该是 2
    ''否，rsTemp.Fields.Count 应该是1 菜单显示的内容就是参与过滤的内容
    If rsTemp.RecordCount = 0 Then
        ReDim mTQuickFilterState.TCmdState(lngIndex).cmdItem(0)
        mTQuickFilterState.TCmdState(lngIndex).intItemCount = 0
        ObjPopMenu.Enabled = False
        Exit Sub
    End If

    '获取选择菜单配置（用于恢复选择情况）
    strTmp = mTQuickFilterState.TCmdState(lngIndex).strRelationChooseMenu
    
    If rsTemp.Fields.Count = 1 Then
    '''1 表示过滤内容就是显示内容
        rsTemp.MoveFirst
        ReDim mTQuickFilterState.TCmdState(lngIndex).cmdItem(rsTemp.RecordCount)
        
        For i = 1 To rsTemp.RecordCount
        
            strItemName = NVL(rsTemp.Fields(0).value)
            strMenuValue = strItemName
            If InStr(1, strItemName, "-") = 0 Then
                strMeneShow = strItemName
                strMenuValue = strItemName
            Else
                strMeneShow = Mid(strItemName, InStr(1, strItemName, "-") + 1, Len(strItemName) - InStr(1, strItemName, "-"))
                strMenuValue = Mid(strItemName, 1, InStr(1, strItemName, "-") - 1)
            End If
        
            Set cbrControl = ObjPopMenu.CommandBar.Controls.Add(xtpControlButton, lngRelationID - 100 + i, strMeneShow)
            cbrControl.Parameter = strMenuValue

            mTQuickFilterState.TCmdState(lngIndex).cmdItem(i).blChoose = IIf(InStr(";" & strTmp & ";", ";" & strMeneShow & ";") > 0, True, False)
            mTQuickFilterState.TCmdState(lngIndex).cmdItem(i).intItemIndex = i
            mTQuickFilterState.TCmdState(lngIndex).cmdItem(i).strName = strMeneShow
            mTQuickFilterState.TCmdState(lngIndex).cmdItem(i).strFilterValue = strMenuValue
            cbrControl.CloseSubMenuOnClick = False
            If i <> rsTemp.RecordCount Then rsTemp.MoveNext
        Next
        mTQuickFilterState.TCmdState(lngIndex).intItemCount = rsTemp.RecordCount
    Else
    '''rsTemp.Fields.Count <> 1(=2) 表示第一个字段是过滤内容 第二个字段是显示内容
        rsTemp.MoveFirst
        intMenuCount = 1

        ReDim mTQuickFilterState.TCmdState(lngIndex).cmdItem(0)
        mTQuickFilterState.TCmdState(lngIndex).intItemCount = 0

        For i = 1 To rsTemp.RecordCount
            strMeneShow = rsTemp.Fields(1).value


            ReDim Preserve mTQuickFilterState.TCmdState(lngIndex).cmdItem(mTQuickFilterState.TCmdState(lngIndex).intItemCount + 1)

            Set cbrControl = ObjPopMenu.CommandBar.Controls.Add(xtpControlButton, lngRelationID - 100 + intMenuCount, strMeneShow)

            cbrControl.Parameter = NVL(rsTemp.Fields(0).value)

            mTQuickFilterState.TCmdState(lngIndex).cmdItem(intMenuCount).blChoose = IIf(InStr(";" & strTmp & ";", ";" & strMeneShow & ";") > 0, True, False)

            mTQuickFilterState.TCmdState(lngIndex).cmdItem(intMenuCount).intItemIndex = i
            mTQuickFilterState.TCmdState(lngIndex).cmdItem(intMenuCount).strName = strMeneShow
            mTQuickFilterState.TCmdState(lngIndex).cmdItem(intMenuCount).strFilterValue = NVL(rsTemp.Fields(0).value)
            mTQuickFilterState.TCmdState(lngIndex).intItemCount = intMenuCount

            intMenuCount = intMenuCount + 1
            cbrControl.CloseSubMenuOnClick = False

             If i <> rsTemp.RecordCount Then rsTemp.MoveNext
        Next
    End If
    Call GetQuickFilterSQLPar(lngIndex)
'
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[刷新快速过滤菜单]" & vbCrLf & err.Description
'    Resume
End Sub

Private Sub InitFunctionFollow()
'初始化功能跟随
On Error GoTo errH
    Dim lngIconW As Long
    Dim lngIconH As Long

    lngIconW = 360
    lngIconH = 360

    Call mobjImgFollow1.Move(0, 0, lngIconW, lngIconH)
    Call mobjImgFollow2.Move(lngIconW, 0, lngIconW, lngIconH)
    Call mobjImgFollow3.Move(2 * lngIconW, 0, lngIconW, lngIconH)
    Call mobjImgFollow4.Move(3 * lngIconW, 0, lngIconW, lngIconH)

    mobjImgFollow1.Visible = True
    mobjImgFollow2.Visible = True
    mobjImgFollow3.Visible = True
    mobjImgFollow4.Visible = True

    If Not mSqlScheme.UseFuncFollow Then mobjPicFollow.Visible = False

    Exit Sub
errH:
    MsgBox "[初始化功能跟随]" & vbCrLf & err.Description, vbOKOnly, gstrSysName
End Sub

Private Function GetQuickFilterSQLPar(ByVal lngIndex As Long) As String
'获取与组合自定义快速过滤条件，类似"头,双眼,甲状腺"这种，用于过滤
'参数： 过滤信息index，菜单ID
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngID As Long, lngIDEnd As Long

    Dim strTmp As String
    Dim cbrPopControl As CommandBarControl
    Dim objControl As CommandBarControl
    Dim blChooseOne As Boolean
    Dim strAry() As String

    lngID = 100 * lngIndex
    lngIDEnd = mTQuickFilterState.TCmdState(lngIndex).intItemCount
    blChooseOne = False

    For i = 1 To lngIDEnd
        Set objControl = mobjFilterCmdBar.FindControl(, lngID - 100 + i, , True)

        If mTQuickFilterState.TCmdState(lngIndex).cmdItem(i).blChoose Then
            blChooseOne = True

            If Len(strTmp) > 0 Then strTmp = strTmp & ","
            If InStr("," & strTmp & ",", "," & objControl.Parameter & ",") = 0 Then strTmp = strTmp & objControl.Parameter

        End If

    Next

    If blChooseOne = False Then
        For i = 1 To lngIDEnd
            Set objControl = mobjFilterCmdBar.FindControl(, lngID - 100 + i, , True)

            If Len(strTmp) > 0 Then strTmp = strTmp & ","
            If InStr("," & strTmp & ",", "," & objControl.Parameter & ",") = 0 Then strTmp = strTmp & objControl.Parameter

        Next
    End If

    strAry = Split(strTmp, ",")
    strTmp = ""
    For i = 0 To UBound(strAry)
        For j = i + 1 To UBound(strAry)
            If strAry(i) = strAry(j) Then strAry(j) = ""
        Next

        If strAry(i) <> "" Then
            If Len(strTmp) > 0 Then strTmp = strTmp & ","
            strTmp = strTmp & strAry(i)
        End If
    Next

    strTmp = Replace(strTmp, ",,,,,,", ",")
    strTmp = Replace(strTmp, ",,,,,", ",")
    strTmp = Replace(strTmp, ",,,,", ",")
    strTmp = Replace(strTmp, ",,,", ",")
    strTmp = Replace(strTmp, ",,", ",")

    mTQuickFilterState.TCmdState(lngIndex).strRelationValueForVBSFilter = Trim(strTmp)

    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[获取快速过滤条件]" & vbCrLf & err.Description
End Function

Public Function ExecuteQuery(ByVal strExecuteType As String) As Boolean
'执行过滤、刷新、查找功能
On Error GoTo errH
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim objCboControl As CommandBarComboBox
    Dim strTmp As String
    Dim i As Long, j As Long
    Dim lngAdviceId As Long
    Dim lngHeight As Long
    Dim t1 As Long
    Dim dtStartTmp As Date
    Dim dtEndTmp As Date
    Dim lngDefaultTime As Long
    Dim value As Variant
    Dim strValue As String
    Dim lngH As Long
    Dim lngRow As Long
    Dim rsData As Recordset
    
'    t1 = GetTickCount
'    Debug.Print "ExecuteQuery 1:" & GetTickCount - t1
    '执行刷新前获取之前的选中行医嘱ID,用于刷新后恢复选中
    lngAdviceId = GetSelectRowAdviceID
    lngHeight = mobjList.Row - mobjList.TopRow
    mObjOwner.MousePointer = 13
    If strExecuteType = C_QUERY_更多过滤 Then
        
        mTqueryType = 过滤
        Set rsData = mObjQuery.EmbedFilter(mObjOwner, True)
        
        If rsData Is Nothing Then
            mObjOwner.MousePointer = 0
            Exit Function
        End If
        
    ElseIf strExecuteType = C_QUERY_数据检索 Or strExecuteType = C_QUERY_刷新 Or strExecuteType = C_QUERY_首次刷新 Then
        
        If strExecuteType = C_QUERY_数据检索 Then
            mobjFindPati.Text = ""
        End If
        
        mTqueryType = 刷新
        Set rsData = mObjQuery.EmbedFilter(mObjOwner)
        
    ElseIf strExecuteType = C_QUERY_查找 Then
        mTqueryType = 查找
        strValue = mobjFindPati.Text

        If mstrFindName = "系统.病人ID" Then
            value = mPatiID
        Else
            value = strValue
        End If

        Set rsData = mObjQuery.ExecuteWithAttach(mstrFindName, value)
    End If
    
    '149208 复制后的记录集后面才能修改里面的内容
    If Not rsData Is Nothing Then
        Set mrsData = zlDatabase.CopyNewRec(rsData)
    End If
    
    mdtStartDate = mObjQuery.GetFilterValue("系统.开始日期")
    mdtEndDate = mObjQuery.GetFilterValue("系统.结束日期")
    
    For lngRow = 1 To mobjList.Rows - 1
         mobjList.RowData(lngRow) = ""
    Next
            
    If UBound(mColConfig) = 0 Then
        ReDim mColConfig(mrsData.Fields.Count - 1)

        '首先判断是否存在个性化设置，若存在，则加载
        Call DoListConfig(mrsData, mColConfig)
    End If
    
    If App.LogMode = 0 Then
        If mrsData Is Nothing Then MsgBox "mrsData nothing "
    End If
    
    If Not mrsData Is Nothing Then
        '145420 pati控件查找增加实时过滤功能
        If strExecuteType = C_QUERY_查找 And Not mSqlScheme.FindRealTimeFilter Then
            Set mrsDataShow = CopyRecordSet(mrsData)
        Else
            Set mrsDataShow = GetFilterFromQuickFilter(mrsData)
        End If
        If mrsData.RecordCount > 0 Then
            
            Set mrsDataShow = mObjQuery.DataConvert(mrsDataShow, mlngSchemeNo)
 
            If Not mrsDataShow Is Nothing Then
                
                mblExitSelChange = True
                Set mobjList.DataSource = mrsDataShow
                mblExitSelChange = False
            
                '列统计
                Call ColStatistics
            Else
                Call ColStatistics(True)
            End If
            
            Call LoadListHeadCfg
    
            Call ResetSort(mlngSortCol, mintSortOrder)
        Else
            mblExitSelChange = True
            Set mobjList.DataSource = mrsData
            mblExitSelChange = False
            '列统计
            Call ColStatistics(True)
            
            Call LoadListHeadCfg
        End If
    End If
    
    If App.LogMode = 0 Then
        If mrsDataShow Is Nothing Then MsgBox "mrsDataShow nothing "
    End If
    
'    Debug.Print "ExecuteQuery 3:" & GetTickCount - t1
    If strExecuteType = C_QUERY_刷新 Then
        Call SetSelectRow(True, False, False, lngHeight)
    ElseIf strExecuteType = C_QUERY_首次刷新 Then
        'SetSelectRow 返回false 需要单独触发selchange
        If Not SetSelectRow(False, True, False, 0) Then Call mobjList_SelChange
    Else
        Call SetSelectRow(False, True, False, lngHeight)
    End If
    
    If mobjList.TopRow <> mobjList.BottomRow Then
        lngH = mObjOwner.TextHeight("字") + 120
        For i = mobjList.TopRow To mobjList.BottomRow
            mobjList.RowHeight(i) = lngH
        Next
    End If

    mObjOwner.MousePointer = 0
'    Debug.Print "ExecuteQuery 4:" & GetTickCount - t1
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[执行主查询]" & vbCrLf & err.Description
   Resume
End Function

Private Function GetSelectRowAdviceID() As Long
'根据当前列选中行获取医嘱ID
On Error GoTo errH
    GetSelectRowAdviceID = 0
    If mobjList.Rows < 1 Or mobjList.Row < 1 Then Exit Function
    
    GetSelectRowAdviceID = Val(mobjList.TextMatrix(mobjList.Row, mobjList.ColIndex(mstrListKeyCol)))
    
    Exit Function
errH:
    GetSelectRowAdviceID = 0
End Function

Private Function GetFilterFromQuickFilter(ByRef rsData As Recordset, Optional ByVal isUpdateRow As Boolean = False) As Recordset
'获取快速过滤条件、并且进行过滤,返回过滤后的记录集
'对 rsdata 操作，运行后不会改变 rsdata,返回经过过滤后的全新记录集
'函数实现简要说明：首先进行固定快速过滤的过滤，获取过滤后的纪律及，然后进行脚本方式过滤。
'脚本方式过滤每个过滤菜单都会新copy一个记录集过滤
On Error GoTo errH
    Dim objControl As CommandBarControl
    Dim strFilter As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim blChooseOne As Boolean '某个菜单是否有过滤项被选中，若没有相当于不过滤
    Dim strFilterTmp As String
    Dim strFilterField As String
    Dim intCustomeFilter As Integer   '是否有自定义快速过滤
    Dim strOldFilter As String  '用于恢复原始记录集
    
    If rsData.RecordCount = 0 Then
        Set GetFilterFromQuickFilter = CopyRecordSet(rsData)
        Exit Function
    End If
    strOldFilter = ""
    If rsData.Filter <> 0 Then strOldFilter = rsData.Filter
    
    intCustomeFilter = 0
    strFilter = ""
    
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
    
        If mTQuickFilterState.TCmdState(i).intFilterWay = EFilterType.未知 Then
            mTQuickFilterState.TCmdState(i).intFilterWay = EFilterType.不过滤
            
            For j = 0 To UBound(mColConfig)
            
                If mTQuickFilterState.TCmdState(i).strName = mColConfig(j).strName Then
                    mTQuickFilterState.TCmdState(i).intFilterWay = EFilterType.普通
                    Exit For
                End If
                
            Next
        End If
        
        If mTQuickFilterState.TCmdState(i).intFilterWay = EFilterType.不过滤 Or mTQuickFilterState.TCmdState(i).intFilterWay = EFilterType.未知 Then
            GoTo NextStep
        End If
        
        blChooseOne = False
        strFilterTmp = ""
        Set objControl = mobjFilterCmdBar.FindControl(, i * 100)

        If mTQuickFilterState.TCmdState(i).intFilterWay = EFilterType.普通 Then
            For j = 1 To mTQuickFilterState.TCmdState(i).intItemCount
                strFilterField = mTQuickFilterState.TCmdState(i).cmdItem(j).strFilterValue

                If Not mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose Then
                '未被选中
                    If Len(strFilterTmp) = 0 Then
                        strFilterTmp = strFilterTmp & mTQuickFilterState.TCmdState(i).strName & " <> '" & strFilterField & "' "
                    Else
                        strFilterTmp = strFilterTmp & " and " & mTQuickFilterState.TCmdState(i).strName & " <> '" & strFilterField & "' "
                    End If
                Else
                    blChooseOne = True
                End If

            Next
        ElseIf mTQuickFilterState.TCmdState(i).intFilterWay = EFilterType.脚本 Then
            intCustomeFilter = intCustomeFilter + 1
        End If
        
        '只有被选中了过滤项，才加入到过滤条件中
        If blChooseOne And Len(strFilterTmp) > 0 Then
            If Len(strFilter) = 0 Then
                strFilter = strFilter & strFilterTmp
            Else
                strFilter = strFilter & " and " & strFilterTmp
            End If
        End If
NextStep:
    Next
    
    If isUpdateRow Then
        If InStr(strFilter, "and 检查过程 <> '已检查'  and") > 0 Then
            strFilter = Replace(strFilter, "检查过程 <> '已检查'  and", "")
        Else
            If InStr(strFilter, "检查过程 <> '已检查'  and") > 0 Then
                strFilter = Replace(strFilter, "检查过程 <> '已检查'  and", "")
            ElseIf InStr(strFilter, "and 检查过程 <> '已检查'") > 0 Then
                strFilter = Replace(strFilter, "and 检查过程 <> '已检查'", "")
            Else
                strFilter = Replace(strFilter, "检查过程 <> '已检查'", "")
            End If
        End If
    End If
    rsData.Filter = strFilter
    '没有自定义快速过滤可以现在退出
    If intCustomeFilter = 0 Then
        Set GetFilterFromQuickFilter = CopyRecordSet(rsData)
        rsData.Filter = strOldFilter
        Exit Function
    End If

    Dim strVBS As String
    Dim rstVBS As Recordset
    Dim rsTmp() As Recordset

    '遍历每个自定义快速过滤条件,根据菜单选中情况执行VBS脚本
    ReDim rsTmp(intCustomeFilter)
    
    j = 0

    'vbs脚本过滤每次需要CopyRecordSet 之前的记录集
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        If mTQuickFilterState.TCmdState(i).intFilterWay = EFilterType.脚本 Then
            j = j + 1
            If j = 1 Then
                Set rsTmp(0) = CopyRecordSet(rsData)
            Else
                Set rsTmp(j - 1) = CopyRecordSet(rsTmp(j - 2))
            End If

            Set rsTmp(j - 1) = zl9PacsQuery.ExecuteScript(mTQuickFilterState.TCmdState(i).strCustomScript, rsTmp(j - 1), mTQuickFilterState.TCmdState(i).strRelationValueForVBSFilter)
        End If
    Next

    Set GetFilterFromQuickFilter = rsTmp(intCustomeFilter - 1)
    rsData.Filter = strOldFilter
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[执行数据过滤]" & vbCrLf & err.Description
    Resume
End Function

Public Sub RefreshRowRelation(ByVal lngRow As Long)
'单独刷新某一行，同时触发行关联处理
On Error GoTo errH
    Dim value As String
    Dim i As Long

    '首先应初始化本行全部行关联设置，（颜色，图标）
    With mobjList
        If .Rows < 2 Then Exit Sub
        .Cell(flexcpBackColor, lngRow, 1, lngRow, .Cols - 1) = vbWhite
        .Cell(flexcpForeColor, lngRow, 1, lngRow, .Cols - 1) = vbBlack '黑色字体
        Set .Cell(flexcpPicture, lngRow, 1, lngRow, .Cols - 1) = Nothing
    End With

    For i = 1 To mobjList.Cols - 1
        value = mobjList.TextMatrix(lngRow, i)
        Call RowRelationConvert(lngRow, i, value)
    Next
    Call SetPatiTyprColor(lngRow, True)
    
    mobjList.Cell(flexcpAlignment, lngRow, 1, lngRow, mobjList.Cols - 1) = flexAlignLeftCenter
    
    mobjList.RowHeight(lngRow) = mObjOwner.TextHeight("字") + 120
    mobjList.RowData(lngRow) = C_BLN_行关联刷新完成
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[RefreshRowRelation]" & vbCrLf & err.Description
    
End Sub

Private Sub ColStatistics(Optional ByVal blClear As Boolean = False)
'列统计处理
On Error GoTo errhandle
    Dim i As Long, j As Long, k As Long
    Dim lngColICount As Long '待统计的数据序号
    Dim lngColIndex() As Integer
    
    Dim strColName As String '"需要处理跟踪列的列" 类似： "影像质量;检查过程;是否阳性"
    Dim strColNameAll As String
    Dim strStateBarInfo As String '最终通过事件传递的列统计信息
    Dim strInfoTmp As String
    
    Dim objTColTotalInfo() As TColTotalInfo
    Dim DictColTotal() As Dictionary
    
    Dim value As Variant
    
    If blClear Then
        RaiseEvent OnColStatistics("")
        Exit Sub
    End If
        
    strColNameAll = ""
    strColName = ""
    lngColICount = 0
    
    If mSqlScheme Is Nothing Then
        RaiseEvent OnColStatistics("")
        Exit Sub
    End If
     
    With mSqlScheme
        For i = 1 To .ShowCfgCount
            If .ShowCfg(i).IsTotal Then
                lngColICount = lngColICount + 1
                strColNameAll = strColNameAll & .ShowCfg(i).Name & ";"
                ReDim Preserve DictColTotal(lngColICount)
                Set DictColTotal(lngColICount) = New Dictionary
            End If
        Next
    End With

    With mobjList
        
        strStateBarInfo = ""
        strColName = ""
        
        For k = 1 To lngColICount
        
            '获取字段名
            strColName = Split(strColNameAll, ";")(k - 1)

            For i = 1 To .Rows - 1
                value = .TextMatrix(i, .ColIndex(strColName))
                
                If Not IsNull(value) Then
                    If Len(Trim(value)) > 0 Then
                        If DictColTotal(k).Exists(value) Then
                             DictColTotal(k).Item(value) = DictColTotal(k).Item(value) + 1
                        Else
                            Call DictColTotal(k).Add(value, 1)
                        End If
                        
                    End If
                End If
            Next
            
            strInfoTmp = "[" & strColName & "]:"
            For j = 1 To DictColTotal(k).Count
                strInfoTmp = strInfoTmp & DictColTotal(k).Keys(j - 1) & ":" & DictColTotal(k).Item(DictColTotal(k).Keys(j - 1)) & " "
            Next
            
            If Len(strStateBarInfo) > 0 Then strStateBarInfo = strStateBarInfo & "|"
            strStateBarInfo = strStateBarInfo & strInfoTmp
            
            RaiseEvent OnColStatistics(strStateBarInfo)
                
        Next
   
    End With
    
    
    RaiseEvent OnColStatistics(strStateBarInfo)
    Exit Sub
errhandle:
    err.Raise -1, "clsPacsQueryWrap", "[列统计]" & vbCrLf & err.Description
End Sub

Private Function LoadListHeadCfg() As Boolean
'根据个性化参数刷新列表配置(排序、宽度、是否显示)
On Error GoTo errH
    

'    功能简介：调整首行对齐方式 判断是否匹配  若不配直接恢复默认（不考虑用户隐藏，宽度，顺序）

    mobjList.Cell(flexcpAlignment, 0, 0, 0, mobjList.Cols - 1) = flexAlignLeftCenter
    
    '主要逻辑：
    '遍历每一列，跟缓存对应位置比较，
    '宽度、是否显示  根据缓存设置
    '顺序：如果已经顺序调整，查找正确顺序的列，调换到指定位置，否则不改变顺序
    
    Call DoListColsOrder(mobjList, mColConfig)
    
    '根据最大行号修改第一列的宽度
    Call SetListColWidth

    If mobjList.Rows > 1 Then Exit Function

    Call DoFunctionFollow(False)

    Set mobjCurStudyInfo = GetNullAdviceInf

    RaiseEvent OnListRowSelClear
        
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[LoadListHeadCfg]" & vbCrLf & err.Description
End Function

Private Sub ResetSort(ByVal lngCol As Long, ByVal lngWay As Long)
'重置排序
'排序参考列功能说明： 参考列的mobjList.ColSort(n)必定是升序，1,3,5,7这种
On Error GoTo errH
    Dim i As Long
    Dim RowIndex As Long
    Dim LngColTmp As Long
    Dim blnstr As Boolean
    Dim lngH As Long
    Dim strColName As String
    
    blnstr = True
    If mobjList.Rows <= 1 Then Exit Sub
    
    strColName = mobjList.ColKey(lngCol)
    mblExitSelChange = True
    
    If mintSortOrder <> -1 Then
    'mintSortOrder=-1 说明还未设置排序列，不进行排序相关操作
        '排序参考列
        LngColTmp = mobjList.ColIndex(GetColSort(strColName))
        
        If mobjList.Col <> LngColTmp Then
            mobjList.Col = LngColTmp
    
            If LngColTmp > 0 Then
                If mrsDataShow.Fields(LngColTmp - 1).type = adVarNumeric Or mrsDataShow.Fields(LngColTmp - 1).type = adNumeric Then
                    blnstr = False
                End If
            End If
            
            If IsNumberSort(strColName) Then
                blnstr = False
            End If
    
            If lngWay Mod 2 = 0 Then
            '当前排序是降序
                If blnstr Then
                    mobjList.Sort = 6
                Else
                    mobjList.Sort = 4
                End If
            Else
            '当前排序是升序
                If blnstr Then
                    mobjList.Sort = 5
                Else
                    mobjList.Sort = 3
                End If
            End If
        Else
            mobjList.Col = lngCol
            mobjList.Sort = lngWay
        End If
    End If
    
    mblExitSelChange = False
    
    For i = 1 To mobjList.Rows - 1
        mobjList.TextMatrix(i, 0) = i
    Next
    
    lngH = mObjOwner.TextHeight("字") + 120
    For RowIndex = mobjList.TopRow To mobjList.BottomRow
        mobjList.RowHeight(RowIndex) = lngH
        Call RefreshRowRelation(RowIndex)
    Next
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[排序异常]" & vbCrLf & err.Description
'    Resume
End Sub

Public Function CopyRecordSet(ByRef rsData As ADODB.Recordset) As ADODB.Recordset
'复制新的数据集使其可以修改
    Dim rsNew As ADODB.Recordset
    Dim strXml As String
    
    Set CopyRecordSet = Nothing
    
    If rsData Is Nothing Then Exit Function
    
    strXml = RecordSetToXml(rsData)
    Set rsNew = XmlToRecordSet(strXml)
    
    Set CopyRecordSet = rsNew
End Function

Private Sub RowRelationConvert(ByVal Row As Long, ByVal Col As Long, value As String)
'行关联处理 主要处理 颜色，图标 每行只需要处理一次行颜色
'行颜色处理 使用col=1 相关方式，主要是保证行关联首先设置航颜色，避免后面其他颜色被行颜色覆盖，导致设置失效。
On Error GoTo errH

    Dim i As Integer, j As Integer, k As Integer
    Dim lngRelationIndex As Long
    Dim objClsRelation As New clsScRowRelation
    Dim blContinue As Boolean '是否继续
    Dim lngColColor As Long '行颜色列
    Dim strColorColValue As String '行颜色列内容 如 "已登记"
    Static TpRowColorInfo As TRowColorInfo
    
    For i = 0 To UBound(mColConfig)
        If Col = mColConfig(i).lngColOrder Then
            lngRelationIndex = mColConfig(i).lngCfgOrder
        End If
    Next
    
    If lngRelationIndex < 1 Then
        If Col <> 1 Then
            Exit Sub
        End If
    Else
        If mSqlScheme.ShowCfg(lngRelationIndex).RowRelationCount < 1 Then
            If Col <> 1 Then
                Exit Sub
            End If
        End If
    End If

    If Col = 1 Then
        '首次加载方案后获取行颜色相关列
        '行颜色处理 使用col=1 相关方式，主要是保证行关联首先设置航颜色，避免后面其他颜色被行颜色覆盖，导致设置失效。
        
        If TpRowColorInfo.LngSchemeNo <> mlngSchemeNo Then
            TpRowColorInfo.blHaveRowColor = False
            TpRowColorInfo.LngSchemeNo = mlngSchemeNo
            
            For i = 1 To mSqlScheme.ShowCfgCount
            
                For j = 1 To mSqlScheme.ShowCfg(i).RowRelationCount
                    If mSqlScheme.ShowCfg(i).RowRelation(j).RowBackColor > 0 Or mSqlScheme.ShowCfg(i).RowRelation(j).RowFontColor > 0 Then
                        TpRowColorInfo.intRowColorIndex = i
                        TpRowColorInfo.blHaveRowColor = True
                        Exit For
                    End If
                Next
                
            Next
    
        End If
        
        If TpRowColorInfo.blHaveRowColor Then
        '存在行颜色才需要进行下面的处理
            If TpRowColorInfo.intRowColorIndex > 0 Then
                With mobjList
                    '首先处理行关联行颜色
                    For i = 1 To mSqlScheme.ShowCfg(TpRowColorInfo.intRowColorIndex).RowRelationCount
                        Set objClsRelation = mSqlScheme.ShowCfg(TpRowColorInfo.intRowColorIndex).RowRelation(i)
    
                        lngColColor = .ColIndex(mSqlScheme.ShowCfg(TpRowColorInfo.intRowColorIndex).Name)
                        strColorColValue = .TextMatrix(Row, lngColColor)
                        If strColorColValue = objClsRelation.TiggerData Then
                            '行背景色
                            If objClsRelation.RowBackColor > 0 Then .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = objClsRelation.RowBackColor
                            '行前景色
                            If objClsRelation.RowFontColor > 0 Then .Cell(flexcpForeColor, Row, 1, Row, .Cols - 1) = objClsRelation.RowFontColor
                        End If
                    Next
                End With
                
            End If
        End If
        
    End If
    
    'lngRelationIndex <1 说明不需要后面的处理
    mobjList.RowData(Row) = C_BLN_行关联刷新完成
    If lngRelationIndex < 1 Then Exit Sub

    With mobjList
        '首先处理行关联
        If mSqlScheme.ShowCfg(lngRelationIndex).RowRelationCount > 0 Then
            For i = 1 To mSqlScheme.ShowCfg(lngRelationIndex).RowRelationCount
            
                Set objClsRelation = mSqlScheme.ShowCfg(lngRelationIndex).RowRelation(i)
                
                If (LTrim(value) = objClsRelation.TiggerData And objClsRelation.TiggerData <> "[非空]" And objClsRelation.TiggerData <> "[空]") _
                    Or (Len(Trim(value)) = 0 And objClsRelation.TiggerData = "[空]") Or (Len(Trim(value)) > 0 And objClsRelation.TiggerData = "[非空]") Then
                 '图标显示及 指定列显示
  
                    If Len(objClsRelation.Icon) > 0 Then
                        '图标显示列
                        If Len(objClsRelation.IconPerformCol) > 0 Then
                            Set .Cell(flexcpPicture, Row, .ColIndex(objClsRelation.IconPerformCol)) = GetIcon(objClsRelation.Icon)
                        Else
                        '图标显示
                            Set .Cell(flexcpPicture, Row, Col) = GetIcon(objClsRelation.Icon)
                        End If
                    End If

                    If Len(objClsRelation.ColorPerformCol) > 0 Then
                        If objClsRelation.CellBackColor > 0 Or objClsRelation.CellFontColor > 0 Then
                            For k = 1 To .Cols - 1
                                If .Cell(flexcpText, 0, k) = objClsRelation.ColorPerformCol Then
                                    '颜色显示列
                                    If objClsRelation.RowBackColor > 0 Then .Cell(flexcpBackColor, Row, k) = objClsRelation.CellBackColor
                                    '颜色显示列
                                    If objClsRelation.RowFontColor > 0 Then .Cell(flexcpForeColor, Row, k) = objClsRelation.CellFontColor
        
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        '当前cell背景色
                        If objClsRelation.CellBackColor > 0 Then .Cell(flexcpBackColor, Row, Col) = objClsRelation.CellBackColor
                        '当前cell前景色
                        If objClsRelation.CellFontColor > 0 Then .Cell(flexcpForeColor, Row, Col) = objClsRelation.CellFontColor
                    End If
                End If
            Next
        End If
    End With
    
    mobjList.RowData(Row) = C_BLN_行关联刷新完成
    Exit Sub

errH:
    err.Raise -1, "clsPacsQueryWrap", "[行关联处理异常]" & vbCrLf & err.Description
'    MsgBox err.Description
'    Resume
End Sub

Public Function GetIcon(ByVal strID As String) As StdPicture
'通过ID获取图标，判断字典中是否已经存在该图标，若存在，使用字典中对象，若不存在。调用mObjQuery.GetIconRes并且添加到字典
On Error GoTo errH
    Dim stdPic As StdPicture
    
    If mPicDictionary.Exists(strID) Then
        Set GetIcon = mPicDictionary.Item(strID)
    Else
        Set stdPic = mObjQuery.GetIconRes(strID)
        Call mPicDictionary.Add(strID, stdPic)
        Set GetIcon = stdPic
    End If
    
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[获取图标异常]" & vbCrLf & err.Description
End Function

Private Function GetNullAdviceInf() As clsStudyInfo
    Dim ObjClsStudyInfo As New clsStudyInfo

    With ObjClsStudyInfo
        .lngPatId = 0
        .strPatientName = ""
        .lngPatDept = 0
        .lngAdviceId = 0
        .lngUnit = 0
        .lngSendNo = 0
        .strStudyUID = ""
        .blnCanPrint = False
        .blnIsInsidePatient = False
        .blnMoved = False
        .intState = -1
        .intStep = -1
        .strRegNo = ""
        .lngExeDepartmentId = 0
        .strExeRoom = ""
        .lngPatientFrom = 0
        .strDoDoctor = ""
        .strStudyNum = ""
        .strMarkNum = "0"
        .lngBaby = 0
        .strPatientSex = ""
        .strPatientAge = ""
        .strNO = ""
        .lngRecordKind = 0
        .intFilmGiveOut = 0
        .intReportGiveOut = 0
        .strAdviceContext = ""
        .strAdviceDepartAndMethod = ""
        .strStuStateDesc = ""
        .blnIsTechincalSure = False
        .intDangerState = 0
        .intEmergentTag = 0
        .intGreenChannel = 0
        .lng门诊执行状态 = 0
    End With
    
    Set GetNullAdviceInf = ObjClsStudyInfo
End Function

Private Function GetColSort(ByVal strColName As String) As String
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
'获取排序列
    GetColSort = strColName
    
    If mTColSort.LngSchemeNo <> mlngSchemeNo Then
        '首先获取一次排序信息
        mTColSort.LngSchemeNo = mlngSchemeNo
        Set mTColSort.dictSortInfo = New Dictionary
        Set mTColSort.dictNumberSort = New Dictionary
        
        If mSqlScheme Is Nothing Then Exit Function
        
        For i = 1 To mSqlScheme.ShowCfgCount
            If Len(mSqlScheme.ShowCfg(i).SortContrastCol) > 0 Then
                Call mTColSort.dictSortInfo.Add(mSqlScheme.ShowCfg(i).Name, mSqlScheme.ShowCfg(i).SortContrastCol)
            End If
            
            If mSqlScheme.ShowCfg(i).IsNumerSort Then
                Call mTColSort.dictNumberSort.Add(mSqlScheme.ShowCfg(i).Name, mSqlScheme.ShowCfg(i).IsNumerSort)
            End If
        Next
    End If
    
    If mTColSort.dictSortInfo Is Nothing Then
        GetColSort = strColName
        Exit Function
    End If
    
    
    If mTColSort.dictSortInfo.Exists(strColName) Then
        GetColSort = mTColSort.dictSortInfo.Item(strColName)
    End If
    
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[排序异常]" & vbCrLf & err.Description
'    Resume
End Function

Public Function RecordSetToXml(ByRef rsData As ADODB.Recordset) As String
'recordset转xml字符
    Dim strDataSource As String
    Dim strReadText As String
    Dim strSchema As String
    Dim lngStartPos As Long
    Dim strNewPro As String
    Dim strXml As String
    Dim strXML_RSDATA As String
    Dim strXML_SCHEMA As String
    Dim i As Long
    Dim objField As ADODB.Field
    Dim strNames As String
    Dim strSplitName As String
    Dim strBufer(100) As String
    Dim lngBufCount As Long
    
    Dim adoSourceStream As ADODB.Stream
    
    Set adoSourceStream = New ADODB.Stream
    adoSourceStream.type = adTypeText
    adoSourceStream.Mode = adModeRead
    
    Call rsData.Save(adoSourceStream, adPersistXML)
    
    adoSourceStream.Position = 0
    
    lngBufCount = 0
    '如果一次性读取所有数据，速度会变很慢
    strReadText = adoSourceStream.ReadText(200000)
    While Trim(strReadText) <> ""
        strBufer(lngBufCount) = strBufer(lngBufCount) & strReadText
        If Len(strBufer(lngBufCount)) >= 2000000 Then
            lngBufCount = lngBufCount + 1
        End If
        
        strReadText = adoSourceStream.ReadText(200000)
    Wend
    
    strDataSource = ""

    For i = 0 To lngBufCount
        strDataSource = strDataSource & strBufer(i)
    Next i
    
    lngStartPos = InStr(strDataSource, "<rs:data>")
    strXML_RSDATA = Mid(strDataSource, lngStartPos, InStr(strDataSource, "</rs:data>") - lngStartPos + 11)
    
    '将日期格式为“2012-03-04T12:13:14”替换为“2012-03-04 12:13:14”
'    strXML_RSDATA = RegReplace(strXML_RSDATA, "(?!\b-\d{1,2})T(?=\d{1,2}:)", " ")
    
    strNewPro = ""
    '增加绑定显示的数据列
    For i = 0 To rsData.Fields.Count - 1
        Set objField = rsData.Fields(i)
        strSplitName = "," & objField.Name & ","
        
        If InStr(strNames, strSplitName) <= 0 Then
            strNames = strNames & strSplitName
        
            If strNewPro <> "" Then strNewPro = strNewPro & vbCrLf
            If objField.type = adDate Or objField.type = adDBTimeStamp Or objField.type = adDBDate Or objField.type = adDBTime Then
                strNewPro = strNewPro & "<s:AttributeType name='" & objField.Name & "' rs:number='" & i + 1 & "' rs:nullable='true' rs:writeunknown='true'>" & _
                        "<s:datatype dt:type='dateTime' rs:dbtype='timestamp' rs:scale='0' rs:precision='3' rs:fixedlength='true'/>" & _
                        "</s:AttributeType>"
            Else
                strNewPro = strNewPro & "<s:AttributeType name='" & objField.Name & "' rs:number='" & i + 1 & "' rs:nullable='true' rs:writeunknown='true'>" & _
                        "<s:datatype dt:type='string' rs:dbtype='str' rs:scale='0' rs:precision='3' rs:fixedlength='true'/>" & _
                        "</s:AttributeType>"
            End If
        End If
    Next i
    
    strXML_SCHEMA = "<s:Schema id='RowsetSchema'>" & vbCrLf & _
                    "    <s:ElementType name='Row' content='eltOnly' rs:CommandTimeout='30'" & vbCrLf & _
                    "    rs:updatable='true' rs:ReshapeName='DSRowset1'>" & vbCrLf & _
                    strNewPro & vbCrLf & _
                    "    <s:extends type='rs:rowbase'/>" & vbCrLf & _
                    "</s:ElementType>" & vbCrLf & _
                    "</s:Schema>"
                    
    strXml = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'" & vbCrLf & _
            "xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'" & vbCrLf & _
            "xmlns:rs='urn:schemas-microsoft-com:rowset'" & vbCrLf & _
            "xmlns:z='#RowsetSchema'>" & vbCrLf & _
            strXML_SCHEMA & vbCrLf & _
            strXML_RSDATA & vbCrLf & _
            "</xml>"
            
    RecordSetToXml = strXml
End Function

Public Function XmlToRecordSet(ByRef strData As String) As ADODB.Recordset
    Dim adoNewStream As ADODB.Stream
    Dim rsData As New ADODB.Recordset
    
    Set adoNewStream = New ADODB.Stream
    adoNewStream.type = adTypeText
    adoNewStream.Mode = adModeWrite
    
    '读取修改后的流数据
    adoNewStream.Open
    adoNewStream.WriteText strData
    adoNewStream.Position = 0
    
    rsData.CursorLocation = adUseClient
    rsData.CursorType = adOpenDynamic
    rsData.LockType = adLockOptimistic
    
    rsData.Open adoNewStream
    
    Set XmlToRecordSet = rsData
End Function

Public Sub Free()
On Error GoTo errhandle:
    '首先保存参数
    SaveLocalPara

    If mlngSchemeNo > 0 And mintShowType = 0 Then Call SaveShemeCustomCfg(mlngSchemeNo)
    
    Call mObjQuery.EmbedFree
    
    Exit Sub
errhandle:
    MsgBox "[释放部件异常]" & vbCrLf & err.Description, vbOKOnly, gstrSysName
End Sub

Private Sub Class_Initialize()
    Set mQuickFIlterSaveDict = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set mObjOwner = Nothing
    Set mobjFind = Nothing
    Set mobjFilterCmdBar = Nothing
    Set mobjList = Nothing
    Set mobjIconPanel = Nothing
    Set mobjtxtHistory = Nothing
    Set mobjImg24 = Nothing
    Set mobjImg16 = Nothing
    Set mobjFindPati = Nothing
    Set mobjFindCmd = Nothing
    Set mobjTab = Nothing
    Set mObjQuery = Nothing
    Set mobjSqlParse = Nothing
    Set mobjSqlParseFilter = Nothing
    Set mPicDictionary = Nothing
    Set mrsData = Nothing
    Set mTColSort.dictSortInfo = Nothing
    Set mcnOracle = Nothing
    Set mobjSquareCard = Nothing
    Set mobjPacsMainPicList = Nothing
    Set mobjTim = Nothing
    Set mobjTimHistory = Nothing
    Set mobjPicHistory = Nothing
    
    Set mobjPicFollow = Nothing
    Set mobjImgFollow1 = Nothing
    Set mobjImgFollow2 = Nothing
    Set mobjImgFollow3 = Nothing
    Set mobjImgFollow4 = Nothing
    
    Set mobjGetDataContainer = Nothing
    Set mobjcmdQuery = Nothing
    Set mobjtxtAppend = Nothing
    Set mobjTab = Nothing
    Set mobjFindCmd = Nothing
    
    Set mrsDataShow = Nothing
    Set mSqlScheme = Nothing
    Set mobjCurStudyInfo = Nothing
    Set mQuickFIlterSaveDict = Nothing
    Set mobjPicTmp = Nothing
    
    Set mobjPublicPatient = Nothing
    Set mobjlabPatiInfo = Nothing
End Sub

Public Sub mobjFilterCmdBar_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'快速过滤 执行
On Error GoTo errhandle
    Dim i As Integer
    Dim strTemp As String
    Dim lngAdviceId As Long
    Dim intIndex As Integer '菜单序号
    Dim intItemIndex As Integer '项目序号
    Dim intTMP As Integer
    Dim lngHeight As Long
    Dim lngRow As Long
    Dim intRelationCount As Integer
    Dim lngH As Long
    
    lngHeight = mobjList.Row - mobjList.TopRow

    intTMP = Control.ID
    If (intTMP Mod 100) <> 0 Then
        intIndex = Int(intTMP / 100) + 1
        intItemIndex = (intTMP Mod 100)

        '根据选择情况更新缓存中的值
        If mTQuickFilterState.TCmdState(intIndex).blSingleChoose Then
            Control.Checked = mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose
            
            For i = 1 To mTQuickFilterState.TCmdState(intIndex).intItemCount
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose = False
            Next
            mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu = ""
            mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose = Not Control.Checked
        Else
            mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose = Not mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose
        End If
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''关联快速过滤快速过滤处理开始
'        '刷新动态菜单设置
        intRelationCount = mTQuickFilterState.TCmdState(intIndex).intRelaitonCount
        If intRelationCount > 0 Then
        '若是关联过滤前者，需要更新后者显示项
            For i = 0 To intRelationCount - 1
                Call RefreshCbrQuickFilter(mTQuickFilterState.TCmdState(intIndex).lngRelationIndex(i), False)
            Next
        End If
        
        If mTQuickFilterState.TCmdState(intIndex).blIsRelationFilter Then
        '若是关联过滤后者，需要更新选择项
            Control.Checked = mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose

            If Control.Checked Then
                If InStr(";" & mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu & ";", ";" & Control.Caption & ";") = 0 Then
                    mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu = mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu & Control.Parameter & ";"
                End If
            Else
                If InStr(";" & mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu & ";", ";" & Control.Caption & ";") > 0 Then
                    mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu = Replace(mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu, Control.Parameter & ";", "")
                End If
            End If

            Call GetQuickFilterSQLPar(intIndex)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''自定义快速过滤处理结束
        
        '排队队列显示处理
        If mTQuickFilterState.TCmdState(intIndex).strKey = CONST_EXT_FILTERVALUE Then
            Call DoQueueSelRoom(intIndex)
        End If

        Call SaveFilterCfg

        ''''''''''''''点击快速过滤后立刻执行过滤并且呈现到列表中
        If Not mrsData Is Nothing Then
            Set mrsDataShow = GetFilterFromQuickFilter(mrsData)

            lngAdviceId = GetSelectRowAdviceID

            mblExitSelChange = True
            Set mobjList.DataSource = mrsDataShow
            mblExitSelChange = False
            For lngRow = 1 To mobjList.Rows - 1
                 mobjList.RowData(lngRow) = ""
            Next

            '列统计
            Call ColStatistics
            '列顺序重排
            Call LoadListHeadCfg

            '行关联
            If mobjList.TopRow <> mobjList.BottomRow Then
                lngH = mObjOwner.TextHeight("字") + 120
                For i = mobjList.TopRow To mobjList.BottomRow
                    mobjList.RowHeight(i) = lngH
                    If mobjList.RowData(i) <> C_BLN_行关联刷新完成 Then Call RefreshRowRelation(i)
                Next
            End If

            '重新排序
            Call ResetSort(mlngSortCol, mintSortOrder)
        Else
            Call ColStatistics(True)
        End If

        Call SetSelectRow(True, False, False, lngHeight)

    End If

    Exit Sub
errhandle:
    MsgBox "[点击快速过滤菜单异常]" & vbCrLf & err.Description, vbOKOnly, gstrSysName
'    Resume
End Sub

Public Sub mobjFilterCmdBar_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
'on error GoTo errHandle 'ID9400 下标越界 原因未知   LSQ329待处理
On Error Resume Next
    Dim i As Integer
    Dim strTemp As String
    Dim intTMP As Integer
    Dim intIndex As Integer
    Dim intItemIndex As Integer

    Static blRun As Boolean
    If Control.ID = 9400 Then Exit Sub
    If blRun Then Exit Sub
    blRun = True
    strTemp = ""
    intTMP = Control.ID

    intItemIndex = (intTMP Mod 100)

    If (intTMP Mod 100) = 0 Then
    
        '菜单只显示6个字符

        intIndex = Int(intTMP / 100)
        '根据子项目选择情况决定菜单名称和使用的图标
        For i = 1 To mTQuickFilterState.TCmdState(intIndex).intItemCount

            If mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose Then
                If Len(strTemp) = 0 Then
                    strTemp = mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName
                Else
                    strTemp = strTemp & "," & mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName
                End If
            End If

        Next

        If Len(strTemp) = 0 Then
            Control.ToolTipText = "根据[" & Control.Parameter & "]进行过滤"
            Control.Caption = GetFilterMenuCaption(Control.Parameter)
            Control.iconid = C_ICON_MENUNOCHOOSE
        Else
            Control.ToolTipText = "显示[" & Control.Parameter & "]为[" & strTemp & "]的检查"
            Control.Caption = GetFilterMenuCaption(strTemp)
            Control.iconid = C_ICON_MENUCHOOSE
        End If

        If mTQuickFilterState.TCmdState(intIndex).intItemCount = 0 Then
            Control.ToolTipText = "根据[" & Control.Parameter & "]进行过滤"
            Control.Caption = GetFilterMenuCaption(Control.Parameter)
            Control.iconid = C_ICON_MENUCHOOSE
            Control.Enabled = False
        Else
            Control.Enabled = True
        End If
        
        If Not mTQuickFilterState.TCmdState(intIndex).blHaveItem Then
            Control.iconid = C_ICON_MENUNOCHOOSE
            Control.Caption = GetFilterMenuCaption(Control.Parameter)
            Control.Enabled = False
        End If

    Else
        '改变图标
        intIndex = Int(intTMP / 100) + 1

        If mTQuickFilterState.TCmdState(intIndex).intItemCount < 1 Then
            Control.Enabled = False
        End If
        
        If mTQuickFilterState.TCmdState(intIndex).blSingleChoose Then
            Control.iconid = IIf(mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose, C_ICON_MENUCHOOSE单选, C_ICON_MENUNOCHOOSE单选)
        Else
            Control.iconid = IIf(mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose, C_ICON_MENUCHOOSE, C_ICON_MENUNOCHOOSE)
        End If
    End If

    blRun = False

    Exit Sub
'errHandle:
'    err.Raise -1, "clsPacsQueryWrap", "[mobjFilterCmdBar_Update]" & vbCrLf & err.Description
'    MsgBox err.Description
''    Resume
End Sub

Private Sub SaveFilterCfg()
'保存快速过滤参数
On Error GoTo errH
    Dim objControl As CommandBarControl
    Dim strName As String
    Dim strValue As String
    Dim lngID As Long
    Dim i As Integer
    Dim j As Integer
    Dim intMenuCount As Integer '菜单数
    Dim intItemCount As Integer '菜单子菜单数
    Dim strValueAll As String

    Dim strOldFilterCfg As String
    Dim strOldFilterCfgTmp As String
    Dim strNameTmp As String

    strValueAll = ""

    intMenuCount = mTQuickFilterState.intQuickFilterMenuCount
    For i = 1 To intMenuCount
        intItemCount = mTQuickFilterState.TCmdState(i).intItemCount

        lngID = 100 * i
        Set objControl = mobjFilterCmdBar.FindControl(, lngID)
        strName = objControl.Parameter


        strOldFilterCfgTmp = mstrSchemeCfg.strFilterCfg

        If strOldFilterCfgTmp = "" Then
            strOldFilterCfg = ""
        Else
            strOldFilterCfg = ""
            On Error Resume Next
                For j = 0 To UBound(Split(strOldFilterCfgTmp, "|")) - 1
                    strOldFilterCfg = Split(strOldFilterCfgTmp, "|")(j) '一项菜单的信息

                    If Split(strOldFilterCfg, ",")(0) = strName Then
                        strOldFilterCfg = Split(strOldFilterCfg, ",")(1)
                        Exit For
                    End If
                Next
            On Error GoTo errH
        End If
        
        ''''
        
        
        '动态快速过滤的保存，保存的菜单名称
        If intItemCount > 0 Then
            For j = 1 To intItemCount
                strNameTmp = mTQuickFilterState.TCmdState(i).cmdItem(j).strName
                If mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose Then
                    If InStr(1, ";" & strOldFilterCfg & ";", ";" & strNameTmp & ";") = 0 Then
                        strOldFilterCfg = strOldFilterCfg & ";" & strNameTmp
                    End If
                Else
                    If InStr(1, ";" & strOldFilterCfg & ";", ";" & strNameTmp & ";") > 0 Then
                        strOldFilterCfg = Replace(strOldFilterCfg, ";" & strNameTmp, "")
                    End If
                End If
            Next
        End If
        '注意此处的"1"用来表示动态快速过滤
        strValue = strName & "," & strOldFilterCfg
        strValueAll = strValueAll & strValue & "|"
        
        If mQuickFIlterSaveDict.Exists(strName) Then
            mQuickFIlterSaveDict.Item(strName) = strOldFilterCfg
        End If

    Next
    
    mstrSchemeCfg.strFilterCfg = strValueAll
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[保存过滤参数异常]" & vbCrLf & err.Description
'    Resume
End Sub

Private Sub mobjFindPati_FindPatiBefore(ByVal objCard As zlOneCardComLib.Card, blnCard As Boolean, strShowText As String, objCardData As zlOneCardComLib.clsPatientInfo, blnFindPatied As Boolean, blnCancel As Boolean)
    On Error GoTo errH
    If mblnDefaultLocate Then
        Call StartFindPati(False, False)
    Else
        Call StartFindPati(True, True)
    End If
    blnCancel = True
    Exit Sub
errH:
    MsgBox "定位功能异常" & vbCrLf & err.Description, vbOKOnly, gstrSysName
End Sub

Private Sub mobjFindPati_ItemClick(Index As Integer, objCard As zlOneCardComLib.Card)
On Error GoTo errhandle
    
    If mblnAssignment Then Exit Sub
    mobjFindPati.Text = ""  '切换Item时，要将输入框清空
    mstrFindName = objCard.名称
    
    mobjFindCmd.Enabled = IsPatiFilterCanFind(mobjFindPati.GetCurCard.接口序号 > 0, objCard.名称)
    If Not mblnDefaultLocate Then mblnDefaultLocate = Not mobjFindCmd.Enabled
    RaiseEvent OnLocateBackColor
    
    mTPatiIdentifyInfo.strLocateItem = mstrFindName

    Call SaveLocalPara_PatiIdentify
    Exit Sub
errhandle:
    MsgBox "[刷卡部件点击异常]" & vbCrLf & err.Description, vbOKOnly, gstrSysName
End Sub

Private Sub mobjFindPati_KeyPress(KeyAscii As Integer)
'录入事件
On Error GoTo errhandle
    Dim blnCard As Boolean
    Dim lngPatientID As Long

    If KeyAscii = 13 Then
        '按下回车，默认先进行定位，如果没有定位到数据则自动查找
        If Trim(mobjFindPati.Text) = "" Then
            Call ExecuteQuery(C_QUERY_刷新)
            Exit Sub
        End If
    
        '按下回车后会进入mobjFindPati_FindPatiArfter
        '单据号补全处理
        Call FillNo(mobjFindPati.Text)
        
        Exit Sub
    End If

    If mobjFindPati.GetCurCard.是否刷卡 Then
        blnCard = mobjFindPati.zlIsBrushCard(mobjFindPati.objTxtInput, KeyAscii)

        If blnCard And Len(mobjFindPati.Text) = mobjFindPati.GetCardNoLen - 1 And KeyAscii <> 8 Then  '刷卡完毕处理
            mobjFindPati.Text = mobjFindPati.Text & Chr(KeyAscii)

            KeyAscii = 0
            
            mobjList.Row = 0

            If mblnDefaultLocate Then
                Call StartFindPati(False, False)
            Else
                Call StartFindPati(True, True)
            End If
        End If
    
    End If
    
Exit Sub
errhandle:
    MsgBox "[刷卡部件内容输入异常]" & vbCrLf & err.Description, vbOKOnly, gstrSysName
End Sub

Private Sub mobjList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    SaveListHeadCfg
End Sub

Private Sub mobjList_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mobjPicFollow.Visible = False
    Call SetSelectRow(False, False, True, 0)
    
End Sub

Private Sub mobjList_AfterSort(ByVal Col As Long, Order As Integer)
'排序后需要更新列表可见行行关联设置
On Error GoTo errH
    Dim RowIndex As Long
    Dim i As Long
    Dim lngH As Long
    
    mlngSortCol = Col
    mintSortOrder = Order
    
    For i = 1 To mobjList.Rows - 1
        mobjList.TextMatrix(i, 0) = i
    Next
    
    If mobjList.TopRow = mobjList.BottomRow Then Exit Sub
    
    lngH = mObjOwner.TextHeight("字") + 120
    For RowIndex = mobjList.TopRow To mobjList.BottomRow
        mobjList.RowHeight(RowIndex) = lngH
        Call RefreshRowRelation(RowIndex)
    Next
    
    Call SetSelectRow(True, False, False, 0)
    
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[排序后处理异常]" & vbCrLf & err.Description
End Sub

Private Sub mobjList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    SaveListHeadCfg
End Sub

Private Sub mobjList_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
On Error GoTo errH
'获取之后的列表显示范围，用于控制执行RefreshRowRelation
    Dim lngHeight As Long
    Dim RowIndex As Long
    Dim LngListBottom As Long
    Dim t1 As Long
    Dim lngH As Long
    
    lngHeight = mobjList.BottomRow - mobjList.TopRow
    
    LngListBottom = NewTopRow + lngHeight
    If LngListBottom > mobjList.Rows - 1 Then LngListBottom = mobjList.Rows - 1
    
    lngH = mObjOwner.TextHeight("字") + 120
    For RowIndex = NewTopRow To LngListBottom
        mobjList.RowHeight(RowIndex) = lngH
        If mobjList.RowData(RowIndex) <> C_BLN_行关联刷新完成 Then Call RefreshRowRelation(RowIndex)
    Next

    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[列表滚动前异常]" & vbCrLf & err.Description
'    Resume
End Sub

Private Sub mobjList_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo errH
    Dim lngOrder As Long
    Dim blnstr As Boolean '排序对照列是否字符型
    Dim lngColSort As Long
    
    blnstr = True
    lngColSort = mobjList.ColIndex(GetColSort(mobjList.ColKey(Col)))
    If Col <> lngColSort Then
        
        If lngColSort > 0 Then
            If mrsDataShow.Fields(lngColSort - 1).type = adVarNumeric Or mrsDataShow.Fields(lngColSort - 1).type = adNumeric Then
                blnstr = False
            End If
        End If

        If mintSortOrder Mod 2 = 0 Then
        '当前排序是降序
            If blnstr Then
                lngOrder = 5
            Else
                lngOrder = 3
            End If
        Else
        '当前排序是升序
            If blnstr Then
                lngOrder = 6
            Else
                lngOrder = 4
            End If
        End If

        '使用了排序列需要单独排序，后面需要将Order 设置为0 避免执行自带的排序
        mblExitSelChange = True
        Call SetOrder(lngColSort, lngOrder)
        mblExitSelChange = False
        
        mlngSortCol = lngColSort
        mintSortOrder = lngOrder
        
        Call SetSelectRow(True, False, False, 0)
        Order = 0
    End If
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[列表排序前异常]" & vbCrLf & err.Description
'    Resume
End Sub

Private Sub mobjList_Click()
    If mlngAdviceId = 0 Then Call mobjList_SelChange
End Sub

Private Sub mobjList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errH
    Dim strCfgNew As String
    Dim blRaiseEvent As Boolean '比如用于pacsMain弹出右键菜单
    Dim blnTemp As Boolean
    Dim pPoint As POINTAPI

    If mobjList.MouseRow = -1 Then
        Exit Sub
    End If
    blRaiseEvent = True

    '处理功能跟随
    If Button = 1 And mobjList.Rows > 1 And mobjList.MouseRow > 0 And mobjList.MouseCol > 0 Then
        '每次点击重新计时
        mobjTim.Enabled = False
        Call DoFunctionFollow(True)
        
        If Not mSqlScheme Is Nothing Then
            If mSqlScheme.UseFuncFollow Then
                blnTemp = mobjPicFollow.Visible
                mobjPicFollow.Visible = False
                Call GetCursorPos(pPoint)
                Call ScreenToClient(mObjOwner.hwnd, pPoint)
                Call mobjPicFollow.Move(Screen.TwipsPerPixelX * pPoint.X + 300, Screen.TwipsPerPixelY * pPoint.Y + 200)
                If blnTemp Then mobjPicFollow.Visible = True
            Else
                mobjPicFollow.Visible = False
            End If
        Else
            mobjPicFollow.Visible = False
        End If
        
    End If

    If mintShowType = 1 Then Exit Sub
    If mobjList.MouseRow = 0 And Button = 2 Then
        '获得初始配置串 和 个性化配置串
        If mstrSchemeCfg.strListCfg = "" Then
            mstrSchemeCfg.strListCfg = GetListHeadString()
        End If

        If frmVsfColsList.ShowVsfColsListWindow(mColConfig, mObjOwner) Then
            Call LoadListHeadCfg
            mstrSchemeCfg.strListCfg = GetListHeadString()
            Call SaveShemeCustomCfg(mlngSchemeNo)
        End If

        blRaiseEvent = False
    End If

    If blRaiseEvent Then RaiseEvent OnMouseUp(Button, Shift, X, Y)

    Exit Sub
errH:
    MsgBox "[鼠标抬起操作异常]" & vbCrLf & err.Description, vbOKOnly, gstrSysName
End Sub

Private Sub mobjList_SelChange()
'列表选中行改变    'LSQ待处理行选中
'1 根据医嘱ID查询基本信息并且保存到列表中
'2 刷新列表下方控件显示内容
'3 更新当前列表结构数据
On Error GoTo errH
    Dim intCol As Integer
    Dim lngAdviceId As Long
    Dim ObjTStudyInfo As TStudyInfo
    Dim lngListSelectRow As Long


    Call DoBackColorSel
    
    If mobjList.Rows = 1 Then mlngAdviceId = 0
    
    If mblExitSelChange Then Exit Sub 'datasource 操作避免进入后面流程

    If mSqlScheme Is Nothing Then Exit Sub

    If mobjList.Row < 1 Then Exit Sub

    lngListSelectRow = mobjList.Row
    
    Call SetPatiTyprColor(lngListSelectRow, False)

    '状态图
    RaiseEvent OnDoStateImage(lngListSelectRow)
    intCol = mobjList.ColIndex(mstrListKeyCol)
    If intCol = -1 Then Exit Sub

    lngAdviceId = Val(mobjList.TextMatrix(lngListSelectRow, intCol))
    
    If lngAdviceId = 0 Or (mlngAdviceId = lngAdviceId And mlngAdviceId <> 0 And mobjList.Row <> 1) Then Exit Sub
    mlngAdviceId = lngAdviceId
  
    Set mobjCurStudyInfo = GetBaseInfo(lngAdviceId, GetMovedState(lngListSelectRow, mobjList))
    
    If mobjList.Cell(flexcpData, lngListSelectRow) = Empty Then
        If mobjCurStudyInfo.blnMoved Then
            mobjList.Cell(flexcpData, lngListSelectRow) = 2
        Else
            mobjList.Cell(flexcpData, lngListSelectRow) = 1
        End If
    End If

    RaiseEvent OnChangeData(True, True)

    '这里需要更新功能跟随的按键
    If mobjPicFollow.Visible = True Then
        Call DoFunctionFollow(True)
    End If

    Call FillAppend(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.blnMoved, mobjtxtAppend)
    Call DoAfterFillAppend
    
    Call SetSelectRow(True, True, True, 0)
    
    Exit Sub
errH:
    MsgBox "[数据选中发生异常]" & vbCrLf & err.Description, vbOKOnly, gstrSysName
End Sub

Private Sub mobjSqlParse_OnGetParameterValue(ByVal strParName As String, value As Variant)
'获取快速过滤的参数值
On Error GoTo errhandle
    Dim i As Integer
    Dim strValue As String
    Dim strValueAll As String
    Dim j As Integer
    Dim blChooseOne As Boolean
    
    blChooseOne = False
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        If mTQuickFilterState.TCmdState(i).strName = strParName Then
            Exit For
        End If
    Next
    
    For j = 1 To mTQuickFilterState.TCmdState(i).intItemCount
        If mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose Then
            strValue = IIf(Len(strValue) = 0, strValue, strValue & ",")
            strValue = strValue & mTQuickFilterState.TCmdState(i).cmdItem(j).strName
            blChooseOne = True
        End If
        strValueAll = IIf(Len(strValueAll) = 0, strValueAll, strValueAll & ",")
        strValueAll = strValueAll & mTQuickFilterState.TCmdState(i).cmdItem(j).strName
    Next
    
    If Not blChooseOne Then
        value = strValueAll
    Else
        value = strValue
    End If
    
    Exit Sub
errhandle:
    err.Raise -1, "clsPacsQueryWrap", "[SqlParse读取参数]" & vbCrLf & err.Description
End Sub

Private Sub mObjQuery_OnGetParameterValue(ByVal strParName As String, value As Variant)
On Error GoTo errH
    Dim strTest As String
    Dim strValue As String
    
    Call SubGetSysPar(strParName, value)

    If mTqueryType = 查找 Then
        If strParName = mstrFindName Then
            strValue = mobjFindPati.Text
            value = IIf(IsNumeric(strValue), Val(strValue), strValue)
        End If
    End If
    
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[Query读取参数]" & vbCrLf & err.Description
End Sub

Private Sub SaveListHeadCfg()
'保存加载列头参数
On Error GoTo errH
    Dim strValue As String
    
    mstrSchemeCfg.strListCfg = GetListHeadString()
    
    Exit Sub
errH:
    MsgBox "[保存列头个性化配置]" & vbCrLf & err.Description, vbOKOnly, gstrSysName
End Sub

Private Function GetListHeadString(Optional ByVal blnCfg As Boolean = False) As String
'得到列名参数: 名称,宽度,是否显示  例如  "类别,1000,1|执行过程,2000,0|"
'blnCfg 是否列头配置操作，如果是，才会更新列隐藏情况
On Error GoTo errH
    Dim i As Integer, j As Integer
    Dim strTemp As String
    Dim strData As String
    Dim strName As String
    
    Dim intCount As Integer 'mColConfig数量

    strTemp = ""
    intCount = UBound(mColConfig)
    
    If intCount = 0 Then Exit Function
    
    For i = 1 To mobjList.Cols - 1
        
        strName = mobjList.TextMatrix(0, i)
        
        '列顺序和用户隐藏属性赋值
        For j = 0 To intCount
            If strName = mColConfig(j).strName Then
                mColConfig(j).lngColOrder = i
                If blnCfg Then mColConfig(j).blnIsUserHide = mobjList.ColHidden(i) And Not mColConfig(j).blnIsSysHide
                mColConfig(j).lngWidth = mobjList.ColWidth(i)
                Exit For
            End If
        Next
    Next
    
    For i = 0 To UBound(mColConfig)
        If Len(strTemp) > 0 Then
            strTemp = strTemp & "|"
        End If
        
        strTemp = strTemp & mColConfig(i).strName & "," & mColConfig(i).lngWidth & "," & mColConfig(i).blnIsUserHide & "," & mColConfig(i).lngColOrder
    Next
    
    GetListHeadString = strTemp
    
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[获取列头个性化配置]" & vbCrLf & err.Description
    Resume
End Function

Public Function SetOrder(ByVal lngCurSortCol As Long, ByVal lngCurOrder As Long) As Long
'单独排序，替代控件自带的排序（参考vsflexgrid的排序demo）
On Error GoTo errH
    Dim i As Long
    Dim RowIndex As Long
    Dim lngH As Long
    
    SetOrder = lngCurOrder
    
     '没有数据时退出排序
    If mobjList.Rows = 1 Then Exit Function
    
    With mobjList
        Dim R&, c&, RS&, cs&
        .GetSelection R, c, RS, cs
        .Redraw = flexRDNone
    
        ' apply sort to non-empty range
        Dim Row%
        
        For Row = .Rows - 1 To .FixedRows Step -1
            '整行数据为空时，不参与排序
            If Len(.TextMatrix(Row, lngCurSortCol)) Or Not Trim(.TextMatrix(Row, .ColIndex(mstrListKeyCol))) = "" Then Exit For
        Next
        
        If Row > .FixedRows Then
            .Select .FixedRows, lngCurSortCol, Row, lngCurSortCol
            .Sort = lngCurOrder
        End If
        
        ' restore selection
        .Select R, c, RS, cs
        .Redraw = flexRDDirect
        
        ' cancel default sort
        SetOrder = 0
    End With
    
    For i = 1 To mobjList.Rows - 1
        mobjList.TextMatrix(i, 0) = i
    Next
    
    If mobjList.TopRow = mobjList.BottomRow Then Exit Function
    
    lngH = mObjOwner.TextHeight("字") + 120
    For RowIndex = mobjList.TopRow To mobjList.BottomRow
        mobjList.RowHeight(RowIndex) = lngH
        Call RefreshRowRelation(RowIndex)
    Next
    

    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[排序]" & vbCrLf & err.Description
End Function

Public Function UpdateRow(ByRef rsData As Recordset, ByVal lngAdviceId As Long, Optional ByVal intSyncDataType As Integer = SyncDataType.rsDataAndrsShow, Optional ByVal blnForceNoRefreshModul As Boolean = False) As Boolean
'列表单行刷新，查询记录集同步，对检查进行操作、后台刷新，等操作可以使用
'intSyncDataType ' 0 同步 mrsdata 和 mrsdataShow  ；1 只同步    mrsdataShow ；2 都不同步(后台刷新 强制刷新)
On Error GoTo errH
    Dim rsTemp As ADODB.Recordset
    Dim rsTempShow As ADODB.Recordset
    Dim i As Long
    Dim lngIndex As Long
    
    Dim lngAdviceIDOld As Long '之前的选中行医嘱ID，用于恢复
    Dim lngHeight As Long

    Dim blnAdviceInRs As Boolean

    Dim intType As Long '内部操作类型  1 新增  2 更新  3 强制显示  4 后台刷新
    Dim lngRowIDOld As Long   '列表更新操作前医嘱ID 项目所在行  可能为-1
    Dim blnDel As Boolean '发生自动刷新删除操作
    Dim blnQuickFilterDel As Boolean '在新增或者更新操作中被快速过滤条件过滤掉
    
    blnDel = False
    UpdateRow = False
    blnQuickFilterDel = False
    If mObjQuery Is Nothing Then Exit Function
    '排序与定位预处理
    
    lngAdviceIDOld = GetSelectRowAdviceID
    lngHeight = mobjList.Row - mobjList.TopRow

    '---------------------------更新记录集 和 列表的处理---------------------------------------------------------------------------------------------
    
    lngRowIDOld = mobjList.FindRow(lngAdviceId, 1, mobjList.ColIndex(mstrListKeyCol))
    
    '更新 记录集和List中医嘱ID=lngAdviceID的记录
    mTqueryType = 刷新

    If lngRowIDOld = -1 Then
        Set rsTemp = mObjQuery.ExecuteWithAttach("[系统.医嘱ID]", lngAdviceId, 0)
    Else
        Set rsTemp = mObjQuery.ExecuteWithAttach("[系统.医嘱ID]", lngAdviceId, GetMovedState(lngRowIDOld, mobjList))
    End If
    
    If rsTemp.EOF Then
        Exit Function
    End If
    
    'LSQTODO 确认 mObjQuery.DataConvert(rsTemp, mlngSchemeNo)几种情况
    Set rsTempShow = mObjQuery.DataConvert(rsTemp, mlngSchemeNo)
    If rsTempShow.EOF Then
        Exit Function
    End If
    '此时 rsTemp 是原始记录  rsTempShow是数据转换后的记录
    
    If intSyncDataType = SyncDataType.rsDataAndrsShow Then
        Call SyncRst(mrsData, rsTemp, blnAdviceInRs)
        Call SyncRst(mrsDataShow, rsTempShow, True)
    ElseIf intSyncDataType = SyncDataType.rsShow Then
        Call SyncRst(mrsData, rsTemp, blnAdviceInRs)
    End If
    
    
    If blnAdviceInRs Then
        intType = EUpDateRowType.更新
        '更新
    Else
        intType = EUpDateRowType.新增
        '新增
    End If
    
    If mSqlScheme.RealTimeFilter Then
        Set rsTempShow = GetFilterFromQuickFilter(rsTempShow, True)
        If rsTempShow.RecordCount = 0 Then blnQuickFilterDel = True
    End If
    
    '只有同步了mrsDataShow 才需要执行UpdateListRow
    If (intSyncDataType = SyncDataType.rsDataAndrsShow Or intSyncDataType = SyncDataType.rsShow) And Not blnQuickFilterDel Then
        Call UpdateListRow(mrsDataShow, lngAdviceId, intType)
    End If
    
    If blnQuickFilterDel Then
    
        If intType = EUpDateRowType.新增 Then
        '在新增操作被过滤不需要刷新界面
            UpdateRow = True
            Exit Function
        Else
        '更新操作检查被过滤掉后，如果列表没有数据，清空列表，如果列表有数据，定位到第一条数据
            If lngRowIDOld > 0 Then
                If DelRow(lngRowIDOld) Then
                    lngRowIDOld = 1
                    lngAdviceIDOld = Val(mobjList.TextMatrix(1, mobjList.ColIndex(mstrListKeyCol)))
                    blnDel = True
                Else
                    lngRowIDOld = -1
                End If
            End If
        End If
    End If
    
    Call ColStatistics
    Call SetListColWidth
    
    If intType = EUpDateRowType.新增 And lngRowIDOld = -1 Then
        '空列表首次新增需要刷新新增的检查的信息，
        Call RefreshDisplay(0, lngAdviceId, intType, blnForceNoRefreshModul)
    ElseIf intType = EUpDateRowType.更新 And lngRowIDOld = 1 And blnQuickFilterDel Then
        '删除后列表中还有数据，定位到第一条数据
        Call RefreshDisplay(lngRowIDOld, lngAdviceIDOld, intType, blnForceNoRefreshModul)
    ElseIf intType = EUpDateRowType.更新 And lngRowIDOld = -1 And blnQuickFilterDel Then
        '删除后列表中没有数据，不再需要刷新界面信息
        UpdateRow = True
        Exit Function
    Else
        '其他情况刷新选中检查的信息。
        Call RefreshDisplay(lngRowIDOld, lngAdviceIDOld, intType, blnForceNoRefreshModul)
    End If
    
    ''重新进行重新排序与定位
    If (lngAdviceIDOld > 0 And intType = EUpDateRowType.更新) Then
    '更新操作
        Call SetSelectRow(True, False, False, lngHeight)
    ElseIf intType = EUpDateRowType.新增 Then
    '登记操作
        Call SetSelectRow(True, False, False, lngHeight)
        '如果新登记检查在列表中，定位并且加粗新登记检查
        lngIndex = mobjList.FindRow(lngAdviceId, 1, mobjList.ColIndex(mstrListKeyCol))
        If lngIndex > 0 Then
            mobjList.Row = lngIndex
            If Not mobjList.RowIsVisible(lngIndex) Then mobjList.TopRow = lngIndex
        End If
    End If

    Call DoListCfgAfterUpDateRow
    UpdateRow = True
    
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[UpdateRow]" & vbCrLf & err.Description
    Resume
End Function

Private Function DoListCfgAfterUpDateRow() As Boolean
'UpdateRow 后执行，用于列表显示处理，目前主要是为了隐藏数据显示功能。
On Error GoTo errH
    Dim i As Long, j As Long
    Dim intColIndex As Integer
    
    If mSqlScheme Is Nothing Then Exit Function
    
    For i = 0 To UBound(mColConfig)
        If mColConfig(i).blnHiddenData Then
            intColIndex = mColConfig(i).lngColOrder
            For j = 1 To mobjList.Rows - 1
                mobjList.Cell(flexcpText, j, intColIndex) = "                                " & mobjList.Cell(flexcpText, j, intColIndex)
            Next
            mobjList.Cell(flexcpAlignment, 1, intColIndex, mobjList.Rows - 1, intColIndex) = flexAlignLeftTop
        End If
    Next
    
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[UpdateRow]" & vbCrLf & err.Description
    Resume
End Function

Private Sub InitPatiIdentify()
'第一次加载方案是需要 获取查找/定位选项，后面可以直接使用
'查找功能：默认数据+过滤数据
'定位功能：列表配置
'同时判断是否需要显示Pati控件
On Error GoTo errH
    Dim i As Integer
    Dim strSQL As String
    
    For i = 1 To mSqlScheme.ShowCfgCount
        If mSqlScheme.ShowCfg(i).UseListLocate Then
            If InStr(";" & mTPatiIdentifyInfo.strLocateItems & ";", ";" & mSqlScheme.ShowCfg(i).Name & ";") = 0 Then
                mTPatiIdentifyInfo.strLocateItems = mTPatiIdentifyInfo.strLocateItems & mSqlScheme.ShowCfg(i).Name & ";"
                mTPatiIdentifyInfo.blShowPatiIdentify = True
            End If
        End If
    Next
        
    mblnAssignment = True

    If Not mTPatiIdentifyInfo.blShowPatiIdentify Then Exit Sub
    
    mobjFindPati.IDKindStr = InitCardType(mTPatiIdentifyInfo.strLocateItems)
    If mTPatiIdentifyInfo.strLocateItems <> "" Then
        mobjFindPati.IDKindIDX = mobjFindPati.GetKindIndex(mTPatiIdentifyInfo.strLocateItem)
    Else
        mTPatiIdentifyInfo.strLocateItem = Split(mTPatiIdentifyInfo.strLocateItems, ";")(0)
        mobjFindPati.IDKindIDX = mobjFindPati.GetKindIndex(mTPatiIdentifyInfo.strLocateItem)
    End If
    
    mblnAssignment = False

    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[初始化刷卡]" & vbCrLf & err.Description
End Sub

Private Sub StartFindPati(ByVal blnFind As Boolean, blnClick As Boolean)
On Error GoTo errhandle
'blnFind 是否查找  TRUE 查找  false 定位
'blnClick 是否鼠标点击
'开始读卡
    Dim lngPatientID As Long
    
    If mobjFindPati.GetCurCard.接口序号 > 0 Then
        Call mobjSquareCard.zlGetPatiID(mobjFindPati.GetCurCard.接口序号, mobjFindPati.Text, , lngPatientID)

        Call DoFilterValue(mobjFindPati.Text, IIf(lngPatientID > 0, lngPatientID, C_LNG_使用医疗卡查找定位但是未找到对应病人ID), blnFind, blnClick)
    Else
        Call DoFilterValue(mobjFindPati.Text, 0, blnFind, blnClick)
    End If
    
    Exit Sub
errhandle:
    err.Raise -1, "clsPacsQueryWrap", "[定位查找]" & vbCrLf & err.Description
'    Resume
End Sub

Private Sub DoFilterValue(ByVal strFilter As String, ByVal lngPatientID As Long, ByVal blnFind As Boolean, blnClick As Boolean)
'调用定位功能
'blnClickFind : 鼠标点击查找功能
On Error GoTo errhandle
    Dim blnClickFind As Boolean
    Dim strTemp As String
    
    blnClickFind = blnFind And blnClick
    
    If lngPatientID = C_LNG_使用医疗卡查找定位但是未找到对应病人ID Then
        If Len(strFilter) > 0 Then
            If blnFind Then
                Call MsgBoxD(mObjOwner, "检查列表中通过" & mobjFindPati.GetCurCard.名称 & "未查找到数据", vbOKOnly, gstrSysName)
                mobjList.Rows = 1
                RaiseEvent OnClearFace
            Else
                Call MsgBoxD(mObjOwner, "检查列表中通过" & mobjFindPati.GetCurCard.名称 & "未定位到数据", vbOKOnly, gstrSysName)
            End If
        End If
        
        Exit Sub
    End If
    
    '单据号补全处理
    Call FillNo(strFilter)
    
    If blnClickFind Then
        If Trim(mobjFindPati.Text) = "" Then
            Call ExecuteQuery(C_QUERY_刷新)
            Exit Sub
        End If
        
        If lngPatientID > 0 Then
            Call SeekNextPatiFromDb("病人ID", lngPatientID, True, False)
        Else
            Call SeekNextPatiFromDb(mobjFindPati.GetCurCard.名称, strFilter, False, False)
        End If
    Else
        If lngPatientID > 0 Then
            Call SeekNextPati(mobjFindPati.tag <> mobjFindPati.Text, "病人ID", lngPatientID, True, blnClick)
        Else
            Call SeekNextPati(mobjFindPati.tag <> mobjFindPati.Text, mobjFindPati.GetCurCard.名称, mobjFindPati.Text, False, blnClick)
        End If
    End If
    Call mobjFindPati.SetFocus
    Exit Sub
errhandle:
    err.Raise -1, "clsPacsQueryWrap", "[定位查找]" & vbCrLf & err.Description
'    Resume
End Sub

Private Sub SeekNextPati(ByVal blnFirst As Boolean, ByVal strName As String, _
    ByVal strFilter As String, ByVal blnIsPatiID As Boolean, ByVal blClick As Boolean)
On Error GoTo errH
'定位功能，若没有定位到数据，则从头开始，若还未定位到数据，返回False。
'blnFirst 是否首次用这个定位条件
'strName 条件名
'strFilter 条件值
'blnIsPatiID 是否病人ID
'blClick 是否鼠标点击 出发的功能

    Dim i As Long
    Dim intB As Long
    Dim lngEndRow As Long
    Dim lngSelRow As Long
    Dim strTemp As String
    Dim lngRowIndex As Long
    Dim lngPatientID As Long
    Dim bl定位找到 As Boolean
    Dim lngColIndex As Long

    
    bl定位找到 = False
    '如果没有记录并且是回车触发的定位，则直接查找
    If mobjList.Rows - 1 <= 0 And Not blClick Then
        Call SeekNextPatiFromDb(strName, strFilter, blnIsPatiID, False)
        Exit Sub
    End If
    
    '如果没有录入需要定位的数据，则不执行后续操作
    If Len(strFilter) <= 0 Then
        Exit Sub
    End If

    intB = 0
    lngRowIndex = -1

    If Not blnFirst Then
        intB = mobjList.Row + 1
        If intB >= mobjList.Rows Then intB = 1
    End If

    lngSelRow = mobjList.Row
    lngEndRow = mobjList.Rows - 1

continue1:
   '根据字段名定位的处理
   lngColIndex = mobjList.ColIndex(strName)
    If lngColIndex > 0 Then
        lngRowIndex = mobjList.FindRow(strFilter, intB, lngColIndex, False, False)
        
        '如果定位项目时姓名，则同时通过简码定位
        If lngRowIndex <= 0 And LenB(StrConv(strFilter, vbFromUnicode)) = Len(strFilter) And _
        (strName = "姓名" Or strName = "姓 名" Or strName = "姓  名" Or strName = "姓   名") Then
            For i = intB To lngEndRow
                If zlCommFun.SpellCode(mobjList.TextMatrix(i, lngColIndex) & "※0") Like UCase(strFilter) & "*" Then
                    lngRowIndex = i
                    Exit For
                End If
            Next i
        End If
    End If

    If lngRowIndex > 0 Then
        '已经找到数据
        mobjFindPati.tag = mobjFindPati.Text
        mobjList.Row = lngRowIndex

        If mobjList.TopRow > mobjList.Row Then mobjList.TopRow = mobjList.Row
        If mobjList.BottomRow - 1 < mobjList.Row Then
            mobjList.TopRow = mobjList.TopRow + (mobjList.Row - mobjList.BottomRow) + 1
        End If
        
        If lngSelRow = mobjList.Row And lngSelRow > 0 Then
            '执行类似双击检查的功能
            RaiseEvent OnSwipeCard
        End If
        
        bl定位找到 = True
    Else
       '未找到数据，从头开始再来一遍
        If intB > 1 Then
            intB = 0
            GoTo continue1:
        End If
    End If
    
    If Not bl定位找到 Then
        If blClick Then
            If Len(strFilter) > 0 Then
                Call MsgBoxD(mObjOwner, "检查列表中通过" & strName & "未定位到数据", vbOKOnly, gstrSysName)
            End If
            
            Exit Sub
        Else
            If Not bl定位找到 Then Call SeekNextPatiFromDb(strName, strFilter, blnIsPatiID, True)
        End If
        
    End If
    
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[定位数据]" & vbCrLf & err.Description
End Sub

Private Sub SeekNextPatiFromDb(ByVal strName As String, _
    ByVal strFilter As String, ByVal blnIsPatiID As Boolean, ByVal blnShowMsg As Boolean)
'定位失败后的查找功能。判断是否配置条件中包含该数据，若有，则提示是否进行查找，若没有，则提示无法检索这个数据
'使用病人ID的情况不必判断直接查找,使用其他条件要先判断是否存在对应过滤条件
'mstrFindName 获得定位条件，用于 过滤参数赋值 注意三方卡的情况这个值= "系统.病人ID"
'blnShowMsg 是否需要弹出是否尝试在数据库中查找这个提示
On Error GoTo errH
    Dim i As Integer
    Dim strTmp As String
    Dim blExist As Boolean '存在条件

    blExist = False
    With mObjQuery.objSqlParse.SqlStruct
        If Not blnIsPatiID Then
            For i = 1 To .ParCount
                If InStr(1, .AllParameter(i), "[" & strName & "]") > 0 Or InStr(1, .AllParameter(i), "." & strName & "]") > 0 Then
                    blExist = True
                    
                    'LocSerachHint: 定位检索提示  true 启用提示
                    If mSqlScheme.LocSerachHint And blnShowMsg Then
                        If Not MsgBoxD(mObjOwner, "检查列表中通过" & strName & "未定位到数据,是否尝试在数据库中查找", vbYesNo, gstrSysName) = vbYes Then
                            Exit Sub
                        End If
                    End If
                    
                    strTmp = .AllParameter(i)
                    strTmp = Replace(strTmp, "[", "")
                    strTmp = Replace(strTmp, "]", "")
                    mstrFindName = strTmp
                    
                    Exit For
                End If
            Next
        Else
            blExist = True
            
            If mSqlScheme.LocSerachHint And blnShowMsg Then
                If Not MsgBoxD(mObjOwner, "检查列表中通过" & strName & "未定位到数据,是否尝试在数据库中查找", vbYesNo, gstrSysName) = vbYes Then
                    Exit Sub
                End If
            End If
                    
            mstrFindName = "系统.病人ID"
            mPatiID = strFilter
        End If

        
        If blExist = False Then
            Call MsgBoxD(mObjOwner, "检查列表中通过" & strName & "未定位到数据", vbOKOnly, gstrSysName)
        Else
            Call ExecuteQuery(C_QUERY_查找)
            If mrsData.RecordCount = 0 Then
                Call MsgBoxD(mObjOwner, "检查列表中通过" & strName & "未查找到数据", vbOKOnly, gstrSysName)
            End If
        End If
    
    End With
    
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "[定位数据]" & vbCrLf & err.Description
End Sub

Private Function SetSelectRow(ByVal blDoRowSelect As Boolean, ByVal blSelectLineOne As Boolean, _
 ByVal blIsScoll As Boolean, ByVal lngHeight As Long) As Boolean
' SetSelectRow 是否实际发生了行选中改变
'Private Sub SetSelectRow(ByVal blDoFind As Boolean, Optional ByVal blAfterSort As Boolean = False, Optional ByVal blFindRow As Boolean = False)
'恢复选中行，查询、快速过滤后都应该执行本过程
'选中行设置字体加粗，其他取消字体加粗.

'blDoRowSelect 是否行保持处理
'blSelectLineOne 是否强制选中首行
'blIsScoll 是否滚动操作
'lngHeight 之前的列表选中位置
On Error GoTo errH
    Dim lngForeColor As Long
    Dim lngR As Long, lngG As Long, lngB As Long
    Dim lngRow As Long
    Dim blNeedChangeSel As Boolean
    Dim lngTmp As Long
    Dim lngTop As Long

    SetSelectRow = False
    If mobjList.Rows < 2 Then
        RaiseEvent OnClearFace
        Exit Function
    End If
    
    If Not blIsScoll Then
    
        If blDoRowSelect Then
            If mlngAdviceId > 0 Then
                lngRow = mobjList.FindRow(mlngAdviceId, 1, mobjList.ColIndex(mstrListKeyCol), False, False)
                If lngRow > 0 Then
                    mobjList.Row = lngRow
                    lngTop = mobjList.Row - lngHeight
        
                    If lngTop < 2 Then lngTop = 1
                    mobjList.TopRow = lngTop
                    
                    If mobjList.Row < mobjList.TopRow Or mobjList.Row > mobjList.BottomRow Then
                        mobjList.TopRow = mobjList.Row
                    End If
        
                Else
                    blNeedChangeSel = True
                End If
            Else
                blNeedChangeSel = True
            End If
        End If
        
        If (blSelectLineOne Or (blDoRowSelect And blNeedChangeSel)) Then
            mobjList.TopRow = 1
            If mobjList.Rows = 1 Then
                Call mobjList_SelChange
                SetSelectRow = True
            Else
                If mobjList.Row <> 1 Then
                    mobjList.Row = 1
                Else
                    '增加判断第一行数据是否是原来数据  如果不是 需要单独mobjList_SelChange
                    If mlngAdviceId <> Val(mobjList.TextMatrix(1, mobjList.ColIndex(mstrListKeyCol))) Then
                        Call mobjList_SelChange
                    End If
                End If
                SetSelectRow = True
            End If
        End If
    End If
    
    With mobjList
        If .Row < 0 Then Exit Function
    
        If .Cols > 2 And .Rows > 1 Then
            .Cell(flexcpFontBold, .TopRow, 1, .BottomRow, .Cols - 1) = False
            .Cell(flexcpFontBold, .Row, 1, .Row, .Cols - 1) = True
            If mblnSelRowTransparent Then
                .Cell(flexcpFontSize, .TopRow, 1, .BottomRow, .Cols - 1) = gbytFontSize
                .Cell(flexcpFontSize, .Row, 1, .Row, .Cols - 1) = gbytFontSize + 1
            End If
        End If
    End With
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''blDoRowSelect blFindRow blDoSelchange blIsScoll
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[SetSelectRow]" & vbCrLf & err.Description
'    Resume
End Function

Private Function GetStudyNumberDisplayName() As String
'获取检查号码显示名称
    GetStudyNumberDisplayName = IIf(mlngModule = G_LNG_PATHSTATION_MODULE, "病理号", "检查号")
End Function

Public Sub LocateRow(Optional ByVal lngRowIndex As Long = -1)
'定位指定行，默认定位为最后一位
    Dim lngRow As Long
    Dim iCol As Long
    
    If mobjList.Rows <= 1 Then Exit Sub
    
    lngRow = lngRowIndex
    If lngRow < 0 Then
        lngRow = mobjList.Rows - 1
    End If
    
    '取得第一个未隐藏的列
    For iCol = 1 To mobjList.Cols - 1
        If Not mobjList.ColHidden(iCol) Then Exit For
    Next iCol
    
    Call mobjList.Select(lngRow, iCol)
    Call mobjList.ShowCell(lngRow, iCol)
End Sub

Private Sub DoFunctionFollow(ByVal blIsShow As Boolean)
'处理功能跟随图标显示情况
'blIsShow 是否显示
'功能跟随
On Error GoTo errH
'C_ICON_STR_报到 C_ICON_STR_书写报告  C_ICON_STR_观片 C_ICON_STR_完成 C_ICON_STR_查看病人信息
'新增 观片 完成 查看详细信息
'初步计划：
'已登记：报到、修改信息 查看详细信息
'已报到：修改信息、写报告 查看详细信息
'已检查,已报告：写报告、修改信息 ，(观片) ，查看详细信息
'已审核：写报告， (观片)，完成 查看详细信息
'处理中：写报告、修改信息 查看详细信息
'报告中：写报告、修改信息 查看详细信息

    Dim lngIconW As Long
    Dim lngIconH As Long

    '未启用功能跟随直接退出
    If Not mSqlScheme.UseFuncFollow Then
        mobjPicFollow.Visible = False
        Exit Sub
    End If
    lngIconW = 360
    lngIconH = 360
    
    If blIsShow Then
    
        If mobjCurStudyInfo.intState <> 2 Then

            If mobjCurStudyInfo.intStep = 0 Or mobjCurStudyInfo.intStep = 1 Then
                '报到、修改信息 查看详细信息
                Set mobjImgFollow1.Picture = mobjImg24.ListImages.Item(C_ICON_STR_报到).Picture
                Set mobjImgFollow2.Picture = mobjImg24.ListImages.Item(C_ICON_STR_查看病人信息).Picture

                mobjImgFollow1.ToolTipText = C_FUNC_STR_报到
                mobjImgFollow2.ToolTipText = C_FUNC_STR_查看病人信息

                Call mobjPicFollow.Move(mobjPicFollow.Left, mobjPicFollow.Top, lngIconW * 2, lngIconH)
                mobjPicFollow.Visible = True
            ElseIf mobjCurStudyInfo.intStep = 2 Then
                '修改信息、写报告 查看详细信息
                Set mobjImgFollow1.Picture = mobjImg24.ListImages.Item(C_ICON_STR_书写报告).Picture
                Set mobjImgFollow2.Picture = mobjImg24.ListImages.Item(C_ICON_STR_查看病人信息).Picture

                mobjImgFollow1.ToolTipText = C_FUNC_STR_书写报告
                mobjImgFollow2.ToolTipText = C_FUNC_STR_查看病人信息

                Call mobjPicFollow.Move(mobjPicFollow.Left, mobjPicFollow.Top, lngIconW * 2, lngIconH)
                mobjPicFollow.Visible = True
            ElseIf mobjCurStudyInfo.intStep = 3 Or mobjCurStudyInfo.intStep = 4 Then
                '写报告、修改信息 ，(观片) ，查看详细信息 C_ICON_STR_书写报告
                If Trim(mobjCurStudyInfo.strStudyUID) <> "" And mlngModule = G_LNG_PACSSTATION_MODULE Then '判断是否显示观片
                    Set mobjImgFollow1.Picture = mobjImg24.ListImages.Item(C_ICON_STR_书写报告).Picture
                    Set mobjImgFollow2.Picture = mobjImg24.ListImages.Item(C_ICON_STR_观片).Picture
                    Set mobjImgFollow3.Picture = mobjImg24.ListImages.Item(C_ICON_STR_查看病人信息).Picture
    
                    mobjImgFollow1.ToolTipText = C_FUNC_STR_书写报告
                    mobjImgFollow2.ToolTipText = C_FUNC_STR_观片
                    mobjImgFollow3.ToolTipText = C_FUNC_STR_查看病人信息
    
                    Call mobjPicFollow.Move(mobjPicFollow.Left, mobjPicFollow.Top, lngIconW * 3, lngIconH)
                
                Else
                    Set mobjImgFollow1.Picture = mobjImg24.ListImages.Item(C_ICON_STR_书写报告).Picture
                    Set mobjImgFollow2.Picture = mobjImg24.ListImages.Item(C_ICON_STR_查看病人信息).Picture
    
                    mobjImgFollow1.ToolTipText = C_FUNC_STR_书写报告
                    mobjImgFollow2.ToolTipText = C_FUNC_STR_查看病人信息
    
                    Call mobjPicFollow.Move(mobjPicFollow.Left, mobjPicFollow.Top, lngIconW * 2, lngIconH)
                
                End If
                
                mobjPicFollow.Visible = True
            ElseIf mobjCurStudyInfo.intStep = 5 Then
            '写报告， (观片)，完成 查看详细信息
                If Trim(mobjCurStudyInfo.strStudyUID) <> "" And mlngModule = G_LNG_PACSSTATION_MODULE Then '判断是否显示观片
                    Set mobjImgFollow1.Picture = mobjImg24.ListImages.Item(C_ICON_STR_书写报告).Picture
                    Set mobjImgFollow2.Picture = mobjImg24.ListImages.Item(C_ICON_STR_观片).Picture
                    Set mobjImgFollow3.Picture = mobjImg24.ListImages.Item(C_ICON_STR_完成).Picture
                    Set mobjImgFollow4.Picture = mobjImg24.ListImages.Item(C_ICON_STR_查看病人信息).Picture
                    
                    mobjImgFollow1.ToolTipText = C_FUNC_STR_书写报告
                    mobjImgFollow2.ToolTipText = C_FUNC_STR_观片
                    mobjImgFollow3.ToolTipText = C_FUNC_STR_完成
                    mobjImgFollow4.ToolTipText = C_FUNC_STR_查看病人信息
                    
                    Call mobjPicFollow.Move(mobjPicFollow.Left, mobjPicFollow.Top, lngIconW * 4, lngIconH)
                Else
                    Set mobjImgFollow1.Picture = mobjImg24.ListImages.Item(C_ICON_STR_书写报告).Picture
                    Set mobjImgFollow2.Picture = mobjImg24.ListImages.Item(C_ICON_STR_完成).Picture
                    Set mobjImgFollow3.Picture = mobjImg24.ListImages.Item(C_ICON_STR_查看病人信息).Picture
                    
                    mobjImgFollow1.ToolTipText = C_FUNC_STR_书写报告
                    mobjImgFollow2.ToolTipText = C_FUNC_STR_完成
                    mobjImgFollow3.ToolTipText = C_FUNC_STR_查看病人信息
                    
                    Call mobjPicFollow.Move(mobjPicFollow.Left, mobjPicFollow.Top, lngIconW * 3, lngIconH)
                End If

                mobjPicFollow.Visible = True
            ElseIf mobjCurStudyInfo.intStep = 6 Then
                Set mobjImgFollow1.Picture = mobjImg24.ListImages.Item(C_ICON_STR_查看病人信息).Picture
                mobjImgFollow1.ToolTipText = C_FUNC_STR_查看病人信息
                
                Call mobjPicFollow.Move(mobjPicFollow.Left, mobjPicFollow.Top, lngIconW, lngIconH)
                mobjPicFollow.Visible = True
            End If
        Else
            mobjPicFollow.Visible = False
        End If
    Else
        mobjPicFollow.Visible = False
    End If

    If mobjPicFollow.Visible Then mobjTim.Enabled = True
            
    Exit Sub
errH:
    MsgBox "[功能跟随]" & vbCrLf & err.Description, vbOKOnly, gstrSysName
End Sub

Public Sub Find(ByVal blFind As Boolean, ByVal blClick As Boolean)
'blFind 是否查找  true 查找   false 定位
    Call StartFindPati(blFind, blClick)
End Sub

Public Sub FillAppend(ByVal lngAdviceId As Long, ByVal lngPatId As Long, ByVal blnMoved As Boolean, rtxtShow As RichTextBox)
On Error GoTo errH
    Dim i As Integer
    Dim rsData As Recordset
    Dim strTemp As String
    Dim strSQL As String
    Dim strFormatContext As String
    Dim strSize As String
    Dim blHave病人ID As Boolean

    rtxtShow.Text = ""
    blHave病人ID = False
    strTemp = mSqlScheme.Detail
    
    If InStr(1, strTemp, "[系统.医嘱ID]") = 0 Then
        rtxtShow.Text = ""
        Exit Sub
    End If
    
    If InStr(1, strTemp, "[系统.病人ID]") = 1 Then
        blHave病人ID = True
    End If
    
    strTemp = Replace(strTemp, "[系统.医嘱ID]", "[1]")
    
    If blnMoved Then
        strSQL = mObjQuery.GetSqlWithMoved(strSQL)
    End If
    
    If blHave病人ID Then
        strTemp = Replace(strTemp, "[系统.病人ID]", "[2]")
        Set rsData = zlDatabase.OpenSQLRecord(strTemp, "获取检查明细", lngAdviceId, lngPatId)
    Else
        Set rsData = zlDatabase.OpenSQLRecord(strTemp, "获取检查明细", lngAdviceId)
    End If
    
    If rsData.RecordCount = 0 Then
        rtxtShow.Text = ""
        Exit Sub
    End If
    
    Call rsData.MoveFirst
    
    strSize = 2 * Round(Val(gbytFontSize))
    strFormatContext = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                       "{\colortbl ;\red255\green104\blue104;\red19\green164\blue160;}" & _
                       "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs" & strSize & " "
                                   
    For i = 0 To rsData.Fields.Count - 1
        strFormatContext = strFormatContext & "\b\cf0\fs" & strSize & " " & rsData.Fields(i).Name & ":" & " \b0\cf0\fs" & strSize & " " & Replace(NVL(rsData.Fields(i).value), vbCrLf, " \par\cf0\fs" & strSize & " ") & "\par"
    Next
    
    strSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱附件", "H病人医嘱附件")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "提取病人附件", lngAdviceId)
    Do Until rsData.EOF
        If NVL(rsData!项目) <> "" Then
            strFormatContext = strFormatContext & "\b\cf0\fs" & strSize & " " & rsData!项目 & ":" & " \b0\cf0\fs" & strSize & " " & Replace(NVL(rsData!内容), vbCrLf, " \par\cf0\fs" & strSize & " ") & "\par"
        End If
        rsData.MoveNext
    Loop
    
    strSQL = "select 信息名,信息值 from 病人信息从表 where 病人ID=[1] and 就诊id=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "提取外院病人信息", lngPatId, lngAdviceId)
    Do Until rsData.EOF
        If NVL(rsData!信息名) <> "" Then
        strFormatContext = strFormatContext & "\b\cf0\fs" & strSize & " " & rsData!信息名 & ":" & " \b0\cf0\fs" & strSize & " " & Replace(NVL(rsData!信息值), vbCrLf, " \par\cf0\fs" & strSize & " ") & "\par"
        End If
        rsData.MoveNext
    Loop
    
    strFormatContext = strFormatContext & "}"
    rtxtShow.SelRTF = strFormatContext
    rtxtShow.SelStart = 0
                       
    Exit Sub
errH:
    MsgBox "[附加信息]" & vbCrLf & err.Description, vbOKOnly, gstrSysName
End Sub

Public Function GetTabSelectName(ByVal blSave) As String
'数据检索 附加信息 历次检查 随访描述 事务提醒
'blSave 是否初始化流程，false:用于使用参数初始化   true：用于退出时保存参数
On Error GoTo errH
    Dim i As Integer
    
    If Not blSave Then
        GetTabSelectName = GetSetting("ZLSOFT", "私有模块\" & UserInfo.姓名 & App.ProductName & "\" & "clsPacsQueryWrap" & "\", "附加信息选中标签", "数据检索")
    Else
        For i = 0 To mobjTab.ItemCount - 1
            If mobjTab.Item(i).Selected = True Then
                GetTabSelectName = mobjTab.Item(i).tag
            End If
        Next
    End If
    
    Exit Function
errH:
    GetTabSelectName = "数据检索"
End Function

Private Sub mobjSqlParseFilter_OnGetParameterValue(ByVal strParName As String, value As Variant)
    Dim i As Integer
    Dim strValue As String
    Dim strValueAll As String
    Dim j As Integer
    Dim blChooseOne As Boolean
    Dim intCount As Integer
    
    Call SubGetSysPar(strParName, value)
     
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        If mTQuickFilterState.TCmdState(i).strName = strParName Then
            '菜单创建状态需要进行下面的操作
            If mblnLoadFilterMenu Then
                intCount = mTQuickFilterState.TCmdState(i).intRelaitonCount
                
                ReDim Preserve mTQuickFilterState.TCmdState(i).lngRelationIndex(intCount)
                mTQuickFilterState.TCmdState(i).lngRelationIndex(intCount) = mTQuickFilterState.intQuickFilterMenuCount
                mTQuickFilterState.TCmdState(mTQuickFilterState.intQuickFilterMenuCount).blIsRelationFilter = True
                mTQuickFilterState.TCmdState(i).intRelaitonCount = intCount + 1
                
            End If

            Exit For
        End If
    Next
        
    If Not mblnLoadFilterMenu Then
    '正常工作状态，才需要根据参数生成菜单
        For j = 1 To mTQuickFilterState.TCmdState(i).intItemCount
            If mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose Then
                strValue = IIf(Len(strValue) = 0, strValue, strValue & ",")
                strValue = strValue & mTQuickFilterState.TCmdState(i).cmdItem(j).strFilterValue
                blChooseOne = True
            End If
            strValueAll = IIf(Len(strValueAll) = 0, strValueAll, strValueAll & ",")
            strValueAll = strValueAll & mTQuickFilterState.TCmdState(i).cmdItem(j).strFilterValue
        Next
        
        If Not blChooseOne Then
            value = strValueAll
        Else
            value = strValue
        End If
    End If
    
    Exit Sub
errhandle:
    err.Raise -1, "clsPacsQueryWrap", "[SqlParseFilter获取参数]" & vbCrLf & err.Description
End Sub

Private Sub mobjTim_Timer()
    mobjPicFollow.Visible = False
    mobjTim.Enabled = False
End Sub

Private Sub mobjTimHistory_Timer()
'mobjPicHistory mobjtxtHistory mobjTimHistory
    
    mobjPicHistory.Visible = False
    mobjTimHistory.Enabled = False
End Sub

Public Sub DepChange(ByVal strID As String, Optional ByVal blnDoRefresh As Boolean = True)
'blnDoRefresh 是否需要执行首次刷新
On Error GoTo errH
    mstrDepartId = strID
    
    Call mObjQuery.SyncSysPar(strID)
    
    Call DoQuickFilterMenu
    If blnDoRefresh Then Call ExecuteQuery(C_QUERY_首次刷新)
    Exit Sub
errH:
    MsgBox "科室切换执行错误" & err.Description, vbOKOnly, gstrSysName
End Sub

Public Sub ReportTypeChange(ByVal value As ReportType)
On Error GoTo errH
    mrtReportType = value
    Exit Sub
errH:
End Sub

Public Function GetBaseInfo(ByVal lngAdviceId As Long, Optional intMovedState As Integer = 0) As clsStudyInfo
''根据医嘱ID获取基本信息，保存在检查列表Cell flexcpData中，如果已经有内容，则不需要进行查询
'同时需要重新刷新列表中该行
'intmovedState 相应lngAdviceID 转存状态  0：未知    1：未转出   2：已经转出
On Error GoTo errH
    Dim strSQL As String, strSqlTmp As String
    Dim rsTemp As ADODB.Recordset
    Dim strSqlMore As String
    Dim rsTempMore As ADODB.Recordset  'rsTemp 下获得的Recordset
    Dim rsTempAge As ADODB.Recordset
    Dim StudyInfo As New clsStudyInfo
    Dim strTmp As String
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim strSQLBak As String
    Dim blSubMoved As Boolean  '是否转出，决定“获取基本检查信息”后面的查询是否查询历史表
    Dim dtAgePar As Date
    Dim dtPatientBirthDay As Date
    
    If lngAdviceId < 1 Then
        Set GetBaseInfo = GetNullAdviceInf
        Exit Function
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        strSQL = "Select A.姓名, A.病人id, A.主页id,A.性别,A.年龄,A.病人来源,A.病人科室id,A.挂号单,Nvl(A.婴儿, 0) 婴儿,A.紧急标志,A.医嘱内容,A.执行科室ID,诊疗项目ID, " & vbNewLine & _
                " B.检查uid,B.报告打印,B.检查号,B.检查技师,B.关联ID,B.图像位置,B.报告人,B.复核人,B.报告操作,B.绿色通道,B.影像类别,B.发放胶片,B.报告发放,B.报告质量,B.影像质量,B.随访描述,B.符合情况,B.危急状态, " & vbNewLine & _
                " D.首次时间,D.发送号,D.NO,D.记录性质,D.执行状态,D.执行过程,D.执行间,D.结果阳性,F.住院号,G.门诊号,G.执行状态 as 门诊执行状态,E.身份证号,E.出生日期,E.当前病区ID,0 as 转出 " & vbNewLine & _
                " From 病人医嘱记录 A, 影像检查记录 B, 病人医嘱发送 D,病人信息 E,病案主页 F,病人挂号记录 G" & vbNewLine & _
                " Where G.no(+)= A.挂号单 and A.病人ID=F.病人ID(+) And A.主页ID=F.主页ID(+) And A.ID = B.医嘱id(+)  And A.ID = D.医嘱id and A.病人ID=E.病人ID and A.ID = [1]"
    Else
    '与上面不同之处：C.病理号
        strSQL = "Select A.姓名,A.病人id,A.主页id, A.性别,A.年龄,A.病人来源,A.病人科室id,A.挂号单,Nvl(A.婴儿, 0) 婴儿,A.紧急标志,A.医嘱内容,A.执行科室ID,诊疗项目ID, " & vbNewLine & _
                " B.检查uid,B.报告打印,B.检查号,B.检查技师,B.关联ID,B.图像位置,B.报告人,B.复核人,B.报告操作,B.绿色通道,B.影像类别,B.发放胶片,B.报告发放,B.报告质量,B.影像质量,B.随访描述,B.符合情况,B.危急状态,  " & vbNewLine & _
                " C.病理号,C.检查类型,D.NO,D.记录性质,D.首次时间,D.发送号,D.执行状态,D.执行过程,D.执行间,D.结果阳性,F.住院号,G.门诊号,G.执行状态 as 门诊执行状态,E.身份证号,E.出生日期,E.当前病区ID,0 as 转出 " & vbNewLine & _
                " From 病人医嘱记录 A, 影像检查记录 B, 病理检查信息 C ,病人医嘱发送 D,病人信息 E,病案主页 F,病人挂号记录 G" & vbNewLine & _
                " Where G.no(+)= A.挂号单 and A.病人ID=F.病人ID(+) And A.主页ID=F.主页ID(+) And A.ID = B.医嘱id(+)  And A.ID = D.医嘱id and A.ID=C.医嘱ID(+) and A.病人ID=E.病人ID and A.ID = [1]"
    End If

    If mObjQuery.IsMoved And (intMovedState = 0 Or intMovedState = 2) Then
        If intMovedState = 0 Then
            '未知是否转存，首先判断非历史表
            strSqlTmp = "select 1 from 病人医嘱记录 where Id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSqlTmp, "获取基本检查信息", lngAdviceId)
            blSubMoved = False
            
            If rsTemp.EOF Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
                strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
                strSQL = Replace(strSQL, "0 as 转出", "1 as 转出")
                
                blSubMoved = True
            End If
        Else
            'intmovedState=2 直接查历史表
            strSQL = mObjQuery.GetSqlWithMoved(strSQL)
            strSQL = Replace(strSQL, "0 as 转出", "1 as 转出")
            blSubMoved = True
        End If
    Else
        blSubMoved = False
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取基本检查信息", lngAdviceId)
    
    If Not rsTemp.EOF Then
        With StudyInfo
            .lngAdviceId = lngAdviceId
            .lngPatId = Val(NVL(rsTemp!病人ID))
            .strPatientName = rsTemp!姓名
            .strPatientSex = NVL(rsTemp!性别)
            .strPatientAge = NVL(rsTemp!年龄)
            .lngPatientFrom = Val(NVL(rsTemp!病人来源, 3))
            .strStudyNum = NVL(rsTemp(GetStudyNumberDisplayName))
            .lngBaby = Val(NVL(rsTemp!婴儿))
            .lngPageID = Val(NVL(rsTemp!主页ID))
            .lngUnit = Val(NVL(rsTemp!当前病区ID))
            .blnIsPrinted = (Val(NVL(rsTemp!报告打印)) > 0)
            .lngClinicId = Val(NVL(rsTemp!诊疗项目ID))
            
            If .lngBaby <> 0 Then
                strSQL = "Select A.开嘱时间,Nvl(B.婴儿姓名, A.姓名 || '之子' || Trim(To_Char(B.序号, '9'))) As 婴儿姓名, B.婴儿性别, B.出生时间, B.死亡时间" & vbNewLine & _
                             "  From 病人医嘱记录 A, 病人新生儿记录 B " & vbNewLine & _
                             "  Where a.病人ID = b.病人ID And b.主页id = [2] And b.序号 = [3] And a.ID = [1]"
                        
                If blSubMoved Then
                    strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                End If
                Set rsTempMore = zlDatabase.OpenSQLRecord(strSQL, "提取婴儿信息", lngAdviceId, .lngPageID, .lngBaby)
                
                If Not rsTempMore.EOF Then
                    .strPatientName = rsTempMore!婴儿姓名
                    .strPatientSex = NVL(rsTempMore!婴儿性别)
                    
                    strSqlMore = "Select Zl_Age_Calc(0,[1],[2]) Age From Dual"
                    
                    If NVL(rsTempMore!死亡时间) <> "" Then
                        dtAgePar = NVL(rsTempMore!死亡时间)
                    Else
                        dtAgePar = NVL(rsTempMore!开嘱时间)
                    End If
                    
                    Set rsTempAge = zlDatabase.OpenSQLRecord(strSqlMore, "计算年龄", NVL(rsTempMore!出生时间), dtAgePar)
                    If rsTempAge.RecordCount > 0 Then
                        .strPatientAge = NVL(rsTempAge!Age)
                        
                        .blnInfancy = GetIsPacsBabyByAge(.strPatientAge)
                    Else
                        .strPatientAge = NVL(rsTempMore!出生时间)
                        
                        .blnInfancy = GetIsPacsBabyByBirthDay(.strPatientAge)
                    End If
                End If
            Else
'                有独立身份需要根据出生日期判断是否婴儿
'                若没有出生日期，则根据年龄判断是否婴儿
'                若病人信息没有年龄，则判断为不是婴儿
                dtPatientBirthDay = NVL(rsTemp!出生日期, Empty)
                If dtPatientBirthDay = Empty Then
                    .blnInfancy = GetIsPacsBabyByAge(.strPatientAge)
                Else
                    .blnInfancy = GetIsPacsBabyByBirthDay(dtPatientBirthDay)
                End If
            End If
            
            .lngLinkId = Val(NVL(rsTemp!关联ID))
            .intState = Val(rsTemp!执行状态)

            
            
            
            If mrtReportType = 报告文档编辑器 Then
                '问题130267 需要对智能报告编辑器的strStuStateDesc做调整，最终效果类似pacs报告编辑器 要有书写中的状态。
                .strStuStateDesc = IIf(Val(NVL(rsTemp!执行状态)) = 2, "已拒绝", Decode(Val(NVL(rsTemp!执行过程, 0)), -1, "已驳回", 0, "已登记", 1, "已登记", _
                                                                                            2, "已报到", 3, "已检查", 4, "已报告", 5, "已审核", "已完成"))
'                其中对于 4-已报告 的情况需要拆分成 书写中和已报告（区别在于是否所有的报告都没有签名）
'                If .strStuStateDesc = "已报告" Then
'                    '遍历所有报告，检查是否有已经签名的情况。
'                    strSQL = "select 1 from 影像报告记录 where 医嘱ID=[1] and 报告状态<>1"
'
'                    '这个查询检查是否有已经签名或者之后流程的报告，若有数据 strStuStateDesc仍然为 "已报告"  若没有数据，修改为 "书写中"
'                    Set rsTempMore = zlDatabase.OpenSQLRecord(strSQL, "智能报告状态查询", lngAdviceID)
'                    If rsTempMore.RecordCount = 0 Then
'                        .strStuStateDesc = "书写中"
'                    End If
'                End If
                                                                            
            Else
                .strStuStateDesc = IIf(Val(NVL(rsTemp!执行状态)) = 2, "已拒绝", Decode(Val(NVL(rsTemp!执行过程, 0)), -1, "已驳回", 0, "已登记", 1, "已登记", _
                                                                                        2, IIf(NVL(rsTemp!报告操作) <> "", "书写中", _
                                                                                                IIf(NVL(rsTemp!报告人) = "", "已报到", "书写中")), _
                                                                                        3, IIf(NVL(rsTemp!报告操作) <> "", "书写中", _
                                                                                                IIf(NVL(rsTemp!报告人) = "", "已检查", "书写中")), _
                                                                                        4, "已报告", _
                                                                                        5, "已审核", "已完成"))
            End If
            
            .blnIsTechincalSure = IIf(NVL(rsTemp!检查技师) <> "", True, False)
            .strDoDoctor = NVL(rsTemp!检查技师)
            .blnIsReported = (InStr("已报告,已审核,已完成,报告中,审核中", .strStuStateDesc) > 0)
            .lngPatDept = Val(NVL(rsTemp!病人科室ID))
            
            .strRegNo = NVL(rsTemp!挂号单)
            .blnIsInsidePatient = (.lngPatientFrom = 1) Or (.lngPatientFrom = 2)
            .strStudyUID = NVL(rsTemp!检查UID)
            .intImageLocation = Val(NVL(rsTemp!图像位置))
            .strImgType = NVL(rsTemp!影像类别)
            .strReportDoctor = NVL(rsTemp!报告人)
            .strReportOperation = NVL(rsTemp!报告操作)
            .intEmergentTag = Val(NVL(rsTemp!紧急标志))
            .intGreenChannel = Val(NVL(rsTemp!绿色通道))
            .intFilmGiveOut = Val(NVL(rsTemp!发放胶片))
            .intReportGiveOut = Val(NVL(rsTemp!报告发放))
            .strFollowUpDescribe = NVL(rsTemp!随访描述)
            .strImageQuality = NVL(rsTemp!影像质量)
            .strReportQuality = NVL(rsTemp!报告质量)
            .strAccord = NVL(rsTemp!符合情况)
            .intDangerState = Val(NVL(rsTemp!危急状态))
            .lngExeDepartmentId = Val(NVL(rsTemp!执行科室ID))
            .dtFirstTime = CDate(NVL(rsTemp!首次时间, "0"))
            .blnCanPrint = False
            .lng门诊执行状态 = NVL(rsTemp!门诊执行状态, 0)
            
            If mrtReportType = 报告文档编辑器 Then '紧急 =1  绿色通道=1 可打印
                If .intEmergentTag Or .intGreenChannel Then
                    .blnCanPrint = True
                Else
                    .blnCanPrint = False
                End If
            Else
                .blnCanPrint = IIf(mblnCanPrint, IIf(.intEmergentTag, NVL(rsTemp!报告人) <> "", NVL(rsTemp!复核人) <> "" Or .intGreenChannel), True)
            End If
            
            If mlngModule = G_LNG_PATHOLSYS_NUM Then .intPathoType = Val(NVL(rsTemp!检查类型))
            
            .lngSendNo = Val(NVL(rsTemp!发送号))
            .blnMoved = NVL(rsTemp!转出, False)
            .intStep = Val(NVL(rsTemp!执行过程))
            .strExeRoom = NVL(rsTemp!执行间)
            .strNO = NVL(rsTemp!no)
            .lngRecordKind = Val(NVL(rsTemp!记录性质))
            .intPositive = Val(NVL(rsTemp!结果阳性))
            
            .strMarkNum = IIf(.lngPatientFrom = 1, NVL(rsTemp!门诊号, "-1"), IIf(.lngPatientFrom = 2, NVL(rsTemp!住院号, "-1"), "0"))
            .strIIDNumber = NVL(rsTemp!身份证号)
            
            If UBound(Split(NVL(rsTemp!医嘱内容), ":")) > 0 Then
                .strAdviceContext = Split(NVL(rsTemp!医嘱内容), ":")(0)
                .strAdviceDepartAndMethod = Split(NVL(rsTemp!医嘱内容), ":")(1)
            Else
                .strAdviceContext = NVL(rsTemp!医嘱内容)
                .strAdviceDepartAndMethod = ""
            End If
            
            .dtRefreshDate = Now
            
            strSqlMore = "Select 收藏类别 From 影像收藏类别 A, 影像收藏内容 B Where a.Id = b.收藏id And b.医嘱id = [1]"
            Set rsTempMore = zlDatabase.OpenSQLRecord(strSqlMore, "查询收藏状态", lngAdviceId)
            strTmp = ""
            
            While Not rsTempMore.EOF
                If strTmp = "" Then
                    strTmp = NVL(rsTempMore!收藏类别)
                Else
                    strTmp = strTmp & "◆" & NVL(rsTempMore!收藏类别)
                End If
                rsTempMore.MoveNext
            Wend
            .strCollectionInfo = strTmp
            
            '费用 注意应该在本过程的最后，因为改变了rsTemp
            strSQL = "Select 病人id " & _
                    " From 电子病历记录 A " & _
                    " Where a.病历种类 = 5 And a.病人id = [1] And a.主页id = [2] And a.病历名称 Like '%传染病%' And Exists " & _
                    " (Select 1 From 疾病阳性记录 B Where b.病人id = [1] And b.主页id = [2]) "
                    
            If blSubMoved Then
                strSQL = Replace(strSQL, "电子病历记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "疾病阳性记录", "H疾病阳性记录")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询传染病", lngAdviceId, .lngPageID)
            
            .blnIsInfectious = rsTemp.RecordCount > 0
            
            If mlngModule = G_LNG_PATHOLSYS_NUM Then
                strSQL = " Select a.Id As 医嘱id, a.相关id, b.记录性质, b.计费状态, " & _
                        " (select count(1) from 病理检查信息 D , 病理申请信息 E where E.病理医嘱ID=D.病理医嘱id and A.id=D.医嘱ID(+) and E.补费状态=1) as 补费" & _
                        " From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C " & _
                        " Where a.Id = b.医嘱id And a.病人id = c.病人id And (a.Id = [1] Or a.相关id = [1]) "
            Else
                strSQL = " Select a.Id As 医嘱id, a.相关id, b.记录性质, b.计费状态 " & _
                        " From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C " & _
                        " Where a.Id = b.医嘱id And a.病人id = c.病人id And (a.Id = [1] Or a.相关id = [1]) "
            End If
            
                                  
            If blSubMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            End If
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询费用状态", lngAdviceId)
            If rsTemp.RecordCount > 0 Then
                Set rsTempMore = rsTemp.Clone

                rsTempMore.Filter = "医嘱ID=" & lngAdviceId
                .lngMoneyState = GetMoneyState(rsTempMore)

            End If

        End With

         Set GetBaseInfo = StudyInfo
    End If

    Exit Function
errH:
    MsgBox "GetBaseInfo-" & err.Description, vbExclamation, gstrSysName
'    Resume
'    If ErrCenter = 1 Then Resume
End Function

Private Function GetMoneyState(ByVal rsMainAdvice As ADODB.Recordset) As Long
'获取费用状态
'0-未收费,1-已收费,2-无费,3-,4-需补费,5-记账

    '判断是否已经收费
    '"病人医嘱发送.记录性质"--- 1是收费的，2是记帐的。
    
    '通过"病人医嘱发送.计费状态"直接判断,原有值：-1-无须计费;0-未计费;1-已计费，对于记帐单（包括门诊记帐单），保持原有值不变。
    '对于收费单的发送记录，增加两种状态：2-部分收费，3-全部收费
    
    '没有对应费用的医嘱有两种情况，一种是"-1-无须计费"，即没有设置收费对照，一种是"0-未计费"，即虽然设置了收费对照，但设置为发送后手工计费，即在医技科室去生成。
    '"1-已计费"就是发送时生成了费用的。但生成了费用单据不表示收费了，生成可能是记帐划价单，或收费划价单，其中收费划价单就多两种状态。
    '"2-部分收费"表示部分收费和部分退费的情况，反正没收得完。
    
    '已收费显示状态：已收费；无费用；未收费：
    '未收费----
    '1、主医嘱是收费单的，满足以下条件算未收费
    '   (1)有一条主医嘱和部位医嘱的 计费状态 in (1,2)算未收费 ------“记录性质=1 and 计费状态 in (1,2)”
    '已收费：
    '1、主医嘱是记账的算收费-------“记录性质=2”
    '2、主医嘱是收费单的，满足以下条件算收费
    '   (1)排除未收费后，有一条主医嘱和部位医嘱的 计费状态 =3 算收费-----“记录性质=1 and 计费状态 = 3”
    '无费用
    '1、主医嘱是收费单的，满足以下条件算无费用
    '   (1)所有主医嘱和部位医嘱的 计费状态 in (-1,0)算无费用 ------“记录性质=1 and 计费状态 in (-1,0)”
    
    Dim lngResult As Long

    GetMoneyState = ChargeState.无费用
    lngResult = ChargeState.无费用 '无费用
    
    '1.门诊或住院患者中，记录性质为1（收费单据）的，当计费状态为-1，0时，表示"无"，1表示"欠"，2表示"调"（暂定，表示有调整改动），3表示"收"，4表示"退"；
    '2.门诊或住院患者中，记录性质为2（记账单据）的，当计费状态为-1，0时，表示"无"；1表示"记"，2表示"调"（暂定，表示有调整改动），----------，4表示"销"；（注：记账患者不使用或不存在3的状态）
    
    If NVL(rsMainAdvice!记录性质, 2) = 2 Then
        If NVL(rsMainAdvice!计费状态, -1) = -1 Or NVL(rsMainAdvice!计费状态, -1) = 0 Then   '无
            lngResult = 2
        Else
            If NVL(rsMainAdvice!计费状态, -1) = 1 Then                                '记
                lngResult = 3
            ElseIf NVL(rsMainAdvice!计费状态, -1) = 2 Then                            '调
                lngResult = 7
            ElseIf NVL(rsMainAdvice!计费状态, -1) = 4 Then                            '销
                lngResult = 6
            End If
        End If
    Else
        If NVL(rsMainAdvice!计费状态, -1) = -1 Or NVL(rsMainAdvice!计费状态, -1) = 0 Then   '无
            lngResult = 2
        Else
            If NVL(rsMainAdvice!计费状态, -1) = 1 Then                                '欠
                lngResult = 0
            ElseIf NVL(rsMainAdvice!计费状态, -1) = 2 Then                            '调
                lngResult = 7
            ElseIf NVL(rsMainAdvice!计费状态, -1) = 3 Then                            '收
                lngResult = 1
            ElseIf NVL(rsMainAdvice!计费状态, -1) = 4 Then                            '退
                lngResult = 5
            End If
        End If
    End If

    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        Dim j As Long
        For j = 0 To rsMainAdvice.Fields.Count - 1
            If "补费" = rsMainAdvice.Fields(j).Name Then
                If NVL(rsMainAdvice!补费) > 0 Then lngResult = ChargeState.已补缴 '需要补费，需补费的检查也是未收费的检查
                GetMoneyState = lngResult
                Exit Function
            End If
        Next j
    End If
    
    GetMoneyState = lngResult
End Function

Private Sub DoAfterFillAppend()
'判断是否存在有效的附加信息配置并做相应处理
    Dim blIsFocus As Boolean
    blIsFocus = mobjTab.Item(1).Selected

    If mobjTab.Item(1).Visible = True Then
        If InStr(1, mSqlScheme.Detail, "[系统.医嘱ID]") = 0 Then
            mobjTab.Item(1).Visible = False
            If blIsFocus Then mobjTab.Item(0).Selected = True
        End If
    Else
        If InStr(1, mSqlScheme.Detail, "[系统.医嘱ID]") > 0 Then
            mobjTab.Item(1).Visible = True
        End If
    End If

End Sub

Private Function DoQuickFilterMenu(Optional ByVal blFirstTime As Boolean = False) As Boolean
'DoQuickFilterMene 处理快速过滤菜单的生成,同时恢复选中情况
'blFirstTime是否初始化
On Error GoTo errH
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer, j As Integer, lngID As Long
    Dim intMenuCount As Integer '快速过滤菜单数
    Dim intItemCount As Integer '快速过滤菜单子项目数
    Dim lngCount As Long
    Dim strSQL As String
    Dim strTemp As String
    Dim strItems As String '保存数据库查询出来的快速查询项目
    Dim strName As String, strValue As String, strValueTmp As String
    Dim blDynamicFilter As Boolean '是否动态过滤

    DoQuickFilterMenu = False
    
    '''''''清除已有快速过滤菜单
    Call LockWindowUpdate(mLngOwnerHwnd)
    For lngCount = mobjFilterCmdBar.Count To 2 Step -1
        mobjFilterCmdBar(lngCount).Delete
    Next
    Call LockWindowUpdate(0)
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    mobjFilterCmdBar.VisualTheme = xtpThemeOfficeXP
    With mobjFilterCmdBar.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    mobjFilterCmdBar.AddImageList mobjImg16 '以VB.ImageList的Tag与ID进行关联
    mobjFilterCmdBar.EnableCustomization False
    mobjFilterCmdBar.ActiveMenuBar.Visible = False
    
    Set objBar = mobjFilterCmdBar.Add("快速过滤", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    
    ReDim mTQuickFilterState.TCmdState(mSqlScheme.FilterCfgCount)
    mTQuickFilterState.intQuickFilterMenuCount = 0

    '创建快速过滤菜单 分为三种类型，
    '1:  普通快速过滤，可选项固定：  如配置：已登记;已报到;已报告
    '2:  普通快速过滤，可选项通过查询得到： 如配置： select distinct 编码 as 影像类别 from 影像检查类别
    '3:  自定义快速过滤，可选项通过前面的快速过滤选择情况得到，：如配置：[影像类别]#(这里是快速过滤可选项的Select查询语句)

    For i = 1 To mSqlScheme.FilterCfgCount
        If Not CreatFilterMenu(objBar.Controls, i, mSqlScheme.FilterCfg(i).Name, mSqlScheme.FilterCfg(i).DataFrom, mSqlScheme.FilterCfg(i).CustomScript, mSqlScheme.FilterCfg(i).SelectWay = swSingle, mSqlScheme.FilterCfg(i).ExtProperty) Then
            Exit Function
        End If
    Next

    For Each objControl In objBar.Controls
        If objControl.type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next

    mobjFilterCmdBar.RecalcLayout

    strValue = mstrSchemeCfg.strFilterCfg
    
    If Len(strValue) <> 0 Or mQuickFIlterSaveDict.Count > 0 Then
    '没有快速过滤选中参数的单独判断 共享本机参数
    
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''以下功能：根据参数恢复以前快速过滤菜单的选中情况
        '根据参数设置快速过滤选中情况
        'strValue 所有的快速过滤信息
        'strValueTmp 一组菜单的信息

        DoQuickFilterMenu = False
        intMenuCount = mTQuickFilterState.intQuickFilterMenuCount
        
        If App.LogMode = 0 Then
            On Error GoTo errH
        Else
            On Error Resume Next
        End If
        
        '处理每个快速过滤菜单
        For i = 1 To intMenuCount
            strValueTmp = ""
            intItemCount = mTQuickFilterState.TCmdState(i).intItemCount
            lngID = 100 * i
            Set objControl = mobjFilterCmdBar.FindControl(, lngID)
            strName = objControl.Parameter
            
            If mQuickFIlterSaveDict.Exists(strName) And mQuickFIlterSaveDict.Item(strName) <> CONST_EXT_FILTERKEY_SAVEWAYDEFALT Then
                strValueTmp = mQuickFIlterSaveDict.Item(strName)
                mTQuickFilterState.TCmdState(i).strRelationChooseMenu = strValueTmp
            Else
                '找到对应的快速过滤菜单,根据菜单名称获取配置,需要判断是否动态过滤
                For j = 0 To UBound(Split(strValue, "|")) - 1
                    strValueTmp = Split(strValue, "|")(j) '一项菜单的信息
    
                    If Split(strValueTmp, ",")(0) = strName Then
                        strValueTmp = Split(strValueTmp, ",")(1)
                        
                        If mQuickFIlterSaveDict.Exists(strName) Then
                            strTemp = mQuickFIlterSaveDict.Item(strName)
                            If strTemp <> "" Then strValueTmp = strTemp
                        End If
                        
                        mTQuickFilterState.TCmdState(i).strRelationChooseMenu = strValueTmp
                        Exit For
                    End If
                Next
            End If
            
            '处理快速过滤菜单子项选中情况
            For j = 1 To intItemCount
                mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose = InStr(1, ";" & strValueTmp & ";", ";" & mTQuickFilterState.TCmdState(i).cmdItem(j).strName & ";") > 0
            Next
            
            If mTQuickFilterState.TCmdState(i).strKey = CONST_EXT_FILTERVALUE Then
                Call DoQueueSelRoom(i)
            End If
        Next
    End If

    On Error GoTo errH
    
    Call RefreshCbrQuickFilterALL
    
    DoQuickFilterMenu = True

    Exit Function
errH:
    err.Raise -1, "创建快速过滤" & err.Description
'    Resume
End Function

Public Sub SetRichtxtFontSize()
    Call FillAppend(mobjCurStudyInfo.lngAdviceId, mobjCurStudyInfo.lngPatId, mobjCurStudyInfo.blnMoved, mobjtxtAppend)
    Call DoAfterFillAppend
End Sub

Private Function CreatFilterMenu(ByVal Mycontrol As CommandBarControls, ByVal intIndex As Integer, ByVal strName As String, _
ByVal strMenuSql As String, ByVal strCustomScript As String, ByVal blSingleChoose As Boolean, ByVal strExt As String) As Boolean

'每个过滤菜单，占用100个ID 例如1号 1~99  2号100~199
'向cbrList中增加快速过滤菜单
'Mycontrol 略
'strName：菜单名 如:检查类型
'intIndex:第几个过滤项
'strMenuSql：菜单获取的SQL  或者是 固定配置的菜单
'strCustomScript 自定义过滤脚本
'blSingleChoose 是否单选

On Error GoTo errH
    Dim objControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim i As Integer
    Dim intCount As Integer
    Dim rsTemp As Recordset
    Dim strMeneShow As String '用于显示的值
    Dim strMenuValue As String '用于过滤的值
    Dim strSQL As String
    Dim blFromSQL As Boolean '菜单是否来源为SQl
    Dim strItemName As String
    
    CreatFilterMenu = False
    '快速过滤菜单数 1
    mTQuickFilterState.intQuickFilterMenuCount = mTQuickFilterState.intQuickFilterMenuCount + 1
    
    '首先判断菜单来源类型  分为 固定值生成  和 sql查询方式
    If InStr(UCase(strMenuSql), "SELECT") > 0 Then
    
        Set mobjSqlParseFilter = New clsSqlParse
        
        mblnLoadFilterMenu = True
        mTQuickFilterState.TCmdState(intIndex).blIsRelationFilter = False
        mTQuickFilterState.TCmdState(intIndex).strName = strName
        Call mobjSqlParseFilter.Init(strMenuSql)
        strSQL = mobjSqlParseFilter.GetQuerySql
        Set rsTemp = ExecuteCore(strSQL, "快速过滤获取项目", mobjSqlParseFilter.ParValues)
        mblnLoadFilterMenu = False
        
        blFromSQL = True
        
        If strCustomScript <> "" Then
            mTQuickFilterState.TCmdState(intIndex).intFilterWay = EFilterType.脚本
        Else
            mTQuickFilterState.TCmdState(intIndex).intFilterWay = EFilterType.未知
        End If
        
        If rsTemp.EOF And Not mTQuickFilterState.TCmdState(intIndex).blIsRelationFilter Then
'            mstrLoadMsg = "快速过滤项目[" & strName & "]菜单未正常产生"
'            Exit Function
            mTQuickFilterState.TCmdState(intIndex).blHaveItem = False
        Else
            mTQuickFilterState.TCmdState(intIndex).blHaveItem = True
        End If
        
        If rsTemp.Fields.Count <> 1 And rsTemp.Fields.Count <> 2 Then
            mstrLoadMsg = "快速过滤项目[" & strName & "]菜单查询生成SQL配置不符合要求，请检查"
            Exit Function
        End If
        
        intCount = rsTemp.RecordCount
    Else
        mTQuickFilterState.TCmdState(intIndex).intFilterWay = EFilterType.普通
        mTQuickFilterState.TCmdState(intIndex).blHaveItem = True
        blFromSQL = False
        intCount = UBound(Split(strMenuSql, ";")) + 1
        
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''前面的代码需要注意的缓存变量：
            'mTQuickFilterState.TCmdState(intIndex).blHaveItem
            'mTQuickFilterState.TCmdState(intIndex).intFilterWay
            'mTQuickFilterState.TCmdState(intIndex).blIsRelationFilter
            'mTQuickFilterState.TCmdState(intIndex).strName
            'mTQuickFilterState.intQuickFilterMenuCount
        
        
    '生成快速菜单菜单项
    Set objControl = Mycontrol.Add(xtpControlButtonPopup, 100 * intIndex, strName)
    objControl.ToolTipText = "根据" & strName & "过滤"
    objControl.Parameter = strName
    mTQuickFilterState.TCmdState(intIndex).intRelaitonCount = 0
    mTQuickFilterState.TCmdState(intIndex).intItemCount = intCount
    mTQuickFilterState.TCmdState(intIndex).lngID = 100 * intIndex
    mTQuickFilterState.TCmdState(intIndex).blSingleChoose = blSingleChoose
    mTQuickFilterState.TCmdState(intIndex).strCustomScript = strCustomScript
    mTQuickFilterState.TCmdState(intIndex).strMenuSql = strMenuSql
    mTQuickFilterState.TCmdState(intIndex).strName = strName
    
    mTQuickFilterState.TCmdState(intIndex).strKey = mSqlScheme.GetFilterExtValue(strName, CONST_EXT_FILTERKEY)
    If mTQuickFilterState.TCmdState(intIndex).strKey = "" Then mTQuickFilterState.TCmdState(intIndex).strKey = strName
    
    '生成子菜单项
    '100 * (intIndex - 1) + i分配给这个快速过滤项的ID
    
    If intCount = 0 Then
        CreatFilterMenu = True
        Exit Function
    End If
    ReDim mTQuickFilterState.TCmdState(intIndex).cmdItem(intCount)
    
    If blFromSQL Then
    '来源是SQL查询
        If rsTemp.Fields.Count = 1 Then
        '一个字段的情况，需要考虑字段中间包含"-"的处理  "-"前面是过滤内容   后面是显示内容
            '注意strMeneShow strMenuValue和的值和用途
            rsTemp.MoveFirst
            While Not rsTemp.EOF
                i = rsTemp.AbsolutePosition
                strItemName = NVL(rsTemp.Fields(0).value)
                strMenuValue = strItemName
                If InStr(1, strItemName, "-") = 0 Then
                    strMeneShow = strItemName
                    strMenuValue = strItemName
                Else
                    strMeneShow = Mid(strItemName, InStr(1, strItemName, "-") + 1, Len(strItemName) - InStr(1, strItemName, "-"))
                    strMenuValue = Mid(strItemName, 1, InStr(1, strItemName, "-") - 1)
                End If
                
                Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, 100 * (intIndex - 1) + i, strMeneShow)
                cbrPopControl.Parameter = strMenuValue
                
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose = False
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).intItemIndex = i
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName = strMeneShow
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strFilterValue = strMenuValue
                cbrPopControl.CloseSubMenuOnClick = False
            
                rsTemp.MoveNext
            Wend
        Else
            '两个字段的情况，第一个字段是过滤内容，第二个字段是显示内容
            '注意 strMenuValue strItemName 的值和用途
            rsTemp.MoveFirst
            While Not rsTemp.EOF
                i = rsTemp.AbsolutePosition
                strMenuValue = NVL(rsTemp.Fields(0).value)
                strItemName = NVL(rsTemp.Fields(1).value)
                
                Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, 100 * (intIndex - 1) + i, strItemName)
                cbrPopControl.Parameter = strMenuValue
                
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose = False
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).intItemIndex = i
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName = strItemName
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strFilterValue = strMenuValue
                cbrPopControl.CloseSubMenuOnClick = False
            
                rsTemp.MoveNext
            Wend
        End If
    Else
    '固定来源
        ReDim mTQuickFilterState.TCmdState(intIndex).cmdItem(intCount)
    
        For i = 1 To intCount
        
            strItemName = Split(strMenuSql, ";")(i - 1)
            strMenuValue = Split(strMenuSql, ";")(i - 1)
            If InStr(1, strItemName, "-") = 0 Then
                strMeneShow = strItemName
                strMenuValue = strItemName
            Else
                strMeneShow = Mid(strItemName, InStr(1, strItemName, "-") + 1, Len(strItemName) - InStr(1, strItemName, "-"))
                strMenuValue = Mid(strItemName, 1, InStr(1, strItemName, "-") - 1)
            End If
            
            If strMeneShow = "" Then
                strMeneShow = "无效条件"
            End If
    
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, 100 * (intIndex - 1) + i, strMeneShow)
            cbrPopControl.Parameter = strMenuValue
            
            mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose = False
            mTQuickFilterState.TCmdState(intIndex).cmdItem(i).intItemIndex = i
            mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName = strMeneShow
            mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strFilterValue = strMenuValue
            cbrPopControl.CloseSubMenuOnClick = False
        Next
    End If

    CreatFilterMenu = True
    Exit Function
errH:
    err.Raise -1, "创建快速过滤菜单异常" & err.Description
End Function

Private Sub SubGetSysPar(ByVal strParName As String, ByRef value As Variant)
On Error GoTo errH
''[系统.系统号],[系统.模块号],[系统.科室ID],[系统.用户ID],[系统.用户账号]
'[系统.服务器日期],[系统.服务器时间],[系统.本地日期],[系统.本地时间]
'[系统.开始日期],[系统.结束日期]"
'
'[系统.病人ID],[系统.医嘱ID]
    Select Case strParName
        Case "系统.系统号"
            value = glngSys
            
        Case "系统.模块号"
            value = mlngModule
            
        Case "系统.科室ID"
            value = mstrDepartId
        
        Case "系统.用户ID"
            value = UserInfo.ID
            
        Case "系统.用户账号"
            value = UserInfo.用户名
            
        Case "系统.用户名称"
            value = UserInfo.姓名
            
        Case "系统.服务器日期"
            value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
            
        Case "系统.服务器时间"
            value = Format(zlDatabase.Currentdate, "HH:mm:ss")
            
        Case "系统.本地日期"
            value = date
            
        Case "系统.本地时间"
            value = time
            
        Case CONST_EXT_PAR_SYNCNEWADVICE
            value = mstrTemp_新增医嘱ID
            
        Case Else
    End Select
errH:
End Sub

Public Sub ExeColStatistics()
    Call ColStatistics
End Sub

Private Function GetFilterMenuCaption(ByVal strOld As String) As String
'功能：返回处理后的过滤菜单名称
'规则：宽度为7 不足空格补齐 超过4的部分使用"..."表示
    Dim strTmp As String
    Dim i As Integer, j As Integer

    strTmp = strOld
    If Len(strOld) > 4 Then
        strTmp = Mid(strOld, 1, 4) & "..."
    End If
    
    i = 7 - Len(strTmp)
    
    For j = 1 To i
        strTmp = strTmp & " "
    Next
    
    GetFilterMenuCaption = strTmp
    Exit Function
End Function



Private Function IsPatiFilterCanFind(ByVal blUsePatiID As Boolean, ByVal strName As String) As Boolean
'pati控件选中的下拉项目是否支持查找。
'strName 过滤项目名称
'blUsePatiID 是否使用病人ID
On Error GoTo errH
    Dim i As Integer
    
    IsPatiFilterCanFind = False
    
    If blUsePatiID Then
        IsPatiFilterCanFind = True
    Else
        With mObjQuery.objSqlParse.SqlStruct
            
            For i = 1 To .ParCount
                If InStr(1, .AllParameter(i), "[" & strName & "]") > 0 Or InStr(1, .AllParameter(i), "." & strName & "]") > 0 Then
                    IsPatiFilterCanFind = True
                    Exit Function
                End If
            Next
    
        End With
    End If
    
    Exit Function
errH:
    IsPatiFilterCanFind = False
End Function

Private Sub UpdateListRow(ByRef rsData As Recordset, ByVal lngAdviceId As Long, ByVal intType As Integer)
'rsData 一般用于后台刷新，后台刷新操作中需要新增的检查
'lngAdviceID 进行操作的检查
'intType 操作类型
'注意，使用前确保rsData 只有一条记录 或者已经move到需要的记录

On Error GoTo errH
    Dim i As Long
    Dim strColName As String
    Dim lngRow As Long

    If intType = EUpDateRowType.新增 Or intType = EUpDateRowType.强制显示 Then
        
        lngRow = mobjList.Rows
        Call mobjList.AddItem("", lngRow)
        mobjList.TextMatrix(lngRow, 0) = lngRow

    ElseIf intType = EUpDateRowType.更新 Then
        
        lngRow = mobjList.FindRow(lngAdviceId, 1, mobjList.ColIndex("医嘱ID"))
        If lngRow = -1 Then Exit Sub
    Else
    '后台刷新
        lngRow = mobjList.FindRow(lngAdviceId, 1, mobjList.ColIndex("医嘱ID"))
        If lngRow = -1 Then
            lngRow = 1
            Call mobjList.AddItem("", lngRow)
            mobjList.Row = mobjList.Row + 1
        End If
    End If
    
    'lngAdviceID=0后台刷新操作
    '这一段处理用于保证记录集先定位到需要的记录在更新到列表中
    If lngAdviceId > 0 Then
        If rsData.Fields("医嘱ID") <> lngAdviceId Then
            Call rsData.Find("医嘱ID=" & lngAdviceId, , , 1)
        End If
    End If
    
    For i = 1 To mobjList.Cols - 1
        strColName = mobjList.TextMatrix(0, i)
        If strColName <> "" Then mobjList.TextMatrix(lngRow, i) = NVL(rsData.Fields(strColName).value)
    Next
    
    Call RefreshRowRelation(lngRow)
    RaiseEvent OnDoStateImage(lngRow)
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "UpdateListRow" & vbCrLf & err.Description
    Resume
End Sub

Private Sub SyncRst(ByRef rsDataNeedAdd As Recordset, ByRef rsDataSource As Recordset, ByRef blnIsAdd As Boolean)
'功能：在 rsDataNeedAdd 中同步更新 ,若已经存在，则更新，不存在，则新建。rsDataSource 当前一条记录  直接操作，记录集基本容错判断在调用前执行
'注意，本过程里面不要 执行movefirst
On Error GoTo errH
    Dim i As Integer
    Call rsDataNeedAdd.Find("医嘱ID=" & rsDataSource!医嘱ID, , , 1)

    blnIsAdd = rsDataNeedAdd.AbsolutePosition > adPosUnknown

    If rsDataNeedAdd.AbsolutePosition < 1 Then
        rsDataNeedAdd.AddNew
        For i = 0 To rsDataSource.Fields.Count - 1
            rsDataNeedAdd.Fields(i).value = rsDataSource.Fields(i).value
        Next
        rsDataNeedAdd.AbsolutePosition = 1
        rsDataNeedAdd.Update
        Exit Sub
    End If
    
    For i = 0 To rsDataSource.Fields.Count - 1
        If rsDataNeedAdd.Fields(i).Name = rsDataSource.Fields(i).Name Then
            rsDataNeedAdd.Fields(i).value = rsDataSource.Fields(i).value
        Else
            rsDataNeedAdd.Fields(rsDataSource.Fields(i).Name).value = rsDataSource.Fields(i).value
        End If
    Next
    rsDataNeedAdd.Update

    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "SyncRst" & vbCrLf & err.Description
    Resume
End Sub

Private Sub SetListColWidth()
On Error GoTo errH
    If mobjList.Rows < 11 Then
        mobjList.ColWidth(0) = mObjOwner.TextWidth("XX")
    ElseIf 10 < mobjList.Rows And mobjList.Rows < 101 Then
        mobjList.ColWidth(0) = mObjOwner.TextWidth("XXX")
    ElseIf 100 < mobjList.Rows And mobjList.Rows < 1001 Then
        mobjList.ColWidth(0) = mObjOwner.TextWidth("XXXX")
    ElseIf 1000 < mobjList.Rows And mobjList.Rows < 10001 Then
        mobjList.ColWidth(0) = mObjOwner.TextWidth("XXXXX")
    Else
        mobjList.ColWidth(0) = mObjOwner.TextWidth("XXXXXX")
    End If
errH:
End Sub

Private Sub DoListConfig(rsData As ADODB.Recordset, ByRef curCfg() As clsQueryPar.TColConfig)
'获取配置信息
'首次执行查询后执行

On Error GoTo errH
    Dim i As Long
    Dim j As Long
    Dim strName As String
    Dim newCfg() As clsQueryPar.TColConfig
    Dim blnIsFindSysCfg As Boolean
    Dim strCol() As String
    Dim blMatch As Boolean
    
    If rsData Is Nothing Then Exit Sub
    If rsData.Fields.Count <= 0 Then Exit Sub
    
    blMatch = True
    If Len(mstrSchemeCfg.strListCfg) > 0 Then
        strCol = Split(mstrSchemeCfg.strListCfg, "|")
        
        If UBound(mColConfig) <> rsData.Fields.Count - 1 Then
            blMatch = False
        End If

        For i = 1 To rsData.Fields.Count
            If rsData.Fields(i - 1).Name <> Split(strCol(i - 1), ",")(0) Then
                blMatch = False
                Exit For
            End If
        Next
    Else
        blMatch = False
    End If
    
    ReDim newCfg(rsData.Fields.Count - 1)
    
    For i = 0 To rsData.Fields.Count - 1
        strName = rsData.Fields(i).Name
        
        newCfg(i).strName = strName
        newCfg(i).lngColOrder = i + 1
        newCfg(i).lngCfgOrder = 0
        
        blnIsFindSysCfg = False
        For j = 1 To mSqlScheme.ShowCfgCount
            If mSqlScheme.ShowCfg(j).Name = strName Then
                newCfg(i).lngCfgOrder = j
                blnIsFindSysCfg = True
                Exit For
            End If
        Next
        
        If blnIsFindSysCfg Then
        '有系统配置时需要判断是否隐藏列 是否隐藏标题
            newCfg(i).blnHideTitle = mSqlScheme.ShowCfg(j).HiddenTitle
            newCfg(i).blnIsSysHide = mSqlScheme.ShowCfg(j).HiddenCol
            newCfg(i).blnHiddenData = mSqlScheme.ShowCfg(j).HiddenData
        Else
            newCfg(i).blnHideTitle = False
            newCfg(i).blnIsSysHide = False
            newCfg(i).blnHiddenData = False
        End If
        
        newCfg(i).lngWidth = GetExtraWidth(strName) + 1.3 * mObjOwner.TextWidth(strName)
        
        If blMatch Then
            newCfg(i).lngColOrder = Split(strCol(i), ",")(3)
            newCfg(i).lngWidth = Split(strCol(i), ",")(1)
            
            newCfg(i).blnIsUserHide = False
            If Not newCfg(i).blnIsSysHide Then
                newCfg(i).blnIsUserHide = Split(strCol(i), ",")(2)
            End If
        End If
        
    Next
    
    curCfg = newCfg
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "DoListConfig" & vbCrLf & err.Description
End Sub

Private Sub DoListColsOrder(objList As VSFlexGrid, objColCfg() As clsQueryPar.TColConfig)
'功能： 调整列头顺序、宽度、是否显示，列图标，据隐藏显示， 隐藏标题
On Error GoTo errH
    Dim i As Long, j As Long, k As Long
    Dim ObjScShowCfg As New clsScShowCfg
    Dim intubound As Long
    Dim blchange As Boolean
    Dim intCfgIndex As Integer

    With objList
        intubound = UBound(objColCfg)
        For i = 1 To .Cols - 1      'for 1   i
        
            For j = 0 To intubound 'for 2  j
                If .TextMatrix(0, i) = objColCfg(j).strName Then
                    
                    If i <> objColCfg(j).lngColOrder And objColCfg(j).lngColOrder <> 0 Then
                        For k = 0 To intubound     'for 3 k
                            If objColCfg(k).lngColOrder = i Then
                                .ColPosition(i) = (.ColIndex(objColCfg(k).strName))
                                .ColWidth(i) = objColCfg(k).lngWidth
                                .ColHidden(i) = objColCfg(k).blnIsSysHide Or objColCfg(k).blnIsUserHide
                                intCfgIndex = objColCfg(k).lngCfgOrder
                                blchange = True
                                Exit For   'for 3 k
                            End If
                        Next
                    Else
                        .ColWidth(i) = objColCfg(j).lngWidth
                        .ColHidden(i) = objColCfg(j).blnIsSysHide Or objColCfg(j).blnIsUserHide
                        intCfgIndex = objColCfg(j).lngCfgOrder
                    End If
                    
                    If intCfgIndex > 0 And Not blchange Then
                        Set ObjScShowCfg = mSqlScheme.ShowCfg(intCfgIndex)
                        '列图标
                        If Len(ObjScShowCfg.Icon) > 0 Then Set .Cell(flexcpPicture, 0, i) = GetIcon(ObjScShowCfg.Icon)
            
                        '隐藏标题
                        If ObjScShowCfg.HiddenTitle Then
                            .Cell(flexcpData, 0, i) = "1|" & ObjScShowCfg.Name
                            .Cell(flexcpForeColor, 0, i) = .BackColorFixed '用于隐藏标题
                        End If

                        '隐藏数据显示 统一处理 在前方增加空格
                        If ObjScShowCfg.HiddenData Then
                            For k = 1 To .Rows - 1
                                .Cell(flexcpText, k, i) = "                                " & .Cell(flexcpText, k, i)
                            Next
                        End If
                    End If
                    
                    '如果 列顺序 改变（缓存中的列往前移动，需要重新判断地洞后当前i的 字段，否则可能漏掉）
                    If blchange Then
                        i = i - 1
                        blchange = False
                    End If
                                
                    Exit For   'for2 j
                End If
            Next
        Next
    End With
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "DoListColsOrder" & vbCrLf & err.Description
    Resume
End Sub

Private Function DelRow(ByVal lngRowIDOld As Long) As Boolean
'DelRow 返回: 列表中是否还有数据
On Error GoTo errH
    Dim i As Long
    Dim lngAdviceIDOld As Long
    
    Call mobjList.RemoveItem(lngRowIDOld)
    For i = 1 To mobjList.Rows - 1
        mobjList.TextMatrix(i, 0) = i
    Next

    If mobjList.Rows < 2 Then
        DelRow = False
        RaiseEvent OnClearFace
        Exit Function
    End If
    
    DelRow = True
    
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "DelRow" & vbCrLf & err.Description
    Resume
End Function

Private Function GetIsPacsBabyByAge(ByVal strAge As String) As Boolean
'根据年龄判断年龄是否小于一岁
On Error GoTo errH
    If strAge <> "" Then
        Dim lngDays As Long
    
        If strAge Like "*岁*月" Or strAge Like "*岁*个月" Then
            lngDays = 365 * Val(strAge) + 30 * Val(Mid(strAge, InStr(1, strAge, "岁") + 1))
        ElseIf strAge Like "*月*天" Or strAge Like "*个月*天" Then
            lngDays = 30 * Val(strAge) + Val(Mid(strAge, InStr(1, strAge, "月") + 1))
        ElseIf strAge Like "*岁" Or IsNumeric(strAge) Then
            lngDays = 365 * Val(strAge)
        ElseIf strAge Like "*月" Or strAge Like "*个月" Then
            lngDays = 30 * Val(strAge)
        ElseIf strAge Like "*天" Then
            lngDays = Val(strAge)
        Else
            lngDays = 1
        End If
    
        GetIsPacsBabyByAge = lngDays < 365
    Else
        GetIsPacsBabyByAge = False
    End If
    Exit Function
errH:
    GetIsPacsBabyByAge = False
End Function

Private Function GetIsPacsBabyByBirthDay(ByVal strDay As Date) As Boolean
'根据出生日期判断年龄是否小于一岁
On Error GoTo errH
    GetIsPacsBabyByBirthDay = False
    If IsDate(strDay) Then
        If DateAdd("d", -365, zlDatabase.Currentdate) < strDay Then
            GetIsPacsBabyByBirthDay = True
        End If
    End If
    Exit Function
errH:
    GetIsPacsBabyByBirthDay = False
End Function

Public Sub DohistoryFollow(ByVal LngAdvice As Long, ByVal X As Long, ByVal Y As Long)
'mobjPicHistory mobjtxtHistory mobjTimHistory
'ByVal lngAdvice As Long, ByVal X As Long, ByVal Y As Long, ByVal blnClera As Boolean
On Error GoTo errH
    Dim pPoint As POINTAPI
    
    '每次点击重新计时
    mobjTimHistory.Enabled = False
    mobjPicHistory.Visible = False

    Call GetCursorPos(pPoint)
    Call ScreenToClient(mObjOwner.hwnd, pPoint)
    Call mobjPicHistory.Move(Screen.TwipsPerPixelX * pPoint.X + 300, Screen.TwipsPerPixelY * pPoint.Y + 100)
    

    mobjPicHistory.Visible = True
    mobjTimHistory.Enabled = True
    Exit Sub
errH:
    If App.LogMode = 0 Then MsgBox err.Description
End Sub

Public Sub ClearHistoryFollow(ByVal blnTimerStart As Boolean)
'mobjPicHistory mobjtxtHistory mobjTimHistory
On Error GoTo errH
    If blnTimerStart Then
        mobjTimHistory.Enabled = True
    Else
        mobjPicHistory.Visible = False
    End If
    Exit Sub
errH:
    If App.LogMode = 0 Then MsgBox err.Description
End Sub

Private Sub FillNo(ByVal strFilter As String)
On Error GoTo errH
    Dim strTemp As String
    
    '单据号补全处理
    If mobjFindPati.GetCurCard.名称 = "单据号" Then
        If Len(Trim(strFilter)) > 0 Then
            If Len(Trim(strFilter)) < 8 And Not IsNumeric(Trim(strFilter)) Then
                strTemp = GetFullNO(0, 0)
                strTemp = Mid(strTemp, 1, Len(strTemp) - Len(strFilter)) & strFilter
            Else
                strTemp = GetFullNO(NVL(strFilter, 0), 0)
            End If
            mblnAssignment = True
            mobjFindPati.Text = strTemp
            mblnAssignment = False
        End If
    End If
    Exit Sub
errH:
    mblnAssignment = True
    mobjFindPati.Text = strFilter
    mblnAssignment = False
    If App.LogMode = 0 Then MsgBox err.Description
End Sub

Private Sub DoQueueSelRoom(ByVal ingIndex As Long)
'排队队列执行间相关处理
'给这种结构的数据
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As Recordset
    Dim strRoom As String ' "执行间1"  或者 "执行间1,执行间2" 这种格式
    Dim i As Integer
    
    mstrSelQueueRooms = ""
    Call GetQuickFilterSQLPar(ingIndex)
    
    strRoom = "," & mTQuickFilterState.TCmdState(ingIndex).strRelationValueForVBSFilter & ","
    If strRoom = "" Then
        For i = 1 To mTQuickFilterState.TCmdState(ingIndex).intItemCount
            strRoom = "," & mTQuickFilterState.TCmdState(ingIndex).cmdItem(i).strName
        Next
        strRoom = strRoom & ","
    End If

    strSQL = "Select 名称,执行间 From 医技执行房间 a, 部门表 b Where a.科室Id=b.Id and instr([1],','||执行间||',')> 0  and instr([2],','||科室ID||',')> 0 "
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "后台检查数据刷新", strRoom, "," & mstrDepartId & ",")
    
    While Not rsData.EOF
    
        If mstrSelQueueRooms <> "" Then mstrSelQueueRooms = mstrSelQueueRooms & ","
        mstrSelQueueRooms = mstrSelQueueRooms & NVL(rsData!名称) & "-" & NVL(rsData!执行间)
        
        rsData.MoveNext
    Wend
    
    RaiseEvent OnQueueRoomChanged
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub RefreshDisplay(ByVal lngRowIDOld As Long, ByVal lngAdviceId As Long, ByVal intType As Integer, Optional ByVal blnRefreshModul As Boolean = False, Optional ByVal blnForceNoRefreshModul As Boolean = False)
'1 更新检查信息
'2 用新的检查信息刷新界面相关信息
'intType update 中使用时，EUpDateRowType 操作类型
'lngRowIDOld需要更新信息的列
'lngAdviceID需要更新信息的医嘱ID
'blnRefreshModul是否刷新Modul
'blnForceNoRefreshModul 强制不刷新
'刷新模块操作说明，传入参数 blnRefreshModul  blnForceNoRefreshModul ，考虑执行状态 和执行过程
' blnForceNoRefreshModul =TRUE 必定不刷新    否则 如果 blnRefreshModul=true 或者操作前后执行过程/状态之一发生改变则会刷新Modul
On Error GoTo errH
    Dim blnRefresh As Boolean
    Dim intOldStep As Integer
    Dim intOldState As Integer
    
    blnRefresh = False
    
    intOldStep = mobjCurStudyInfo.intStep
    intOldState = mobjCurStudyInfo.intState
    
    If lngRowIDOld > 0 Then
        mobjList.Cell(flexcpBackColor, lngRowIDOld, 0) = C_COLOR_LISTCOL0
        Set mobjCurStudyInfo = GetBaseInfo(lngAdviceId, GetMovedState(lngRowIDOld, mobjList))
        If mobjList.Cell(flexcpData, lngRowIDOld) = Empty Then
            If mobjCurStudyInfo.blnMoved Then
                mobjList.Cell(flexcpData, lngRowIDOld) = 2
            Else
                mobjList.Cell(flexcpData, lngRowIDOld) = 1
            End If
        End If
    Else
        Set mobjCurStudyInfo = GetBaseInfo(lngAdviceId, 0)
    End If

    If blnRefreshModul Or intOldStep <> mobjCurStudyInfo.intStep Or intOldState <> mobjCurStudyInfo.intState Then blnRefresh = True
    If blnForceNoRefreshModul Then blnRefresh = False
    
    RaiseEvent OnChangeData(blnRefresh, False)
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "RefreshDisPlay" & vbCrLf & err.Description
    Resume
End Sub

Private Sub RefreshTabBottom(ByVal ObjSqlScheme As clsSqlScheme, ByRef objTab As TabControl)
'刷新 列表下方tab控件
On Error GoTo errH
    Dim i As Integer
    Dim blnSelectNextVisibleItem As Boolean
    
    If ObjSqlScheme.EmbedFind Then
        objTab.Item(0).Visible = True
    Else
        If objTab.Item(0).Selected Then blnSelectNextVisibleItem = True
        objTab.Item(0).Visible = False
    End If

    If blnSelectNextVisibleItem Then
        For i = 0 To objTab.ItemCount - 1
            If objTab.Item(i).Visible = True Then
                objTab.Item(i).Selected = True
                Exit Sub
            End If
        Next
    End If
    Exit Sub
errH:
    err.Raise -1, "clsPacsQueryWrap", "RefreshTabBottom" & vbCrLf & err.Description
    Resume
End Sub

Public Sub RefreshTabLeft(ByRef tabScheme As TabControl, ByVal strChooseItem As String)
'切换方案后刷新改变快速方案显示的选中项目，比如加载默认方案，菜单切换方案
On Error GoTo errH
    Dim i As Integer
    Dim blnSelectHideItem As Boolean
    blnSelectHideItem = True
    
    For i = 1 To tabScheme.ItemCount - 1
        If tabScheme.Item(i).Caption = strChooseItem Then
            tabScheme.Item(i).Selected = True
            blnSelectHideItem = False
        End If
    Next
    
    If blnSelectHideItem Then tabScheme.Item(0).Selected = True
    Exit Sub
errH:
    If App.LogMode = 0 Then MsgBox err.Description
End Sub

Private Sub initTabScheme(ByRef tabScheme As TabControl, ByVal strItems As String, ByVal blnShowLeftTab As Boolean)
'blnShowLeftTab tab参数不启用
On Error GoTo errH
    Dim i As Integer
    Dim iCount As Integer
    Dim strName() As String
    
    Dim CtlFont As StdFont
    
    '用于表示
    If strItems = "" Or Not blnShowLeftTab Then
        tabScheme.tag = "0"
        tabScheme.Visible = False
    Else
        tabScheme.tag = "1"
    End If
        
    strName = Split(strItems, "|")
    iCount = UBound(strName)
    
    With tabScheme
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionLeft
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.ColorSet.ButtonNormal = &HE0E0E0
        
        .RemoveAll
        
        .InsertItem 1, "隐藏tab", mobjPicTmp.hwnd, 0
        .Item(tabScheme.ItemCount - 1).tag = "-1"
        .Item(tabScheme.ItemCount - 1).Visible = False
        
        For i = 1 To iCount + 1
            .InsertItem i + 1, Split(strName(iCount + 1 - i), "%")(0), mobjPicTmp.hwnd, 0
            .Item(tabScheme.ItemCount - 1).tag = Split(strName(iCount + 1 - i), "%")(1)
        Next
        
    End With
    Exit Sub
errH:
End Sub

Private Sub DoBackColorSel()
'处理透明行
On Error GoTo errH
    Dim lngRow As Long
    Dim intCol As Integer

    If Not mblnSelRowTransparent Then Exit Sub
    intCol = mobjList.ColIndex("医嘱ID")
    If intCol = -1 Then Exit Sub
    If mobjList.Rows < 2 Then Exit Sub
    
    lngRow = mobjList.Row
    mobjList.BackColorSel = mobjList.Cell(flexcpBackColor, lngRow, intCol)
    
    Exit Sub
errH:
    If App.LogMode = 0 Then MsgBox "透明" & err.Description
End Sub

Private Sub SetPatiTyprColor(ByVal lngRow As Long, ByVal IsSetList As Boolean)
'处理病人类型颜色
On Error GoTo errH
    Dim lngCol As Long
    Dim strType As String
    Dim lngColor As Long
    Dim lngSetRow As Long
    
    If mobjlabPatiInfo Is Nothing Then Exit Sub
    
    If Not IsSetList Then mobjlabPatiInfo.ForeColor = vbBlack
    
    If mSqlScheme.PatiColor = "" Then Exit Sub
    lngSetRow = mobjList.ColIndex(mSqlScheme.PatiColor)
    lngCol = mobjList.ColIndex("病人类型")
    
    If lngCol = -1 Or lngSetRow = -1 Then Exit Sub
    
    strType = mobjList.TextMatrix(lngRow, lngCol)
    If strType = "" Then Exit Sub
    
    lngColor = mobjPublicPatient.GetPatiColor(strType)
    
    If IsSetList Then
        mobjList.Cell(flexcpForeColor, lngRow, lngSetRow) = lngColor
    Else
        mobjlabPatiInfo.ForeColor = lngColor
    End If
    Exit Sub
errH:
    If App.LogMode = 0 Then MsgBox "SetPatiTyprColor" & err.Description
End Sub

Private Sub InitPatiComponent()
On Error GoTo errhandle
    Set mobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
    
    If Not mobjPublicPatient Is Nothing Then Call mobjPublicPatient.zlInitCommon(gcnOracle, glngSys)
 
Exit Sub
errhandle:
    Set mobjPublicPatient = Nothing
End Sub

Private Function IsNumberSort(ByVal strColName As String) As Boolean
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
'获取排序列TODO
    IsNumberSort = False
    
    If mTColSort.dictNumberSort Is Nothing Then Exit Function
    If mTColSort.dictNumberSort.Exists(strColName) Then
        IsNumberSort = mTColSort.dictNumberSort.Item(strColName)
    End If
    
    Exit Function
errH:
    err.Raise -1, "clsPacsQueryWrap", "[IsNumberSort]" & vbCrLf & err.Description
'    Resume
End Function











