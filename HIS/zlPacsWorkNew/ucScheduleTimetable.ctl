VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.UserControl ucScheduleTimetable 
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   ScaleHeight     =   6480
   ScaleWidth      =   9240
   ToolboxBitmap   =   "ucScheduleTimetable.ctx":0000
   Begin VB.Timer timColor 
      Interval        =   100
      Left            =   480
      Top             =   6120
   End
   Begin VB.CommandButton btnSchLabel 
      Height          =   180
      Index           =   0
      Left            =   1080
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   90
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfTime 
      Height          =   5835
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8535
      _cx             =   15055
      _cy             =   10292
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   25
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Menu menuTimePopup 
      Caption         =   "时间表右键菜单"
      Begin VB.Menu menuPopupTimeProjectAdd 
         Caption         =   "新增时间计划"
      End
      Begin VB.Menu menuPopupTimeProjectModi 
         Caption         =   "修改时间计划"
      End
      Begin VB.Menu menuPopupTimeProjectDel 
         Caption         =   "删除时间计划"
      End
      Begin VB.Menu menuPopupTimeSplit 
         Caption         =   "-"
      End
      Begin VB.Menu menuPopupTimeProjectColor 
         Caption         =   "时间表颜色设置"
      End
   End
   Begin VB.Menu menuSchedulePopup 
      Caption         =   "预约标签右键菜单"
      Begin VB.Menu menuPopupScheduleModi 
         Caption         =   "修改预约"
      End
      Begin VB.Menu menuPopupSchedulePrint 
         Caption         =   "打印预约单"
      End
   End
End
Attribute VB_Name = "ucScheduleTimetable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
'-------------------------------------------------------------
'   控件的公共事件
'-------------------------------------------------------------
Public Event OnMenuTimeProjectAdd()
Public Event OnMenuTimeProjectModify(ByVal lngTimeProjectID As Long)
Public Event OnMenuTimeProjectBeforeDel(ByRef blnCancel As Boolean)
Public Event OnMenuTimeProjectSetColor()
Public Event OnMenuScheduleModify()
Public Event OnMenuSchedulePrint()
Public Event OnChangeOrder(ByVal lngOrderID As Long, ByVal strOrderInfo As String)  '更改检查医嘱的标签
Public Event OnSchLabelModifed(ByVal iIndex As Integer)                             '预约标签被移动或者修改长度
'-------------------------------------------------------------
'   控件的私有属性
'-------------------------------------------------------------
Private mblnDragingLabel As Boolean     '正在拖拽预约标签的标记
Private mlngOriginLeft As Long          '标签拖拽之前，标签组中第一个标签，所在的Left位置
Private mlngOriginTop As Long           '标签拖拽之前，标签组中第一个标签，所在的Top位置
Private mlngOriginMinute As Long        '标签拖拽之前，时间长度
Private mlngBaseX As Long               '移动标签的基准X，对于标签组，不是MouseDown的X
Private mlngBaseY As Long               '移动标签的基准Y
Private mlngDownX As Long               'MouseDown时的X
Private mblnRestorePos As Boolean       '预约标签是否需要回到原来的位置

Private mblnSizingLabel As Boolean      '正在调整预约标签的宽度
Private mlngOriginWidth As Long         '调整之前，当前预约标签的宽度

Private mlngColorTabRest As Long        '预约时间表，休息时间颜色
Private mlngColorTabWork As Long        '预约时间表，工作时间颜色
Private mlngColorLblWaiting As Long     '预约标签，预约等候颜色
Private mlngColorLblDone As Long        '预约标签，完成颜色
Private mlngColorLblPassed As Long      '预约标签，过号颜色

Private mlngCurrTimePorjectID As Long   '当前鼠标所在位置的时间计划ID
Private mlngSchPlanID  As Long          '当前时间表的预约方案ID
Private mlngSchDeviceID As Long         '当前预约设备ID
Private mlngOrderID As Long             '当前医嘱ID
Private mdtSchDate As Date              '当前日期
Private mblnMoved As Boolean            '是否已经转储
Private mstrOrderInfo As String         '当前医嘱的预约信息
Private mColordict As New Dictionary
Private mlngCol As Long                 '当前鼠标所在位置的列号
Private mlngRow As Long                 '当前鼠标所在位置的行号

Private mlngMouseX As Long               '用于鼠标点击设置时间功能X坐标
Private mlngMouseY As Long               '用于鼠标点击设置时间功能Y坐标

Private mlngSchLabelIndex As Long       '预约标签的最大索引
Private mlngPoolIndex As Long           '当前选中的预约池索引
Private mlngBtnIndex As Long            '当前选中的预约标签索引和预约池索引保持一致
Private mrsCalendar As ADODB.Recordset  '保存预约日历数据集
Private mIsReadOnly As Boolean          '是否只读模式
Private mstrModifiedOrderID As String   '保存过预约信息的医嘱ID串，用“,”连接
Private mlngFontSize As Long            '保存预约标签字体的大小

Private mlngTimeProjectID As Long '当前时间块所属时间ID,点击鼠标清0

Private Const SchPreTime = 5            '当天预约的提前时间，5分钟
'预约控件使用模式
Private Enum constUseType
    Sch_UseType_检查预约 = 1            '只有当前新建的预约标签可以使用
    Sch_UseType_预约管理 = 2            '所有预约标签都可以使用
    Sch_UseType_预约设置 = 3            '不显示预约标签，只显示时间计划
End Enum
Private mlngUseType As constUseType     '记录当前的预约模式

'检查预约方案类型
Private Enum constSchedulePlanType
    Sch_PlanType_每天 = 1
    Sch_PlanType_每周 = 2
    Sch_PlanType_每月 = 3
    Sch_PlanType_一天 = 4
End Enum

'预约标签的信息
Private Type TYPE_SchLabel
    lng序号 As Long             '预约序号,预约池按照序号来排序
    lngRow As Long              '标签所在的行
    lngCol As Long              '标签所在的列
    lng医嘱ID As Long           '医嘱ID
    str医嘱内容 As String       '医嘱内容
    str姓名 As String           '姓名
    dtStartTime As Date         '预约开始时间
    dtEndTime As Date           '预约结束时间
    lngBtnCount As Long         '一个预约标签包含的按钮数量，拐弯标签包含2个以上按钮
    lngBtnIndex As Long         '预约标签的index，如果是标签组，记录第一个标签的index
    lngTimeProjectID As Long    '时间计划ID
    dt开始时间段 As Date        '开始时间段
    dt结束时间段 As Date        '结束时间段
    isModified As Boolean       '是否已经被修改
    bln已执行 As Boolean        '是否已经执行，病人医嘱发送.执行过程=0 or =1 or =2, 执行过程: -1-驳回；0或1-已登记；2-已报到；3-已检查；4-已报告；5-已审核；6-已完成
    bln已保存 As Boolean        '此预约是否已经保存到数据库
End Type
Private mSchLabelPool() As TYPE_SchLabel

'预约时间计划
Private Type Type_TimeProject
    lngID As Long               '时间计划ID
    lngSchPlanID As Long        '预约方案ID
    dtStartTime As Date         '开始时间
    dtEndTime As Date           '结束时间
    lngSum As Long              '预约容量
    lngCalType As Long          '计算方法
End Type
Private mSchTimeProject() As Type_TimeProject   '按照开始时间排序
Private mlngSchSum As Long      '预约总容量

'-------------------------------------------------------------
'   控件的公共属性
'-------------------------------------------------------------
'设置只读模式
Property Get IsReadOnly() As Boolean
    IsReadOnly = mIsReadOnly
End Property

Property Let IsReadOnly(value As Boolean)
    mIsReadOnly = value
    
    Call setReadOnly
End Property

'当前被选中的时间计划ID
Property Get CurrTimeProjectID() As Long
     CurrTimeProjectID = mlngCurrTimePorjectID
End Property

'当前的预约方案ID
Property Get SchedulePlanID() As Long
    SchedulePlanID = mlngSchPlanID
End Property

'控件的使用模式
Property Get UseType() As Long
    UseType = mlngUseType
End Property

'当前选中预约标签的序号
Property Get Label序号() As Long
    Label序号 = mSchLabelPool(mlngPoolIndex).lng序号
End Property

'当前选中预约标签的开始时间段
Property Get Label开始时间段() As Date
    Label开始时间段 = mSchLabelPool(mlngPoolIndex).dt开始时间段
End Property

'当前选中预约标签的结束时间段
Property Get Label结束时间段() As Date
    Label结束时间段 = mSchLabelPool(mlngPoolIndex).dt结束时间段
End Property

'当前选中预约标签的开始时间
Property Get Label开始时间() As Date
    Label开始时间 = mSchLabelPool(mlngPoolIndex).dtStartTime
End Property

'当前选中预约标签的结束时间
Property Get Label结束时间() As Date
    Label结束时间 = mSchLabelPool(mlngPoolIndex).dtEndTime
End Property

'当前选中预约标签的医嘱ID
Property Get LabelOrderID() As Long
    LabelOrderID = mlngOrderID
End Property

'当前选中预约标签的全部预约信息
Property Get LabelOrderInfo() As String
    LabelOrderInfo = mstrOrderInfo
End Property

'当前选中预约标签的患者姓名
Property Get Label姓名() As String
    Label姓名 = mSchLabelPool(mlngPoolIndex).str姓名
End Property

'保存过预约信息的医嘱ID串，用“,”连接
Property Get strModifiedOrderID() As String
    If Trim(mstrModifiedOrderID) = "" Then
        strModifiedOrderID = mstrModifiedOrderID
    Else
        strModifiedOrderID = Mid(mstrModifiedOrderID, 2)
    End If
End Property

'-------------------------------------------------------------
'   控件的公共方法
'-------------------------------------------------------------
Public Function funSaveSchedule(dtSchStartTime As Date, dtSchEndTime As Date, lngOrderID As Long, _
    strName As String, lngNumber As Long, lngSchDeviceID As Long, dtSegStartTime As Date, _
    dtSegEndTime As Date, Optional strNotice As String = "*Nothing*") As Boolean
'------------------------------------------------
'功能：保存预约信息
'参数： dtSchStartTime -- 预约开始时间
'       dtSchEndTime -- 预约结束时间
'       lngOrderID -- 医嘱ID
'       strName -- 患者姓名
'       lngNumber -- 预约序号
'       lngSchDeviceID -- 预约设备ID
'       dtSegStartTime -- 开始时间段
'       dtSegEndTime -- 结束时间段
'       strNotice -- 检查注意
'返回：True - 成功 ； False - 失败
'------------------------------------------------
    Dim strStartTime As String
    Dim strEndTime As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strNoticeSql As String
    
    On Error GoTo err
    
    '每次保存前，先调用自动删除不合格预约的方法
    strSQL = "Zl_影像预约自动删除"
    zlDatabase.ExecuteProcedure strSQL, "自动删除不合格的预约记录"
    
    '先本次预约时间是否和其他检查存在重复
    '开始时间加一分钟，结束时间减一分钟？
    strStartTime = Format(DateAdd("n", 1, dtSchStartTime), "yyyy-MM-dd hh:mm:ss")
    strEndTime = Format(DateAdd("n", -1, dtSchEndTime), "yyyy-MM-dd hh:mm:ss")
    
    strSQL = "Select 预约设备名称,预约开始时间,预约结束时间 " _
        & " From 影像预约记录 Where 医嘱ID In " _
        & " (Select ID From 病人医嘱记录 Where 病人ID =  " _
        & " (Select 病人ID From 病人医嘱记录 Where ID = [1])) And 医嘱ID <> [1] And " _
        & " (预约开始时间 Between [2] And [3] or 预约结束时间 Between [2] And [3] or " _
        & " [2] between 预约开始时间 and 预约结束时间 ) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询在其他设备上的预约", lngOrderID, _
        CDate(strStartTime), CDate(strEndTime))

    If rsTemp.EOF = False Then
        If MsgBox("患者 " & strName & " 在 " & Format(rsTemp!预约开始时间, "hh:mm:ss") & " 到 " & Format(rsTemp!预约结束时间, "hh:mm:ss") & " 这个时间段已经预约了“" & nvl(rsTemp!预约设备名称) & "”上的检查，" _
            & vbCrLf & vbCrLf & "建议修改本次检查预约的时间，避免发生检查时间冲突。" & vbCrLf & vbCrLf & "是否取消保存患者 " & strName & " 的预约时间？", vbYesNo, "检查预约提示") = vbYes Then
            Exit Function
        End If
    End If

    If strNotice = "*Nothing*" Then
        strNoticeSql = ""
    Else
        strNoticeSql = ",'" & strNotice & "'"
    End If
    
    strSQL = "Zl_影像预约记录_更新(" & lngOrderID & ",'" & lngNumber & "'," _
        & lngSchDeviceID & "," & zlStr.To_Date(dtSegStartTime) _
        & "," & zlStr.To_Date(dtSegEndTime) & "," & zlStr.To_Date(dtSchStartTime) _
        & "," & zlStr.To_Date(dtSchEndTime) & strNoticeSql & ")"
    zlDatabase.ExecuteProcedure strSQL, "保存检查预约"
            
    funSaveSchedule = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub RefreshTimeProject(lngSchPlanID As Long)
'------------------------------------------------
'功能：装载预约时间表
'参数：lngSchPlanID -- 预约方案ID
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strStartTime As String
    Dim strEndTime As String
    Dim lngStartRow As Long
    Dim lngStartCol As Long
    Dim lngEndRow As Long
    Dim lngEndCol As Long
    Dim lngTimeProjectID As Long
    Dim iCount As Integer
    Dim blnChange As Boolean
    
    On Error GoTo err
    
    '清空预约标签
    Call unloadSchLabel
    
    mlngSchPlanID = lngSchPlanID
    
    ReDim Preserve mSchTimeProject(0) As Type_TimeProject
    iCount = 0
    mlngSchSum = 0
    
    strSQL = "select ID,开始时间,结束时间,预约容量,计算方法 from 影像预约时间计划 where 预约方案ID =[1] order by 开始时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "加载预约时间表", mlngSchPlanID)
    
    With vsfTime
        '先设置整个时间表都显示为休息时间
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, .Cols - 1) = mlngColorTabRest
        .Cell(flexcpData, 1, 1, .Rows - 1, .Cols - 1) = 0
        '再逐个显示工作时间
        Set mColordict = New Dictionary
        While rsTemp.EOF = False
            '提取时间
            strStartTime = Format(nvl(rsTemp!开始时间, Now), "hh:mm")
            strEndTime = Format(nvl(rsTemp!结束时间, Now), "hh:mm")
            lngTimeProjectID = rsTemp!ID
            
            '将时间计划保存到变量中
            iCount = iCount + 1
            ReDim Preserve mSchTimeProject(iCount) As Type_TimeProject
            mSchTimeProject(iCount).lngID = rsTemp!ID
            mSchTimeProject(iCount).lngSchPlanID = mlngSchPlanID
            mSchTimeProject(iCount).dtStartTime = strStartTime
            mSchTimeProject(iCount).dtEndTime = strEndTime
            mSchTimeProject(iCount).lngSum = rsTemp!预约容量
            mSchTimeProject(iCount).lngCalType = rsTemp!计算方法
            mlngSchSum = mlngSchSum + mSchTimeProject(iCount).lngSum
            
            
            If strEndTime > strStartTime Then
                Call getRowColFromTime(strStartTime, True, lngStartRow, lngStartCol)
                Call getRowColFromTime(strEndTime, False, lngEndRow, lngEndCol)
                
                '画时间表的背景，需要分成三部分来显示一个时间计划
                If (lngStartRow = lngEndRow) Then
                    '在一个小时之内，只画一行
                    .Cell(flexcpBackColor, lngStartRow, lngStartCol, lngEndRow, lngEndCol) = mlngColorTabWork
                    .Cell(flexcpData, lngStartRow, lngStartCol, lngEndRow, lngEndCol) = lngTimeProjectID
                Else
                    '先画第一行
                    .Cell(flexcpBackColor, lngStartRow, lngStartCol, lngStartRow, .Cols - 1) = mlngColorTabWork
                    .Cell(flexcpData, lngStartRow, lngStartCol, lngStartRow, .Cols - 1) = lngTimeProjectID '
                    If (lngEndRow - lngStartRow = 1) Then
                        '只画两行，画第二行
                        .Cell(flexcpBackColor, lngEndRow, 1, lngEndRow, lngEndCol) = mlngColorTabWork
                        .Cell(flexcpData, lngEndRow, 1, lngEndRow, lngEndCol) = lngTimeProjectID
                    Else
                        '画中间的一段
                        .Cell(flexcpBackColor, lngStartRow + 1, 1, lngEndRow - 1, .Cols - 1) = mlngColorTabWork
                        .Cell(flexcpData, lngStartRow + 1, 1, lngEndRow - 1, .Cols - 1) = lngTimeProjectID
                        '画最后一行
                        .Cell(flexcpBackColor, lngEndRow, 1, lngEndRow, lngEndCol) = mlngColorTabWork
                        .Cell(flexcpData, lngEndRow, 1, lngEndRow, lngEndCol) = lngTimeProjectID
                    End If
                End If
                
                If Not mColordict.Exists(lngTimeProjectID) Then
                    blnChange = Not blnChange
                    Call mColordict.Add(lngTimeProjectID, blnChange)
                End If

            Else
                '如果结束时间小于开始时间，则不画这个时间计划
            End If
            rsTemp.MoveNext
        Wend
    End With
    '进入预约界面没有马上刷新两种颜色
    Call ShowMouseTime(1, 1)
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub Init(lngUseType As Long)
'------------------------------------------------
'功能：从外部调用的初始化方法，控件的部分内容需要从数据库读取数据初始化，要在运行的时候进行
'参数：
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    Call loadColor      '加载颜色
    Call LoadCalendar   '加载日历数据集
    mlngUseType = lngUseType
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function RefreshSchedule(lngSchDeviceID As Long, dtDate As Date, lngOrderID As Long, _
    Optional lngSchPlanID As Long = 0, Optional blnMoved As Boolean = False) As Boolean
'------------------------------------------------
'功能：刷新预约时间表
'参数： lngSchDeviceID -- 预约设备ID
'       dtDate -- 预约时间
'       lngOrderID -- 医嘱ID
'       lngSchPlanID -- 可选参数，0，表示需要重新查询
'       blnMoved -- 可选参数，是否已经转储
'返回：True -- 检查预约模式下，有当前医嘱的预约标签；False -- 检查预约模式下，无当前医嘱的预约标签；或者是预约管理模式
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim strFilter As String
    Dim iIndex As Integer
    Dim intPoolCount As Integer
    
    On Error GoTo err
    
    RefreshSchedule = False
    mlngSchDeviceID = lngSchDeviceID
    mlngOrderID = lngOrderID
    mdtSchDate = dtDate
    mblnMoved = blnMoved
    
    If lngSchDeviceID = 0 Then
        '更新时间计划表的时间方案
        Call RefreshTimeProject(0)
        Exit Function
    End If
    
    If lngSchPlanID = 0 Then
        mlngSchPlanID = GetSchPlanID(mlngSchDeviceID, dtDate, False, False)
    ElseIf lngSchPlanID = -1 Then
        lngSchPlanID = 0
        mlngSchPlanID = lngSchPlanID
    Else
        mlngSchPlanID = lngSchPlanID
    End If
    
    '更新时间计划表的时间方案
    Call RefreshTimeProject(mlngSchPlanID)
    
    '显示当天的预约情况
    '查询预约记录
    strSQL = "Select d.ID, d.医嘱ID, d.序号, d.诊室名称, d.预约开始时间, d.预约结束时间, " _
        & " d.预约开始时间段, d.预约结束时间段, b.姓名, b.医嘱内容, b.婴儿, c.执行过程 " _
        & " From 病人医嘱记录 B, 病人医嘱发送 C,影像预约记录 D Where b.id in " _
        & " (Select  a.医嘱ID From 影像预约记录 A Where a.预约设备ID = [1] And " _
        & " a.预约开始时间 Between [2] And [3] )And c.医嘱id = b.id and d.医嘱id=B.id Order By cast(d.序号 as int)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询今天的预约记录", mlngSchDeviceID, CDate(Format(dtDate, "yyyy-MM-dd 00:00:00")), CDate(Format(dtDate, "yyyy-MM-dd 23:59:59")))
        
    If rsTemp.EOF = False Then
        While rsTemp.EOF = False
            '记录预约信息
            intPoolCount = UBound(mSchLabelPool) + 1
            ReDim Preserve mSchLabelPool(intPoolCount) As TYPE_SchLabel
            
            '提取婴儿信息
            If rsTemp!婴儿 <> 0 Then
                strSQL = "Select A.开嘱时间,Nvl(B.婴儿姓名, A.姓名 || '之子' || Trim(To_Char(B.序号, '9'))) As 婴儿姓名, B.婴儿性别, B.出生时间" & vbNewLine & _
                                 "  From 病人医嘱记录 A, 病人新生儿记录 B " & vbNewLine & _
                                 "  Where a.病人ID = b.病人ID  And b.序号 = [2] And a.ID = [1]"
                Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "提取婴儿信息", CLng(rsTemp!医嘱ID), CLng(rsTemp!婴儿))
                mSchLabelPool(intPoolCount).str姓名 = "婴儿：" & rsBaby!婴儿姓名
            Else
                mSchLabelPool(intPoolCount).str姓名 = rsTemp!姓名
            End If
            mSchLabelPool(intPoolCount).lng序号 = rsTemp!序号
            mSchLabelPool(intPoolCount).lng医嘱ID = rsTemp!医嘱ID
            mSchLabelPool(intPoolCount).dtStartTime = rsTemp!预约开始时间
            mSchLabelPool(intPoolCount).dtEndTime = rsTemp!预约结束时间
            mSchLabelPool(intPoolCount).str医嘱内容 = rsTemp!医嘱内容
            mSchLabelPool(intPoolCount).dt开始时间段 = rsTemp!预约开始时间段
            mSchLabelPool(intPoolCount).dt结束时间段 = rsTemp!预约结束时间段
            mSchLabelPool(intPoolCount).lngBtnCount = 1
            mSchLabelPool(intPoolCount).isModified = False
            mSchLabelPool(intPoolCount).bln已执行 = IIf(nvl(rsTemp!执行过程, 0) = 0 Or nvl(rsTemp!执行过程, 0) = 1 Or nvl(rsTemp!执行过程, 0) = 2, False, True)
            mSchLabelPool(intPoolCount).bln已保存 = True
            '创建一个预约标签
            iIndex = CreateNewSchLabel()
            
            '摆放预约标签
            Call PutSchLabel(iIndex, intPoolCount)
            '
            '如果是检查预约模式，这些从数据库中读取的预约标签，全部都是只读显示
            If mlngUseType = Sch_UseType_检查预约 Then
                If mSchLabelPool(intPoolCount).lng医嘱ID <> mlngOrderID Then
                    Call setSchLabelEnable(iIndex, False)
                Else
                    '如果今天有这个医嘱的预约记录，记录模块变量
                    mlngPoolIndex = intPoolCount
                    mlngBtnIndex = iIndex
                    'Call setSchLabelEnable(iIndex, IIf(mdtSchDate < Format(Now, "yyyy-mm-dd"), False, IIf(mSchLabelPool(intPoolCount).bln已保存 = True, False, True)))
                    Call setSchLabelEnable(iIndex, IIf(mdtSchDate < Format(Now, "yyyy-mm-dd"), False, True))
                    Call setSchLabelToolTipText(iIndex)
                    RefreshSchedule = True  '返回成功
                End If
            ElseIf mlngUseType = Sch_UseType_预约管理 Then
                Call ClearLocalParas
                Call setSchLabelEnable(iIndex, IIf(mdtSchDate < Format(Now, "yyyy-mm-dd"), False, True))
            End If
            rsTemp.MoveNext
        Wend
        If mlngUseType = Sch_UseType_检查预约 And RefreshSchedule = False Then
            Call ClearLocalParas
        End If
    Else
        '没有记录
        ReDim mSchLabelPool(0) As TYPE_SchLabel
        Call ClearLocalParas
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function NewSchedule(ByVal lngSchDeviceID As Long, ByRef dtSchDate As Date, _
    ByVal lngOrderID As Long, ByVal blnChangeDate As Boolean) As Boolean
'------------------------------------------------
'功能：根据医嘱ID，创建一个新的检查预约标签，并且将标签自动放在合适的位置
'参数： lngSchDeviceID -- 预约设备ID
'       dtSchDate -- 预约日期
'       lngOrderID -- 医嘱ID
'       blnChangeDate -- 是否改变预约日期，true--从这天开始寻找最适合的预约日期；False--就在这天预约
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim rsSchedule As ADODB.Recordset
    Dim lngSchPlanID As Long
    Dim i As Integer
    Dim lngNumber As Long       '最小可用的预约序号
    Dim lngBaseNumber As Long   '前面几个时间段可用预约容量总和
    Dim dtLastEndTime As Date   '上一个预约的结束时间
    Dim dtNextStartTime As Date '下一个预约的开始时间
    Dim dtStartTime As Date     '新预约的开始时间
    Dim dtEndTime As Date       '新预约的结束时间
    Dim iIndex As Integer       '新预约标签的索引
    Dim iPoolIndex As Integer   '新预约标签在预约池的索引
    Dim iTimeType As Integer    '预约时长计算方法
    Dim lngTimeLength As Long   '预约时长
    Dim dtBeforeTime As Date    '预约当天的最早时间，如果是今天，则是now+5分钟
    Dim blnFill As Boolean      '是否填满两个序号的空隙时间
    Dim blnTimeOK As Boolean    '查找最新时间完成
    
    On Error GoTo err
    If Format(dtSchDate, "YYYY-MM-DD") < Format(Now, "YYYY-MM-DD") Then
        MsgBox "今天以前的日期，不能新建预约。 ", vbOKOnly, "检查预约提示"
        Exit Function
    End If
    blnFill = False
    
    If blnChangeDate = True Then
        '查找从 这天 开始往后，第一个有预约容量的日子
        lngSchPlanID = FindFirstSchDay(lngSchDeviceID, dtSchDate)
    Else
        lngSchPlanID = GetSchPlanID(lngSchDeviceID, dtSchDate, False, False)
    End If
    
    '找到合适的预约日期后，刷新时间表
    '首先刷新预约时间表，显示当前已经预约好的情况
    If RefreshSchedule(lngSchDeviceID, dtSchDate, lngOrderID, IIf(lngSchPlanID = 0, -1, lngSchPlanID)) = True Then
        Exit Function   '如果这个检查已经有预约，就不再显示新的预约标签
    End If
    
    If lngSchPlanID = 0 Then
        Exit Function
    End If
    
    '根据序号，找到最适合的预约位置
    If UBound(mSchTimeProject) = 0 Then
        MsgBox "无法预约，此预约方案没有时间计划，请先设置好预约时间计划，预约方案ID=" & lngSchPlanID, vbOKOnly, "检查预约提示"
        Exit Function
    End If
    
    If Format(dtSchDate, "YYYY-MM-DD") < Format(Now, "YYYY-MM-DD") Then
        '本段提示移动到最前面
        'MsgBox "今天以前的日期，不能新建预约。 ", vbOKOnly, "检查预约提示"
        'Exit Function
    ElseIf Format(dtSchDate, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
        dtBeforeTime = DateAdd("n", 5, Now)
    Else
        dtBeforeTime = Format(dtSchDate, "YYYY-MM-DD") & " 00:00:00"
    End If
    
    '按照时间顺序，逐个查找时间计划，查找可以预约的最早时间
    '查找最小的空预约序号
    strSQL = "select 序号,预约开始时间,预约结束时间 from 影像预约记录 where 预约设备ID=[1] and 预约开始时间 between [2] and [3] order by cast(序号 as int)"
    Set rsSchedule = zlDatabase.OpenSQLRecord(strSQL, "查找最小的空预约序号", mlngSchDeviceID, _
        CDate(Format(dtSchDate, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtSchDate, "yyyy-MM-dd") & " 23:59:59"))
    
    While blnTimeOK = False
        dtStartTime = 0
        dtNextStartTime = Format(dtSchDate, "yyyy-MM-dd") & " 23:59:59"
        iTimeType = 0
        lngTimeLength = 0
        If Not rsSchedule.EOF Then rsSchedule.MoveFirst
        lngNumber = 1
        
        Do While rsSchedule.EOF = False
            If rsSchedule!序号 = lngNumber Then
                dtLastEndTime = rsSchedule!预约结束时间
                lngNumber = lngNumber + 1       '如果出现序号跳号，这里刚好查到第一个空号
            ElseIf rsSchedule!预约开始时间 < dtBeforeTime Then
                '时间太早，跳号之后，继续往后查
                dtLastEndTime = rsSchedule!预约结束时间
                lngNumber = rsSchedule!序号 + 1
            Else
                '可能是紧跟着的下一个号码，如果是则记录下来
'                If rsSchedule!序号 = lngNumber + 1 Then
'                    blnFill = True
'                End If
                dtNextStartTime = rsSchedule!预约开始时间
                Exit Do
            End If
            rsSchedule.MoveNext
        Loop
            
        If dtLastEndTime < dtBeforeTime Then dtLastEndTime = dtBeforeTime
        
        '根据最小序号，预约开始时间，查找最早的时间段
        i = 1
        lngBaseNumber = 0
        While i <= UBound(mSchTimeProject) And dtStartTime = 0
            If (lngNumber > mSchTimeProject(i).lngSum + lngBaseNumber) Or _
                (Format(dtBeforeTime, "HH:MM") > Format(mSchTimeProject(i).dtEndTime, "HH:MM")) Then
                lngBaseNumber = lngBaseNumber + mSchTimeProject(i).lngSum
                i = i + 1
            Else
                '找到这个时间段，跳出循环
                If (Format(dtSchDate, "YYYY-MM-DD") & " " & Format(mSchTimeProject(i).dtStartTime, "HH:MM:SS") > dtLastEndTime) Then
                    dtStartTime = Format(dtSchDate, "YYYY-MM-DD") & " " & Format(mSchTimeProject(i).dtStartTime, "HH:MM:SS")
                Else
                    dtStartTime = dtLastEndTime
                End If
                iTimeType = mSchTimeProject(i).lngCalType
                lngTimeLength = DateDiff("n", mSchTimeProject(i).dtStartTime, mSchTimeProject(i).dtEndTime) / mSchTimeProject(i).lngSum
'                If blnFill = True Then
'                    If Format(dtNextStartTime, "HH:MM") < mSchTimeProject(i).dtStartTime _
'                        Or mSchTimeProject(i).dtEndTime < Format(dtNextStartTime, "HH:MM") Then
'                        blnFill = False
'                    End If
'                End If
            End If
        Wend
        
        '如果时间超过了时间计划，提示用户是否做计划外预约
        If dtStartTime = 0 Then
            '最后一个时间段是否结束？
            If Format(dtBeforeTime, "HH:MM") < Format(mSchTimeProject(UBound(mSchTimeProject)).dtEndTime, "HH:MM") Then
                dtStartTime = dtLastEndTime
                iTimeType = mSchTimeProject(UBound(mSchTimeProject)).lngCalType
                lngTimeLength = DateDiff("n", mSchTimeProject(UBound(mSchTimeProject)).dtStartTime, mSchTimeProject(UBound(mSchTimeProject)).dtEndTime) / mSchTimeProject(UBound(mSchTimeProject)).lngSum
            Else
                If MsgBox("今天已经没有可以预约时间计划，是否在计划外预约？", vbYesNo, "检查预约提示") = vbNo Then
                    Exit Function
                End If
                dtStartTime = dtBeforeTime
            End If
        End If
        
        '计算预约的结束时间
        '如果是跳号的，刚好是两个号中间，则直接填满这个号码的所有时间空隙
'        If blnFill = True Then
'            dtEndTime = dtNextStartTime
'        Else
            '查数据库，获取结束时间，按人次平均，在前面已经计算过了
            If lngTimeLength <> 0 And iTimeType = 1 Then
                dtEndTime = DateAdd("n", lngTimeLength, dtStartTime)
            Else
                '按项目累加
                strSQL = "select b.检查时长 from 病人医嘱记录 a ,影像预约项目 b where a.诊疗项目id = b.诊疗项目id and a.id =[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询预约项目检查时长", lngOrderID)
                
                If rsTemp.EOF = False Then
                    dtEndTime = DateAdd("n", rsTemp!检查时长, dtStartTime)
                End If
            End If
'        End If
        
        '如果时间计划的长度不够，则往后寻找下一个位置和序号
        If dtEndTime > dtNextStartTime Then
            dtBeforeTime = dtEndTime
        Else
            blnTimeOK = True
        End If
    Wend
    
    '从数据库中查询医嘱信息
    strSQL = "select id,姓名,医嘱内容,婴儿  from 病人医嘱记录  where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱内容", lngOrderID)
    
    '增加预约池容量，将预约信息记录到预约池中
    iPoolIndex = UBound(mSchLabelPool) + 1
    ReDim Preserve mSchLabelPool(iPoolIndex) As TYPE_SchLabel
    
    '提取婴儿信息
    If rsTemp!婴儿 <> 0 Then
        strSQL = "Select A.开嘱时间,Nvl(B.婴儿姓名, A.姓名 || '之子' || Trim(To_Char(B.序号, '9'))) As 婴儿姓名, B.婴儿性别, B.出生时间" & vbNewLine & _
                    "  From 病人医嘱记录 A, 病人新生儿记录 B " & vbNewLine & _
                    "  Where a.病人ID = b.病人ID  And b.序号 = [2] And a.ID = [1]"
        Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "提取婴儿信息", lngOrderID, CLng(rsTemp!婴儿))
        mSchLabelPool(iPoolIndex).str姓名 = "婴儿：" & rsBaby!婴儿姓名
    Else
        mSchLabelPool(iPoolIndex).str姓名 = rsTemp!姓名
    End If
    mSchLabelPool(iPoolIndex).lng序号 = lngNumber
    mSchLabelPool(iPoolIndex).lng医嘱ID = rsTemp!ID
    mSchLabelPool(iPoolIndex).dtStartTime = dtStartTime
    mSchLabelPool(iPoolIndex).dtEndTime = dtEndTime
    mSchLabelPool(iPoolIndex).str医嘱内容 = rsTemp!医嘱内容
    mSchLabelPool(iPoolIndex).lngBtnCount = 1
    mSchLabelPool(iPoolIndex).bln已执行 = False
    mSchLabelPool(iPoolIndex).bln已保存 = False
    
    '重新调整预约池顺序
    iPoolIndex = ResortSchLabelPool(iPoolIndex)
    
    '加载一个预约标签，摆放在dtStartTime位置
    iIndex = CreateNewSchLabel()
    
    '摆放预约标签
    Call PutSchLabel(iIndex, iPoolIndex)
    
    '记录模块变量
    mlngPoolIndex = btnSchLabel(iIndex).tag
    mlngBtnIndex = iIndex
    
    Call setSchLabelToolTipText(iIndex)
    
    NewSchedule = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'-------------------------------------------------------------
'   内部私有方法
'-------------------------------------------------------------

Private Sub btnSchLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mlngTimeProjectID = 0
    If Button = 1 And btnSchLabel(Index).Visible = True Then
        mlngOriginLeft = btnSchLabel(mSchLabelPool(btnSchLabel(Index).tag).lngBtnIndex).Left
        mlngOriginTop = btnSchLabel(mSchLabelPool(btnSchLabel(Index).tag).lngBtnIndex).Top
        mlngOriginMinute = DateDiff("n", mSchLabelPool(btnSchLabel(Index).tag).dtStartTime, mSchLabelPool(btnSchLabel(Index).tag).dtEndTime)
        mlngOriginWidth = btnSchLabel(Index).Width
        mlngPoolIndex = btnSchLabel(Index).tag
        
        mlngBaseX = X
        mlngBaseY = Y
        mlngDownX = X
        
        Call setSchLabelSelectTag(Index)
        
        '把标签显示到最前面
        Call setSchLabelZorder(Index)
        '如果医嘱ID发生改变，则触发OnChangeOrder事件
        If mlngOrderID <> mSchLabelPool(btnSchLabel(Index).tag).lng医嘱ID Then
            RaiseEvent OnChangeOrder(mSchLabelPool(btnSchLabel(Index).tag).lng医嘱ID, btnSchLabel(Index).ToolTipText)
        End If
        mlngOrderID = mSchLabelPool(btnSchLabel(Index).tag).lng医嘱ID
        
        If btnSchLabel(Index).MousePointer = vbSizeWE Then
            '调整预约标签宽度
            If CanResizeLabel(Index) = True Then
                mblnSizingLabel = True
            Else
                mblnSizingLabel = False
            End If
        Else
            '拖拽预约标签
            mblnRestorePos = False
            mblnDragingLabel = True
        End If
    End If
End Sub

Private Sub btnSchLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngWidth As Long
    Dim iPoolIndex As Long
    
    

    On Error GoTo err
    If mblnDragingLabel = True And Button = 1 Then
        If mlngBaseX <> X Or mlngBaseY <> Y Then
            Call MoveBtnLabels(Index, mlngBaseX, mlngBaseY, IIf(X < 0, 0, X), Y)
        End If
    ElseIf mblnSizingLabel = True And Button = 1 Then
        '正在调整预约标签的宽度
        
        lngWidth = mlngOriginWidth + (X - mlngBaseX)
        If lngWidth > 0 And lngWidth < (vsfTime.Width - vsfTime.ColWidth(0)) Then
            btnSchLabel(Index).Width = lngWidth
        End If
    Else
        '如果鼠标只是经过预约标签，如果鼠标在标签的右边，显示左右调整的鼠标指针
        If X > btnSchLabel(Index).Width - 100 Then
            btnSchLabel(Index).MousePointer = vbSizeWE
        Else
            btnSchLabel(Index).MousePointer = vbDefault
        End If
        
    On Error Resume Next
    
    '重新加载标签对应的提示信息
    Call setSchLabelToolTipText(Index)
    
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub btnSchLabel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngFirstIndex As Long
    lngFirstIndex = mSchLabelPool(btnSchLabel(Index).tag).lngBtnIndex
    
    If Button = 1 Then mlngTimeProjectID = 0
    Call SetColor(0, True)
    If mblnDragingLabel = True Then
        
'        '对于标签组，由于 mlngBaseX在每一次移动后，都会调整，所以这里使用mlngDownX来判断标签是否真的被移动过
'        If mlngDownX <> X Or mlngBaseY <> Y Then
'            '修正标签位置
'            Call AdjustLabelPos(Index)
'        End If
'
'        mlngPoolIndex = btnSchLabel(lngFirstIndex).tag
        mblnDragingLabel = False
'        mSchLabelPool(btnSchLabel(Index).tag).bln已保存 = False
'        Call setSchLabelToolTipText(lngFirstIndex)
'        RaiseEvent OnSchLabelModifed(lngFirstIndex)
'
'        If Button = 1 Then
'            If mlngBaseX <> X Or mlngBaseY <> Y Then Call MoveBtnLabelsAuto(Index, mlngBaseX, mlngBaseY, IIf(X < 0, 0, X), Y)
'        End If
        
        Call SetMouseTimePro(0, 0, True)
    ElseIf mblnSizingLabel = True Then
        '结束调整预约标签的宽度
        mblnSizingLabel = False
        
        '将新的预约时长记录到预约池中
        '不论是调整单个标签或者是标签组，由于只能调整标签的右边界，所以这个位置肯定是结束时间
        mSchLabelPool(btnSchLabel(Index).tag).dtEndTime = Format(mSchLabelPool(btnSchLabel(Index).tag).dtEndTime, "YYYY-MM-DD") & " " & Format(GetTimeFromXY(btnSchLabel(Index).Left + btnSchLabel(Index).Width, btnSchLabel(Index).Top), "HH:MM")
            
        If IsLabelOverlap(Index) = True Then
            '不用提示，直接将标签移动回原来的位置
            Call RestoreLabelPos(Index)
        End If
        
        mSchLabelPool(btnSchLabel(Index).tag).bln已保存 = False
        Call setSchLabelToolTipText(lngFirstIndex)
        RaiseEvent OnSchLabelModifed(lngFirstIndex)
    ElseIf Button = 2 And (mlngUseType = Sch_UseType_检查预约 Or mlngUseType = Sch_UseType_预约管理) Then
        '弹出右键菜单
        If mlngUseType = Sch_UseType_检查预约 Then
            menuPopupScheduleModi.Visible = False
        Else
            menuPopupScheduleModi.Visible = True
        End If
        Call PopupMenu(menuSchedulePopup)
    End If

End Sub

Private Sub loadTimeTable()
'------------------------------------------------
'功能：装载时间表的表格格式和基础内容
'参数：
'返回：无
'------------------------------------------------
    Dim i As Integer
    
    With vsfTime
        .Rows = 25
        .Cols = 13
        .FixedRows = 1
        .FixedCols = 1
        .AllowUserResizing = flexResizeNone
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarNone
        .CellAlignment = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .AllowSelection = False
        
        For i = 0 To 23
            .TextMatrix(i + 1, 0) = IIf(i < 10, "0" & i, i) & ":00"
        Next i
        
        For i = 0 To 11
            .TextMatrix(0, i + 1) = IIf(i * 5 < 10, "0" & i * 5, i * 5)
        Next i
        
    End With
End Sub

Private Sub ResizeTimeTable()
'------------------------------------------------
'功能：调整时间表的位置
'参数：
'返回：无
'------------------------------------------------
    Dim lngTemp  As Long
    Dim i As Integer
    
    On Error GoTo err
    
    '先调整时间表的位置
    With vsfTime
        .Left = ScaleLeft
        .Top = ScaleTop
        .Height = ScaleHeight
        .Width = ScaleWidth - 20
        .RowHeight(0) = .Height / 25
        If .RowHeight(0) < 400 Then .RowHeight(0) = 400
        lngTemp = (.Height - .RowHeight(0) - 100) / 24
        For i = 0 To 23
            .RowHeight(i + 1) = lngTemp
        Next i
        
        .RowHeight(0) = .Height - (.RowHeight(1) * 24) - 80
        
        '首列的宽度
        .ColWidth(0) = .Width / 13
        If .ColWidth(0) < 600 Then .ColWidth(0) = 600
        lngTemp = (.Width - .ColWidth(0)) / 12
        For i = 0 To 11
            .ColWidth(i + 1) = lngTemp
        Next i
        .ColWidth(0) = .Width - (.ColWidth(1) * 12)
    End With
    
    '再调整预约标签的位置
    For i = 1 To UBound(mSchLabelPool)
        Call PutSchLabel(mSchLabelPool(i).lngBtnIndex, i)
    Next i
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub timColor_Timer()
'刷新当前移动到的时间段
    '避免列表中出现蓝色选中框
    vsfTime.Row = 0
    vsfTime.Col = 0
    If mblnDragingLabel Then
        Call ShowNowSchTimeProject
    Else
        If mlngUseType = Sch_UseType_检查预约 Then
            Call ShowMouseTime(mlngMouseX, mlngMouseY)
        End If
    End If
End Sub

Private Sub ShowNowSchTimeProject()
'展现当前鼠标所在的时间段
    On Error GoTo errH
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngTimeProjectID As Long
    Dim i As Integer
    Dim intTMP As Integer, iFirstIndex As Integer, iPoolIndex As Integer
    
    intTMP = -1
    
    On Error Resume Next
    For i = 0 To btnSchLabel.Count - 1
        If btnSchLabel(i).ToolTipText = mstrOrderInfo Then
            intTMP = i
            Exit For
        End If
    Next
    
    On Error GoTo errH
    If intTMP = -1 Then Exit Sub
    If btnSchLabel(intTMP).HelpContextID <> 0 Then   '是标签组
        '首先找到这个标签组中的第一个标签索引
        iFirstIndex = mSchLabelPool(btnSchLabel(intTMP).tag).lngBtnIndex
    Else    '不是标签组，当前索引就是第一个索引
        iFirstIndex = intTMP
    End If
    iPoolIndex = btnSchLabel(iFirstIndex).tag
    
    lngRow = GetRowsFromY(btnSchLabel(intTMP).Top)
    lngCol = GetColsFromX(btnSchLabel(intTMP).Left)
    lngTimeProjectID = vsfTime.Cell(flexcpData, lngRow, lngCol)
    If lngTimeProjectID = 0 Then
        Call SetColor(lngTimeProjectID, True)
    Else
        mSchLabelPool(iPoolIndex).lngTimeProjectID = lngTimeProjectID
        If mlngTimeProjectID <> lngTimeProjectID Then
            Call SetColor(lngTimeProjectID, False)
        End If
    End If
    
    Exit Sub
errH:
    If InStr(err.Description, "控件数组元素") > 0 Then
        Resume Next
    Else
        Call err.Raise(err.Number, , err.Description)
        Resume
    End If
End Sub

Private Sub UserControl_Initialize()
    Call InitSchedule
End Sub

Private Sub UserControl_Resize()
    Call ResizeTimeTable
End Sub

Private Sub UserControl_Terminate()
    Set mrsCalendar = Nothing
    Set mColordict = Nothing
End Sub

Private Sub vsfTime_DragDrop(Source As Control, X As Single, Y As Single)
    
    Dim lngTop As Long
    
    On Error GoTo err
    
    If Source.Name = "btnSchLabel" Then
        Source.Left = X - Source.Width / 2
        
        '计算预约标签摆放在时间表中的高度，需要跟时间表的行高平齐
        lngTop = AdjustSchLabelTop(Y - Source.Height / 2)
        
        Source.Top = lngTop
        Source.Visible = True
    End If
    mblnDragingLabel = False
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function AdjustSchLabelTop(ByVal lngTop As Long) As Long
'------------------------------------------------
'功能：根据当前的位置，微调预约标签的TOP，把标签摆放在时间表的某一行内
'参数： lngTop --- 当前鼠标所在的Y位置
'返回：在时间表内，最接近鼠标位置的行首的Y值
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
        
    '如果是在首行，则自动放到第一行
    If lngTop < vsfTime.RowHeight(0) Or lngTop > vsfTime.Height Then
        AdjustSchLabelTop = mlngOriginTop
        mblnRestorePos = True
    Else
        For i = 0 To 22
            If lngTop < vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * i Then
                Exit For
            End If
        Next i
        If Abs(lngTop - (vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * (i - 1))) > _
            Abs(lngTop - (vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * i)) Then
            AdjustSchLabelTop = vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * i
        Else
            AdjustSchLabelTop = vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * (i - 1)
        End If
        AdjustSchLabelTop = AdjustSchLabelTop + 30
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function AdjustSchLabelLeft(ByVal lngLeft As Long) As Long
'------------------------------------------------
'功能：根据当前的位置，微调预约标签的Left，确保标签不会移出时间表范围
'参数： lngLeft --- 当前鼠标所在的X位置
'返回：在时间表内，最鼠标位置的X值
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
        
    '如果是在首列，则放回原来的位置
    If lngLeft <= vsfTime.ColWidth(0) Or lngLeft >= vsfTime.Width Then
        AdjustSchLabelLeft = mlngOriginLeft
        mblnRestorePos = True
    Else
        AdjustSchLabelLeft = lngLeft
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function getRowColFromTime(ByVal strTime As String, ByVal blnStart As Boolean, _
    ByRef lngRow As Long, ByRef lngCol As Long) As Boolean
'------------------------------------------------
'功能：根据时间，计算对应的格子行和列位置
'参数： strTime -- 输入的时间，格式为“hh:mm”
'       blnStart -- 是否开始时间 true--开始时间；false--结束时间
'       lngRow -- 在时间表中的行号
'       lngCol -- 在时间表中的列号
'返回：无
'------------------------------------------------
    Dim lngHour As Long
    Dim lngMinute As Long
    
    On Error GoTo err
    
    If UBound(Split(strTime, ":")) <> 1 Then
        getRowColFromTime = False
        Exit Function
    End If
    
    lngHour = Split(strTime, ":")(0)
    lngMinute = Split(strTime, ":")(1)
    
    '如果是结束时间，且时间的分钟等于0，则算是在前一个小时结束
    If blnStart = False And lngMinute = 0 Then
        lngRow = lngHour
        lngCol = 12
    Else
        lngRow = lngHour + 1
        lngCol = Int(lngMinute / 5 + 1) '确保5的倍数被放在上一列
    End If
    
    If blnStart = False And (lngMinute Mod 5 = 0) And (lngMinute <> 0) Then
        '如果是结束时间，正好是5的整数，则需要往前算一行
        lngCol = lngCol - 1
    End If
    
    getRowColFromTime = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsfTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngUseType = Sch_UseType_检查预约 Then
        If mblnDragingLabel Then
            mlngMouseX = 0
            mlngMouseY = 0
        Else
            mlngMouseX = X
            mlngMouseY = Y
        End If
    End If
End Sub

Private Sub vsfTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngTimeProjectID As Long
    
    If Button = 2 And mlngUseType = Sch_UseType_预约设置 And mlngSchPlanID <> 0 Then       '鼠标右键单击，弹出右键菜单
        '先控制菜单的可见性
        
        With vsfTime
            mlngCol = GetColsFromX(X)
            mlngRow = GetRowsFromY(Y)
            If mlngCol >= 1 And mlngRow >= 1 Then
                lngTimeProjectID = .Cell(flexcpData, mlngRow, mlngCol, mlngRow, mlngCol)
                If lngTimeProjectID = 0 Then
                    menuPopupTimeProjectModi.Visible = False
                    menuPopupTimeProjectDel.Visible = False
                Else
                    menuPopupTimeProjectModi.Visible = True
                    menuPopupTimeProjectDel.Visible = True
                End If
            End If
        End With
    
        Call PopupMenu(menuTimePopup)
    End If
    
    If Button = 1 Then
        If mlngUseType = Sch_UseType_检查预约 Then
            Call SetMouseTimePro(CLng(X), CLng(Y), False)
        End If
    End If

End Sub

Private Sub menuPopupTimeProjectAdd_Click()
    On Error Resume Next
    
    RaiseEvent OnMenuTimeProjectAdd
    
    Call RefreshTimeProject(mlngSchPlanID)
    err.Clear
End Sub

Private Sub menuPopupTimeProjectModi_Click()
    Dim lngTimeProjectID As Long
    
    On Error Resume Next
    
    '从时间表中，提取时间计划ID
    If mlngCol >= 1 And mlngRow >= 1 Then
        lngTimeProjectID = vsfTime.Cell(flexcpData, mlngRow, mlngCol, mlngRow, mlngCol)
        RaiseEvent OnMenuTimeProjectModify(lngTimeProjectID)
    End If
    Call RefreshTimeProject(mlngSchPlanID)
    err.Clear
End Sub

Private Sub menuPopupTimeProjectDel_Click()
    '删除时间计划
    Dim strSQL As String
    Dim lngTimeProjectID As Long
    Dim blnCancel As Boolean
    
    On Error Resume Next
    
    RaiseEvent OnMenuTimeProjectBeforeDel(blnCancel)
    
    If blnCancel = True Then
        Exit Sub
    End If
    
    '从时间表中，提取时间计划ID
    If mlngCol >= 1 And mlngRow >= 1 Then
        lngTimeProjectID = vsfTime.Cell(flexcpData, mlngRow, mlngCol, mlngRow, mlngCol)
        strSQL = "Zl_影像预约时间计划_删除(" & lngTimeProjectID & ")"
        zlDatabase.ExecuteProcedure strSQL, "删除时间计划"
        Call RefreshTimeProject(mlngSchPlanID)
    End If
    
    err.Clear
End Sub

Private Sub menuPopupTimeProjectColor_Click()
    On Error Resume Next
    
    RaiseEvent OnMenuTimeProjectSetColor
    Call loadColor
    
    err.Clear
End Sub

Private Sub menuPopupScheduleModi_Click()
    On Error Resume Next
    
    RaiseEvent OnMenuScheduleModify
    
    err.Clear
End Sub

Private Sub menuPopupSchedulePrint_Click()
    On Error Resume Next
    
    RaiseEvent OnMenuSchedulePrint
    
    err.Clear
End Sub

Private Sub InitSchedule()
'------------------------------------------------
'功能：初始化预约时间表控件
'参数：
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    btnSchLabel(0).Width = 0
    btnSchLabel(0).Height = 0
    btnSchLabel(0).Visible = False
    
    Call loadTimeTable
    mlngSchLabelIndex = 0
    mIsReadOnly = False
    ReDim Preserve mSchLabelPool(0) As TYPE_SchLabel
    mstrModifiedOrderID = ""
    mlngFontSize = btnSchLabel(0).Font.Size
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub loadColor()
'------------------------------------------------
'功能：从数据库读取时间表颜色设置
'参数：
'返回：无
'------------------------------------------------
    
    On Error GoTo err
    
    '从数据库中读取设置过的颜色
    mlngColorTabWork = zlDatabase.GetPara("检查预约时间表工作时间颜色", glngSys, 1292, "8421376")
    mlngColorTabRest = zlDatabase.GetPara("检查预约时间表休息时间颜色", glngSys, 1292, "16777215")
    mlngColorLblWaiting = zlDatabase.GetPara("检查预约标签已预约颜色", glngSys, 1292, "0")
    mlngColorLblDone = zlDatabase.GetPara("检查预约标签已完成颜色", glngSys, 1292, "12632256")
    mlngColorLblPassed = zlDatabase.GetPara("检查预约标签已过号颜色", glngSys, 1292, "255")
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub unloadSchLabel()
'------------------------------------------------
'功能：清空预约标签，预约标签的索引是不连续的
'参数：
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim btnLabel As CommandButton
    
    On Error GoTo err
    
    '卸载预约标签
    For Each btnLabel In btnSchLabel
        If btnLabel.Index <> 0 Then
            Call Unload(btnLabel)
        End If
    Next
    
    '把计数器清零
    mlngSchLabelIndex = btnSchLabel.Count - 1
    
    '清空预约池
    ReDim Preserve mSchLabelPool(0) As TYPE_SchLabel
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub PutSchLabel(ByVal iBtnIndex As Integer, ByVal iPoolIndex As Integer)
'------------------------------------------------
'功能：根据输入的开始时间和结束时间，摆放预约标签，然后按照序号调整预约池的顺序
'参数： iBtnIndex — 预约标签的索引
'       iPoolIndex -- 预约池的索引
'返回：无
'------------------------------------------------
    Dim strStartTime As String
    Dim strEndTime As String
    Dim intSHour As Integer
    Dim intSMinute As Integer
    Dim intEHour As Integer
    Dim intEMinute As Integer
    Dim lngSX As Long
    Dim lngSY As Long
    Dim lngEX As Long
    Dim lngEY As Long
    Dim iNewIndex As Integer
    Dim iPreIndex As Integer
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngTimeProjectID As Long        '时间计划ID
    Dim lngColor As Long
    
    
    On Error GoTo err
    '设置预约池和预约标签的内容
    mSchLabelPool(iPoolIndex).lngBtnIndex = iBtnIndex
        
    btnSchLabel(iBtnIndex).Caption = mSchLabelPool(iPoolIndex).str姓名
    btnSchLabel(iBtnIndex).tag = iPoolIndex
    
    '从预约池获取开始时间和结束时间
    strStartTime = Format(mSchLabelPool(iPoolIndex).dtStartTime, "HH:MM")
    strEndTime = Format(mSchLabelPool(iPoolIndex).dtEndTime, "HH:MM")
    
    '设置颜色
    If mSchLabelPool(iPoolIndex).bln已执行 = True Then
        lngColor = mlngColorLblDone
    ElseIf Format(mdtSchDate, "YYYY-MM-DD") & " " & Format(strStartTime, "HH:MM:SS") < Format(Now, "YYYY-MM-DD HH:MM:SS") Then
        lngColor = mlngColorLblPassed
    Else
        lngColor = mlngColorLblWaiting
    End If
    
    '根据开始时间，计算标签的开始位置
    intSHour = Val(Left(strStartTime, 2))
    intSMinute = Val(Mid(strStartTime, 4))
    intEHour = Val(Left(strEndTime, 2))
    intEMinute = Val(Mid(strEndTime, 4))
    
    '使用统一的方法，从时间提取X,Y的位置，结束时间是否需要专门处理？如果是在下一行，应该提前到上一行结束
    Call GetXYFromTime(strStartTime, lngSX, lngSY)
    Call GetXYFromTime(strEndTime, lngEX, lngEY)
        
    '标签是否在同一行显示
    If (intEHour = intSHour) Or (intEHour - intSHour = 1 And intEMinute = 0) Then '标签只有一行,首先判断开始和结束时间，是否为同一个小时

        btnSchLabel(iBtnIndex).Left = lngSX
        btnSchLabel(iBtnIndex).Top = lngSY
        If intEMinute = 0 Then
            btnSchLabel(iBtnIndex).Width = vsfTime.Width - lngSX
        Else
            btnSchLabel(iBtnIndex).Width = lngEX - lngSX
        End If
        
        btnSchLabel(iBtnIndex).Height = vsfTime.RowHeight(1)
        '如果原先是个标签组，则卸载多余的标签
        If btnSchLabel(iBtnIndex).HelpContextID <> 0 Then
            Call DelUnUseLabel(iBtnIndex, btnSchLabel(iBtnIndex).HelpContextID)
        End If
        btnSchLabel(iBtnIndex).HelpContextID = 0      '设置单行标签的标记
        btnSchLabel(iBtnIndex).BackColor = lngColor
    Else    '标签有多行
        '创建一个标签组
        '先摆放第一行标签
        btnSchLabel(iBtnIndex).Left = lngSX
        btnSchLabel(iBtnIndex).Top = lngSY
        btnSchLabel(iBtnIndex).Width = vsfTime.Width - lngSX
        btnSchLabel(iBtnIndex).Height = vsfTime.RowHeight(1)
        btnSchLabel(iBtnIndex).BackColor = lngColor
        iPreIndex = iBtnIndex
        '如果大于2个小时，则循环摆放中间满格的标签
        For i = 1 To IIf(intEMinute = 0, intEHour - intSHour - 2, intEHour - intSHour - 1)
            '如果这套标签原先已经有标签组，而且有足够的其他行标签, 则直接使用现有的标签，没有则创建新标签
            If btnSchLabel(iPreIndex).HelpContextID = 0 _
                Or (btnSchLabel(iPreIndex).HelpContextID = iBtnIndex) Then
                iNewIndex = CreateNewSchLabel()
            Else
                iNewIndex = btnSchLabel(iPreIndex).HelpContextID
            End If
            
            '设置标签的位置
            btnSchLabel(iNewIndex).Left = vsfTime.ColWidth(0)
            btnSchLabel(iNewIndex).Top = lngSY + i * vsfTime.RowHeight(1)
            btnSchLabel(iNewIndex).Width = vsfTime.Width - vsfTime.ColWidth(0)
            btnSchLabel(iNewIndex).Height = vsfTime.RowHeight(1)
            '设置标签的基本信息
            btnSchLabel(iNewIndex).Caption = btnSchLabel(iBtnIndex).Caption
            btnSchLabel(iNewIndex).tag = btnSchLabel(iBtnIndex).tag
            btnSchLabel(iNewIndex).BackColor = lngColor
            '设置标签组
            btnSchLabel(iPreIndex).HelpContextID = iNewIndex
            iPreIndex = iNewIndex
        Next i
        
        '摆放最后一行的标签
        '如果这套标签原先已经有标签组，而且有足够的其他行标签, 则直接使用现有的标签，没有则创建新标签
        If btnSchLabel(iPreIndex).HelpContextID = 0 _
            Or (btnSchLabel(iPreIndex).HelpContextID = iBtnIndex) Then
            iNewIndex = CreateNewSchLabel()
        Else
            iNewIndex = btnSchLabel(iPreIndex).HelpContextID
            '处理原来标签组里面的多余标签，设置为不可见，并卸载
            If btnSchLabel(iNewIndex).HelpContextID <> iBtnIndex Then
                Call DelUnUseLabel(iBtnIndex, btnSchLabel(iNewIndex).HelpContextID)
            End If
        End If
        '设置标签的位置
        btnSchLabel(iNewIndex).Left = vsfTime.ColWidth(0)
        If intEMinute = 0 Then
            btnSchLabel(iNewIndex).Top = lngEY - vsfTime.RowHeight(1)
            btnSchLabel(iNewIndex).Width = vsfTime.Width - vsfTime.ColWidth(0)
        Else
            btnSchLabel(iNewIndex).Top = lngEY
            btnSchLabel(iNewIndex).Width = lngEX - vsfTime.ColWidth(0)
        End If
        
        btnSchLabel(iNewIndex).Height = vsfTime.RowHeight(1)
        '设置标签的基本信息
        btnSchLabel(iNewIndex).Caption = btnSchLabel(iBtnIndex).Caption
        btnSchLabel(iNewIndex).tag = btnSchLabel(iBtnIndex).tag
        btnSchLabel(iNewIndex).BackColor = lngColor
        '设置第一个标签的HelpContextID，让标签组形成闭环
        btnSchLabel(iPreIndex).HelpContextID = iNewIndex
        btnSchLabel(iNewIndex).HelpContextID = iBtnIndex
    End If
    
    '将预约标签的数量，和行列位置，记录到预约池中
    mSchLabelPool(iPoolIndex).lngBtnCount = IIf(intEMinute = 0, intEHour - intSHour, intEHour - intSHour + 1)
    mSchLabelPool(iPoolIndex).lngRow = GetRowsFromY(btnSchLabel(iBtnIndex).Top)
    mSchLabelPool(iPoolIndex).lngCol = GetColsFromX(btnSchLabel(iBtnIndex).Left)
    
    Call getRowColFromTime(strStartTime, True, lngRow, lngCol)
    lngTimeProjectID = vsfTime.Cell(flexcpData, lngRow, lngCol)
    mSchLabelPool(iPoolIndex).lngTimeProjectID = lngTimeProjectID
    
    '默认设置开始和结束时间段，为开始和结束时间，如果在事件计划外预约，就会没有开始和结束时间段
    mSchLabelPool(iPoolIndex).dt开始时间段 = mSchLabelPool(iPoolIndex).dtStartTime
    mSchLabelPool(iPoolIndex).dt结束时间段 = mSchLabelPool(iPoolIndex).dtEndTime
    If lngTimeProjectID <> 0 Then
        For i = 1 To UBound(mSchTimeProject)
            If mSchTimeProject(i).lngID = lngTimeProjectID Then
                mSchLabelPool(iPoolIndex).dt开始时间段 = Format(mSchLabelPool(iPoolIndex).dtStartTime, "YYYY-MM-DD") & " " & Format(mSchTimeProject(i).dtStartTime, "HH:MM:SS")
                mSchLabelPool(iPoolIndex).dt结束时间段 = Format(mSchLabelPool(iPoolIndex).dtStartTime, "YYYY-MM-DD") & " " & Format(mSchTimeProject(i).dtEndTime, "HH:MM:SS")
            End If
        Next i
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function GetSchPlanID(ByVal lngSchDeviceID As Long, ByRef dtDate As Date, _
    ByVal blnFindNextDay As Boolean, ByVal blnSilent As Boolean) As Long
'------------------------------------------------
'功能：根据设备ID和日期，查找从这天开始，最接近一天的可用方案ID
'参数： lngSchDeviceID -- 预约设备ID
'       dtDate -- 【返回参数】预约日期。先查找dtDate这天的方案ID，如果blnFindNextDay=True，且没有可用方案，自动返回下一个有预约方案的日期
'       blnFindNextDay -- 如果dtDate没有可用的预约方案，是否查找下一个可用的预约方案ID
'       blnSilent -- 是否静默，不提示对话框
'返回： 预约方案ID
'------------------------------------------------
    Dim lngSchPlanID As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim str方案内容 As String
    Dim strFilter As String
    Dim i As Integer            '计数器，只查找一年之内的方案
    
    On Error GoTo err
    
    '根据预约设备ID和日期，查找对应的时间方案ID
    strSQL = "select ID,方案名称,方案类型,方案内容,是否启用,开始时间,间隔,是否按日历休息 " _
        & " from 影像预约方案 where 预约设备ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询今天的预约方案", lngSchDeviceID)
    
    i = 0
    If rsTemp.EOF = False Then
        '先查找是否有
        Do
            '先查找今天的一天方案
            strFilter = "方案内容='" & Format(dtDate, "YYYYMMDD") & "'"
            rsTemp.Filter = strFilter
            
            If rsTemp.EOF = False Then
                lngSchPlanID = rsTemp!ID
                blnFindNextDay = False  '找到一个一天方案，退出循环
            Else
                '再查找已经启用的方案
                strFilter = "是否启用=1 "
                rsTemp.Filter = strFilter
                If rsTemp.EOF = False Then
                    '查找是否有合适今天的方案，按照 “每天”，“每周”，“每月”的顺序判断
                    If rsTemp!方案类型 = Sch_PlanType_每天 And rsTemp!开始时间 < dtDate Then
                        If rsTemp!间隔 <> 0 Then
                            If DateDiff("d", rsTemp!开始时间, dtDate) Mod rsTemp!间隔 = 0 Then
                                lngSchPlanID = rsTemp!ID
                                blnFindNextDay = False  '找到一个启用方案，退出程序
                            End If
                        Else
                            lngSchPlanID = rsTemp!ID
                            blnFindNextDay = False  '找到一个启用方案，退出程序
                        End If
                    ElseIf rsTemp!方案类型 = Sch_PlanType_每周 And rsTemp!开始时间 < dtDate Then
                        '“每周”方案需要判断今天是周几
                        '如果有多个每周方案，需要寻找适合dtDate的每周方案
                        '如果使用预约日历，休息日不能预约
                        rsTemp.MoveFirst
                        While rsTemp.EOF = False
                            str方案内容 = rsTemp!方案内容
                            If nvl(rsTemp!是否按日历休息, 0) = 1 And IsDayOff(dtDate) = True Then
                                '什么都不做
                            Else
                                If InStr(str方案内容, Weekday(dtDate, vbMonday)) > 0 Then
                                    If rsTemp!间隔 <> 0 Then
                                        If DateDiff("w", rsTemp!开始时间, dtDate) Mod rsTemp!间隔 = 0 Then
                                            lngSchPlanID = rsTemp!ID
                                            blnFindNextDay = False  '找到一个启用方案，退出程序
                                        End If
                                    Else
                                        lngSchPlanID = rsTemp!ID
                                        blnFindNextDay = False  '找到一个启用方案，退出程序
                                    End If
                                End If
                            End If
                            rsTemp.MoveNext
                        Wend
                    Else    '每月方案，查询预约日历
                        If IsDayOff(dtDate) = False Then
                            lngSchPlanID = rsTemp!ID
                            blnFindNextDay = False  '找到一个启用方案，退出程序
                        End If
                    End If
                    
                ElseIf blnFindNextDay = False Then
                    lngSchPlanID = 0
                End If
            End If
            
            '没有找到预约方案，查找下一天
            If blnFindNextDay = True Then
                dtDate = dtDate + 1
                i = i + 1
            End If
        Loop Until (blnFindNextDay = False) Or i > 365
    Else
        lngSchPlanID = 0
    End If
    
    If lngSchPlanID = 0 And blnSilent = False Then
        MsgBox dtDate & " 没有可用的预约方案，请先设置预约方案。", vbOKOnly, "检查预约提示"
    End If
    GetSchPlanID = lngSchPlanID
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FindFirstSchDay(ByVal lngSchDeviceID As Long, ByRef dtSchDate As Date) As Long
'------------------------------------------------
'功能：根据设备ID和日期，查找从这天开始往后，第一个有预约容量，可以预约的日子
'参数： lngSchDeviceID -- 预约设备ID
'       dtSchDate -- 返回参数，最接近 dtSchDate 的预约日期。
'返回： 预约方案ID
'------------------------------------------------
    Dim lngSchPlanID As Long        '预约方案ID
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngSchCapacity As Long      '预约容量
    Dim blnFindNext As Boolean      '查找下一个日期
    Dim i As Integer                '计数器，只查询一年之内的可预约日期
    
    On Error GoTo err
    
    blnFindNext = True
    FindFirstSchDay = 0
    i = 0
    
    While blnFindNext = True And i < 365
        '首先根据预约设备ID 和 时间，查找第一个可以预约的日期和预约方案ID
        lngSchPlanID = GetSchPlanID(lngSchDeviceID, dtSchDate, True, False)
        If lngSchPlanID = 0 Then
            Exit Function
        End If
        
        '查看预约方案是否有预约容量，如果有，直接预约，如果没有，往后找到最近的一天预约
        strSQL = "select sum(预约容量) as 容量,count(id) as 数量  from 影像预约时间计划 where 预约方案ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询预约容量", lngSchPlanID)
        
        If nvl(rsTemp!数量, 0) = 0 = True Then
            MsgBox "无法预约，此预约方案没有时间计划，请先设置好预约时间计划，预约方案ID=" & lngSchPlanID, vbOKOnly, "检查预约提示"
            Exit Function
        End If
        
        If nvl(rsTemp!容量, 0) = 0 Then
            MsgBox "无法预约，时间计划中，预约容量为0，请联系管理员重新设置预约容量。", vbOKOnly, "检查预约提示"
            Exit Function
        End If
        
        lngSchCapacity = rsTemp!容量
        
        strSQL = "select " & lngSchCapacity & "- count(ID) as 剩余容量 from 影像预约记录 where 预约设备ID=[1] and 预约开始时间 between [2] and [3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询最佳预约日期", mlngSchDeviceID, CDate(Format(dtSchDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(dtSchDate, "yyyy-mm-dd") & " 23:59:59"))
        
        If rsTemp!剩余容量 > 0 Then
            '这天可以预约
            '如果是今天，那就还需要增加判断，当前时间之后，是否还有预约计划
            If Format(dtSchDate, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
                strSQL = "Select Sum(a.预约容量) As 容量 From 影像预约时间计划 A " _
                    & " Where a.预约方案ID = [1] And to_char(a.结束时间, 'hh24:mi:ss') > to_char(sysdate+ 2 / 24, 'hh24:mi:ss')"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断今天此时是否还有预约容量", lngSchPlanID)
                If nvl(rsTemp!容量, 0) = 0 Then
                    '查找下一个日期
                    dtSchDate = dtSchDate + 1
                    i = i + 1
                Else
                    FindFirstSchDay = lngSchPlanID
                    blnFindNext = False     '退出循环
                End If
            Else
                FindFirstSchDay = lngSchPlanID
                blnFindNext = False     '退出循环
            End If
        Else
            '查找下一个日期
            dtSchDate = dtSchDate + 1
            i = i + 1
        End If
    Wend
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateNewSchLabel() As Long
'------------------------------------------------
'功能：创建一个新的预约标签
'参数：无
'返回：新预约标签的索引
'------------------------------------------------
    On Error GoTo err
    
    mlngSchLabelIndex = mlngSchLabelIndex + 1
    Load btnSchLabel(mlngSchLabelIndex)
    btnSchLabel(mlngSchLabelIndex).Visible = True
    btnSchLabel(mlngSchLabelIndex).ZOrder
    CreateNewSchLabel = mlngSchLabelIndex
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetRowsFromY(ByVal lngY As Long) As Long
'------------------------------------------------
'功能：根据当前的Y位置，计算出在时间表中的行数
'参数： lngY --- 当前鼠标所在的Y位置
'返回：在时间表内，最接近鼠标位置的行数
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
        
    '如果是在首行，则自动放到第一行
    If lngY <= vsfTime.RowHeight(0) Then
        GetRowsFromY = 1
    ElseIf lngY >= vsfTime.Height Then
        GetRowsFromY = 24
    Else
        For i = 0 To 23
            If lngY < vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * i Then
                Exit For
            End If
        Next i
        GetRowsFromY = i
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetColsFromX(ByVal lngX As Long) As Long
'------------------------------------------------
'功能：根据当前的X位置，计算出在时间表中的列数
'参数： lngX --- 当前鼠标所在的X位置
'返回：在时间表内，最接近鼠标位置的列数
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
        
    '如果是在首列，则放回第一行
    If lngX < vsfTime.ColWidth(0) Or lngX >= vsfTime.Width Then
        GetColsFromX = 0
    Else
        For i = 0 To 10
            If lngX < vsfTime.ColWidth(0) + vsfTime.ColWidth(1) * i Then
                Exit For
            End If
        Next i
        GetColsFromX = i
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ResortSchLabelPool(iPoolIndex As Integer) As Long
'------------------------------------------------
'功能：根据新增加的序号，重新给预约池排序，预约池内，按照序号排序
'参数： iPoolIndex --- 序号和位置发生改变的标签，在预约池内的索引
'返回：直接在预约池内，对iPoolIndex的位置做排序，返回值是新的序号
'------------------------------------------------
    Dim iSwapIndex As Integer
    Dim i As Integer
    Dim tmpSchLabel As TYPE_SchLabel
    
    On Error GoTo err
        
    '先设置默认值，序号的位置没有发生改变
    ResortSchLabelPool = iPoolIndex
    
    For i = 1 To UBound(mSchLabelPool)
        If i <> iPoolIndex Then
            If mSchLabelPool(iPoolIndex).lng序号 > mSchLabelPool(i).lng序号 Then
                '过
            Else
                Exit For
            End If
        End If
    Next i
    iSwapIndex = i
    If iPoolIndex < iSwapIndex Then iSwapIndex = iSwapIndex - 1
    
    If iSwapIndex <> 0 And iSwapIndex <> iPoolIndex Then
        '新序号小，所以要重新排序，将iPoolIndex指向的内容，排在iSwapIndex的位置
        tmpSchLabel = mSchLabelPool(iPoolIndex)
        If iPoolIndex < iSwapIndex Then
            '往后移动标签，则其他标签需要往前移
            For i = iPoolIndex To iSwapIndex - 1
                '向前移动预约池块的位置
                mSchLabelPool(i) = mSchLabelPool(i + 1)
                '重新设置预约池和预约标签的关系
                Call setSchLabelTag(mSchLabelPool(i).lngBtnIndex, i)
            Next i
        Else
            '往前移动标签，则其他标签需要往后移
            For i = iPoolIndex To iSwapIndex + 1 Step -1
                '向后移动预约池块的位置
                mSchLabelPool(i) = mSchLabelPool(i - 1)
                '重新设置预约池和预约标签的关系
                Call setSchLabelTag(mSchLabelPool(i).lngBtnIndex, i)
            Next i
        End If
        
        mSchLabelPool(iSwapIndex) = tmpSchLabel
        Call setSchLabelTag(mSchLabelPool(iSwapIndex).lngBtnIndex, iSwapIndex)
    End If
    
    ResortSchLabelPool = iSwapIndex
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetTimeFromXY(lngX As Long, lngY As Long) As Date
'------------------------------------------------
'功能：根据X,Y的位置，计算在时间表内对应的时间
'参数： lngX --- X位置
'       lngY --- Y位置
'返回：对应的时间
'------------------------------------------------
    Dim lngMinute As Long
    Dim lngHour As Long
    Dim i As Integer
    
    On Error GoTo err
    
    '根据Y，得到所在的行，计算小时
    lngHour = GetRowsFromY(lngY) - 1
    
    '根据X在时间表中得到比例，得到分钟
    '如果是在首列，则为0
    If lngX <= vsfTime.ColWidth(0) Then
        lngMinute = 0
    ElseIf lngX >= vsfTime.Width Then
        lngMinute = 60  '如果大于宽度，则要在小时上面进位
    Else
        For i = 0 To 10
            If lngX < vsfTime.ColWidth(0) + vsfTime.ColWidth(1) * i Then
                Exit For
            End If
        Next i
        i = i - 1
        lngMinute = i * 5 + ((lngX - (vsfTime.ColWidth(0) + vsfTime.ColWidth(1) * i)) / vsfTime.ColWidth(1) * 5)
    End If
    
    
    '如果大于等于60,需要在小时上面进位
    If lngMinute >= 60 Then
        lngMinute = 0
        lngHour = lngHour + 1
    End If
    
    GetTimeFromXY = lngHour & ":" & lngMinute
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetXYFromTime(ByVal strTime As String, ByRef lngX As Long, ByRef lngY As Long) As Boolean
'------------------------------------------------
'功能：根据时间表内的时间，计算对应的X,Y位置
'参数： strTime --- 时间表内的时间，格式为“HH:MM”
'       lngX --- 返回参数，X位置
'       lngY --- 返回参数，Y位置
'返回：对应的时间
'------------------------------------------------
    Dim lngMinute As Long
    Dim lngHour As Long
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo err
    
    If UBound(Split(strTime, ":")) <> 1 Then
        GetXYFromTime = False
        Exit Function
    End If
    
    lngHour = Split(strTime, ":")(0)
    lngMinute = Split(strTime, ":")(1)
    
    lngRow = lngHour + 1
    lngY = vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * (lngRow - 1)
    
    lngCol = Int(lngMinute / 5 + 1) '确保5的倍数被放在下一列
    lngX = vsfTime.ColWidth(0) + vsfTime.ColWidth(1) * (lngCol - 1) + (lngMinute Mod 5) / 5 * vsfTime.ColWidth(1)
    
    GetXYFromTime = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'暂时保留，确认功能全部正常后删除
'Private Function MoveBtnLabelsAuto(iBtnIndex As Integer, lngBaseX As Long, lngBaseY As Long, ByVal lngX As Long, ByVal lngY As Long) As Boolean
''------------------------------------------------
''功能：调整标签组的位置
''参数： iBtnIndex --- 移动标签索引
''       lngBaseX --- 移动前的X位置
''       lngBaseY --- 移动前的Y位置
''       lngX --- 移动后的X位置
''       lngY --- 移动后的Y位置
''返回：True--成功；False--失败
''------------------------------------------------
'    Dim lngLeft As Long
'    Dim lngTop As Long
'    Dim iFirstIndex As Integer      '标签组中第一个标签的索引
'    Dim iLastIndex As Integer       '标签组中最后一个标签的索引
'    Dim dtStartTime As Date         '开始时间
'    Dim dtEndTime As Date           '结束时间
'    Dim lngRight As Long
'    Dim dtStartTimeSQL As Date
'    Dim dtEndTimeSQL As Date
'    Dim blnFind As Boolean '找到时间段
'    Dim blnTimeOK As Boolean '时间符合条件
'    Dim blnNeedTestTime As Boolean '需要验证时间
'    Dim tmTmp As Date
'
'    Dim intPoolIndex As Integer
'
'
'    Dim i As Integer, j As Integer
'
'    On Error GoTo err
'
'    blnFind = False
'    blnNeedTestTime = False
'    '移动标签的几种情况：
'    '1、移动单个标签，不拐弯
'    '2、移动单个标签，开始拐弯，往上或者往下拐弯
'    '3、移动标签组
'
'    '拐弯标签的处理，如果预约标签超过了日历的左右宽度，则自动增加一个新的预约标签，形成一套标签组
'    '标签组使用 HelpContexID作为连接标记，组内的标签，互相记录HelpContextID，形成闭环
'    '标签组中第一个标签的索引，记录在 标签池 mSchLabelPool 的 lngBtnIndex 中
'    '标签组的数量，记录在标签池 mSchLabelPool 的 lngBtnCount 中
'    '首先判断是否涉及到标签拐弯的问题
'
'    If btnSchLabel(iBtnIndex).HelpContextID <> 0 Then   '是标签组
'        '首先找到这个标签组中的第一个标签索引
'        iFirstIndex = mSchLabelPool(btnSchLabel(iBtnIndex).tag).lngBtnIndex
'        iLastIndex = iFirstIndex
'        While btnSchLabel(iLastIndex).HelpContextID <> iFirstIndex
'            iLastIndex = btnSchLabel(iLastIndex).HelpContextID
'        Wend
'
'    Else    '不是标签组，当前索引就是第一个索引
'        iFirstIndex = iBtnIndex
'        iLastIndex = iBtnIndex
'    End If
'
'    intPoolIndex = btnSchLabel(iBtnIndex).tag
'
'    '计算第一个标签的新位置
'    lngLeft = btnSchLabel(iFirstIndex).Left + (lngX - lngBaseX)
'    lngTop = btnSchLabel(iFirstIndex).Top + (lngY - lngBaseY)
'
'    '判断此标签是否会超出时间表范围
'    '如果在鼠标移动的过程中，标签超出了时间表范围，则取消本次移动，让标签保留在当前位置
'    If lngTop < vsfTime.RowHeight(0) Or lngTop > vsfTime.Height Then
'        '标签已经从上方或下方，超出时间表，停止移动标签
'        Exit Function
'    ElseIf lngLeft < vsfTime.ColWidth(0) Then
'        lngLeft = vsfTime.ColWidth(0)
'    ElseIf lngLeft > vsfTime.Width Then
'        '标签在右边超出时间表，需要往下移动一行
'        lngLeft = vsfTime.ColWidth(0) + (lngLeft - vsfTime.Width)
'        lngTop = lngTop + vsfTime.RowHeight(1)
'    End If
'
'    '计算新的开始时间
'    dtStartTime = GetTimeFromXY(lngLeft, lngTop)
'    dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
'    If DateDiff("d", dtStartTime, dtEndTime) = 1 Then
'        '新标签超过时间表下限，禁止移动，直接退出
'        Exit Function
'    End If
'
'    '根据开始时间判断处于哪一个时间段，并且根据预约池调整新的开始
'
''   '根据鼠标位置判断当前属于哪个时间段
'    For i = 1 To UBound(mSchTimeProject)
'         If mSchTimeProject(i).dtStartTime <= dtStartTime And mSchTimeProject(i).dtEndTime > dtStartTime Then
'            dtStartTimeSQL = mSchTimeProject(i).dtStartTime
'            dtEndTimeSQL = mSchTimeProject(i).dtEndTime
'            blnFind = True
'            Exit For
'         End If
'    Next
'
'    If Not blnFind Then
'        dtStartTimeSQL = mSchTimeProject(1).dtStartTime
'        dtEndTimeSQL = mSchTimeProject(1).dtEndTime
'    End If
'
'    blnNeedTestTime = True
'
'    blnTimeOK = True
'    '如果日期是当前，选中的时间段恰好跨越当前时间，则以当前时间作为开始时间判断是否符合条件，如果符合，直接设置为自动放置的时间
'    If Format(Now, "YYYY-MM-DD") = Format(mdtSchDate, "YYYY-MM-DD") Then
'        tmTmp = Format(Now, "hh:mm:ss")
'        tmTmp = DateAdd("n", 5, tmTmp)
'        If dtStartTimeSQL <= tmTmp And dtEndTimeSQL >= tmTmp Then
'            dtStartTime = tmTmp
'
'            dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
'            For i = 1 To UBound(mSchLabelPool) - 1
'                '避开当前时间块
'                If intPoolIndex <> i Then '需要修改这个条件
'                    If Not ((Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") < dtStartTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") <= dtStartTime) Or (Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtEndTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") > dtEndTime)) Then
'                        blnTimeOK = False
'                    End If
'                End If
'            Next
'        End If
'    End If
'
'    If UBound(mSchLabelPool) = 1 Then
'        If Not blnTimeOK Then
'            dtStartTime = dtStartTimeSQL
'            dtEndTime = dtEndTimeSQL
'        End If
'    Else
'
'        If mSchLabelPool(1).dtStartTime >= Format(mdtSchDate, "YYYY-MM-DD") & dtEndTimeSQL Then
'            '满足条件退出while循环
'            dtStartTime = dtStartTimeSQL
'            dtEndTime = dtEndTimeSQL
'            blnNeedTestTime = False
'        End If
'
'        '首先仍然假设时间段开始时间作为预约开始时间有效
'
'        dtStartTime = dtStartTimeSQL
'        dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
'
'
'        If Not blnTimeOK Then
'            For i = 1 To UBound(mSchLabelPool) - 1
'                '避开当前时间块
'                If intPoolIndex <> i Then '需要修改这个条件
'                    If Not ((Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") < dtStartTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") <= dtStartTime) Or (Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtEndTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") > dtEndTime)) Then
'                        blnTimeOK = False
'                    End If
'                End If
'            Next
'        End If
'
'        If Not blnTimeOK Then
'
'            For i = 1 To UBound(mSchLabelPool) - 1
'                If blnNeedTestTime And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtStartTimeSQL And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") <= dtEndTimeSQL And iBtnIndex <> i Then
'                    blnTimeOK = True
'
'                    If blnTimeOK And blnNeedTestTime Then
'                        dtStartTime = Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") 'rsSchedule!预约结束时间 'Format(rsSchedule!预约结束时间, "hh-mm-ss")
'                        dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
'
'                        For j = 1 To UBound(mSchLabelPool)
'                            If intPoolIndex <> j Then '需要修改这个条件
'                                If Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtStartTimeSQL And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") <= dtEndTimeSQL Then
'                                    If Not ((Format(mSchLabelPool(j).dtStartTime, "hh:mm:ss") < dtStartTime And Format(mSchLabelPool(j).dtEndTime, "hh:mm:ss") <= dtStartTime) Or (Format(mSchLabelPool(j).dtStartTime, "hh:mm:ss") >= dtEndTime And Format(mSchLabelPool(j).dtEndTime, "hh:mm:ss") > dtEndTime)) Then
'                                        blnTimeOK = False
'                                    End If
'                                End If
'                            End If
'                        Next
'
'                        If blnTimeOK Then
'                            blnNeedTestTime = False
'                        End If
'                    End If
'                End If
'            Next
'        End If
'    End If
'
'    Call ModifySchPoolTime(btnSchLabel(iFirstIndex).tag, dtStartTime)
'    '重新摆放标签
'    Call PutSchLabel(iFirstIndex, btnSchLabel(iFirstIndex).tag)
'
'    If iFirstIndex <> iBtnIndex Then
'        mlngBaseX = lngLeft
'    End If
'
'    Exit Function
'err:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

Private Function MoveBtnLabels(iBtnIndex As Integer, lngBaseX As Long, lngBaseY As Long, ByVal lngX As Long, ByVal lngY As Long) As Boolean
'------------------------------------------------
'功能：调整标签组的位置
'参数： iBtnIndex --- 移动标签索引
'       lngBaseX --- 移动前的X位置
'       lngBaseY --- 移动前的Y位置
'       lngX --- 移动后的X位置
'       lngY --- 移动后的Y位置
'返回：True--成功；False--失败
'------------------------------------------------
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim iFirstIndex As Integer      '标签组中第一个标签的索引
    Dim iLastIndex As Integer       '标签组中最后一个标签的索引
    Dim dtStartTime As Date         '开始时间
    Dim dtEndTime As Date           '结束时间
    Dim lngRight As Long
    
    Dim i As Integer
    
    On Error GoTo err
    
    '移动标签的几种情况：
    '1、移动单个标签，不拐弯
    '2、移动单个标签，开始拐弯，往上或者往下拐弯
    '3、移动标签组
    
    '拐弯标签的处理，如果预约标签超过了日历的左右宽度，则自动增加一个新的预约标签，形成一套标签组
    '标签组使用 HelpContexID作为连接标记，组内的标签，互相记录HelpContextID，形成闭环
    '标签组中第一个标签的索引，记录在 标签池 mSchLabelPool 的 lngBtnIndex 中
    '标签组的数量，记录在标签池 mSchLabelPool 的 lngBtnCount 中
    
    '首先判断是否涉及到标签拐弯的问题
    If btnSchLabel(iBtnIndex).HelpContextID <> 0 Then   '是标签组
        '首先找到这个标签组中的第一个标签索引
        iFirstIndex = mSchLabelPool(btnSchLabel(iBtnIndex).tag).lngBtnIndex
        iLastIndex = iFirstIndex
        While btnSchLabel(iLastIndex).HelpContextID <> iFirstIndex
            iLastIndex = btnSchLabel(iLastIndex).HelpContextID
        Wend
        
    Else    '不是标签组，当前索引就是第一个索引
        iFirstIndex = iBtnIndex
        iLastIndex = iBtnIndex
    End If
        
    '计算第一个标签的新位置
    lngLeft = btnSchLabel(iFirstIndex).Left + (lngX - lngBaseX)
    lngTop = btnSchLabel(iFirstIndex).Top + (lngY - lngBaseY)

    '判断此标签是否会超出时间表范围
    '如果在鼠标移动的过程中，标签超出了时间表范围，则取消本次移动，让标签保留在当前位置
    If lngTop < vsfTime.RowHeight(0) Or lngTop > vsfTime.Height Then
        '标签已经从上方或下方，超出时间表，停止移动标签
        Exit Function
    ElseIf lngLeft < vsfTime.ColWidth(0) Then
        '标签在左边超出时间表，停止移动
        lngLeft = vsfTime.ColWidth(0) + (lngLeft - vsfTime.Width)
    ElseIf lngLeft > vsfTime.Width Then
        '标签在右边超出时间表，需要往下移动一行
        lngLeft = vsfTime.ColWidth(0) + (lngLeft - vsfTime.Width)
'        lngTop = lngTop + vsfTime.RowHeight(1)
    End If
    
    '计算新的开始时间
    dtStartTime = GetTimeFromXY(lngLeft, lngTop)
    dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
    If DateDiff("d", dtStartTime, dtEndTime) = 1 Then
        '新标签超过时间表下限，禁止移动，直接退出
        Exit Function
    End If
    '修正缓冲池中的时间
    Call ModifySchPoolTime(btnSchLabel(iFirstIndex).tag, dtStartTime)
    '重新摆放标签
    Call PutSchLabel(iFirstIndex, btnSchLabel(iFirstIndex).tag)
    
    '如果直接移动标签组中第二个以后的标签，由于实际上这些标签是没有被移动的，所以需要对移动量做一个修正
    '重新记录mlngBaseX就可以修正了。mlngBaseY不需要修正，因为鼠标在Y方向的实际位移量是0
    If iFirstIndex <> iBtnIndex Then
        mlngBaseX = lngX
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function DelUnUseLabel(ByVal lngBtnIndex As Long, ByVal lngDelIndex As Long) As Boolean
'------------------------------------------------
'功能：删除一个标签组内，多余的标签。当标签组的标签数量减少时调用
'参数： lngBtnIndex --- 标签组中，第一个标签的索引
'       lngDelIndex --- 要开始删除的标签索引，这个标签以及其后的所有标签都将删除
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    Dim i  As Integer
    Dim iNextIndex As Integer
    
    On Error GoTo err
    
    
    i = lngDelIndex
    
    Do
        iNextIndex = btnSchLabel(i).HelpContextID

        btnSchLabel(i).Visible = False
        Unload btnSchLabel(i)   '不用的标签，直接卸载掉
        
        i = iNextIndex
    Loop While (btnSchLabel(i).HelpContextID <> lngBtnIndex) And (btnSchLabel(i).HelpContextID <> 0) And (i <> lngBtnIndex)
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ModifySchPoolTime(ByVal lngPoolIndex As Long, ByVal dtStartTime As Date) As Boolean
'------------------------------------------------
'功能：根据传入的开始时间，重新调整预约池中对应索引的开始和结束时间，保持预约的总时长不变
'参数： lngPoolIndex --- 预约池中的索引
'       dtStartTime --- 新的开始时间
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    Dim lngMinuteDiff As Long   '总时长
    
    On Error GoTo err
    lngMinuteDiff = DateDiff("n", mSchLabelPool(lngPoolIndex).dtStartTime, mSchLabelPool(lngPoolIndex).dtEndTime)
    mSchLabelPool(lngPoolIndex).dtStartTime = Format(mSchLabelPool(lngPoolIndex).dtStartTime, "YYYY-MM-DD") & " " & Format(dtStartTime, "HH:MM:SS")
    mSchLabelPool(lngPoolIndex).dtEndTime = DateAdd("n", lngMinuteDiff, mSchLabelPool(lngPoolIndex).dtStartTime)
    mSchLabelPool(lngPoolIndex).isModified = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function AdjustLabelPos(Index As Integer) As Boolean
'------------------------------------------------
'功能：拖拽鼠标结束，鼠标抬起来的时候，修正预约标签的位置
'参数： Index -- 标签的索引
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Integer
    Dim iPrePoolIndex As Integer
    Dim iNextPoolIndex As Integer
    Dim lngNewNumber As Long
    Dim iPoolIndex As Integer
    Dim blnIsOutOfTime As Boolean   '预约标签是否在预约时间计划内
    Dim iFirstIndex As Integer
    
    On Error GoTo err
    
    blnIsOutOfTime = False
    
    '首先判断是否涉及到标签拐弯的问题
    If btnSchLabel(Index).HelpContextID <> 0 Then   '是标签组
        '首先找到这个标签组中的第一个标签索引
        iFirstIndex = mSchLabelPool(btnSchLabel(Index).tag).lngBtnIndex
    Else    '不是标签组，当前索引就是第一个索引
        iFirstIndex = Index
    End If
    iPoolIndex = btnSchLabel(iFirstIndex).tag
    
    '如果标签被放在其他标签上面，给出提示，禁止预约
    If IsLabelOverlap(iFirstIndex) = True Then
        '不用提示，直接将标签移动回原来的位置
        
        mblnRestorePos = True
    End If
    
    '如果标签被拖拽到了非预约时间段，给出提示，但是允许预约
    If mblnRestorePos = False Then
        lngRow = GetRowsFromY(btnSchLabel(iFirstIndex).Top)
        lngCol = GetColsFromX(btnSchLabel(iFirstIndex).Left)
        
        If (Format(mdtSchDate, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD")) _
                And (Format(mSchLabelPool(iPoolIndex).dtStartTime, "HH:MM") < Format(Now, "HH:MM")) Then
            MsgBox "现在已经无法做检查，请更换一个时间段重新预约。", vbOKOnly, "检查预约提示"
            mblnRestorePos = True   '将标签摆放回原来的位置
        ElseIf vsfTime.Cell(flexcpData, lngRow, lngCol) = 0 Then
            blnIsOutOfTime = True
            MsgBox "不在可以预约的时间段内，不能继续预约。", vbOKOnly, "检查预约提示"
            mblnRestorePos = True   '将标签摆放回原来的位置

        End If
    End If
    
    '分配序号，开始预约
    If mblnRestorePos = False Then
        '给标签分配序号
        lngNewNumber = GetNewNumber(iFirstIndex)
        
        If lngNewNumber = 0 Then
            MsgBox "这个时间段已经没有空的预约序号，无法预约，请更换一个时间段重新预约。", vbOKOnly, "检查预约提示"
            mblnRestorePos = True   '将标签摆放回原来的位置
        ElseIf lngNewNumber = -1 Then
            mblnRestorePos = True   '将标签摆放回原来的位置
        End If
        
        '按照新序号预约
        If mblnRestorePos = False Then
            mSchLabelPool(iPoolIndex).lng序号 = lngNewNumber
            '重新调整预约池顺序
            Call ResortSchLabelPool(iPoolIndex)
        End If
    End If
    
     '如果标签被拖拽到了时间表之外，需要恢复标签的原来位置
    If mblnRestorePos = True Then
        Call RestoreLabelPos(iFirstIndex)
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsLabelOverlap(Index As Integer) As Boolean
'------------------------------------------------
'功能：判断当前标签，是否和其他标签在时间上覆盖了
'参数： Index --- 标签的索引
'返回：True -- 有覆盖； False -- 无覆盖
'------------------------------------------------
    Dim i As Integer
    Dim strS1 As String
    Dim strS2 As String
    Dim strE1 As String
    Dim strE2 As String
    Dim iFirstIndex As Integer
    
    On Error GoTo err
    
    iFirstIndex = mSchLabelPool(btnSchLabel(Index).tag).lngBtnIndex
    strS2 = Format(mSchLabelPool(btnSchLabel(Index).tag).dtStartTime, "HH:MM")
    strE2 = Format(mSchLabelPool(btnSchLabel(Index).tag).dtEndTime, "HH:MM")
    
    '根据时间来判断是否存在标签覆盖
    For i = 1 To UBound(mSchLabelPool)
        If mSchLabelPool(i).lngBtnIndex <> iFirstIndex Then
            strS1 = Format(mSchLabelPool(i).dtStartTime, "HH:MM")
            strE1 = Format(mSchLabelPool(i).dtEndTime, "HH:MM")
            
            If (strS1 <= strS2 And strS2 < strE1) _
                Or (strS1 < strE2 And strE2 <= strE1) _
                Or (strS2 <= strS1 And strS1 < strE2) Then
                IsLabelOverlap = True
                Exit Function
            End If
        End If
    Next i
    
    IsLabelOverlap = False
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetNewNumber(iBtnIndex As Integer) As Long
'------------------------------------------------
'功能：根据当前预约标签所在的位置，计算出新的预约序号
'参数： iBtnIndex --- 预约标签的索引
'返回：新的预约序号,-1表示用户主动停止预约；0表示没有序号；其他返回序号
'------------------------------------------------
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngTimeProjectID As Long
    Dim dtStartTime As Date
    Dim lngBaseNumber As Long       '当前时间计划之前的预约容量总和，序号基数
    Dim lngNumber As Long           '计算出来的新序号
    Dim iNumCount As Integer        '可用的序号总数
    Dim iTimeLength As Integer      '每一个序号所占的预计时间长度
    Dim dtProjectStartTime As Date  '预约时间计划表中，这个时间段的开始时间
    Dim dtProjectEndTime As Date    '预约时间计划表中，这个时间段的结束时间
    Dim lngPlanId As Long           '预约时间计划表中，记录的方案ID
    Dim iPoolIndex As Integer       '预约池中的索引
    Dim i As Integer                '循环变量
    Dim iNextPoolIndex As Integer   '新序号之后的第一个序号
    Dim iPrePoolIndex As Integer    '新序号之前的最后一个序号，跟自己的序号不同
    Dim strMsg As String            '信息
    
    GetNewNumber = 0
    
    On Error GoTo err
    
    iPoolIndex = btnSchLabel(iBtnIndex).tag
    
    '在预约池中循环对比，找到新预约序号的位置，按照时间进行对比
    For i = 1 To UBound(mSchLabelPool)
        If i <> iPoolIndex Then
            If DateDiff("n", mSchLabelPool(iPoolIndex).dtEndTime, mSchLabelPool(i).dtStartTime) >= 0 Then
                Exit For
            End If
        End If
    Next i
    
    '找到下一个和上一个索引，确定新序号所在的位置
    iNextPoolIndex = IIf(i > UBound(mSchLabelPool), 0, i)
    If iNextPoolIndex <> 0 And iNextPoolIndex - 1 = iPoolIndex Then
        iPrePoolIndex = iPoolIndex - 1
    ElseIf iNextPoolIndex <> 0 Then
        iPrePoolIndex = iNextPoolIndex - 1
    Else
        iPrePoolIndex = UBound(mSchLabelPool) - 1
    End If
    
    '如果比序号1小，则无号码
    If iNextPoolIndex = 1 And mSchLabelPool(1).lng序号 = 1 Then
        Exit Function   '没有序号，退出
    ElseIf iNextPoolIndex = 0 And iPoolIndex <> UBound(mSchLabelPool) Then
        '如果是移动到最后一个序号，且超出序号容量，提示是否加号，加号，或者无号码
        '如果是移动到最后一个序号，没有超出预约容量，继续计算
        If mSchLabelPool(UBound(mSchLabelPool)).lng序号 >= mlngSchSum Then
            If MsgBox("预约序号已经超过今天的预约总容量，是否继续加号？", vbYesNo, "检查预约提示") = vbYes Then
                GetNewNumber = mSchLabelPool(UBound(mSchLabelPool)).lng序号 + 1
                Exit Function   '找到序号，退出
            Else
                GetNewNumber = -1
                Exit Function   '没有序号，退出
            End If
        End If
    Else
        '如果是中间号码，前后号码之间无空号，则无号码
        If mSchLabelPool(iNextPoolIndex).lng序号 - mSchLabelPool(iPrePoolIndex).lng序号 = 1 Then
            Exit Function   '没有序号，退出
        ElseIf mSchLabelPool(iNextPoolIndex).lng序号 - mSchLabelPool(iPrePoolIndex).lng序号 = 2 Then
            '刚好有一个空号，返回这个空号
            GetNewNumber = mSchLabelPool(iNextPoolIndex).lng序号 - 1
            Exit Function   '找到序号，退出
        End If
    End If
    
    '是中间的序号，计算新的序号
    '根据标签的位置，就算出时间计划ID
    lngRow = GetRowsFromY(btnSchLabel(iBtnIndex).Top)
    lngCol = GetColsFromX(btnSchLabel(iBtnIndex).Left)
    lngTimeProjectID = vsfTime.Cell(flexcpData, lngRow, lngCol)
    'lngTimeProjectID = 0 预约标签不在时间计划之内，使用上一个时间段的最后一个号码，这样子就不会影响到下一个时间段的序号编号了。
    
    dtStartTime = mSchLabelPool(iPoolIndex).dtStartTime
    lngBaseNumber = 0
    '先判断当前标签所在位置上，预约时间计划的基础开始序号
    For i = 1 To UBound(mSchTimeProject)
        If Format(dtStartTime, "HH:MM:SS") >= Format(mSchTimeProject(i).dtEndTime, "HH:MM:SS") Then
            lngBaseNumber = lngBaseNumber + mSchTimeProject(i).lngSum
        End If
    Next i
    
    '如果lngTimeProjectID=0 ,表示没有在时间计划中做预约，则直接使用序号基数中的下一个序号
    If lngTimeProjectID = 0 Then
        lngNumber = IIf(lngBaseNumber = 0, 1, lngBaseNumber + 1)  '如果是0 ，则修正成1
    Else
        '在预约时间计划中，计算出这个位置的序号
        For i = 1 To UBound(mSchTimeProject)
            If mSchTimeProject(i).lngID = lngTimeProjectID Then
                iNumCount = mSchTimeProject(i).lngSum
                dtProjectStartTime = mSchTimeProject(i).dtStartTime
                dtProjectEndTime = mSchTimeProject(i).dtEndTime
                lngPlanId = mSchTimeProject(i).lngSchPlanID
                Exit For
            End If
        Next i
        
        '判断时间计划是否已经约满了
        If (SegmentCanUse(dtProjectStartTime, dtProjectEndTime, iNumCount, strMsg)) = 0 Then
            If MsgBox(strMsg & " 是否继续加号？", vbYesNo, "检查预约提示") = vbNo Then
                GetNewNumber = -1
                Exit Function   '没有序号，退出
            End If
        End If
        
        If iNumCount <> 0 Then
            iTimeLength = DateDiff("n", dtProjectStartTime, dtProjectEndTime) / iNumCount
            For i = 1 To iNumCount
                '根据时间段的长度和序号容量，计算出新序号
                If Format(dtStartTime, "HH:MM") < Format(DateAdd("n", iTimeLength * i, dtProjectStartTime), "HH:MM") Then
                    Exit For
                End If
            Next i
            If i > iNumCount Then
                i = iNumCount
            End If
            lngNumber = lngBaseNumber + i
        End If
    End If
    
    '计算出来的新序号，需要比当前位置之前预约标签的序号大，
    '这种情况应该是不存在的，除非是用户手工调整了很多个预约标签的宽度，导致很多序号聚集在较短的时间内
    If lngNumber <= mSchLabelPool(iPrePoolIndex).lng序号 Then
        lngNumber = mSchLabelPool(iPrePoolIndex).lng序号 + 1
    ElseIf lngNumber >= mSchLabelPool(iNextPoolIndex).lng序号 And iNextPoolIndex <> 0 Then
        lngNumber = mSchLabelPool(iNextPoolIndex).lng序号 - 1
    End If
    
    '再次往后寻找一个空的序号
    For i = iPrePoolIndex To UBound(mSchLabelPool)
        If mSchLabelPool(i).lng序号 = lngNumber And i <> iPoolIndex Then
            lngNumber = lngNumber + 1
        End If
    Next i
    
    If lngNumber > mlngSchSum Then
        If MsgBox("预约序号已经超过今天的预约总容量，是否继续加号？", vbYesNo, "检查预约提示") = vbNo Then
            GetNewNumber = -1
            Exit Function   '没有序号，退出
        End If
    End If
        
    GetNewNumber = lngNumber

    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub SaveAllSchedule()
'------------------------------------------------
'功能：保存所有被修改过的预约信息
'参数：
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strStartTime As String
    Dim strEndTime As String
    
    On Error GoTo err
    
    For i = 1 To UBound(mSchLabelPool)
        If mSchLabelPool(i).isModified = True Then
            If funSaveSchedule(mSchLabelPool(i).dtStartTime, mSchLabelPool(i).dtEndTime, _
                mSchLabelPool(i).lng医嘱ID, mSchLabelPool(i).str姓名, mSchLabelPool(i).lng序号, _
                mlngSchDeviceID, mSchLabelPool(i).dt开始时间段, mSchLabelPool(i).dt结束时间段) = False Then
                Exit Sub
            End If
            
            If InStr(mstrModifiedOrderID, CStr(mSchLabelPool(i).lng医嘱ID)) = 0 Then
                mstrModifiedOrderID = mstrModifiedOrderID & "," & CStr(mSchLabelPool(i).lng医嘱ID)
            End If
        End If
    Next i
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SegmentCanUse(dtProjectStartTime As Date, dtProjectEndTime As Date, iCapacity As Integer, ByRef strMsg As String) As Boolean
'------------------------------------------------
'功能：判断当前时间段是否可用
'参数： dtProjectStartTime -- 开始时间
'       dtProjectEndTime -- 结束时间
'       iCapacity -- 预约容量
'       strMsg -- 【OUT】时间段不可用时，返回原因
'返回：True -- 可用；False -- 不可用
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim dtStartTime As Date
    Dim dtEndTime As Date
       
    On Error GoTo err
    
    '小于今天，返回False
    If Format(mdtSchDate, "YYYY-MM-DD") < Format(Now, "YYYY-MM-DD") Then
        strMsg = "今天以前的时间，已经不能预约。"
        Exit Function
    End If
    
    dtStartTime = CDate(Format(mdtSchDate, "YYYY-MM-DD") & " " & Format(dtProjectStartTime, "hh:mm:ss"))
    dtEndTime = CDate(Format(mdtSchDate, "YYYY-MM-DD") & " " & Format(dtProjectEndTime, "hh:mm:ss"))
    
    '是今天，如果当前时间，距离结束时间不足2小时，也返回False
    If Format(mdtSchDate, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
        If DateDiff("n", Now, dtEndTime) <= 120 Then
            strMsg = "现在距离本时间段的结束时间不足2小时。"
            Exit Function
        End If
    End If
    
    '判断这个时间段的预约总数是否大于预约时间计划的容量
    strSQL = "Select Count(a.序号) as SchCount From 影像预约记录 A Where a.预约设备id = [1] And " _
            & " a.预约开始时间 >= [2]  And a.预约结束时间 <=[3] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "时间段内已经预约的数量", mlngSchDeviceID, _
        CDate(dtStartTime), CDate(dtEndTime))
    If rsTemp!SchCount >= iCapacity Then
        strMsg = "当前时间段预约容量已满。"
        Exit Function
    End If
        
    SegmentCanUse = True
    
    Exit Function
    
err:
   If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ClearLocalParas()
'------------------------------------------------
'功能：清空模块变量
'参数：
'返回：无
'------------------------------------------------
    
    On Error GoTo err
    
    mlngPoolIndex = 0
    mlngBtnIndex = 0
    mlngOrderID = 0
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadCalendar()
'------------------------------------------------
'功能：从数据库读取预约日历
'参数：
'返回：无
'------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo err
    
    strSQL = "select 年月,休息日 from 影像预约日历 where 年月>=[1]"
    Set mrsCalendar = zlDatabase.OpenSQLRecord(strSQL, "查询预约日历", Format(Now, "YYYYMM"))
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsDayOff(dtDate As Date) As Boolean
'------------------------------------------------
'功能：判断是否休息日
'参数： dtDate -- 判断的日期
'返回： True -- 是休息日 ； False -- 是工作日
'------------------------------------------------
    Dim strFilter As String
    
    On Error GoTo err
    
    IsDayOff = False
    mrsCalendar.Filter = 0
    If mrsCalendar.RecordCount = 0 Then
        Exit Function
    End If
    
    strFilter = "年月=" & Format(dtDate, "YYYYMM")
    mrsCalendar.Filter = strFilter
    If mrsCalendar.RecordCount = 1 Then
        If InStr(mrsCalendar!休息日, Format(dtDate, "DD")) > 0 Then
            IsDayOff = True
        End If
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub setReadOnly()
'------------------------------------------------
'功能：设置是否只读模式
'参数：
'返回： 无
'------------------------------------------------
        
    On Error GoTo err
    
    vsfTime.Enabled = Not mIsReadOnly
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub setSchLabelTag(iIndex As Long, iPoolIndex As Integer)
'------------------------------------------------
'功能：设置预约标签的TAG值，如果是拐弯的多个标签，同时设置所有标签的TAG值
'参数： iIndex -- 预约标签的索引
'       iPoolIndex -- 预约标签对应缓冲池的索引
'返回： 无
'------------------------------------------------
    Dim iTempIndex As Integer
        
    On Error GoTo err
    
    btnSchLabel(iIndex).tag = iPoolIndex
    
    If btnSchLabel(iIndex).HelpContextID <> 0 Then
        iTempIndex = btnSchLabel(iIndex).HelpContextID
        While btnSchLabel(iIndex).HelpContextID <> btnSchLabel(iTempIndex).HelpContextID
            btnSchLabel(iTempIndex).tag = iPoolIndex
            iTempIndex = btnSchLabel(iTempIndex).HelpContextID
        Wend
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub setSchLabelEnable(iIndex As Integer, blnEnable As Boolean)
'------------------------------------------------
'功能：设置预约标签的可用性，如果是拐弯的多个标签，同时设置所有标签的可用性
'参数： iIndex -- 预约标签的索引
'       blnEnable -- 按钮是否可用
'返回： 无
'------------------------------------------------
    Dim iTempIndex As Integer
        
    On Error GoTo err
    
    btnSchLabel(iIndex).Enabled = blnEnable
    
    If btnSchLabel(iIndex).HelpContextID <> 0 Then
        iTempIndex = btnSchLabel(iIndex).HelpContextID
        While btnSchLabel(iIndex).HelpContextID <> btnSchLabel(iTempIndex).HelpContextID
            btnSchLabel(iTempIndex).Enabled = blnEnable
            iTempIndex = btnSchLabel(iTempIndex).HelpContextID
        Wend
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub setSchLabelZorder(iIndex As Integer)
'------------------------------------------------
'功能：设置预约标签的Zorder，显示到最前面
'参数： iIndex -- 预约标签的索引
'返回： 无
'------------------------------------------------
    Dim iTempIndex As Integer
        
    On Error GoTo err
    
    Call btnSchLabel(iIndex).ZOrder
    
    If btnSchLabel(iIndex).HelpContextID <> 0 Then
        iTempIndex = btnSchLabel(iIndex).HelpContextID
        While btnSchLabel(iIndex).HelpContextID <> btnSchLabel(iTempIndex).HelpContextID
            Call btnSchLabel(iTempIndex).ZOrder
            iTempIndex = btnSchLabel(iTempIndex).HelpContextID
        Wend
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub setSchLabelToolTipText(ByVal iIndex As Integer)
'------------------------------------------------
'功能：设置预约标签的Zorder，显示到最前面
'参数： iIndex -- 预约标签的索引
'返回： 无
'------------------------------------------------
    Dim iPoolIndex As Integer
    
    On Error GoTo err
    
    iPoolIndex = Val(btnSchLabel(iIndex).tag)
    btnSchLabel(iIndex).ToolTipText = "  序号：" & mSchLabelPool(iPoolIndex).lng序号 & vbCrLf & "  姓名：" & mSchLabelPool(iPoolIndex).str姓名 _
            & vbCrLf & "  医嘱内容：" & mSchLabelPool(iPoolIndex).str医嘱内容 & vbCrLf & "  开始时间：" & Format(mSchLabelPool(iPoolIndex).dtStartTime, "HH:MM") _
            & vbCrLf & "  结束时间：" & Format(mSchLabelPool(iPoolIndex).dtEndTime, "HH:MM")
            
    mstrOrderInfo = btnSchLabel(iIndex).ToolTipText
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub setSchLabelSelectTag(iIndex As Integer)
'------------------------------------------------
'功能：设置预约标签的被选中标记
'参数： iIndex -- 预约标签的索引
'返回： 无
'------------------------------------------------
    Dim iTempIndex As Integer
    Dim i As Integer
    Dim btnLabel As CommandButton
    
    On Error GoTo err
    
    For Each btnLabel In btnSchLabel
        If btnLabel.Index <> 0 Then
            btnLabel.Font.Bold = False
            btnLabel.Font.Size = mlngFontSize
        End If
    Next
    
    btnSchLabel(iIndex).Font.Bold = True
    btnSchLabel(iIndex).Font.Size = mlngFontSize + 2
    
    If btnSchLabel(iIndex).HelpContextID <> 0 Then
        iTempIndex = btnSchLabel(iIndex).HelpContextID
        While btnSchLabel(iIndex).HelpContextID <> btnSchLabel(iTempIndex).HelpContextID
            btnSchLabel(iTempIndex).Font.Bold = True
            btnSchLabel(iTempIndex).Font.Size = mlngFontSize + 3
            iTempIndex = btnSchLabel(iTempIndex).HelpContextID
        Wend
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CanResizeLabel(iIndex As Integer) As Boolean
'------------------------------------------------
'功能：判断这个预约标签是否可以改变宽度，如果是一组标签，只允许修改最后一个标签的宽度
'参数： iIndex -- 预约标签的索引
'返回： True -- 可以修改；False -- 不能修改
'------------------------------------------------
        
    On Error GoTo err
    
    If btnSchLabel(iIndex).HelpContextID = 0 Or (btnSchLabel(iIndex).HelpContextID = mSchLabelPool(btnSchLabel(iIndex).tag).lngBtnIndex) Then
        CanResizeLabel = True
    Else
        CanResizeLabel = False
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RestoreLabelPos(iIndex As Integer)
'------------------------------------------------
'功能：恢复标签原先的位置
'参数： iIndex -- 预约标签的索引
'返回：
'------------------------------------------------
    Dim iFirstIndex As Integer
    Dim iPoolIndex As Integer
    
    On Error GoTo err
    iPoolIndex = btnSchLabel(iIndex).tag
    iFirstIndex = mSchLabelPool(iPoolIndex).lngBtnIndex
    
    '先计算结束时间，再计算开始时间
    'lngMinutes = DateDiff("n", mSchLabelPool(iPoolIndex).dtStartTime, mSchLabelPool(iPoolIndex).dtEndTime)
    
    mSchLabelPool(iPoolIndex).dtStartTime = Format(mSchLabelPool(iPoolIndex).dtStartTime, "YYYY-MM-DD") & " " & Format(GetTimeFromXY(mlngOriginLeft, mlngOriginTop), "HH:MM")
    '结束时间，需要根据标签组来计算
    mSchLabelPool(iPoolIndex).dtEndTime = DateAdd("n", mlngOriginMinute, mSchLabelPool(iPoolIndex).dtStartTime)
    Call PutSchLabel(iFirstIndex, iPoolIndex)
   
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetColor(ByVal lngID As Long, ByVal blnClear As Boolean)
On Error GoTo errH
    '所有方案重新上色
    Dim i As Long, j As Long
    
    Dim lngColorTabWorkSel As Long
    Dim intR As Integer, intG As Integer, intB As Integer
    Dim intTMP As Integer
    Dim lngColorTabWork(1) As Long
    
    With vsfTime
        intR = (mlngColorTabWork And &HFF) Mod 256
        intG = ((mlngColorTabWork And &HFF00) \ &H100) Mod 256
        intB = ((mlngColorTabWork And &HFF0000) \ &H10000) Mod 256
        
        intR = intR - 50
        If intR < 1 Then intR = 100
        
        intG = intG - 50
        If intG < 1 Then intG = 100
        
        intB = intB - 50
        If intB < 1 Then intB = 100
        
        lngColorTabWorkSel = RGB(intR, intG, intB)
        lngColorTabWork(0) = mlngColorTabWork
        
        
        intR = (mlngColorTabWork And &HFF) Mod 256
        intG = ((mlngColorTabWork And &HFF00) \ &H100) Mod 256
        intB = ((mlngColorTabWork And &HFF0000) \ &H10000) Mod 256
        
        intR = intR + 40
        If intR > 255 Then intR = 165
        
        intG = intG + 40
        If intG > 255 Then intG = 165
        
        intB = intB + 40
        If intB > 255 Then intB = 165
        lngColorTabWork(1) = RGB(intR, intG, intB)
        
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                If .Cell(flexcpData, i, j) = lngID Then
                    '突出显示
                    If Not blnClear Then
                        .Cell(flexcpBackColor, i, j) = lngColorTabWorkSel
                    End If
                ElseIf .Cell(flexcpData, i, j) <> 0 Then
                    '还原工作颜色
                    
                    If mColordict.Item(.Cell(flexcpData, i, j)) = True Then
                        .Cell(flexcpBackColor, i, j) = lngColorTabWork(0)
                    Else '
                        .Cell(flexcpBackColor, i, j) = lngColorTabWork(1)
                    End If
                    
                Else
                    '什么都不做
                End If
            Next
        Next
    End With

    Exit Sub
errH:
    MsgBox err.Description, vbOKOnly, "检查预约提示"
End Sub

Private Sub ShowMouseTime(lngX As Long, lngY As Long)
'未拖动时间状态下根据鼠标位置显示当前所属方案
On Error GoTo errH
    Dim lngRow As Long, lngCol As Long, lngTimeProjectID As Long
        
    If lngX <> 0 And lngY <> 0 Then
        lngRow = GetRowsFromY(lngY)
        lngCol = GetColsFromX(lngX)
        lngTimeProjectID = vsfTime.Cell(flexcpData, lngRow, lngCol)
        
        If lngTimeProjectID = 0 Then
            Call SetColor(0, True)
        Else
            Call SetColor(lngTimeProjectID, False)
        End If
    End If
    Exit Sub
errH:
    MsgBox err.Description
End Sub

Private Sub SetMouseTimePro(ByVal lngX As Long, ByVal lngY As Long, ByVal IsMove As Boolean)
'根据鼠标在list中的X,Y,和当前操作的时间标签，自动把时间标签移动到点击位置对应的时间，如果已经放不下，弹出提示 换一个时间段
    
    Dim dtStartTime As Date         '开始时间
    Dim dtEndTime As Date           '结束时间
    Dim blnTimeOK As Boolean '时间符合条件
    Dim blnNeedTestTime As Boolean '需要验证时间
    Dim dtStartTimeSQL As Date         '开始时间
    Dim dtEndTimeSQL As Date           '结束时间
    Dim intPoolIndex As Integer
    Dim i As Integer, j As Integer
    Dim iFirstIndex As Integer
    Dim iBtnIndex As Integer
    Dim blnFind As Boolean
    
    Dim lngLeft As Long
    Dim lngTop As Long
    On Error GoTo err
    
    blnNeedTestTime = False
    If mlngOrderID = 0 Then Exit Sub
    For i = 1 To UBound(mSchLabelPool)
        If mSchLabelPool(i).lng医嘱ID = mlngOrderID Then
            iBtnIndex = mSchLabelPool(i).lngBtnIndex
            Exit For
        End If
    Next
    
    If btnSchLabel(iBtnIndex).HelpContextID <> 0 Then   '是标签组
        '首先找到这个标签组中的第一个标签索引
        iFirstIndex = mSchLabelPool(btnSchLabel(iBtnIndex).tag).lngBtnIndex
    Else    '不是标签组，当前索引就是第一个索引
        iFirstIndex = iBtnIndex
    End If

    Call AdjustLabelPos(iFirstIndex)

    mlngPoolIndex = btnSchLabel(iFirstIndex).tag
    mSchLabelPool(btnSchLabel(iFirstIndex).tag).bln已保存 = False
    Call setSchLabelToolTipText(iFirstIndex)
    
    If IsMove Then
        '根据控件当前最左边位置计算所属时间块
        lngLeft = btnSchLabel(iFirstIndex).Left
        lngTop = btnSchLabel(iFirstIndex).Top + 0.5 * (btnSchLabel(iFirstIndex).Height)
        dtStartTimeSQL = GetTimeFromXY(lngLeft, lngTop)
    Else
        '根据X ,Y 计算所属时间段
        dtStartTimeSQL = GetTimeFromXY(lngX, lngY)
    End If
    dtEndTimeSQL = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTimeSQL)
    intPoolIndex = btnSchLabel(iBtnIndex).tag
    
    For i = 1 To UBound(mSchTimeProject)
         If mSchTimeProject(i).dtStartTime <= dtStartTimeSQL And mSchTimeProject(i).dtEndTime > dtStartTimeSQL Then
            dtStartTimeSQL = mSchTimeProject(i).dtStartTime
            dtEndTimeSQL = mSchTimeProject(i).dtEndTime
            blnFind = True
            Exit For
         End If
    Next
    
    If Not blnFind Then
        Exit Sub
    End If
  
    blnNeedTestTime = True
    If UBound(mSchLabelPool) = 1 Then
        dtStartTime = dtStartTimeSQL
        dtEndTime = dtEndTimeSQL
    Else
    
        If mSchLabelPool(1).dtStartTime >= Format(mdtSchDate, "YYYY-MM-DD") & dtEndTimeSQL Then
            '满足条件退出while循环
            dtStartTime = dtStartTimeSQL
            dtEndTime = dtEndTimeSQL
            blnNeedTestTime = False
        End If
        
        '首先仍然假设时间段开始时间作为预约开始时间有效
        blnTimeOK = True
        dtStartTime = dtStartTimeSQL
        dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
    
        For i = 1 To UBound(mSchLabelPool)
            '避开当前时间块
            If intPoolIndex <> i Then '需要修改这个条件
                If Not ((Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") < dtStartTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") <= dtStartTime) Or (Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtEndTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") > dtEndTime)) Then
                    blnTimeOK = False
                    Exit For
                End If
            End If
        Next
        
        If Not blnTimeOK Then
        
            For i = 1 To UBound(mSchLabelPool)
                If blnNeedTestTime And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtStartTimeSQL And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") <= dtEndTimeSQL Then  'And iBtnIndex <> i
                    blnTimeOK = True
                    
                    If blnTimeOK And blnNeedTestTime Then
                        dtStartTime = Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") 'rsSchedule!预约结束时间 'Format(rsSchedule!预约结束时间, "hh-mm-ss")
                        dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
                    
                        For j = 1 To UBound(mSchLabelPool)
                            If intPoolIndex <> j Then '需要修改这个条件
                                If Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtStartTimeSQL And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") <= dtEndTimeSQL Then
                                    If Not ((Format(mSchLabelPool(j).dtStartTime, "hh:mm:ss") < dtStartTime And Format(mSchLabelPool(j).dtEndTime, "hh:mm:ss") <= dtStartTime) Or (Format(mSchLabelPool(j).dtStartTime, "hh:mm:ss") >= dtEndTime And Format(mSchLabelPool(j).dtEndTime, "hh:mm:ss") > dtEndTime)) Then
                                        blnTimeOK = False
                                    End If
                                End If
                            End If
                        Next
    
                        If blnTimeOK Then
                            blnNeedTestTime = False
                        End If
                    End If
                End If
            Next
        End If
    End If

    Call ModifySchPoolTime(btnSchLabel(iFirstIndex).tag, dtStartTime)
    '重新摆放标签
    Call PutSchLabel(iFirstIndex, btnSchLabel(iFirstIndex).tag)

    If iFirstIndex <> iBtnIndex Then
        mlngBaseX = lngX
    End If
    
    RaiseEvent OnSchLabelModifed(iFirstIndex)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
