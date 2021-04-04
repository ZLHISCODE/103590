Attribute VB_Name = "mdlClinicPlanGridFun"
Option Explicit
Public Enum gPlanGrid_ColIndex '网格固定列索引
    COL_图标
    COL_出诊ID
    COL_号源ID
    COL_安排ID
    COL_号类
    COL_号码
    COL_科室
    COL_项目
    COL_医生
    Col_医生职称
    COL_是否建病案
    COL_预约天数
    COL_出诊频次
    COL_假日控制状态
    COL_假日换休
    COL_排班方式
    COL_开始时间
    COL_终止时间
    COL_登记时间
    COL_临时安排
    COL_是否审核
    COL_是否临床排班
End Enum

'安排列表固定列数
Public Const gPlanGrid_FixedCols = 22

'安排表类型
Public Enum gPlanGrid_DataStyle
    Data_Templet = 0
    Data_FixedRule = 1
    Data_Plan = 2
    Data_MonthTemplet = 3 '月模板，月出诊表生成的模板
End Enum

Public Sub InitPlanGrid(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    Optional ByVal dtMinDate As Date, Optional ByVal dtMaxDate As Date, _
    Optional ByVal blnPublished As Boolean)
    '功能：初始化安排数据表格
    '   vsfGrid - VSF表格
    '   bytDataStyle - 数据类型
    Dim strHead As String, varData As Variant
    Dim strHeadSub As String, varDataSub As Variant
    Dim i As Long, lngCol As Long
    Dim arrDate As Variant, strTemp As String
    Dim dtCurdate As Date, intDays As Integer

    Err = 0: On Error GoTo errHandler
    With vsfGrid
        .Redraw = False
        .Rows = 2
        
        '固定列
        strHead = " ,4,300|出诊ID,4,0|号源ID,4,0|安排ID,4,0|号类,4,0|号码,4,500|科室,1,1000|项目,1,0|医生,1,850|医生职称,1,0|" & _
                "建档,4,0|预约天数,4,0|出诊频次,4,0|假日控制状态,1,0|假日换休,4,0|排班方式,4,0|开始时间,1,0|终止时间,1,0|" & _
                "登记时间,1,0|临时安排,4,0|是否审核,4,0|是否临床排班,4,0"
        strHeadSub = " ,出诊ID,号源ID,安排ID,号类,号码,科室,项目,医生,医生职称," & _
                "建档,预约天数,出诊频次,假日控制状态,假日换休,排班方式,开始时间,终止时间," & _
                "登记时间,临时安排,是否审核,是否临床排班"
        varData = Split(strHead, "|")
        varDataSub = Split(strHeadSub, ",")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0): .TextMatrix(1, i) = varDataSub(i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .ColKey(i) = Split(varData(i), ",")(0)
        Next
        .FixedCols = 9: .FixedRows = 2
        '动态列
        Select Case bytDataStyle
        Case Data_Templet, Data_FixedRule   '模板,固定规则
            strHead = "周一,1,450|周一,4,550|周一,4,550|周二,1,450|周二,4,450|周二,4,550|周三,1,450|周三,4,550|周三,4,550|" & _
                    "周四,1,450|周四,4,550|周四,4,550|周五,1,450|周五,4,550|周五,4,550|周六,1,450|周六,4,550|周六,4,550|" & _
                    "周日,1,450|周日,4,550|周日,4,550"
            strHeadSub = "时段,限号,限约,时段,限号,限约,时段,限号,限约," & _
                    "时段,限号,限约,时段,限号,限约,时段,限号,限约," & _
                    "时段,限号,限约"
            If bytDataStyle = Data_Templet Then
                strHead = strHead & "|其他规则,1,1150|其他规则,1,450|其他规则,4,550|其他规则,4,550"
                strHeadSub = strHeadSub & ",限制项目,时段,限号,限约"
            End If
            varData = Split(strHead, "|")
            varDataSub = Split(strHeadSub, ",")
            lngCol = .Cols
            .Cols = .Cols + UBound(varData) + 1
            For i = 0 To UBound(varData)
                .TextMatrix(0, lngCol) = Split(varData(i), ",")(0): .TextMatrix(1, lngCol) = varDataSub(i)
                .Cell(flexcpData, 0, lngCol) = CStr(Split(varData(i), ",")(0))
                .ColAlignment(lngCol) = Split(varData(i), ",")(1)
                .ColWidth(lngCol) = Split(varData(i), ",")(2)
                lngCol = lngCol + 1
            Next
            .FixedAlignment(-1) = flexAlignCenterCenter
            .RowHeight(0) = 420: .RowHeight(1) = 300
            
            .AllowSelection = False
        Case Data_Plan, Data_MonthTemplet '安排记录
            intDays = DateDiff("d", dtMinDate, dtMaxDate) + 1 '天数
            If intDays < 0 Then intDays = 0
            dtCurdate = dtMinDate
            lngCol = .Cols
            .Cols = .Cols + intDays * 3
            For i = 1 To intDays
                If bytDataStyle = Data_MonthTemplet Then
                    .Cell(flexcpText, 0, lngCol, 0, lngCol + 2) = Day(dtCurdate) & "日 "
                Else
                    strTemp = Decode(bytDataStyle, Data_MonthTemplet, Day(dtCurdate) & "日", Format(dtCurdate, "mm月dd日")) & _
                              Chr(13) & GetWeekName(Weekday(dtCurdate, vbMonday) - 1)
                    .TextMatrix(0, lngCol) = strTemp
                    .TextMatrix(0, lngCol + 1) = strTemp
                    .TextMatrix(0, lngCol + 2) = strTemp
                End If
                .Cell(flexcpData, 0, lngCol, 0, lngCol + 2) = Format(dtCurdate, "yyyy-MM-dd") '日期
                .Cell(flexcpText, 1, lngCol, 1, lngCol + 2) = "时段" & vbTab & "限号" & vbTab & "限约"
                .ColAlignment(lngCol) = 1: .ColAlignment(lngCol + 1) = 4: .ColAlignment(lngCol + 2) = 4
                .ColWidth(lngCol) = 450
                .ColWidth(lngCol + 1) = IIf(bytDataStyle = Data_MonthTemplet Or blnPublished = False, 550, 650)
                .ColWidth(lngCol + 2) = IIf(bytDataStyle = Data_MonthTemplet Or blnPublished = False, 550, 650)
                dtCurdate = DateAdd("d", 1, dtCurdate)
                lngCol = lngCol + 3
            Next
            .FixedAlignment(-1) = flexAlignCenterCenter
            If bytDataStyle = Data_MonthTemplet Then
                .RowHeight(0) = 420: .RowHeight(1) = 300
            Else
                .RowHeight(0) = 500: .RowHeight(1) = 300
            End If
            
            .AllowSelection = blnPublished
        End Select
        
        .AllowBigSelection = False
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .GridLines = flexGridFlat
        .PicturesOver = True '文字在图片上面
        
        '列属性设置,用于用户选择显示列
        For i = 0 To .Cols - 1
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)|列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            Select Case i
            Case COL_出诊ID, COL_号源ID, COL_安排ID, COL_临时安排, COL_是否审核
                 .ColData(i) = "-1|1"
            Case COL_号码, COL_科室, COL_医生
                .ColData(i) = "1|0"
            End Select
        Next
        '非固定规则时，开始时间和终止时间不显示
        If bytDataStyle <> Data_FixedRule Then
            .ColData(COL_开始时间) = "-1|1": .ColData(COL_终止时间) = "-1|1": .ColData(COL_登记时间) = "-1|1"
        End If

        '合并设置
        .MergeCells = flexMergeRestrictColumns
        .MergeRow(0) = True: .MergeCol(-1) = True
        .Redraw = True
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function LoadPlanDataByRecordset(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    ByVal rsData As ADODB.Recordset, Optional ByVal bytTableMode As Byte, _
    Optional ByRef lngSignalCount As Long, Optional ByVal blnSortLoad As Boolean, _
    Optional ByVal blnPublished As Boolean, _
    Optional ByVal strStartDate As String, Optional ByVal strEndDate As String) As Boolean
    '功能：根据Recordset对象加载数据
    '入参：
    '   bytTableMode 0-固定出诊表,1-月出诊表,2-周出诊表,3-模板
    '   strStartDate,strEndDate 出诊表日期范围，周出诊表时传入
    '出参：
    '   lngSignalCount - 号源数量
    '说明：数据必须是按"号类,号码,科室,项目,医生"进行排序了的，否则可能显示不正确
    Dim i As Long, j As Long, lngCurRow As Long, lngCurCol As Long
    Dim strGroupKey As String '用于纵向按"号类,号码,科室,项目,医生"分组
    Dim lngBackColor As Long '设置纵向组的交替色
    Dim strTemp As String, blnAddRow As Boolean
    Dim lngRowStart As Long, lngRowEnd As Long
    Dim lngoldCol As Long, blnFindCol As Boolean
    Dim lngOldRow As Long, str限约数 As String
    Dim strRecordInfo As String
    
    Err = 0: On Error GoTo errHandle
    lngSignalCount = 0
    '记录当前选择单元格，用于恢复选择
    lngOldRow = vsfGrid.Row: lngoldCol = vsfGrid.Col
    '先清空数据
    vsfGrid.Clear 1: vsfGrid.Rows = 2
    
    If rsData Is Nothing Then Exit Function
    If rsData.RecordCount = 0 Then Exit Function
    
    rsData.MoveFirst
    With vsfGrid
        lngCurRow = 2
        strGroupKey = ""
        lngBackColor = G_AlternateColor
        .Redraw = flexRDNone
        Do While Not rsData.EOF
            blnFindCol = False
            '1.纵向分组
            strTemp = Nvl(rsData!号码) & "," & Nvl(rsData!科室) & "," & Nvl(rsData!收费项目) & "," & Nvl(rsData!医生姓名)
            If bytDataStyle = Data_FixedRule Then strTemp = strTemp & "," & Nvl(rsData!安排ID)
            If strGroupKey <> strTemp Then
                lngSignalCount = lngSignalCount + 1
                strGroupKey = strTemp
                lngCurCol = gPlanGrid_FixedCols  '用以判断是否确定了列
                lngBackColor = IIf(lngBackColor = vbWindowBackground, G_AlternateColor, vbWindowBackground)
                
                .Rows = .Rows + 1: lngCurRow = .Rows - 1
                .RowData(lngCurRow) = -1 '标记，用于判断是否为隐藏空行
                
                lngCurRow = lngCurRow + 1
            End If
            '2.横向分组
            '2.1确定当前列
            Select Case bytDataStyle
            Case Data_Templet  '模板
                If Nvl(rsData!排班规则) <> 1 Then '其它规则
                    '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
                    lngCurCol = .Cols - 4: blnFindCol = True
                Else
                    If Nvl(rsData!限制项目) <> "" Then
                        strTemp = Nvl(rsData!限制项目)
                        For i = lngCurCol To .Cols - 1 Step 3
                            If strTemp = .Cell(flexcpData, 0, i) Then
                                lngCurCol = i: blnFindCol = True
                                Exit For
                            End If
                        Next
                        '没找到再从开始重新找，主要是按限制项目排序跟界面顺序不一致
                        If blnFindCol = False Then
                            For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
                                If strTemp = .Cell(flexcpData, 0, i) Then
                                    lngCurCol = i: blnFindCol = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            Case Data_FixedRule  '固定规则
                If Nvl(rsData!限制项目) <> "" Then
                    strTemp = Nvl(rsData!限制项目)
                    For i = lngCurCol To .Cols - 1 Step 3
                        If strTemp = .Cell(flexcpData, 0, i) Then
                            lngCurCol = i: blnFindCol = True
                            Exit For
                        End If
                    Next
                    '没找到再从开始重新找，主要是按限制项目排序跟界面顺序不一致
                    If blnFindCol = False Then
                        For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
                            If strTemp = .Cell(flexcpData, 0, i) Then
                                lngCurCol = i: blnFindCol = True
                                Exit For
                            End If
                        Next
                    End If
                End If
            Case Else '安排记录
                If Nvl(rsData!出诊日期) = "" Then
                    '出诊日期为空,取安排的开始时间,主要针对整周跨月的周出诊表无出诊记录时,两个安排显示在同一行
                    strTemp = Format(Nvl(rsData!开始时间), "yyyy-mm-dd")
                    lngCurCol = gPlanGrid_FixedCols
                Else
                    strTemp = Format(Nvl(rsData!出诊日期), "yyyy-mm-dd")
                End If
                If IsDate(strTemp) Then
                    For i = lngCurCol To .Cols - 1 Step 3
                        If DateDiff("d", strTemp, .Cell(flexcpData, 0, i)) = 0 Then
                            lngCurCol = i: blnFindCol = True
                            Exit For
                        End If
                    Next
                End If
            End Select
            
            If blnFindCol Then
                '2.2确定当前行
                For i = IIf(.Rows - 1 > lngCurRow, lngCurRow, .Rows - 1) To 2 Step -1
                    If .RowData(i) = -1 Or .TextMatrix(i, lngCurCol) <> "" Then  '是隐藏空行或者无数据行
                        lngCurRow = i + 1: Exit For
                    End If
                Next
            End If
            
            '3.加载数据
            blnAddRow = False
            If .Rows - 1 < lngCurRow Then
                '已有行不够，加1行
                .Rows = .Rows + 1: lngCurRow = .Rows - 1
                .RowData(lngCurRow) = lngBackColor '用于设置交替色
                
                Select Case bytTableMode
                Case 0 '固定出诊表
                    If Val(Nvl(rsData!是否有效)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("FixedItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("InvalidFixedItem")
                    End If
                Case 1 '月出诊表
                    If Val(Nvl(rsData!是否有效)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("MonthItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("InvalidMonthItem")
                    End If
                Case 2 '周出诊表
                    If Val(Nvl(rsData!是否有效)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("WeekItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("InvalidWeekItem")
                    End If
                Case 3 '模板
                    If Nvl(rsData!排班方式) = "按月排班" Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("MonthItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("WeekItem")
                    End If
                End Select
                .Cell(flexcpPictureAlignment, lngCurRow, COL_图标) = flexAlignCenterCenter

                If Not (bytDataStyle = Data_Plan And IsDate(strStartDate) And IsDate(strEndDate)) Then
                    .TextMatrix(lngCurRow, COL_安排ID) = Nvl(rsData!安排ID)
                End If
                .TextMatrix(lngCurRow, COL_号源ID) = Nvl(rsData!号源ID)
                .TextMatrix(lngCurRow, COL_号类) = Nvl(rsData!号类)
                .TextMatrix(lngCurRow, COL_号码) = Nvl(rsData!号码)
                .TextMatrix(lngCurRow, COL_科室) = Nvl(rsData!科室)
                .TextMatrix(lngCurRow, COL_项目) = Nvl(rsData!收费项目)
                .TextMatrix(lngCurRow, COL_医生) = Nvl(rsData!标识符) & Nvl(rsData!医生姓名)
                .Cell(flexcpData, lngCurRow, COL_医生) = Nvl(rsData!医生姓名)
                .TextMatrix(lngCurRow, Col_医生职称) = Nvl(rsData!医生职称)

                .TextMatrix(lngCurRow, COL_是否建病案) = IIf(Val(Nvl(rsData!是否建病案)) = 1, "√", "")
                .TextMatrix(lngCurRow, COL_预约天数) = Nvl(rsData!预约天数)
                .TextMatrix(lngCurRow, COL_出诊频次) = Nvl(rsData!出诊频次)
                .TextMatrix(lngCurRow, COL_假日控制状态) = Nvl(rsData!假日控制状态)
                .TextMatrix(lngCurRow, COL_假日换休) = IIf(Val(Nvl(rsData!是否假日换休)) = 1, "√", "")
                .TextMatrix(lngCurRow, COL_排班方式) = Nvl(rsData!排班方式)
                .TextMatrix(lngCurRow, COL_开始时间) = Format(Nvl(rsData!开始时间), "yyyy-mm-dd hh:mm:ss")
                .TextMatrix(lngCurRow, COL_终止时间) = Format(Nvl(rsData!终止时间), "yyyy-mm-dd hh:mm:ss")
                If bytDataStyle = Data_FixedRule Then
                    .TextMatrix(lngCurRow, COL_登记时间) = Format(Nvl(rsData!登记时间), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngCurRow, COL_临时安排) = Val(Nvl(rsData!临时安排))
                    .TextMatrix(lngCurRow, COL_是否审核) = Val(Nvl(rsData!是否审核))
                End If
                .TextMatrix(lngCurRow, COL_是否临床排班) = IIf(Val(Nvl(rsData!是否临床排班)) = 1, "√", "")
                blnAddRow = True
            End If
                
            If bytDataStyle = Data_Plan And IsDate(strStartDate) And IsDate(strEndDate) Then
                '针对整周跨月的周出诊表无出诊，行安排ID为当前选择出诊表中号源的安排ID
                If IsDate(Nvl(rsData!开始时间)) And IsDate(Nvl(rsData!终止时间)) Then
                    If DateDiff("d", Nvl(rsData!开始时间), strStartDate) <= 0 And DateDiff("d", Nvl(rsData!终止时间), strEndDate) >= 0 Then
                        .TextMatrix(lngCurRow, COL_安排ID) = Nvl(rsData!安排ID)
                    End If
                End If
            End If
            
            If blnFindCol Then
                '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
                '预约控制方式：0-不作预约限制;1-该号码禁止预约;2-仅禁止三方机构平台的预约
                If Nvl(rsData!上班时段) <> "" Then
                    str限约数 = IIf(Nvl(rsData!预约控制方式) = 1, "-", _
                        IIf(Val(Nvl(rsData!限约数)) = 0, IIf(Val(Nvl(rsData!限号数)) = 0, "∞", _
                            Val(Nvl(rsData!限号数))), Val(Nvl(rsData!限约数))))
                    Select Case bytDataStyle
                    Case Data_Templet
                        If Nvl(rsData!排班规则) = 1 Then
                            .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!上班时段)
                            .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!记录ID)
                            .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!限号数)) = 0, "∞", Nvl(rsData!限号数))
                            .TextMatrix(lngCurRow, lngCurCol + 2) = str限约数
                        Else
                            .TextMatrix(lngCurRow, lngCurCol) = _
                                IIf(Nvl(rsData!排班规则) = 4 Or Nvl(rsData!排班规则) = 5, "轮循(" & Val(Nvl(rsData!限制项目)) & "天)", Nvl(rsData!限制项目))
                            .TextMatrix(lngCurRow, lngCurCol + 1) = Nvl(rsData!上班时段)
                            .Cell(flexcpData, lngCurRow, lngCurCol + 1) = Nvl(rsData!记录ID)
                            .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Val(Nvl(rsData!限号数)) = 0, "∞", Nvl(rsData!限号数))
                            .TextMatrix(lngCurRow, lngCurCol + 3) = str限约数
                        End If
                    Case Data_FixedRule
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!上班时段)
                        .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!记录ID)
                        .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!限号数)) = 0, "∞", Nvl(rsData!限号数))
                        .TextMatrix(lngCurRow, lngCurCol + 2) = str限约数
                    Case Data_Plan
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!上班时段)
                        .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!记录ID)
                        
                        '限号列flexcpData标记出诊记录类型，格式"是否临时出诊|是否锁定|是否停诊|是否替诊"
                        strRecordInfo = IIf(Val(Nvl(rsData!是否临时出诊)) = 1, 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Val(Nvl(rsData!是否锁定)) = 1, 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Nvl(rsData!停诊开始时间) <> "", 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Nvl(rsData!替诊医生姓名) <> "", 1, 0)
                        If Nvl(rsData!替诊医生姓名) <> "" Then '替诊号用蓝色字体显示并显示替诊医生
                            .TextMatrix(lngCurRow, lngCurCol) = .TextMatrix(lngCurRow, lngCurCol) & vbCrLf & "(" & Nvl(rsData!替诊医生姓名) & ")"
                        End If
                        .Cell(flexcpData, lngCurRow, lngCurCol + 1) = strRecordInfo
                        '未发布的不显示已挂数和已约数
                        .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(blnPublished, Nvl(rsData!已挂数, "0") & "/", "") & IIf(Nvl(rsData!限号数) = "", "∞", Nvl(rsData!限号数))
                        
                        .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Nvl(rsData!预约控制方式) = 1, "-", IIf(blnPublished, Val(Nvl(rsData!已约数)) & "/", "") & str限约数)
                        .Cell(flexcpData, lngCurRow, lngCurCol + 2) = Val(Nvl(rsData!出诊ID)) & "," & Val(Nvl(rsData!安排ID))
                    End Select
                ElseIf bytDataStyle = Data_Plan Then
                    .Cell(flexcpData, lngCurRow, lngCurCol + 2) = Val(Nvl(rsData!出诊ID)) & "," & Val(Nvl(rsData!安排ID))
                End If
                If blnAddRow Then lngCurRow = lngCurRow + 1
            End If
            rsData.MoveNext
        Loop
        
        Call SetGridFormat(vsfGrid, bytDataStyle, , , blnSortLoad, strStartDate, strEndDate)
        .Redraw = flexRDBuffered
        
        On Error Resume Next
        If .Rows > .FixedRows And .Cols > .FixedCols Then     '缺省定位行
            .Row = -1 '保证在选择行不变的情况下也触发RowColChange事件
            .Row = IIf(lngOldRow < .FixedRows Or lngOldRow > .Rows - 1, IIf(.Rows > .FixedRows, .FixedRows + 1, .FixedRows), lngOldRow)
            .Col = IIf(lngoldCol = 0 Or lngoldCol > .Cols - 1, .FixedCols, lngoldCol)
            .ShowCell .Row, .Col  '立刻显示到指定单元
        End If
    End With
    LoadPlanDataByRecordset = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function RefreshOnePlanData(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    Optional ByVal rsData As ADODB.Recordset, Optional ByVal lngCurRowOld As Long = -1, _
    Optional ByVal blnPublished As Boolean, Optional ByVal bytTableMode As Byte, _
    Optional ByVal strStartDate As String, Optional ByVal strEndDate As String) As Boolean
    '刷新指定行号源数据
    '入参：
    '   vsfGrid VSFlexGrid表格对象
    '   bytDataStyle 安排表类型
    '   rsData 出诊安排记录集,为Nothing表示无出诊记录
    '   bytTableMode 0-固定出诊表,1-月出诊表,2-周出诊表,3-模板
    '   strStartDate,strEndDate 出诊表日期范围，周出诊表时传入
    Dim lngCurRow As Long, lngCurCol As Long
    Dim lngStartRow As Long, lngEndRow As Long
    Dim i As Long, j As Long
    Dim blnFindCol As Boolean, strTemp As String
    Dim str限约数 As String, strRecordInfo As String
    Dim blnFindRow As Boolean, blnRefrashed As Boolean
    Dim blnHaveData As Boolean
    
    Err = 0: On Error GoTo errHandle
    With vsfGrid
'        .Redraw = flexRDNone
        '1.清除该号源的出诊记录
        lngCurRow = IIf(lngCurRowOld = -1, .Row, lngCurRowOld)
        '108641，当前行如果是空行，则不用处理
        If .RowData(lngCurRow) = -1 Then RefreshOnePlanData = True: Exit Function
        If GetPlanGroupRange(vsfGrid, lngCurRow, lngStartRow, lngEndRow) = False Then Exit Function
        For i = lngEndRow To lngStartRow + 1 Step -1
            .RemoveItem i  '移除多余行，只保留一行
        Next
        lngCurRow = lngStartRow: lngEndRow = lngStartRow
        lngCurCol = gPlanGrid_FixedCols
        
        .Cell(flexcpText, lngCurRow, gPlanGrid_FixedCols, lngCurRow, .Cols - 1) = ""
        .Cell(flexcpData, lngCurRow, gPlanGrid_FixedCols, lngCurRow, .Cols - 1) = ""
        .Cell(flexcpForeColor, lngCurRow, gPlanGrid_FixedCols, lngCurRow, .Cols - 1) = .ForeColor
        Set .Cell(flexcpPicture, lngCurRow, gPlanGrid_FixedCols, lngCurRow, .Cols - 1) = Nothing
        .TextMatrix(lngCurRow, COL_安排ID) = ""
        If bytDataStyle = Data_FixedRule Then
            .TextMatrix(lngCurRow, COL_开始时间) = ""
            .TextMatrix(lngCurRow, COL_终止时间) = ""
            .TextMatrix(lngCurRow, COL_登记时间) = ""
            .TextMatrix(lngCurRow, COL_临时安排) = ""
            .TextMatrix(lngCurRow, COL_是否审核) = ""
        End If
        
        '2.重新加载数据
        blnHaveData = True
        If rsData Is Nothing Then blnHaveData = False
        If rsData.RecordCount = 0 Then blnHaveData = False
        If blnHaveData = False Then
            .RemoveItem lngCurRow
            If .RowData(lngCurRow - 1) = -1 And lngCurRow - 1 >= .FixedRows Then
                .RemoveItem lngCurRow - 1
            End If
            RefreshOnePlanData = True
            Exit Function
        End If
        
        Do While Not rsData.EOF
'            lngCurRow = lngStartRow
            blnFindRow = False: blnFindCol = False
            '2.1确定当前列
            Select Case bytDataStyle
            Case Data_Templet  '模板
                If Nvl(rsData!排班规则) <> 1 Then '其它规则
                    '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
                    lngCurCol = .Cols - 4: blnFindCol = True
                Else
                    If Nvl(rsData!限制项目) <> "" Then
                        strTemp = Nvl(rsData!限制项目)
                        For i = lngCurCol To .Cols - 1 Step 3
                            If strTemp = .Cell(flexcpData, 0, i) Then
                                lngCurCol = i: blnFindCol = True
                                Exit For
                            End If
                        Next
                        '没找到再从开始重新找，主要是按限制项目排序跟界面顺序不一致
                        If blnFindCol = False Then
                            For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
                                If strTemp = .Cell(flexcpData, 0, i) Then
                                    lngCurCol = i: blnFindCol = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            Case Data_FixedRule  '固定规则
                If Nvl(rsData!限制项目) <> "" Then
                    strTemp = Nvl(rsData!限制项目)
                    For i = lngCurCol To .Cols - 1 Step 3
                        If strTemp = .Cell(flexcpData, 0, i) Then
                            lngCurCol = i: blnFindCol = True
                            Exit For
                        End If
                    Next
                    '没找到再从开始重新找，主要是按限制项目排序跟界面顺序不一致
                    If blnFindCol = False Then
                        For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
                            If strTemp = .Cell(flexcpData, 0, i) Then
                                lngCurCol = i: blnFindCol = True
                                Exit For
                            End If
                        Next
                    End If
                End If
            Case Else '安排记录
                If Nvl(rsData!出诊日期) = "" Then
                    '出诊日期为空,取安排的开始时间,主要针对整周跨月的周出诊表无出诊记录时,两个安排显示在同一行
                    strTemp = Format(Nvl(rsData!开始时间), "yyyy-mm-dd")
                    lngCurCol = gPlanGrid_FixedCols
                Else
                    strTemp = Format(Nvl(rsData!出诊日期), "yyyy-mm-dd")
                End If
                If IsDate(strTemp) Then
                    For i = lngCurCol To .Cols - 1 Step 3
                        If DateDiff("d", strTemp, .Cell(flexcpData, 0, i)) = 0 Then
                            lngCurCol = i: blnFindCol = True
                            Exit For
                        End If
                    Next
                End If
            End Select
            
            If blnFindCol Then
                '2.2确定当前行
                For i = lngEndRow To lngStartRow Step -1
                    If .TextMatrix(i, lngCurCol) <> "" Then   '是隐藏空行或者无数据行
                        lngCurRow = i + 1: blnFindRow = True
                        Exit For
                    End If
                Next
                If blnFindRow = False Then lngCurRow = lngStartRow
            End If
            
            '2.3加载数据
            If blnRefrashed = False Then
                blnRefrashed = True
                If Not (bytDataStyle = Data_Plan And IsDate(strStartDate) And IsDate(strEndDate)) Then
                    .TextMatrix(lngCurRow, COL_安排ID) = Val(Nvl(rsData!安排ID))
                End If
                If bytDataStyle = Data_FixedRule Then
                    .TextMatrix(lngCurRow, COL_开始时间) = Format(Nvl(rsData!开始时间), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngCurRow, COL_终止时间) = Format(Nvl(rsData!终止时间), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngCurRow, COL_登记时间) = Format(Nvl(rsData!登记时间), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngCurRow, COL_临时安排) = Val(Nvl(rsData!临时安排))
                    .TextMatrix(lngCurRow, COL_是否审核) = Val(Nvl(rsData!是否审核))
                End If
                
                '新增号源时从无效变成有效
                Select Case bytTableMode
                Case 0 '固定出诊表
                    If Val(Nvl(rsData!是否有效)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("FixedItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("InvalidFixedItem")
                    End If
                Case 1 '月出诊表
                    If Val(Nvl(rsData!是否有效)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("MonthItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("InvalidMonthItem")
                    End If
                Case 2 '周出诊表
                    If Val(Nvl(rsData!是否有效)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("WeekItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_图标) = GetPlanItemImage("InvalidWeekItem")
                    End If
                End Select
                .Cell(flexcpPictureAlignment, lngCurRow, COL_图标) = flexAlignCenterCenter
            End If
            If lngEndRow < lngCurRow Then
                '已有行不够，加1行
                lngEndRow = lngEndRow + 1: .AddItem "", lngEndRow
                
                lngCurRow = lngEndRow
                .RowData(lngCurRow) = .RowData(lngCurRow - 1) '用于设置交替色
                
                For j = 0 To gPlanGrid_FixedCols - 1
                    .TextMatrix(lngCurRow, j) = .TextMatrix(lngCurRow - 1, j)
                Next
                Set .Cell(flexcpPicture, lngCurRow, COL_图标) = .Cell(flexcpPicture, lngCurRow - 1, COL_图标)
                .Cell(flexcpPictureAlignment, lngCurRow, COL_图标) = flexAlignCenterCenter
            End If
            
            If bytDataStyle = Data_Plan And IsDate(strStartDate) And IsDate(strEndDate) Then
                '针对整周跨月的周出诊表无出诊，行安排ID为当前选择出诊表中号源的安排ID
                If IsDate(Nvl(rsData!开始时间)) And IsDate(Nvl(rsData!终止时间)) Then
                    If DateDiff("d", Nvl(rsData!开始时间), strStartDate) <= 0 And DateDiff("d", Nvl(rsData!终止时间), strEndDate) >= 0 Then
                        .TextMatrix(lngCurRow, COL_安排ID) = Nvl(rsData!安排ID)
                    End If
                End If
            End If
            
            If blnFindCol Then
                '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
                '预约控制方式：0-不作预约限制;1-该号码禁止预约;2-仅禁止三方机构平台的预约
                If Nvl(rsData!上班时段) <> "" Then
                    str限约数 = IIf(Nvl(rsData!预约控制方式) = 1, "-", _
                        IIf(Val(Nvl(rsData!限约数)) = 0, IIf(Val(Nvl(rsData!限号数)) = 0, "∞", _
                            Val(Nvl(rsData!限号数))), Val(Nvl(rsData!限约数))))
                    Select Case bytDataStyle
                    Case Data_Templet
                        If Nvl(rsData!排班规则) = 1 Then
                            .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!上班时段)
                            .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!记录ID)
                            .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!限号数)) = 0, "∞", Nvl(rsData!限号数))
                            .TextMatrix(lngCurRow, lngCurCol + 2) = str限约数
                        Else
                            .TextMatrix(lngCurRow, lngCurCol) = _
                                IIf(Nvl(rsData!排班规则) = 4 Or Nvl(rsData!排班规则) = 5, "轮循(" & Val(Nvl(rsData!限制项目)) & "天)", Nvl(rsData!限制项目))
                            .TextMatrix(lngCurRow, lngCurCol + 1) = Nvl(rsData!上班时段)
                            .Cell(flexcpData, lngCurRow, lngCurCol + 1) = Nvl(rsData!记录ID)
                            .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Val(Nvl(rsData!限号数)) = 0, "∞", Nvl(rsData!限号数))
                            .TextMatrix(lngCurRow, lngCurCol + 3) = str限约数
                        End If
                    Case Data_FixedRule
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!上班时段)
                        .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!记录ID)
                        .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!限号数)) = 0, "∞", Nvl(rsData!限号数))
                        .TextMatrix(lngCurRow, lngCurCol + 2) = str限约数
                    Case Data_Plan
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!上班时段)
                        .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!记录ID)
                        
                        '限号列flexcpData标记出诊记录类型，格式"是否临时出诊|是否锁定|是否停诊|是否替诊"
                        strRecordInfo = IIf(Val(Nvl(rsData!是否临时出诊)) = 1, 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Val(Nvl(rsData!是否锁定)) = 1, 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Nvl(rsData!停诊开始时间) <> "", 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Nvl(rsData!替诊医生姓名) <> "", 1, 0)
                        If Nvl(rsData!替诊医生姓名) <> "" Then '替诊号用蓝色字体显示并显示替诊医生
                            .TextMatrix(lngCurRow, lngCurCol) = .TextMatrix(lngCurRow, lngCurCol) & vbCrLf & "(" & Nvl(rsData!替诊医生姓名) & ")"
                        End If
                        .Cell(flexcpData, lngCurRow, lngCurCol + 1) = strRecordInfo
                        .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(blnPublished, Nvl(rsData!已挂数, "0") & "/", "") & IIf(Nvl(rsData!限号数) = "", "∞", Nvl(rsData!限号数))
                        
                        .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Nvl(rsData!预约控制方式) = 1, "-", IIf(blnPublished, Val(Nvl(rsData!已约数)) & "/", "") & str限约数)
                        .Cell(flexcpData, lngCurRow, lngCurCol + 2) = Val(Nvl(rsData!出诊ID)) & "," & Val(Nvl(rsData!安排ID))
                    End Select
                ElseIf bytDataStyle = Data_Plan Then
                    .Cell(flexcpData, lngCurRow, lngCurCol + 2) = Val(Nvl(rsData!出诊ID)) & "," & Val(Nvl(rsData!安排ID))
                End If
            End If
            rsData.MoveNext
        Loop
            
        '3.设置格式
        Call SetGridFormat(vsfGrid, bytDataStyle, lngStartRow, lngEndRow, False, strStartDate, strEndDate)
'        .Redraw = flexRDBuffered
    End With
    RefreshOnePlanData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetGridFormat(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    Optional ByVal lngStartRow As Long = -1, Optional ByVal lngEndRow As Long = -1, _
    Optional ByVal blnSortLoad As Boolean, _
    Optional ByVal strStartDate As String, Optional ByVal strEndDate As String)
    '设置单元格格式
    '入参：
    '   lngStartRow 不传时则取.FixedRows
    '   lngEndRow 不传时则取.Rows-1
    '   strStartDate,strEndDate 出诊表日期范围，周出诊表时传入
    Dim i As Long, j As Long
    Dim lngCurRowInGroup As Long, strSpace As String
    Dim intDataType As Integer
    Dim varRecordType As Variant
    Dim lngSortCol As Long, strSort As String
    
    With vsfGrid
        If lngStartRow = -1 Then lngStartRow = .FixedRows
        If lngEndRow = -1 Then lngEndRow = .Rows - 1
        
        '特殊处理，以便能够合并行列
        lngCurRowInGroup = 0 '纵向组内行号
        For i = lngStartRow To lngEndRow
            If .RowData(i) = -1 Then lngCurRowInGroup = 0
            For j = 0 To .Cols - 1
                If .RowData(i) = -1 Then Exit For
                If .TextMatrix(i, j) = "" Then .TextMatrix(i, j) = " " '防止内容为空不能合并
                If .RowData(i - 1) <> -1 And j >= gPlanGrid_FixedCols Then '是否为星期数据
                    If (j - gPlanGrid_FixedCols) Mod 3 = 0 Then '"时间段"列
                        If .TextMatrix(i, j) = " " Then '合并后面的空行
                            .Cell(flexcpAlignment, i - 1, j, i, j) = flexAlignLeftCenter
                            .TextMatrix(i, j) = .TextMatrix(i - 1, j)
                            .TextMatrix(i, j + 1) = .TextMatrix(i - 1, j + 1)
                            .TextMatrix(i, j + 2) = .TextMatrix(i - 1, j + 2)
                        Else
                            strSpace = Space(lngCurRowInGroup Mod 2) '填充空格，防止内容相同合并
                            .TextMatrix(i, j + 1) = strSpace & .TextMatrix(i, j + 1) & strSpace
                            .TextMatrix(i, j + 2) = strSpace & .TextMatrix(i, j + 2) & strSpace
                        End If
                        '模板非星期排班
                        If bytDataStyle = Data_Templet And j = .Cols - 4 Then
                            If .TextMatrix(i, j) = " " Then '合并后面的空行
                                .TextMatrix(i, j + 3) = .TextMatrix(i - 1, j + 3)
                            Else
                                strSpace = Space(lngCurRowInGroup Mod 2) '填充空格，防止内容相同合并
                                .TextMatrix(i, j + 1) = LTrim(.TextMatrix(i, j + 1)) '去除左边的空格
                                .TextMatrix(i, j + 3) = strSpace & .TextMatrix(i, j + 3) & strSpace
                            End If
                            j = j + 1
                        End If
                        j = j + 2
                    End If
                End If
            Next
            If .RowData(i) <> -1 Then lngCurRowInGroup = lngCurRowInGroup + 1
        Next
        
        '行背景色
        If .FixedCols <= .Cols - 1 Then
            For i = lngStartRow To lngEndRow
                .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = .RowData(i)
                .RowHeight(i) = 420
                If .RowData(i) = -1 Then .RowHeight(i) = 0 '设置隐藏行高度
            Next
        End If
        
        If bytDataStyle = Data_Plan Then
            '限号列flexcpData标记出诊记录类型，格式"是否临时出诊|是否锁定|是否停诊|是否替诊"
            For i = lngStartRow To lngEndRow
                For j = gPlanGrid_FixedCols To .Cols - 1 Step 3
                    varRecordType = Split(.Cell(flexcpData, i, j + 1) & "|||", "|")
                    If varRecordType(0) = 1 Then '临时出诊号用蓝色字体显示
                        .Cell(flexcpForeColor, i, j, i, j + 2) = vbBlue
                    End If
                    If varRecordType(1) = 1 Then '锁定号显示锁的图标
                        .Cell(flexcpPicture, i, j) = GetLockImage
                        .Cell(flexcpPictureAlignment, i, j) = flexAlignRightBottom
                    End If
                    If varRecordType(2) = 1 Then '停诊号用红色背景显示
                        .Cell(flexcpBackColor, i, j, i, j + 2) = vbRed
                    End If
                    If varRecordType(3) = 1 Then '替诊号用蓝色字体显示并显示替诊医生
                        .Cell(flexcpForeColor, i, j, i, j + 2) = vbBlue
                    End If
                Next
            Next
        End If
        
        '针对周出诊表，不是当前出诊表的日期用灰色字体显示
        If bytDataStyle = Data_Plan And IsDate(strStartDate) And IsDate(strEndDate) Then
            For j = gPlanGrid_FixedCols To .Cols - 1 Step 3
                If Not (CDate(.Cell(flexcpData, 0, j)) >= CDate(strStartDate) _
                    And CDate(.Cell(flexcpData, 0, j)) <= CDate(strEndDate)) Then
                    .Cell(flexcpForeColor, 0, j, .Rows - 1, j + 2) = &HC0C0C0
                End If
            Next
        End If
        
        '设置排序图标
        If blnSortLoad = False Then
            Call SetSortFlexcpData(vsfGrid) '清空排序标识
            vsfGrid.Cell(flexcpData, 1, COL_号码) = "ASC" '缺省按号码升序排列
        End If
        
        If .Cell(flexcpData, 1, COL_号码) <> "-" Then strSort = .Cell(flexcpData, 1, COL_号码): lngSortCol = COL_号码
        If .Cell(flexcpData, 1, COL_号类) <> "-" Then strSort = .Cell(flexcpData, 1, COL_号类): lngSortCol = COL_号类
        If .Cell(flexcpData, 1, COL_科室) <> "-" Then strSort = .Cell(flexcpData, 1, COL_科室): lngSortCol = COL_科室
        If .Cell(flexcpData, 1, COL_项目) <> "-" Then strSort = .Cell(flexcpData, 1, COL_项目): lngSortCol = COL_项目
        If .Cell(flexcpData, 1, COL_医生) <> "-" Then strSort = .Cell(flexcpData, 1, COL_医生): lngSortCol = COL_医生
        If lngSortCol = 0 Then lngSortCol = COL_号码: strSort = "ASC"
        
        If strSort = "ASC" Then
            Set .Cell(flexcpPicture, 0, lngSortCol, 1, lngSortCol) = GetSortIcon("ASC")
        ElseIf strSort = "DESC" Then
            Set .Cell(flexcpPicture, 0, lngSortCol, 1, lngSortCol) = GetSortIcon("DESC")
        Else
            Set .Cell(flexcpPicture, 0, lngSortCol, 1, lngSortCol) = Nothing
        End If
        .Cell(flexcpPictureAlignment, 0, lngSortCol) = flexAlignCenterBottom
        
        On Error Resume Next
        If .Row < .FixedRows And .Rows > .FixedRows Then .Row = .FixedRows + 1
        If .RowData(.Row) = -1 And .Rows > .Row + 1 Then .Row = .Row + 1
        If .Col < .FixedCols And .Cols > .FixedCols Then .Col = .FixedCols
        Call SetPlanGridRangeColor(vsfGrid, bytDataStyle)
    End With
End Sub

Public Sub ShowHolidayToPlan(vsfGrid As VSFlexGrid, ByVal dtStartDate As Date, ByVal dtEndStart As Date)
    '显示法定节假日
    Dim strSQL As String, rsHoliday As ADODB.Recordset
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandler
    '法定节假日
    strSQL = "Select 开始日期, 终止日期, 节日名称" & vbNewLine & _
            " From 法定假日表" & vbNewLine & _
            " Where 性质 = 0 And 年份 = To_Number(To_Char([1], 'yyyy'))" & vbNewLine & _
            "       And Not(开始日期 > [2] Or 终止日期 < [1])"
    Set rsHoliday = zlDatabase.OpenSQLRecord(strSQL, "获取节假日数据", dtStartDate, dtEndStart)
    If rsHoliday.RecordCount = 0 Then Exit Sub
    
    With vsfGrid
        For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
            If IsDate(.Cell(flexcpData, 0, i)) Then
                rsHoliday.MoveFirst
                Do While Not rsHoliday.EOF
                    If CDate(.Cell(flexcpData, 0, i)) >= CDate(Nvl(rsHoliday!开始日期)) _
                        And CDate(.Cell(flexcpData, 0, i)) <= CDate(Nvl(rsHoliday!终止日期)) Then
                        .TextMatrix(0, i) = .TextMatrix(0, i) & "(" & Nvl(rsHoliday!节日名称) & ")"
                        .TextMatrix(0, i + 1) = .TextMatrix(0, i + 1) & "(" & Nvl(rsHoliday!节日名称) & ")"
                        .TextMatrix(0, i + 2) = .TextMatrix(0, i + 2) & "(" & Nvl(rsHoliday!节日名称) & ")"
                        Exit Do
                    End If
                    rsHoliday.MoveNext
                Loop
            End If
        Next
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ShowStopVisitPlan(vsfGrid As VSFlexGrid, ByVal dtStartDate As Date, ByVal dtEndStart As Date, _
    Optional ByVal lng号源Id As Long)
    '显示停诊安排
    Dim strSQL As String, rsStopVisitPlan As ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo errHandler
    '停诊安排
    strSQL = "Select b.Id As 号源id, Trunc(a.开始时间) As 开始时间, a.终止时间, a.停诊原因" & vbNewLine & _
            " From 临床出诊停诊记录 A, 临床出诊号源 B" & vbNewLine & _
            " Where a.申请人 = b.医生姓名 And b.医生id Is Not Null" & vbNewLine & _
            "       And a.记录id Is Null And a.审批人 Is Not Null And a.取消人 Is Null" & vbNewLine & _
            "       And Not (a.开始时间 > [2] Or a.终止时间 < [1])" & _
                    IIf(lng号源Id = 0, "", " And b.ID = [3]")

    Set rsStopVisitPlan = zlDatabase.OpenSQLRecord(strSQL, "获取停诊安排", dtStartDate, dtEndStart, lng号源Id)
    If rsStopVisitPlan.RecordCount = 0 Then Exit Sub
    
    rsStopVisitPlan.MoveFirst
    With vsfGrid
        For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
            blnFind = False
            If IsDate(.Cell(flexcpData, 0, i)) Then
                For j = .FixedRows To .Rows - 1
                    If blnFind And lng号源Id <> 0 And lng号源Id <> Val(.TextMatrix(j, COL_号源ID)) Then
                        Exit For
                    End If
                    rsStopVisitPlan.MoveFirst
                    Do While Not rsStopVisitPlan.EOF
                        If CDate(.Cell(flexcpData, 0, i)) >= CDate(Nvl(rsStopVisitPlan!开始时间)) _
                            And CDate(.Cell(flexcpData, 0, i)) <= CDate(Nvl(rsStopVisitPlan!终止时间)) _
                            And Val(.TextMatrix(j, COL_号源ID)) = Val(Nvl(rsStopVisitPlan!号源ID)) Then
                            blnFind = True
                            If .Cell(flexcpBackColor, j, i) <> vbRed Then
                                .Cell(flexcpForeColor, j, i, j, i + 2) = vbRed
                            End If
                            .Cell(flexcpText, j, i) = Trim(.TextMatrix(j, i)) & IIf(Trim(.TextMatrix(j, i)) = "", "", vbCrLf) & "(" & Nvl(rsStopVisitPlan!停诊原因) & ")"
                            Exit Do
                        End If
                        rsStopVisitPlan.MoveNext
                    Loop
                Next
            End If
        Next
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub SetPlanGridSelRange(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle)
    '功能：设置选择行列范围
    Dim lngRowStart As Long, lngRowEnd As Long '起始行和终止行
    Dim lngColStart As Long, lngColEnd As Long '起始列和终止列
    
    On Error Resume Next
    With vsfGrid
        If Not .Visible Then Exit Sub
        If .Row < .FixedRows Or .RowSel < .FixedRows Then Exit Sub
        If .Col < gPlanGrid_FixedCols And .ColSel < gPlanGrid_FixedCols Then Exit Sub
            
        '选择行范围
        lngRowStart = .Row: lngRowEnd = .RowSel
        
        '选择列范围
        If .Col >= gPlanGrid_FixedCols And .ColSel < gPlanGrid_FixedCols Then
            '起始列为日期列，末尾列为非日期列，则只选择非日期列
            lngColStart = .ColSel
            lngColEnd = lngColStart
        End If
        
        If .Col < gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
            '起始列为非日期列，末尾列为日期列，则选择日期列安排范围
            lngColStart = GetPlanItemNameCol(.ColSel) '确定"时间段"列
            lngColEnd = lngColStart + 2
        End If
        
        If .Col >= gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
            lngColStart = GetPlanItemNameCol(.Col) '确定"时间段"列
            lngColEnd = GetPlanItemNameCol(.ColSel)
            If lngColStart > lngColEnd Then
                lngColStart = lngColStart + 2
            Else
                lngColEnd = lngColEnd + 2
            End If
        End If
        
        '模板最后一列特殊处理
        If bytDataStyle = Data_Templet And lngColStart = .Cols - 4 Then
            lngColEnd = lngColEnd + 1: lngColStart = lngColStart + 1
        End If
        
        '重新选择
        .ForeColorSel = .Cell(flexcpForeColor, .RowSel, .ColSel)
        .Select lngRowStart, lngColStart, lngRowEnd, lngColEnd
    End With
End Sub

Public Sub SetPlanGridRangeColor(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    Optional ByVal strOldSelRange As String)
    '功能：设置选择行颜色,.RowData中存了颜色值
    'strOldSelRange:上一次选择网格区域，格式"开始行|结束行|开始列|结束列"
    Dim lngRowStart As Long, lngRowEnd As Long '起始行和终止行
    Dim lngColStart As Long, lngColEnd As Long '起始列和终止列
    Dim varRecordType As Variant
    Dim i As Long, j As Long, lngTemp As Long
    Dim varTemp As Variant
    
    On Error Resume Next
    With vsfGrid
        If Not .Visible Then Exit Sub
        '恢复颜色
        If strOldSelRange <> "" Then
            varTemp = Split(strOldSelRange & "|||", "|")
            lngRowStart = varTemp(0): lngRowEnd = varTemp(1)
            lngColStart = varTemp(2): lngColEnd = varTemp(3)
            
            If lngRowStart < .FixedRows Or lngRowEnd < .FixedRows Then
            ElseIf lngColStart < gPlanGrid_FixedCols And lngColEnd < gPlanGrid_FixedCols Then
                .Cell(flexcpBackColor, lngRowStart, lngColStart) = .RowData(lngRowStart)
            Else
                If lngRowStart > lngRowEnd Then lngTemp = lngRowStart: lngRowStart = lngRowEnd: lngRowEnd = lngTemp
                If lngColStart > lngColEnd Then lngTemp = lngColStart: lngColStart = lngColEnd: lngColEnd = lngTemp
                If lngRowStart < .FixedRows Then lngRowStart = .FixedRows
                If lngColStart < .FixedCols Then lngColStart = .FixedCols
                For i = lngRowStart To lngRowEnd
                    .Cell(flexcpBackColor, i, lngColStart, i, lngColEnd) = .RowData(i)
                    For j = lngColStart To lngColEnd
                        '限号列flexcpData标记出诊记录类型，格式"是否临时出诊|是否锁定|是否停诊|是否替诊"
                        varRecordType = Split(.Cell(flexcpData, i, j + 1) & "|||", "|")
                        If Val(varRecordType(2)) = 1 Then   '停诊
                            .Cell(flexcpBackColor, i, j, i, j + 2) = vbRed
                            '停诊时改变了字体颜色的，在这里进行恢复
                            If Val(varRecordType(0)) = 1 Or Val(varRecordType(3)) = 1 Then '临时出诊/替诊
                                .Cell(flexcpForeColor, i, j, i, j + 2) = vbBlue
                            Else
                                .Cell(flexcpForeColor, i, j, i, j + 2) = vbBlack
                            End If
                        End If
                        j = j + 2
                    Next
                Next
            End If
        End If
        
        If .Row < .FixedRows Or .RowSel < .FixedRows Then Exit Sub
        If .Col < gPlanGrid_FixedCols And .ColSel < gPlanGrid_FixedCols Then
            lngRowStart = .Row: lngColStart = .Col
            If lngRowStart >= .FixedRows And lngColStart >= .FixedCols Then
                .Cell(flexcpBackColor, lngRowStart, lngColStart) = .BackColorSel
            End If
        Else
            '选择行范围
            lngRowStart = .Row: lngRowEnd = .RowSel
            
            '选择列范围
            If .Col >= gPlanGrid_FixedCols And .ColSel < gPlanGrid_FixedCols Then
                '起始列为日期列，末尾列为非日期列，则只选择非日期列
                lngColStart = .ColSel
                lngColEnd = lngColStart
                lngRowEnd = lngRowStart
            End If
            
            If .Col < gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
                '起始列为非日期列，末尾列为日期列，则选择日期列安排范围
                lngColStart = GetPlanItemNameCol(.ColSel) '确定"时间段"列
                lngColEnd = lngColStart + 2
            End If
            
            If .Col >= gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
                lngColStart = GetPlanItemNameCol(.Col) '确定"时间段"列
                lngColEnd = GetPlanItemNameCol(.ColSel)
                If lngColStart > lngColEnd Then
                    lngColStart = lngColStart + 2
                Else
                    lngColEnd = lngColEnd + 2
                End If
            End If
            
            '模板最后一列特殊处理
            If bytDataStyle = Data_Templet And lngColStart = .Cols - 4 Then
                lngColEnd = lngColEnd + 1: lngColStart = lngColStart + 1
            End If
            
            '重新选择
            If lngRowStart <> .Row Or lngColStart <> .Col _
                Or lngRowEnd <> .RowSel Or lngColEnd <> .ColSel Then
                .Select lngRowStart, lngColStart, lngRowEnd, lngColEnd
            End If
            
            If lngRowStart > lngRowEnd Then lngTemp = lngRowStart: lngRowStart = lngRowEnd: lngRowEnd = lngTemp
            If lngColStart > lngColEnd Then lngTemp = lngColStart: lngColStart = lngColEnd: lngColEnd = lngTemp
            If lngRowStart < .FixedRows Then lngRowStart = .FixedRows
            If lngColStart < .FixedCols Then lngColStart = .FixedCols
            For i = lngRowStart To lngRowEnd
                .Cell(flexcpBackColor, i, lngColStart, i, lngColEnd) = .BackColorSel
                For j = lngColStart To lngColEnd
                    '限号列flexcpData标记出诊记录类型，格式"是否临时出诊|是否锁定|是否停诊|是否替诊"
                    varRecordType = Split(.Cell(flexcpData, i, j + 1) & "|||", "|")
                    If Val(varRecordType(2)) = 1 Then   '停诊
                        .Cell(flexcpForeColor, i, j, i, j + 2) = vbRed
                    End If
                    j = j + 2
                Next
            Next
        End If
    End With
End Sub

Public Function GetPlanItemNameCol(ByVal lngCurCol As Long) As Long
    On Error GoTo errHandle
    '确定"时段"列的列索引
    GetPlanItemNameCol = lngCurCol - Choose(((lngCurCol - gPlanGrid_FixedCols) Mod 3) + 1, 0, 1, 2)
    Exit Function
errHandle:
    Err.Clear
    GetPlanItemNameCol = 0
End Function

Public Function GetPlanGroupRange(vsfGrid As VSFlexGrid, _
    ByVal lngCurRow As Long, ByRef lngRowStart As Long, ByRef lngRowEnd As Long) As Boolean
    '当前组的行索引范围
    Dim i As Integer
    
    With vsfGrid
        lngRowStart = .FixedRows
        For i = lngCurRow To .FixedRows Step -1
            If .RowData(i) = -1 Then lngRowStart = i + 1: Exit For
        Next
        lngRowEnd = .Rows - 1
        For i = lngCurRow + 1 To .Rows - 1
            If .RowData(i) = -1 And i <> .Rows - 1 Then lngRowEnd = i - 1: Exit For
        Next
    End With
    GetPlanGroupRange = True
End Function

Public Function GetPlanSortCircleStr(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    lngRow As Long, lngCol As Long) As String
    '获取排序串
    Dim i As Long, strSort As String
    
    On Error GoTo errHandle
    '按列排序
    Select Case lngCol
    Case COL_号类
        strSort = SortCircle(vsfGrid, lngCol, "号类") & "排序号码,"
    Case COL_号码
        strSort = SortCircle(vsfGrid, lngCol, "排序号码")
    Case COL_科室
        strSort = SortCircle(vsfGrid, lngCol, "科室") & "排序号码,"
    Case COL_项目
        strSort = SortCircle(vsfGrid, lngCol, "收费项目") & "排序号码,"
    Case COL_医生
        strSort = SortCircle(vsfGrid, lngCol, "医生姓名") & "排序号码,"
    End Select

    If strSort <> "" Then
        Call SetSortFlexcpData(vsfGrid, lngCol) '清空排序标识
        
        Select Case bytDataStyle
        Case Data_FixedRule
            strSort = strSort & "开始时间,终止时间,限制项目,上班时段"
        Case Data_Templet
            strSort = strSort & "限制项目,上班时段"
        Case Data_Plan
            strSort = strSort & "出诊日期,上班时段"
        End Select
    End If
    GetPlanSortCircleStr = strSort
    Exit Function
errHandle:
    Err.Clear
End Function

Private Sub SetSortFlexcpData(vsfGrid As VSFlexGrid, Optional ByVal lngSortCol As Long)
    '清空排序标识
    On Error GoTo errHandle
    If lngSortCol <> COL_号码 Then vsfGrid.Cell(flexcpData, 1, COL_号码) = "-"
    If vsfGrid.Cell(flexcpData, 1, COL_号码) = "-" Then Set vsfGrid.Cell(flexcpPicture, 0, COL_号码, 1, COL_号码) = Nothing
    If lngSortCol <> COL_号类 Then vsfGrid.Cell(flexcpData, 1, COL_号类) = "-"
    If vsfGrid.Cell(flexcpData, 1, COL_号类) = "-" Then Set vsfGrid.Cell(flexcpPicture, 0, COL_号类, 1, COL_号类) = Nothing
    If lngSortCol <> COL_科室 Then vsfGrid.Cell(flexcpData, 1, COL_科室) = "-"
    If vsfGrid.Cell(flexcpData, 1, COL_科室) = "-" Then Set vsfGrid.Cell(flexcpPicture, 0, COL_科室, 1, COL_科室) = Nothing
    If lngSortCol <> COL_项目 Then vsfGrid.Cell(flexcpData, 1, COL_项目) = "-"
    If vsfGrid.Cell(flexcpData, 1, COL_项目) = "-" Then Set vsfGrid.Cell(flexcpPicture, 0, COL_项目, 1, COL_项目) = Nothing
    If lngSortCol <> COL_医生 Then vsfGrid.Cell(flexcpData, 1, COL_医生) = "-"
    If vsfGrid.Cell(flexcpData, 1, COL_医生) = "-" Then Set vsfGrid.Cell(flexcpPicture, 0, COL_医生, 1, COL_医生) = Nothing
    Exit Sub
errHandle:
    Err.Clear
End Sub

Private Function SortCircle(vsfGrid As VSFlexGrid, ByVal lngCol As Long, ByVal strColName As String) As String
    'Cell(flexcpData, 1, lngCol)记录了当前排序方式，注意在重新加载数据时清除
    Select Case vsfGrid.Cell(flexcpData, 1, lngCol)
    Case ""
        If lngCol = COL_号码 Then '号码列初始时就是以升序排列的
            vsfGrid.Cell(flexcpData, 1, lngCol) = "DESC"
            SortCircle = strColName & " DESC,"
        Else
            vsfGrid.Cell(flexcpData, 1, lngCol) = "ASC"
            SortCircle = strColName & " Asc,"
        End If
    Case "ASC" '升序
        vsfGrid.Cell(flexcpData, 1, lngCol) = "DESC"
        SortCircle = strColName & " Desc,"
    Case "DESC" '降序
        If lngCol = COL_号码 Then '号码列要么升序要么降序
            vsfGrid.Cell(flexcpData, 1, lngCol) = "ASC"
            SortCircle = strColName & " Asc,"
        Else
            vsfGrid.Cell(flexcpData, 1, lngCol) = "-"
            SortCircle = ""
        End If
    Case "-" '不排序
        vsfGrid.Cell(flexcpData, 1, lngCol) = "ASC"
        SortCircle = strColName & " Asc,"
    End Select
End Function

'限号列flexcpData标记出诊记录类型，格式"是否临时出诊|是否锁定|是否停诊|是否替诊"
Public Function PlanIsLocked(vsfGrid As VSFlexGrid, _
    Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1) As Boolean
    '是否锁号状态
    Dim lngCurRow As Long, lngCurCol As Long
    Dim varRecordType As Variant
    
    With vsfGrid
        lngCurRow = IIf(lngRow = -1, .Row, lngRow)
        lngCurCol = IIf(lngCol = -1, GetPlanItemNameCol(.Col), GetPlanItemNameCol(lngCol))
                    
        varRecordType = Split(.Cell(flexcpData, lngCurRow, lngCurCol + 1) & "|||", "|")
        If Val(varRecordType(1)) = 1 Then
            PlanIsLocked = True
        End If
    End With
End Function

Public Function PlanIsStopVisit(vsfGrid As VSFlexGrid) As Boolean
    '是否停诊状态
    Dim lngCurRow As Long, lngCurCol As Long
    Dim varRecordType As Variant
    
    lngCurRow = vsfGrid.Row
    lngCurCol = GetPlanItemNameCol(vsfGrid.Col)
    
    varRecordType = Split(vsfGrid.Cell(flexcpData, lngCurRow, lngCurCol + 1) & "|||", "|")
    If Val(varRecordType(2)) = 1 Then
        PlanIsStopVisit = True
    End If
End Function

Public Function PlanIsReplaceDoctor(vsfGrid As VSFlexGrid) As Boolean
    '是否替诊诊状态
    Dim lngCurRow As Long, lngCurCol As Long
    Dim varRecordType As Variant
    
    lngCurRow = vsfGrid.Row
    lngCurCol = GetPlanItemNameCol(vsfGrid.Col)
    
    varRecordType = Split(vsfGrid.Cell(flexcpData, lngCurRow, lngCurCol + 1) & "|||", "|")
    If Val(varRecordType(3)) = 1 Then
        PlanIsReplaceDoctor = True
    End If
End Function

Public Function PlanIsSelOne(vsfGrid As VSFlexGrid) As Boolean
    '是否只选择了一个时段
    Dim lngRowStart As Long, lngRowEnd As Long '起始行和终止行
    Dim lngColStart As Long, lngColEnd As Long '起始列和终止列
    
    With vsfGrid
        '选择行范围
        lngRowStart = .Row: lngRowEnd = .RowSel
        
        '选择列范围
        If .Col >= gPlanGrid_FixedCols And .ColSel < gPlanGrid_FixedCols Then
            '起始列为日期列，末尾列为非日期列，则只选择非日期列
            lngColStart = .ColSel
            lngColEnd = lngColStart
        End If
        
        If .Col < gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
            '起始列为非日期列，末尾列为日期列，则选择日期列安排范围
            lngColStart = GetPlanItemNameCol(.ColSel) '确定"时间段"列
            lngColEnd = lngColStart + 2
        End If
        
        If .Col >= gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
            lngColStart = GetPlanItemNameCol(.Col) '确定"时间段"列
            lngColEnd = GetPlanItemNameCol(.ColSel)
            If lngColStart > lngColEnd Then
                lngColStart = lngColStart + 2
            Else
                lngColEnd = lngColEnd + 2
            End If
        End If
    End With
    PlanIsSelOne = Not (Abs(lngRowEnd - lngRowStart) > 0 Or Abs(lngColEnd - lngColStart) > 2)
End Function

Public Function SelectedIsNotNull(ByVal vsfGrid As VSFlexGrid) As Boolean
    '判断当前选择单元格是否不是空值
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Function
    End With
    SelectedIsNotNull = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Is禁止预约(ByVal vsfGrid As VSFlexGrid) As Boolean
    '判断当前选择安排是否禁止预约
    On Error GoTo errHandler
    Is禁止预约 = True
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        If Trim(.TextMatrix(.Row, GetPlanItemNameCol(.Col) + 2)) = "-" Then Exit Function
    End With
    Is禁止预约 = False
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsVerified(ByVal vsfGrid As VSFlexGrid) As Boolean
    '判断当前选择安排是否已审核
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        If Val(.TextMatrix(.Row, COL_是否审核)) = 0 Then Exit Function
    End With
    IsVerified = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsTempPlan(ByVal vsfGrid As VSFlexGrid) As Boolean
    '判断当前选择安排是否临时安排
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        If Val(.TextMatrix(.Row, COL_临时安排)) = 0 Then Exit Function
    End With
    IsTempPlan = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPlanItemImage(ByVal strKey As String) As IPictureDisp
    '获取安排号源类型图像
    '入参：
    '   strKey 图像索引:
    '       InvalidFixedItem '无效按固定排班号源
    '       FixedItem '正常按固定排班号源
    '       InvalidMonthItem '无效按月排班号源
    '       MonthItem '正常按月排班号源
    '       InvalidWeekItem '无效按周排班号源
    '       WeekItem '正常按周排班号源
    Set GetPlanItemImage = frmClinicPlanTemp.GetPlanItemImage(strKey)
End Function

Public Function GetSortIcon(ByVal strKey As String) As IPictureDisp
    '获取排序图标
    '入参：
    '   strKey 图像索引
    '       ASC '升序
    '       DESC '降序
    Set GetSortIcon = frmClinicPlanTemp.GetSortIcon(strKey)
End Function

Private Function GetLockImage() As IPictureDisp
    '获取锁号图像
    Set GetLockImage = frmClinicPlanTemp.GetLockPicture
End Function

Public Sub RegistPlan_KeyDown(vsfGrid As VSFlexGrid, KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    On Error Resume Next
    With vsfGrid
        Select Case KeyCode
        Case vbKeyRight '向右移动键
            If Shift = vbShiftMask Then
                If .ColSel + 3 >= gPlanGrid_FixedCols Then
                    .ColSel = .ColSel + 3
                    KeyCode = 0 '屏蔽键值
                Else
                    
                End If
            Else
                If .Col < gPlanGrid_FixedCols Then
                    i = 1
                    Do While .Col + i < .Cols - 1
                        If .ColHidden(.Col + i) Or .ColWidth(.Col + i) = 0 Then
                            i = i + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    .Col = .Col + i
                    KeyCode = 0 '屏蔽键值
                Else
                    .Col = .Col + 3
                End If
            End If
        Case vbKeyLeft '向左移动键
            If Shift = vbShiftMask Then
                If .ColSel - 3 >= gPlanGrid_FixedCols Then
                    .ColSel = .ColSel - 3
                    KeyCode = 0 '屏蔽键值
                Else
                    
                End If
            End If
        End Select
    End With
End Sub

Public Function GetSelectRange(vsfGrid As VSFlexGrid, ByVal strSelRange As String, _
    ByRef lngRowStart As Long, ByRef lngRowEnd As Long, _
    ByRef lngColStart As Long, ByRef lngColEnd As Long) As Boolean
    '分解出选择网格区域，
    '入参：
    '   vsfGrid:网格控件
    '   strSelRange:格式"开始行|结束行|开始列|结束列"
    '出参：
    '   lngRowStart 开始行
    '   lngRowEnd 结束行
    '   lngColStart 开始列
    '   lngColEnd 结束列
    Dim varTemp As Variant, lngTemp As Long
    
    Err = 0: On Error GoTo errHandler
    If InStr(strSelRange, "|") <= 0 Then Exit Function
    
    varTemp = Split(strSelRange & "|||", "|")
    lngRowStart = varTemp(0): lngRowEnd = varTemp(1)
    lngColStart = varTemp(2): lngColEnd = varTemp(3)
    With vsfGrid
        If lngRowStart < .FixedRows Or lngRowStart > .Rows - 1 Then Exit Function
        If lngRowEnd < .FixedRows Or lngRowEnd > .Rows - 1 Then Exit Function
        If lngColStart < .FixedRows Or lngColStart > .Cols - 1 Then Exit Function
        If lngColEnd < .FixedRows Or lngColEnd > .Cols - 1 Then Exit Function
    End With
    
    If lngRowStart > lngRowEnd Then lngTemp = lngRowStart: lngRowStart = lngRowEnd: lngRowEnd = lngTemp
    If lngColStart > lngColEnd Then lngTemp = lngColStart: lngColStart = lngColEnd: lngColEnd = lngTemp
    GetSelectRange = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function vsGrid_Para_Restore_Plan(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional bln强制恢复保存 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:从数据库中恢复网格的宽度等信息
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     strKey-主建
    '     blnSaveToDataBase-是否是往数据库中保存参数(如果是往数据库中保存,则强制保存为true,否则根据是否使用个性化风格来确定)
    '     bln强制恢复保存-决定是否将保存注册表的参数值,进行强制恢复
    '返回:恢复成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '说明：
    '       单独写该过程是因为临床出诊安排的表格比较特殊，只恢复某些列，同时列数也是动态变化的
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, arrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    If blnSaveToDataBase = False Then
        '只有在本地注册表中才会处理个性化设置
        vsGrid_Para_Restore_Plan = True
        If bln强制恢复保存 = False Then
            If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g私有模块, strCaption, strKey, strParaValue)
    Else
        strParaValue = zlDatabase.GetPara(strKey, glngSys, lngModule)
    End If
    
    vsGrid_Para_Restore_Plan = False
    If strParaValue = "" Then Exit Function
    'strParaValue:保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    Err = 0: On Error GoTo Errhand:
    arrReg = Split(strParaValue, "|")
'    If vsGrid.Cols <> UBound(arrReg) + 1 Then Exit Function
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrTemp(0)
            If strColName <> "" Then
                intTemp = .ColIndex(strColName)
                If intTemp <> -1 Then
                    .ColWidth(intTemp) = Val(arrTemp(1))
                    If Val(arrTemp(2)) = 1 Then
                        .ColHidden(intTemp) = True
                    Else
                        .ColHidden(intTemp) = False
                    End If
                    If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                    .ColPosition(.ColIndex(strColName)) = intCol
                End If
            End If
        Next
    End With
    vsGrid_Para_Restore_Plan = True
    Exit Function
Errhand:
End Function

Public Function VSFlexGridCopyTo(ByVal vsfSource As VSFlexGrid, ByRef vsfNew As VSFlexGrid, _
    Optional ByVal bytMode As Byte) As Boolean
    '功能: 将vsfSource的数据复制到vsfNew中，包括显示格式，用于打印\预览
    '参数:
    '     vsfNew-复制后的对象
    '     vsfSource-被复制的对象
    '     bytMode=1 打印;2 预览;3 输出到EXCEL
    '返回：复制成功，返回True；否则，返回False
    VSFlexGridCopyTo = frmClinicPlanTemp.VSFlexGridCopyTo(vsfSource, vsfNew, bytMode)
End Function
