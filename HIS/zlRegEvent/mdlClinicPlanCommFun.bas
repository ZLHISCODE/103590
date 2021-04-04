Attribute VB_Name = "mdlClinicPlanCommFun"
Option Explicit
Public Const G_AlternateColor As Long = 16772055   '行交替色
Public Const G_LostFocusColor As Long = &HE0E0E0   '失去焦点时的网格背景色

Public Enum G_Enum_Fun
    Fun_View = 0 '查看
    Fun_Add = 1 '新增
    Fun_Update = 2 '编辑
    Fun_Delete = 3 '删除
    Fun_TempPlan = 4 '临时安排(固定出诊表)
    Fun_UpdateUnit = 5 '调整合作单位预约挂号
    Fun_AddSignalSourcePlan = 6 '新增号源
    Fun_TempPlanRecord = 7 '发布后临时出诊
    Fun_TempPlanVerify = 8 '临时安排(固定出诊表)审核
    Fun_TempPlanCancel = 9 '临时安排(固定出诊表)取消审核
    Fun_UpdatePlan = 10 '调整已发布后的安排
End Enum

Public Enum gRegistPlanEditMode
    ED_RegistPlan_Edit = 0
    ED_RegistPlan_View = 1
    ED_RegistPlan_UpdateUnit = 2
    ED_RegistPlan_NumLimitModify = 3
End Enum

Public Enum PictureTextAlignmentSettings
    '图片内文本显示位置
    pictxtAlignLeftTop = 0
    pictxtAlignLeftCenter = 1
    pictxtAlignLeftBottom = 2
    pictxtAlignCenterTop = 3
    pictxtAlignCenterCenter = 4
    pictxtAlignCenterBottom = 5
    pictxtAlignRightTop = 6
    pictxtAlignRightCenter = 7
    pictxtAlignRightBottom = 8
End Enum

Public Function GetPlanKey(ByVal strItem As String) As String
    '根据项目数据类型获取Key值
    GetPlanKey = "K" & IIf(IsDate(strItem), Format(strItem, "yyyy-mm-dd"), strItem)
End Function

Public Function HavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String, Optional ByVal blnAnd As Boolean) As Boolean
    '判断权限
    '   strMyPriv 多个用分号";"分隔
    '   blnAnd True-多个权限是And关系，False-多个权限是Or的关系
    Dim varPrivs As Variant, i As Integer
    Dim blnHave As Boolean
    
    If InStr(strMyPriv, ";") > 0 Then
        varPrivs = Split(strMyPriv, ";")
        blnHave = IIf(blnAnd, True, False)
        For i = 0 To UBound(varPrivs)
            If blnAnd Then
                If zlStr.IsHavePrivs(strPrivs, varPrivs(i)) = False Then
                    blnHave = False: Exit For
                End If
            Else
                If zlStr.IsHavePrivs(strPrivs, varPrivs(i)) Then
                    blnHave = True: Exit For
                End If
            End If
        Next
    Else
        blnHave = zlStr.IsHavePrivs(strPrivs, strMyPriv)
    End If
    HavePrivs = blnHave
End Function

Public Function ZDate(ByVal varValue As Variant, Optional ByVal varDefault As Variant = "", _
    Optional ByVal blnDataBase As Boolean = True) As String
'功能：将缺省时间转换为"NULL"串,在生成SQL语句时用
    Dim strTemp As String
    
    If blnDataBase Then
        If IsDate(varValue) Then
            If DateDiff("s", varValue, "1899-12-30") = 0 Then
                'varValue为空
                If IsDate(varDefault) Then
                    If DateDiff("s", varDefault, "1899-12-30") = 0 Then
                        'varDefault为空
                        ZDate = "NULL"
                    Else
                        ZDate = "To_Date('" & Format(varDefault, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                    End If
                Else
                    ZDate = IIf(CStr(varDefault) = "", "NULL", CStr(varDefault))
                End If
            Else
                ZDate = "To_Date('" & Format(varValue, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
            End If
        Else
            If CStr(varValue) = "" Then
                'varValue为空
                If IsDate(varDefault) Then
                    If DateDiff("s", varDefault, "1899-12-30") = 0 Then
                        'varDefault为空
                        ZDate = "NULL"
                    Else
                        ZDate = "To_Date('" & Format(varDefault, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
                    End If
                Else
                    ZDate = IIf(CStr(varDefault) = "", "NULL", CStr(varDefault))
                End If
            Else
                ZDate = CStr(varValue)
            End If
        End If
    Else
        If IsDate(varValue) Then
            If DateDiff("s", varValue, "1899-12-30") = 0 Then
                'varValue为空
                If IsDate(varDefault) Then
                    If DateDiff("s", varDefault, "1899-12-30") = 0 Then
                        'varDefault为空
                        ZDate = ""
                    Else
                        ZDate = Format(varDefault, "yyyy-mm-dd hh:mm:ss")
                    End If
                Else
                    ZDate = CStr(varDefault)
                End If
            Else
                ZDate = Format(varValue, "yyyy-mm-dd hh:mm:ss")
            End If
        Else
            If CStr(varValue) = "" Then
                'varValue为空
                If IsDate(varDefault) Then
                    If DateDiff("s", varDefault, "1899-12-30") = 0 Then
                        'varDefault为空
                        ZDate = ""
                    Else
                        ZDate = Format(varDefault, "yyyy-mm-dd hh:mm:ss")
                    End If
                Else
                    ZDate = CStr(varDefault)
                End If
            Else
                ZDate = CStr(varValue)
            End If
        End If
    End If
End Function


Public Function GetWorkTrueDate(ByVal dtStart As Date, ByVal dtCur As Date, _
    Optional ByVal blnNextDate As Boolean = True, Optional ByVal blnEqualNextDay As Boolean = True) As String
    '根据开始时间，确定指定时间的日期，即当前时间是否加一天或减一天（转换为同一天比较时分）
    '入参：
    '   dtCur - 转换日期
    '   dtStart - 开始时间
    '   dtCur - 当前时间
    '   blnNextDate - True表示加一天，False表示减一天
    '   blnEqualNextDay - 时间相等时是否为下一天
    If blnNextDate Then
        If DateDiff("n", Format(dtStart, "hh:mm:ss"), Format(dtCur, "hh:mm:ss")) < 0 Then
            dtStart = DateAdd("d", 1, dtStart)
        End If
        
        If blnEqualNextDay Then
            If DateDiff("n", Format(dtStart, "hh:mm:ss"), Format(dtCur, "hh:mm:ss")) = 0 Then
                dtStart = DateAdd("d", 1, dtStart)
            End If
        End If
    Else
        If DateDiff("n", Format(dtStart, "hh:mm:ss"), Format(dtCur, "hh:mm:ss")) > 0 Then
            dtStart = DateAdd("d", -1, dtStart)
        End If
    End If
    
    GetWorkTrueDate = Format(dtStart, "yyyy-mm-dd ") & Format(dtCur, "hh:mm:ss")
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlNumStrSort(ByVal strNums As String, Optional ByVal blnDelRepeat As Boolean, _
    Optional blnAsc As Boolean = True, Optional ByVal strSplit As String = ",") As String
    '将数字字符串排序输出
    '入参：
    '   blnDelRepeat - 去除重复的
    '   blnAsc - 是否按升序排序
    '   strSplit - 分隔符
    '说明：插入法排序,默认升序
    Dim varData As Variant, arrNum() As Long
    Dim varTemp As Variant, strReturn As String
    Dim i As Long, j As Long, k As Long
    Dim blnFind As Boolean
    
    If InStr(strNums, strSplit) = 0 Then ZlNumStrSort = Val(strNums): Exit Function
    varData = Split(strNums, strSplit)
    ReDim arrNum(0)
    For i = 0 To UBound(varData)
        blnFind = False
        If i = 0 Then
            blnFind = True
            arrNum(0) = Val(varData(i))
        Else
            For j = 0 To UBound(arrNum)
                If blnAsc Then '升序
                    If arrNum(j) > Val(varData(i)) Then blnFind = True
                Else '降序
                    If arrNum(j) < Val(varData(i)) Then blnFind = True
                End If
                If blnFind Then
                    ReDim Preserve arrNum(UBound(arrNum) + 1)
                    For k = UBound(arrNum) - 1 To j Step -1
                        arrNum(k + 1) = arrNum(k)
                    Next
                    arrNum(j) = Val(varData(i))
                    Exit For
                End If
            Next
        End If
        If blnFind = False Then
            ReDim Preserve arrNum(UBound(arrNum) + 1)
            arrNum(UBound(arrNum)) = Val(varData(i))
        End If
    Next
    
    For i = 0 To UBound(arrNum)
        If blnDelRepeat And i > 0 Then '去重
            If arrNum(i) <> arrNum(i - 1) Then
                strReturn = strReturn & strSplit & arrNum(i)
            End If
        Else
            strReturn = strReturn & strSplit & arrNum(i)
        End If
    Next
    If strReturn <> "" Then strReturn = Mid(strReturn, Len(strSplit) + 1)
    
    ZlNumStrSort = strReturn
End Function

Public Sub SetEnabled(ByVal objControls As Object, ByVal blnEnabled As Boolean)
    '设置控件可用状态
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To objControls.Count - 1
        If UCase(objControls(i).Name) <> UCase("cmdHelp") _
            And UCase(objControls(i).Name) <> UCase("cmdOk") _
            And UCase(objControls(i).Name) <> UCase("cmdCancel") _
            And UCase(TypeName(objControls(i))) <> UCase("Label") _
            And UCase(TypeName(objControls(i))) <> UCase("Frame") _
            And UCase(TypeName(objControls(i))) <> UCase("TabControl") _
            And UCase(TypeName(objControls(i))) <> UCase("PictureBox") _
            And UCase(TypeName(objControls(i))) <> UCase("VSFlexGrid") Then
            objControls(i).Enabled = blnEnabled
        End If
    Next
End Sub

Public Function GetClientPoint(ByVal objFrm As Form) As POINTAPI
'获取当前指针对应在控件中的位置
    Dim pRet As POINTAPI
    Dim lngReturn As Long
    
    pRet = zlControl.GetCursorPosition()
    pRet.X = objFrm.ScaleX(pRet.X, vbPixels, vbTwips)
    pRet.Y = objFrm.ScaleY(pRet.Y, vbPixels, vbTwips)
    GetClientPoint = pRet
End Function

Public Sub SetEnabledBackColor(ByVal objControls As Object)
    '设置控件可用状态与不可用状态的背景颜色
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To objControls.Count - 1
        If UCase(TypeName(objControls(i))) = UCase("TextBox") _
            Or UCase(TypeName(objControls(i))) = UCase("ComboBox") Then
            objControls(i).BackColor = IIf(objControls(i).Enabled, vbWindowBackground, vbButtonFace)
        End If
    Next
End Sub

Public Sub SetBackColor(ByVal objControls As Object, ByVal lngBackColor As OLE_COLOR)
    '设置控件背景颜色
    Dim i As Integer
    On Error Resume Next
    
    For i = 0 To objControls.Count - 1
        If UCase(TypeName(objControls(i))) <> UCase("VSFlexGrid") _
            And UCase(TypeName(objControls(i))) <> UCase("TextBox") _
            And UCase(TypeName(objControls(i))) <> UCase("ComboBox") Then
            objControls(i).BackColor = lngBackColor
        End If
    Next
End Sub


Public Function GetSelectedIndex(ByVal OptionButtons As Object) As Integer
    '获取单选按钮组的选中项的索引
    Dim i As Integer
    
    For i = OptionButtons.LBound To OptionButtons.UBound
        If OptionButtons(i).Value Then
            GetSelectedIndex = i: Exit For
        End If
    Next
End Function

Public Sub SetReportControlBackColorAlternate(rptData As ReportControl, Optional CustomColor As OLE_COLOR = -1)
    '设置ReportControl的行交替色
    Dim i As Long, ObjItem As ReportRecordItem
    Dim lngRowCount As Long '组内行号
    
    On Error Resume Next
    For i = 0 To rptData.Rows.Count - 1
        If rptData.Rows(i).GroupRow Then
            lngRowCount = 0
        Else
            For Each ObjItem In rptData.Rows(i).Record
                If lngRowCount Mod 2 = 0 Then
                    ObjItem.BackColor = rptData.PaintManager.BackColor
                Else
                    ObjItem.BackColor = IIf(CustomColor = -1, G_AlternateColor, CustomColor)
                End If
            Next
            lngRowCount = lngRowCount + 1
        End If
    Next
End Sub

Public Sub SetVsGridRowChangeBackColor(ByVal vsGrid As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngNewRow As Long, _
    ByVal lngoldCol As Long, ByVal lngNewCol As Long, Optional CustomColor As OLE_COLOR = -1, _
    Optional ByVal lngStartCol As Long = -1, Optional ByVal lngEndCol As Long = -1)
    '功能：行列改变时,设置相关的颜色
    '入参：
    '   CustomColor - 自定义颜色，选中行颜色
    '   lngStartCol - 开始列
    '   lngEndCol - 终止列
    '说明：RowData中可能存了颜色值
    Dim lngColStart As Long, lngColEnd As Long
    Err = 0: On Error Resume Next
    With vsGrid
        If lngOldRow > .FixedRows - 1 Then
            lngColStart = IIf(lngStartCol = -1 Or .IsSubtotal(lngOldRow), .FixedCols, lngStartCol)
            lngColEnd = IIf(lngEndCol = -1 Or .IsSubtotal(lngOldRow), .Cols - 1, lngEndCol)
            .Cell(flexcpBackColor, lngOldRow, lngColStart, lngOldRow, lngColEnd) = IIf(Val(.RowData(lngOldRow)) = 0, .BackColor, Val(.RowData(lngOldRow)))
        End If
        If lngNewRow > .FixedRows - 1 Then
            lngColStart = IIf(lngStartCol = -1 Or .IsSubtotal(lngNewRow), .FixedCols, lngStartCol)
            lngColEnd = IIf(lngEndCol = -1 Or .IsSubtotal(lngNewRow), .Cols - 1, lngEndCol)
            .Cell(flexcpBackColor, lngNewRow, lngColStart, lngNewRow, lngColEnd) = IIf(CustomColor <> -1, CustomColor, -2147483635) '16772055
        End If
    End With
End Sub

Public Function GetVsfGridData(rptData As ReportControl, _
    Optional ByVal strHiddenCols As String) As VSFlexGrid
    '功能:将ReportControl转换为VSFlexGrid
    '入参:
    '   strHiddenCols 隐藏列索引(索引从0开始)，格式：列1,列2,列3,...
    Set GetVsfGridData = frmClinicPlanTemp.GetVsfGrid(rptData, strHiddenCols)
End Function

Public Function GetFirstCommandBar(ByRef objControls As CommandBarControls) As Long
'功能：获取工具栏打印预览按钮后的第一个按钮的index
    Dim objControl As CommandBarControl, idx As Long
    
    For Each objControl In objControls
        If objControl.ID = conMenu_File_Preview Then
            idx = objControl.index + 1
        End If
    Next
    GetFirstCommandBar = idx
End Function

Public Function FindNodeByKey(ByVal objNodes As Nodes, ByVal strKey As String) As Node
    '功能：根据Key值查找树节点
    Dim objNode As Node
    
    For Each objNode In objNodes
        If objNode.Key = strKey Then
            Set FindNodeByKey = objNode
            Exit For
        End If
    Next
End Function

Public Function CollExitsValue(ByVal coll As Collection, ByVal strKey As String) As Boolean
'功能：根据关键字判断元素是否存在于集合中
    Dim blnExits As Boolean
    
    If coll Is Nothing Then Exit Function
    CollExitsValue = True
    Err = 0: On Error Resume Next
    blnExits = IsObject(coll(strKey))
    If Err <> 0 Then Err = 0: CollExitsValue = False
End Function

Public Function AddRange(objOld As Collection, objAdd As Collection) As Collection
    '拼接集合
    Dim i As Long
    Dim objCol As New Collection
    
    Err = 0: On Error GoTo Errhand:
    If objOld Is Nothing And objAdd Is Nothing Then Exit Function
    If Not objOld Is Nothing Then
        For i = 1 To objOld.Count
            objCol.Add objOld(i)
        Next
    End If
    If Not objAdd Is Nothing Then
        For i = 1 To objAdd.Count
            objOld.Add objAdd(i)
        Next
    End If
    Set AddRange = objOld
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CalculatTimeInterval(ByVal BytMode As Byte, ByVal bln序号控制 As Boolean, _
    ByRef intInterval As Integer, ByVal lngMaxStartSN As Long, _
    ByVal dtStartTime As Date, ByVal dtEndTime As Date, Optional ByVal str休息时段 As String, _
    Optional ByVal lngStartSN As Long = 1, Optional ByVal lng限约数 As Long, _
    Optional ByVal strOriginalStartTime As String) As Collection
    '功能：根据间隔时间或者最大序号数计算时段
    '入参：
    '   bytMode - 计算模式：0-根据传入间隔分钟数计算，1-根据最大序号计算
    '   intInterval - 间隔时间(分钟)，bytMode=0时传入
    '   lngMaxStartSN - 最大序号
    '   dtStartTime - 开始时间
    '   dtStartTime - 终止时间
    '   intStartSN - 开始序号，用于多个时段连续编号
    '   dtOriginalStartTime 原始的时间段的开始时间，主要用于有休息时段的时间段，
    '               在进行加号时，传入的开始时间dtStartTime不是时间段的开始时间，导致无法判断休息时段是当日还是第二日
    '出参：
    '   intInterval - 间隔时间(分钟)，bytMode=1时返回
    '返回：
    '   Array(序号,开始时间,终止时间,预约数量)
    Dim objCol As Collection, lngCurSN As Long
    Dim dtCur As Date, dtTemp As Date
    Dim lngCount As Long, lngOverplus As Long
    Dim dtStart As Date, dtEnd As Date, varTimes As Variant
    Dim i As Long, j As Long
    Dim dtCurStart As Date, dtCurEnd As Date
    Dim objColAll As Collection
    Dim dtOriginalStartTime As Date
    
    Set objColAll = New Collection
    If strOriginalStartTime = "" Then
        dtOriginalStartTime = dtStartTime
    Else
        dtOriginalStartTime = CDate(strOriginalStartTime)
    End If
    If BytMode = 0 Then
        If str休息时段 = "" Then
            dtCur = dtStartTime
            lngCurSN = lngStartSN
            Do While True
                If bln序号控制 And lngCurSN > lngMaxStartSN Then Exit Do '大于最大序号时结束
                If DateDiff("n", dtEndTime, dtCur) >= 0 Then Exit Do '大于等于终止时间时结束
                dtTemp = DateAdd("n", intInterval, dtCur) '加间隔时间
                If DateDiff("n", dtTemp, dtEndTime) <= 0 Then
                    objColAll.Add Array(lngCurSN, dtCur, dtEndTime, IIf(lng限约数 > 0, 1, 0)), "K_" & lngCurSN
                    Exit Do '大于等于终止时间时结束
                Else
                    objColAll.Add Array(lngCurSN, dtCur, dtTemp, IIf(lng限约数 > 0, 1, 0)), "K_" & lngCurSN
                End If
                lngCurSN = lngCurSN + 1 '序号加1
                lng限约数 = lng限约数 - 1
                dtCur = dtTemp
            Loop
        Else
            varTimes = Split(str休息时段, ";")
            dtCurStart = dtStartTime
            For i = 0 To UBound(varTimes)
                '如果休息时段的开始时间小于当前时段的开始时间，则表示是第二天，休息时段的开始时间和终止时间都要加一天
                dtStart = CDate(Format(dtOriginalStartTime, "yyyy-mm-dd ") & Split(varTimes(i), "-")(0))
                dtEnd = CDate(Format(dtOriginalStartTime, "yyyy-mm-dd ") & Split(varTimes(i), "-")(1))
                If DateDiff("n", dtStart, dtOriginalStartTime) > 0 Then dtStart = DateAdd("d", 1, dtStart): dtEnd = DateAdd("d", 1, dtEnd)
                '休息时段的终止时间小于休息时段的开始时间，则休息时段的终止时间加一天
                If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
                
                If DateDiff("n", dtCurStart, dtStart) > 0 Then '休息时段的开始时间大于开始时间才有效
                    dtCurEnd = dtStart '当前时段的终止时间等于休息时段的开始时间
                    Set objCol = CalculatTimeInterval(0, bln序号控制, intInterval, lngMaxStartSN, dtCurStart, dtCurEnd, "", _
                        lngStartSN + objColAll.Count, lng限约数 - objColAll.Count)
                    Set objColAll = AddRange(objColAll, objCol)
                    
                    '下一个时段的开始时间等于当前休息时段的终止时间
                    dtCurStart = dtEnd
                End If
            Next
            dtCurEnd = dtEndTime
            Set objCol = CalculatTimeInterval(0, bln序号控制, intInterval, lngMaxStartSN, dtCurStart, dtCurEnd, "", _
                lngStartSN + objColAll.Count, lng限约数 - objColAll.Count)
            Set objColAll = AddRange(objColAll, objCol)
        End If
    ElseIf BytMode = 1 Then
        lngCount = GetMinuteCount(dtStartTime, dtEndTime, str休息时段)
        intInterval = lngCount \ lngMaxStartSN '每个时段的平均分钟数
        lngOverplus = lngCount - intInterval * lngMaxStartSN '剩余未分配完的分钟数，将分配到后面的序号上
        
        '间隔时间必须大于零
        If intInterval = 0 Then lngMaxStartSN = lngOverplus
        
        If str休息时段 = "" Then
            dtCur = dtStartTime
            For lngCurSN = 1 To lngMaxStartSN
                dtTemp = DateAdd("n", intInterval, dtCur) '加间隔时间
                If lngCurSN > lngMaxStartSN - lngOverplus Then
                    dtTemp = DateAdd("n", 1, dtTemp) '分配未分配完的时间，在间隔时间基础上再加1
                End If
                If DateDiff("n", dtTemp, dtEndTime) <= 0 Then
                    objColAll.Add Array(lngCurSN, dtCur, dtEndTime, IIf(lng限约数 > 0, 1, 0)), "K_" & lngCurSN
                    Exit For
                Else
                    objColAll.Add Array(lngCurSN, dtCur, dtTemp, IIf(lng限约数 > 0, 1, 0)), "K" & lngCurSN
                End If
                lng限约数 = lng限约数 - 1
                dtCur = dtTemp
            Next
        Else
            lngCurSN = 1
            varTimes = Split(str休息时段, ";")
            dtCurStart = dtStartTime: lngCount = 0
            For j = 0 To UBound(varTimes)
                '如果休息时段的开始时间小于当前时段的开始时间，则表示是第二天，休息时段的开始时间和终止时间都要加一天
                dtStart = CDate(Format(dtOriginalStartTime, "yyyy-mm-dd ") & Split(varTimes(j), "-")(0))
                dtEnd = CDate(Format(dtOriginalStartTime, "yyyy-mm-dd ") & Split(varTimes(j), "-")(1))
                If DateDiff("n", dtStart, dtOriginalStartTime) > 0 Then dtStart = DateAdd("d", 1, dtStart): dtEnd = DateAdd("d", 1, dtEnd)
                '休息时段的终止时间小于休息时段的开始时间，则休息时段的终止时间加一天
                If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
                
                If DateDiff("n", dtCurStart, dtStart) > 0 Then '休息时段的开始时间大于开始时间才有效
                    dtCurEnd = dtStart '当前时段的终止时间等于休息时段的开始时间
                    For i = lngCurSN To lngMaxStartSN
                        dtCur = DateAdd("n", intInterval, dtCurStart) '加间隔时间
                        If i > lngMaxStartSN - lngOverplus Then
                            dtCur = DateAdd("n", 1, dtCur) '分配未分配完的时间，在间隔时间基础上再加1
                        End If
                        '当前时段的终止时间大于等于休息时段的开始时间，则当前时段的终止时间等于休息时段的开始时间
                        If DateDiff("n", dtCur, dtCurEnd) <= 0 Then
                            objColAll.Add Array(lngCurSN, dtCurStart, dtCurEnd, IIf(lng限约数 > 0, 1, 0)), "K" & i
                            lngCurSN = lngCurSN + 1
                            lng限约数 = lng限约数 - 1
                            dtCurStart = dtEnd
                            Exit For
                        Else
                            objColAll.Add Array(lngCurSN, dtCurStart, dtCur, IIf(lng限约数 > 0, 1, 0)), "K" & i
                            lngCurSN = lngCurSN + 1
                            lng限约数 = lng限约数 - 1
                            dtCurStart = dtCur
                        End If
                    Next
                End If
            Next
            dtCurEnd = dtEndTime
            For i = lngCurSN To lngMaxStartSN
                dtCur = DateAdd("n", intInterval, dtCurStart) '加间隔时间
                If i > lngMaxStartSN - lngOverplus Then
                    dtCur = DateAdd("n", 1, dtCur) '分配未分配完的时间，在间隔时间基础上再加1
                End If
                '当前时段的终止时间大于等于休息时段的开始时间，则当前时段的终止时间等于休息时段的开始时间
                If DateDiff("n", dtCur, dtCurEnd) <= 0 Then
                    objColAll.Add Array(lngCurSN, dtCurStart, dtCurEnd, IIf(lng限约数 > 0, 1, 0)), "K" & i
                    lng限约数 = lng限约数 - 1
                    Exit For
                Else
                    objColAll.Add Array(lngCurSN, dtCurStart, dtCur, IIf(lng限约数 > 0, 1, 0)), "K" & i
                    lngCurSN = lngCurSN + 1
                    lng限约数 = lng限约数 - 1
                    dtCurStart = dtCur
                End If
            Next
        End If
    End If
    Set CalculatTimeInterval = objColAll
End Function


Public Function GetMinuteCount(ByVal dtStartTime As Date, ByVal dtEndTime As Date, ByVal str休息时段 As String) As Long
    '获取时段的总的分钟数
    Dim dtCurStart As Date, dtCurEnd As Date
    Dim lngCount As Long, varTimes As Variant
    Dim i As Long, dtStart As Date, dtEnd As Date
    
    Err = 0: On Error GoTo Errhand:
    dtStartTime = Format(dtStartTime, "yyyy-mm-dd") & " " & Format(dtStartTime, "HH:MM")
    dtEndTime = GetWorkTrueDate(dtStartTime, dtEndTime)
    
    If str休息时段 = "" Then
        lngCount = DateDiff("n", dtStartTime, dtEndTime) '总的分钟数
    Else
        varTimes = Split(str休息时段, ";")
        dtCurStart = dtStartTime: lngCount = 0
        For i = 0 To UBound(varTimes)
            '如果休息时段的开始时间小于当前时段的开始时间，则表示是第二天，休息时段的开始时间和终止时间都要加一天
            dtStart = CDate(Format(dtCurStart, "yyyy-mm-dd ") & Split(varTimes(i), "-")(0))
            dtEnd = CDate(Format(dtCurStart, "yyyy-mm-dd ") & Split(varTimes(i), "-")(1))
            If DateDiff("n", dtStart, dtCurStart) > 0 Then dtStart = DateAdd("d", 1, dtStart): dtEnd = DateAdd("d", 1, dtEnd)
            '休息时段的终止时间小于休息时段的开始时间，则休息时段的终止时间加一天
            If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
            
            dtCurEnd = dtStart '当前时段的终止时间等于休息时段的开始时间
            lngCount = lngCount + DateDiff("n", dtCurStart, dtCurEnd) '总的分钟数
            
            '下一个时段的开始时间等于当前休息时段的终止时间
            dtCurStart = dtEnd
        Next
        dtCurEnd = dtEndTime
        lngCount = lngCount + DateDiff("n", dtCurStart, dtCurEnd) '总的分钟数
    End If
    GetMinuteCount = lngCount
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetMaxLocalCode(ByVal strTableName As String, _
    Optional ByVal strColumn As String = "编码", Optional ByVal strWhere As String) As String
    '功能：获取指定表的最大编码(不涉及上级编码)
    '入参：表名
    '出参：成功返回 最大编码; 否者返回 空
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim strNextCode As String
    
    Err = 0: On Error GoTo ErrHandler
    '最大号码长度
    strSQL = "Select Nvl(Max(Length(" & strColumn & ")), 1) CodeLen" & vbNewLine & _
            " From " & strTableName & " Where 1=1 " & strWhere
    '最大号码
    strSQL = "Select Max(LPad(" & strColumn & ", CodeLen, '0')) As MaxCode" & vbNewLine & _
            " From " & strTableName & " A, (" & strSQL & ") B" & vbNewLine & _
            " Where 1=1 " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetMaxLocalCode")
    If rsTemp.EOF Then GetMaxLocalCode = "1": Exit Function
    
    strNextCode = zlStr.Increase(Nvl(rsTemp!MaxCode, "0")) '加1
    If strNextCode = "" Then strNextCode = "1"
    
    GetMaxLocalCode = strNextCode
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    GetMaxLocalCode = ""
End Function

Public Function GetWeekCount(ByVal intYear As Integer, ByVal intMonth As Integer) As Integer
    '功能：计算某年某月一共有多少周
    Dim dtCur As Date, intCount As Integer
    
    dtCur = CDate(intYear & "-" & intMonth & "-01")
    Do While Month(dtCur) = intMonth
        intCount = intCount + 1
        dtCur = DateAdd("ww", 1, dtCur)
    Loop
    '若日期小于星期，则再加一周
    'Weekday(dtStartDate, vbMonday):1-星期一,2-星期二,3-星期三,4-星期四,5-星期五,6-星期六,7-星期日
    If Day(dtCur) < Weekday(dtCur, vbMonday) Then intCount = intCount + 1
    GetWeekCount = intCount
End Function

Public Function GetDateWeek(ByVal dtStart As Date) As Integer
    '功能：根据当前日期计算是第几周
    Dim dtCur As Date, intWeek As Integer
    Dim dtFirstDayWeek As Integer
    
    dtCur = Format(dtStart, "yyyy-mm-01")
    dtFirstDayWeek = Weekday(dtCur, vbMonday) '记录第一天星期几
    Do While Month(dtCur) = Month(dtStart) And DateDiff("d", dtCur, dtStart) > 0
        If Weekday(dtCur, vbMonday) = dtFirstDayWeek Then '等于第一天的星期则加一周
            intWeek = intWeek + 1
        End If
        dtCur = DateAdd("d", 1, dtCur)
    Loop
    '当前日期的星期小于等于本月第一天的星期，则再加一周
    If Weekday(dtCur, vbMonday) <= dtFirstDayWeek Then intWeek = intWeek + 1
    GetDateWeek = intWeek
End Function

Public Function GetDateRange(ByVal intYear As Integer, ByVal intMonth As Integer, Optional ByVal intWeek As Integer) As Variant
    '功能：根据年月周信息获取日期范围
    '返回：数组，0-开始时间，1-终止时间
    Dim datStart As Date, datEnd As Date, datFirstDay As Date
    
    datFirstDay = intYear & "-" & intMonth & "-01"
    If intWeek = 0 Then
        '本月第一天加一月再减一天就是本月的开始日期和结束日期
        datStart = datFirstDay
        datEnd = DateAdd("d", -1, DateAdd("m", 1, datStart))
    Else
        datStart = datFirstDay
        datEnd = datFirstDay
        '如果第一天就是周一，就加一天
        If Weekday(datEnd) = vbMonday Then datEnd = DateAdd("d", 1, datEnd)
        Do While Weekday(datEnd, vbMonday) < 7
            datEnd = DateAdd("d", 1, datEnd)
        Loop
        If intWeek > 1 Then
            datStart = DateAdd("ww", intWeek - 2, DateAdd("d", 1, datEnd))
            datEnd = DateAdd("ww", intWeek - 1, datEnd)
            If DateDiff("m", datFirstDay, datEnd) > 0 Then '跳月了,直接取最后一天
                datEnd = DateAdd("d", -1, DateAdd("m", 1, datFirstDay))
            End If
        End If
    End If
    '开始时间小于当前日期，则从明天开始
    'If DateDiff("d", datStart, CDate(Format(Now, "yyyy-mm-dd"))) >= 0 Then datStart = CDate(Format(DateAdd("d", Now, 1), "yyyy-mm-dd"))
    GetDateRange = Array(datStart, datEnd)
End Function

Public Function GetWeekIndex(ByVal strWeek As String) As Integer
    '根据星期名称获取索引
    Select Case strWeek
    Case "周一", "星期一"
        GetWeekIndex = 0
    Case "周二", "星期二"
        GetWeekIndex = 1
    Case "周三", "星期三"
        GetWeekIndex = 2
    Case "周四", "星期四"
        GetWeekIndex = 3
    Case "周五", "星期五"
        GetWeekIndex = 4
    Case "周六", "星期六"
        GetWeekIndex = 5
    Case "周日", "星期日"
        GetWeekIndex = 6
    Case Else
        GetWeekIndex = -1
    End Select
End Function

Public Function GetWeekName(ByVal intIndex As Integer, Optional ByVal blnShort As Boolean = True) As String
    '根据索引获取星期名称
    '   blnShort True-如"周一"，False-如"星期一"
    Select Case intIndex
    Case 0
        GetWeekName = IIf(blnShort, "周一", "星期一")
    Case 1
        GetWeekName = IIf(blnShort, "周二", "星期二")
    Case 2
        GetWeekName = IIf(blnShort, "周三", "星期三")
    Case 3
        GetWeekName = IIf(blnShort, "周四", "星期四")
    Case 4
        GetWeekName = IIf(blnShort, "周五", "星期五")
    Case 5
        GetWeekName = IIf(blnShort, "周六", "星期六")
    Case 6
        GetWeekName = IIf(blnShort, "周日", "星期日")
    Case Else
        GetWeekName = ""
    End Select
End Function

Public Function GetTempImage(ByVal strTxt As String, _
    Optional ByVal dblWidth As Double, Optional ByVal dblHight As Double, _
    Optional ByVal lngBackColor As OLE_COLOR = vbButtonFace, _
    Optional ByVal lngForeColor As OLE_COLOR = vbBlue, _
    Optional ByVal objFont As StdFont, _
    Optional ByVal intAlignment As PictureTextAlignmentSettings = pictxtAlignCenterCenter, _
    Optional ByVal strSubTxt As String, _
    Optional ByVal lngSubForeColor As OLE_COLOR = vbBlack, _
    Optional ByVal objSubFont As StdFont, _
    Optional ByVal intSubAlignment As PictureTextAlignmentSettings = pictxtAlignCenterCenter) As IPictureDisp
    '功能：根据参数生成图片
    '入参：
    '   strTxt - 主要显示文本
    '   dblWidth,dblHight - 图片大小，缺省为文本打印出来后的宽度和高度
    '   lngBackColor - 图片背景色，缺省为按钮表面颜色
    '   lngForeColor - 主要文本的前景色
    '   objFont - 主要文本的字体
    '   intAlignment - 主要文本的相对位置
    '   strSubTxt - 附加文本
    '   lngSubForeColor - 附加文本的前景色
    '   objSubFont - 附加文本的字体
    '   intSubAlignment - 附加文本的相对位置
    '返回：图片对象
    Set GetTempImage = frmClinicPlanTemp.GetTempPicture(strTxt, dblWidth, dblHight, lngBackColor, _
        lngForeColor, objFont, intAlignment, strSubTxt, lngSubForeColor, objSubFont, intSubAlignment)
End Function

Public Function FormatApplyToStr(ByVal strApply As String) As String
    '格式化被应用于字符串
    '入参：
    '   strApply yyyy-mm-dd|yyyy-mm-dd|yyyy-mm-dd|...
    '说明：用于按天安排出诊的月模板
    Dim varStr As Variant, i As Integer
    Dim strNewStr As String
    
    If InStr(strApply, "|") = 0 And IsDate(strApply) = False Then
        FormatApplyToStr = strApply
        Exit Function
    End If
    
    varStr = Split(strApply, "|")
    For i = 0 To UBound(varStr)
        strNewStr = strNewStr & "|" & Day(varStr(i)) & "日"
    Next
    If strNewStr <> "" Then strNewStr = Mid(strNewStr, 2)
    FormatApplyToStr = strNewStr
End Function

Public Function CheckTimeBucketIsCross(ByVal dtStartA As Date, ByVal dtEndA As Date, _
    ByVal dtStartB As Date, ByVal dtEndB) As Boolean
    '检查时间段是否有交叉
    '说明：
    '   1.如果时间段跨了天，则需要再分段检查
    '   2.传入时间必须是已处理为参考同一天的
    Dim blnTwoDayA As Boolean, blnTwoDayB As Boolean
    Dim blnHaveCross As Boolean
    Dim dtStartTemp As Date, dtEndTemp As Date
    
    blnTwoDayA = DateDiff("d", dtStartA, dtEndA) > 0
    blnTwoDayB = DateDiff("d", dtStartB, dtEndB) > 0
    If blnTwoDayA And blnTwoDayB Then '两个时间段均跨天,肯定有交叉
        CheckTimeBucketIsCross = True
        Exit Function
    End If
    
    blnHaveCross = Not (DateDiff("n", dtStartA, dtEndB) <= 0 Or DateDiff("n", dtEndA, dtStartB) >= 0)
    If blnHaveCross Then CheckTimeBucketIsCross = True: Exit Function '已知有交叉，直接退出
    
    If blnTwoDayA And blnTwoDayB = False Then 'A时间段跨天，B时间段不跨天
        '将A时间段第二天的部分与B比较
        dtStartTemp = CDate(Format(dtStartA, "yyyy-mm-dd 00:00"))
        dtEndTemp = CDate(Format(dtStartA, "yyyy-mm-dd ") & Format(dtEndA, "HH:mm"))
        blnHaveCross = Not (DateDiff("n", dtStartTemp, dtEndB) <= 0 Or DateDiff("n", dtEndTemp, dtStartB) >= 0)
    ElseIf blnTwoDayA = False And blnTwoDayB Then 'A时间段不跨天，B时间段跨天
        '将B时间段第二天的部分与A比较
        dtStartTemp = CDate(Format(dtStartB, "yyyy-mm-dd 00:00"))
        dtEndTemp = CDate(Format(dtStartB, "yyyy-mm-dd ") & Format(dtEndB, "HH:mm"))
        blnHaveCross = Not (DateDiff("n", dtStartA, dtEndTemp) <= 0 Or DateDiff("n", dtEndA, dtStartTemp) >= 0)
    Else '两个时间段均不跨天，已比较出结果
        'blnHaveCross
    End If
    CheckTimeBucketIsCross = blnHaveCross
End Function

Public Function IsDoubleMonthWeekPlan(ByRef intYear As Integer, ByRef intMonth As Integer, _
    ByRef intWeek As Integer, ByRef dtStartDate As Date, ByRef dtEndDate As Date) As Boolean
    '判断并获取跨月的周安排的另一个出诊表的年月周
    '入参：
    '   dtStartDate、dtEndDate 要判断的出诊表的时间范围
    '返回：如果存在跨月则返回True，否则返回False
    '说明：如果返回True，则
    '       intWeek返回另一个出诊表的年月周
    '       dtStartDate、dtEndDate分别返回完整周(七天)的开始时间和结束时间
    
    If DateDiff("d", dtStartDate, dtEndDate) >= 6 Then Exit Function
    
    '存在跨月的，查找另一个出诊表的年月周
    If Month(DateAdd("d", -1, dtStartDate)) <> Month(dtStartDate) Then
        '当前是第一周
        dtStartDate = DateAdd("d", DateDiff("d", dtStartDate, dtEndDate) - 6, dtStartDate)
        intYear = Year(dtStartDate): intMonth = Month(dtStartDate)
        intWeek = GetWeekCount(intYear, intMonth)
    ElseIf Month(DateAdd("d", 1, dtEndDate)) <> Month(dtEndDate) Then
        '当前是最后一周
        dtEndDate = DateAdd("d", 6 - DateDiff("d", dtStartDate, dtEndDate), dtEndDate)
        intYear = Year(dtEndDate): intMonth = Month(dtEndDate)
        intWeek = 1
    End If
    IsDoubleMonthWeekPlan = True
End Function

Public Function GetPopupCommandBar(frmMain As Form, cbsMain As CommandBars, _
    Optional ByVal lngControlPopupID As Long = conMenu_EditPopup) As CommandBar
    '构建弹出菜单
    Dim objPopup As CommandBarPopup, cbCommandBar As CommandBar
    Dim cbrControl As CommandBarControl, cbrControlNew As CommandBarControl
    Dim i As Integer
    
    Set objPopup = cbsMain.FindControl(xtpControlPopup, lngControlPopupID, , True)
    If objPopup Is Nothing Then Exit Function
    Set cbCommandBar = cbsMain.Add("Popup", xtpBarPopup) '弹出菜单
    If cbCommandBar Is Nothing Then Exit Function
    
    For i = 1 To objPopup.CommandBar.Controls.Count
        Set cbrControl = objPopup.CommandBar.Controls(i)
        Call frmMain.zlUpdateCommandBars(cbrControl) '判断是否可见，因为第一次时菜单还没有执行Update
        If cbrControl.Visible Then
            Set cbrControlNew = cbCommandBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
            cbrControlNew.BeginGroup = cbrControl.BeginGroup
            cbrControlNew.Enabled = cbrControl.Enabled
        End If
    Next
    Set GetPopupCommandBar = cbCommandBar
End Function

