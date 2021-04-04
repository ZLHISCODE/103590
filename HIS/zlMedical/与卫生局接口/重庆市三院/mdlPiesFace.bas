Attribute VB_Name = "mdlPiesFace"
Option Explicit

Public glngTXTProc As Long
Public gstrSysName As String
Public gstrDBUser As String
Public gstrSQL As String
Public gstrPrive As String

Public gcnAccess As New Connection

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    部门 As String
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type

Public UserInfo As TYPE_USER_INFO

Public Sub OpenRecordSet(rsTemp As ADODB.Recordset, Optional ByVal strFormCaption As String)
'功能：打开记录。同时保存SQL语句
    If rsTemp.State = adStateOpen Then rsTemp.Close

    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly

End Sub

Public Sub OpenAccessRecordSet(rsTemp As ADODB.Recordset, Optional ByVal strFormCaption As String)
'功能：打开记录。同时保存SQL语句
    If rsTemp.State = adStateOpen Then rsTemp.Close

    rsTemp.Open gstrSQL, gcnAccess, adOpenStatic, adLockReadOnly

End Sub

Public Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    
    Dim blnForeColor As Boolean
    Dim blnBkColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rsData("前景色").Name = "前景色")
    blnBkColor = (rsData("背景色").Name = "背景色")
    
    On Error GoTo 0
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
        lngRow = 0
    Else
        
        If Val(objMsf.RowData(objMsf.Rows - 1)) <= 0 Then
            lngRow = objMsf.Rows - 2
        Else
            lngRow = objMsf.Rows - 1
        End If
                
    End If
    
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            If Trim(objMsf.TextMatrix(0, lngLoop)) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If strMask <> "" Then
                    objMsf.TextMatrix(lngRow, lngLoop) = Format(NVL(rsData(objMsf.TextMatrix(0, lngLoop))), strMask)
                Else
                    objMsf.TextMatrix(lngRow, lngLoop) = NVL(rsData(objMsf.TextMatrix(0, lngLoop)))
                End If
            End If
            
        Next
        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("前景色").Value)
        If blnBkColor Then objMsf.Cell(flexcpBackColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("背景色").Value)
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
errHand:
    
End Function

Public Sub LocationObj(ByRef objTxt As Object)
    On Error Resume Next
    
    TxtSelAll objTxt
    objTxt.SetFocus
End Sub



Public Sub LocationGrid(ByRef vsf As Object, Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1)
    
    On Error Resume Next
    
    If lngRow <> -1 Then vsf.Row = lngRow
    If lngCol <> -1 Then vsf.Col = lngCol
    
    vsf.SetFocus
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

Public Function GetCol(ByVal objVsf As Object, ByVal strData As String) As Long
    
    Dim lngLoop As Long
    
    GetCol = -1
    For lngLoop = 0 To objVsf.Cols - 1
        If objVsf.Cell(flexcpText, 0, lngLoop, 0, lngLoop) = strData Then
            GetCol = lngLoop
            Exit Function
        End If
    Next
    
End Function


Public Function CreateVsf(ByRef objVsf As Object, ByVal strVsf As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim varArray As Variant
    Dim varItem As Variant
    Dim i As Integer
    
    On Error GoTo errHand
    
    objVsf.Cols = 0
    
    varArray = Split(strVsf, ";")
    For lngLoop = 0 To UBound(varArray)
        varItem = Split(varArray(lngLoop), ",")
                
        objVsf.Cols = objVsf.Cols + 1
        i = objVsf.Cols - 1
    
        objVsf.TextMatrix(0, i) = varItem(0)
        objVsf.ColWidth(i) = Val(varItem(1))
        objVsf.ColAlignment(i) = Val(varItem(2))
        objVsf.ColHidden(i) = (Val(varItem(4)) = 0)
        objVsf.Cell(flexcpData, 0, i) = IIf(varItem(5) = "", varItem(0), varItem(5))
        
    Next
    
    CreateVsf = True
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then Resume
End Function

Public Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    
    On Error GoTo errHand
    
'    Exit Function
    
    If objVsf.Rows = 0 Then Exit Function
    lngTop = objVsf.Cell(flexcpTop, objVsf.Rows - 1, 0) + objVsf.RowHeight(objVsf.Rows - 1)
    
    '1.隐藏所有的线
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.重新计算需要的纵线
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.重新计算需要的横线
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeightMin) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeightMin + IIf(lngIndex = 1, 30, 0)
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
errHand:
    
End Function

Public Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '功能：重新定义数组
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function

Public Function NextNo(intBillID As Integer, Optional ByVal intStep As Integer = 1) As Variant
'功能：根据特定规则产生新的号码,规则如下：
'   一、项目序号：
'   1   病人ID         数字
'   2   住院号         数字
'   3   门诊号         数字
'   10  医嘱发送号     数字,顺序递增编号
'   x   其它单据号     字符,根据编号规则顺序递增编号,不自动补缺
'   二、年度位确定原则:
'       以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码

    Dim rsCtrl As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim vntNo As Variant, strSQL As String
    Dim intYear, strYear As String
    Dim curDate As Date, blnByDate As Boolean
ReStart:
    Err = 0
    On Error GoTo errHand

    If intBillID = 1 Then '病人ID
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From 号码控制表 Where 项目序号=" & intBillID
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            strSQL = "Select Nvl(Max(病人ID),0)+1 as 病人ID From 病人信息 Where 病人ID>=" & vntNo
            
            With rsTmp
                If .State = adStateOpen Then .Close
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 2 Then '住院号
        '顺序编号还是日期编号
        strSQL = "Select A.*,Sysdate as 日期 From 系统参数表 A Where A.参数号=27"
        With rsTmp
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            
            If Not .EOF Then
                blnByDate = (IIf(IsNull(!参数值), 1, !参数值) = 2)
                curDate = !日期
            End If
        End With
        
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From 号码控制表 Where 项目序号=" & intBillID
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(住院号),0)+1 as 住院号 From 病人信息 Where 住院号>=" & vntNo
            Else
                strSQL = "Select Nvl(Max(住院号),To_Number(To_Char(Sysdate,'YYMM')||'0000'))+1 as 住院号" & _
                    " From 病人信息 Where 住院号 Like To_Number(To_Char(Sysdate,'YYMM'))||'%' And 住院号>=" & vntNo
            End If
            
            With rsTmp
                If .State = adStateOpen Then .Close
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            If Not blnByDate Then
                .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "最大号码", IIf(vntNo - 10 > Val(Format(curDate, "YYMM0000")), vntNo - 10, Val(Format(curDate, "YYMM0001")))
            End If
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 3 Then '门诊号
        '顺序编号还是日期编号
        strSQL = "Select A.*,Sysdate as 日期 From 系统参数表 A Where A.参数号=46"
        With rsTmp
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            
            If Not .EOF Then
                blnByDate = (IIf(IsNull(!参数值), 1, !参数值) = 2)
                curDate = !日期
            End If
        End With
    
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From 号码控制表 Where 项目序号=" & intBillID
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!最大号码), 0, !最大号码)
            
            If Not blnByDate Then
                strSQL = "Select Nvl(Max(门诊号),0)+1 as 门诊号 From 病人信息 Where 门诊号>=" & vntNo
            Else
                strSQL = "Select Nvl(Max(门诊号),To_Number(To_Char(Sysdate,'YYMMDD')||'0000'))+1 as 门诊号" & _
                    " From 病人信息 Where 门诊号 Like To_Number(To_Char(Sysdate,'YYMMDD'))||'%' And 门诊号>=" & vntNo
            End If
            
            With rsTmp
                If .State = adStateOpen Then .Close
                
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            If Not blnByDate Then
                .Update "最大号码", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            Else
                .Update "最大号码", IIf(vntNo - 10 > Val(Format(curDate, "YYMMDD0000")), vntNo - 10, Val(Format(curDate, "YYMMDD0001")))
            End If
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    ElseIf intBillID = 10 Then '医嘱发送号
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From 号码控制表 C Where C.项目序号=" & intBillID
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            vntNo = Val(IIf(IsNull(!最大号码), 0, !最大号码)) + 1
            
            On Error Resume Next
            .Update "最大号码", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
        
    ElseIf intBillID = 81 Then      '体检编号
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From 号码控制表 C Where C.项目序号=" & intBillID
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            vntNo = Val(IIf(IsNull(!最大号码), 0, !最大号码)) + 1
            
            On Error Resume Next
            .Update "最大号码", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    Else
        
        Dim vtnEndNo As Variant
        
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From 号码控制表 C Where C.项目序号=" & intBillID
            If .State = adStateOpen Then .Close
            
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            intYear = Format(!Today, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            vntNo = IIf(IsNull(!最大号码), "", !最大号码)
            
            If IIf(IsNull(!编号规则), 0, !编号规则) = 1 Then
                '按日顺序编号
                If vntNo < strYear & Format(CDate("1992-" & Format(!Today, "MM-dd")) - CDate("1992-01-01") + 1, "000") & "0000" Then
                    vntNo = strYear & Format(CDate("1992-" & Format(!Today, "MM-dd")) - CDate("1992-01-01") + 1, "000") & "0000"
                End If
                vtnEndNo = Left(vntNo, 4) & Right(String(4, "0") & CStr(Val(Mid(vntNo, 5)) + intStep), 4)
                vntNo = Left(vntNo, 4) & Right(String(4, "0") & CStr(Val(Mid(vntNo, 5)) + 1), 4)
            Else
                '按年顺序编号
                If Left(vntNo, 1) < strYear Then
                    vntNo = strYear & "0000000"
                End If
                vtnEndNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + intStep), 7)
                vntNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + 1), 7)
            End If
            
            If Not (UCase(strYear) >= "A" And UCase(strYear) <= "Z") Or ActualLen(vntNo) > 8 Then GoTo ReStart
            
            On Error Resume Next
            .Update "最大号码", vtnEndNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    End If
    Exit Function
errHand:
    'If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
    NextNo = Null
End Function

Public Function GetNextPatientID() As Long
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select max(病人id) as 病人id From 病人信息 "
    rs.Open strSQL, gcnOracle
    If rs.BOF = False Then
        GetNextPatientID = NVL(rs("病人id"), 0)
    End If
    GetNextPatientID = GetNextPatientID + 1
    
End Function

Public Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'功能: 装载数据入指定的组合下拉框或网格中的下拉框中
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            
            If rsTemp1.Fields.Count > 2 Then
                If Val(rsTemp1.Fields(2).Value) = 1 Then
                    objSource.ListIndex = objSource.NewIndex
                End If
            End If
            
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Public Function ShowTxtSelectDialog(ByVal frmParent As Object, _
                                    ByVal objTxt As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional blnMuliSel As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开树型+列表结构
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
        
    If Trim(strSQL) = "" Then Exit Function
    
    On Error GoTo errHand
    
    Call OpenRecord(rs, strSQL, frmParent.Caption, adOpenStatic, adLockBatchOptimistic)
    If rs.BOF Then
        MsgBox "没有可选择的数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objTxt.hWnd, objPoint)
                
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTxt.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
    
    If frmSelectDialog.ShowSelect(frmParent, 3, rs, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTxt.Height, , strSavePath, , False, blnMuliSel) Then
                            
        Set rsResult = rs
        ShowTxtSelectDialog = True
        
    End If
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then Resume
    
End Function

Public Function ShowTxtFilterDialog(ByVal frmParent As Object, _
                                    ByVal objTxt As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能;显示文本输入选择列表(只用于文本框控件)
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    Dim strInput As String
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo errHand

    If InStr(objTxt.Text, "'") > 0 Then Exit Function
    
    '参数初始化
    strInput = "'%" & UCase(objTxt.Text) & "%'"
    Call ClientToScreen(objTxt.hWnd, objPoint)
    
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTxt.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        
    '执行查询
    Call OpenRecord(rs, strSQL, frmParent.Caption)
    If rs.BOF Then
        If blnPrompt Then MsgBox "没有找到相匹配的结果！", , gstrSysName
        Exit Function                            '没有结果，直接返回
    End If
            
    If rs.RecordCount = 1 And blnFilter Then GoTo over                    '因为是输入查找，如果只有一条，则直接返回
    'If frmSelectList.ShowSelect(frmParent, rs, strLvw, lngX, lngY, lngCX, lngCY, strSavePath, strDescrible, , , objTxt.Height) Then GoTo Over
    
    If frmSelectDialog.ShowSelect(frmParent, 2, rs, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTxt.Height, , strSavePath, , False, False) Then GoTo over
    
    Exit Function
    
over:
    
    Set rsResult = rs
    
    ShowTxtFilterDialog = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then Resume
End Function

Public Function OpenRecord(rsTmp As ADODB.Recordset, strSQL As String, ByVal strTitle As String, _
    Optional CursorType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    
    rsTmp.Open strSQL, gcnOracle, CursorType, LockType
        
    Set OpenRecord = rsTmp
End Function

Public Sub ExecuteProc(ByVal strSQL As String, ByVal strCaption As String)
'功能：执行SQL语句
    
    If UCase(Left(strSQL, 3)) = "ZL_" Then
        gcnOracle.Execute strSQL, , adCmdStoredProc
    Else
        gcnOracle.Execute strSQL
    End If
    
End Sub

Public Function CloseChildWindows(ByVal frmMain As Object, ByVal FrmSon As Object) As Boolean
    '功能:关闭所有子窗口
    
    Dim frmThis As Form
    
    On Error Resume Next
    
    CloseChildWindows = True
    
    For Each frmThis In Forms
        If frmThis.Caption <> frmMain.Caption And frmThis.Caption <> FrmSon.Caption Then Unload frmThis
    Next
    
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '-----------------------------------------------------------------------------------------
    '功能:获取特殊时间
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    
    Select Case strMode
    Case "当  时"      '当时
        GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetDateTime = Format(Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetDateTime = Format(Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetDateTime = Format(Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    End Select
    
End Function

Public Function LoadGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True, Optional ByVal objIls As Object, Optional ByVal blnCharge As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    Dim strField As String
    Dim strIcon As String
    Dim blnField As Boolean
    Dim blnForeColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rsData("前景色").Name = "前景色")
    
    On Error GoTo 0
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            strField = objMsf.Cell(flexcpData, 0, lngLoop)
            
            If Trim(strField) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If Left(strField, 1) = "[" Then
                
                    strField = Mid(strField, 2, Len(strField) - 2)
                    strIcon = ""
                    
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                    If Not (objIls Is Nothing) Then
                        strIcon = NVL(rsData(strField))
                        If strIcon <> "" Then
                            Set objMsf.Cell(flexcpPicture, lngRow, lngLoop) = objIls.ListImages(strIcon).Picture
                        End If
                    End If
                    
                    objMsf.Cell(flexcpData, lngRow, lngLoop) = strIcon
                    objMsf.TextMatrix(lngRow, lngLoop) = strIcon
                Else
                
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                     If strMask <> "" Then
                        objMsf.TextMatrix(lngRow, lngLoop) = Format(NVL(rsData(strField)), strMask)
                    Else
                        objMsf.TextMatrix(lngRow, lngLoop) = NVL(rsData(strField))
                    End If
                
                    objMsf.Cell(flexcpData, lngRow, lngLoop, lngRow, lngLoop) = objMsf.TextMatrix(lngRow, lngLoop)
                End If
                
            End If
NextCol:
            '下一列
        Next
        
pointNext:
        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("前景色").Value)
        
        rsData.MoveNext
    Loop
    
    LoadGrid = True
    Exit Function
    
errHand:
'
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Function

Public Function SQLInit(ByRef rs As ADODB.Recordset) As Boolean
    
    Set rs = New ADODB.Recordset
    
    With rs
        .Fields.Append "SQL", adVarChar, 4000
        .Open
    End With
    
    SQLInit = True
    
End Function

Public Function SQLAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String) As Boolean
    
    With rs
        .AddNew
        .Fields("SQL").Value = strSQL
    End With
    
    SQLAdd = True
    
End Function

Public Sub ResetVsf(objVsf As Object)
    '
    objVsf.Rows = 2
    objVsf.RowData(1) = ""
    objVsf.Cell(flexcpText, 1, 0, 1, objVsf.Cols - 1) = ""
    
    On Error Resume Next
    
    Set objVsf.Cell(flexcpPicture, 1, 0, 1, objVsf.Cols - 1) = Nothing
End Sub

Public Function ConnectAccess(ByVal strFile As String) As Boolean
    
    If gcnAccess.State = adStateOpen Then gcnAccess.Close
    
    Set gcnAccess = New ADODB.Connection
    gcnAccess.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFile & ";Persist Security Info=False"
    
    If gcnAccess.State <> adStateOpen Then
        ShowSimpleMsg "连接'" & strFile & "'失败！"
        Exit Function
    End If
    
    ConnectAccess = True
    
End Function

Public Function AcceptPackage(frmMain As Object, ByVal strFile As String, Optional ByVal strTitle As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:接受任务包
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    Dim rs As New ADODB.Recordset
    Dim rsTask As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim bytNew As Byte
    Dim intCount As Long
    Dim intCount1 As Long
    Dim intCount2 As Long
    Dim lngTotal As Long
    
    Dim lngKey As Long
    Dim lng门诊号 As Long
    Dim lng病人id As Long
            
    If ConnectAccess(strFile) = False Then Exit Function
                
    '1.任务
    gstrSQL = "Select taskcode,taskname,taskyear,builddate From htask"
    Call OpenAccessRecordSet(rsTask, strTitle)
    If rsTask.BOF Then Exit Function
         
    '检查此任务包是否已接受
    gstrSQL = "Select b.ID,b.体检状态 From 体检登记记录_干保 a,体检登记记录 b Where a.登记id=b.ID AND a.任务包号=[1]"
    Set rs = OpenSQLRecord(gstrSQL, strTitle, rsTask("taskcode").Value)
    If rs.BOF = False Then
        
        If NVL(rs("体检状态").Value) >= 4 Then
            ShowSimpleMsg "此任务包已经开始体检，不能重新接受！"
        Else
            If MsgBox("是否重新接受当前任务包？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
                GoTo over
            Else
                '删除原来的
                lngKey = rs("ID").Value
            End If
        End If
        
    End If
    
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    
    frmWait.OpenWait frmMain, "接受任务包"
    
    If lngKey > 0 Then
        frmWait.WaitInfo = "正在删除原来接受的任务包..."
        gstrSQL = "ZL_体检登记记录_DELETE(" & lngKey & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
    End If
    
                
    frmWait.WaitInfo = "正在接受任务包..."
    
    lngKey = GetNextId("体检登记记录")
    
    
    gstrSQL = "ZL_体检登记记录_INSERT(" & lngKey & ",'" & _
                                        NextNo(78) & "'," & _
                                        "1," & _
                                        "1," & _
                                        "NULL," & _
                                        "NULL," & _
                                        "NULL," & _
                                        "NULL," & _
                                        Val(GetSetting("ZLSOFT", "公共全局\干保接口", "合约单位", 0)) & "," & _
                                        "1," & _
                                        "Sysdate+1," & _
                                        Val(GetSetting("ZLSOFT", "公共全局\干保接口", "体检部门", 0)) & "," & _
                                        "NULL," & _
                                        "Sysdate," & _
                                        "NULL," & _
                                        "1," & _
                                        "1," & _
                                        "NULL)"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    gstrSQL = "Insert Into 体检登记记录_干保(登记id,任务包号,任务包名,发送状态) Values ("
    gstrSQL = gstrSQL & lngKey
    gstrSQL = gstrSQL & ",'" & rsTask("taskcode").Value & "'"
    gstrSQL = gstrSQL & ",'" & rsTask("taskname").Value & "'"
    gstrSQL = gstrSQL & ",0"
    gstrSQL = gstrSQL & ")"
    gcnOracle.Execute gstrSQL
    
    '2.组别（套餐）
    gstrSQL = "Select taskcode,asmcode,asmseq,asmname,asmsex,asmdesc From htaskasm"
    
    Call OpenAccessRecordSet(rs, strTitle)
    If rs.BOF Then GoTo over
    Do While Not rs.EOF
        
        gstrSQL = "ZL_体检组别_INSERT(" & lngKey & ",'" & rs("asmname").Value & "')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        gstrSQL = "Insert Into 体检组别_干保(登记id,组别名称,套餐编码,套餐序号,套餐名称) Values ("
        gstrSQL = gstrSQL & lngKey & ","
        gstrSQL = gstrSQL & "'" & rs("asmname").Value & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("asmcode").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("asmseq").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("asmname").Value) & "')"
        gcnOracle.Execute gstrSQL
        
        rs.MoveNext
    Loop
    
    '3.人员
    
    frmWait.WaitInfo = "正在接受体检人员..."
    frmWait.ShowProgress = True
    
    gstrSQL = "Select a.asmname,b.* From htaskasm a,htaskmemb b Where b.asmcode=a.asmcode and b.asmseq=a.asmseq"
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = adUseClient
    Call OpenAccessRecordSet(rs, strTitle)
    
    If rs.BOF Then GoTo over
    
    intCount1 = 0
    lngTotal = rs.RecordCount
    Do While Not rs.EOF
        
        lng病人id = 0
        lng门诊号 = 0
        intCount1 = intCount1 + 1
        
        frmWait.WaitProgress = Format(100 * intCount1 / lngTotal, "0.00")
        
        gstrSQL = "Select * From 病人信息 Where 健康号='" & rs("membcode").Value & "'"
        Call OpenRecord(rsTmp, gstrSQL, "接受任务包")
        If rsTmp.BOF = False Then
            lng门诊号 = NVL(rsTmp("门诊号").Value, 0)
            lng病人id = NVL(rsTmp("病人id").Value, 0)
        End If
        
        bytNew = 0
        If lng病人id = 0 Then
            lng病人id = GetNextPatientID + intCount
            intCount = intCount + 1
            bytNew = 1
        End If
        
        If lng门诊号 = 0 Then
            intCount2 = intCount2 + 1
            lng门诊号 = NextNo(3) + intCount2
        End If
        
        If rsTmp.BOF = False Then
            gstrSQL = "ZL_体检人员档案_INSERT(" & lngKey & "," & _
                                                lng病人id & "," & _
                                                "'" & rs("asmname").Value & "','" & _
                                                rs("a0101").Value & "','" & _
                                                NVL(rsTmp("身份证号").Value) & "','" & _
                                                rs("a0107").Value & "'," & _
                                                IIf(rsTmp("出生日期").Value = "", "NULL", "TO_DATE('" & rsTmp("出生日期").Value & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                                NVL(rsTmp("婚姻状况").Value) & "','" & _
                                                NVL(rsTmp("民族").Value) & "','" & _
                                                NVL(rsTmp("国籍").Value) & "','" & _
                                                rs("a0405").Value & "','" & _
                                                NVL(rsTmp("职业").Value) & "','" & _
                                                NVL(rsTmp("联系人姓名").Value) & "','" & _
                                                NVL(rsTmp("联系人电话").Value) & "','" & _
                                                "','" & _
                                                NVL(rsTmp("联系人地址").Value) & "','" & _
                                                rs("b0105").Value & "','" & _
                                                rs("age").Value & "'," & _
                                                lng门诊号 & ",'" & _
                                                NVL(rsTmp("IC卡号").Value) & "','" & _
                                                rs("membcode").Value & "',''," & _
                                                "1," & _
                                                IIf(intCount1 = rs.RecordCount, "1", "0") & ",0," & bytNew & _
                                                ")"
        Else
            gstrSQL = "ZL_体检人员档案_INSERT(" & lngKey & "," & _
                                                lng病人id & "," & _
                                                "'" & rs("asmname").Value & "','" & _
                                                rs("a0101").Value & "'," & _
                                                "NULL,'" & _
                                                rs("a0107").Value & "'," & _
                                                "NULL," & _
                                                "NULL," & _
                                                "NULL," & _
                                                "NULL,'" & _
                                                rs("a0405").Value & "'," & _
                                                "NULL," & _
                                                "NULL," & _
                                                "NULL," & _
                                                "NULL," & _
                                                "NULL,'" & _
                                                rs("b0105").Value & "','" & _
                                                rs("age").Value & "'," & _
                                                lng门诊号 & "," & _
                                                "NULL,'" & _
                                                rs("membcode").Value & "',''," & _
                                                "1," & _
                                                IIf(intCount1 = rs.RecordCount, "1", "0") & ",0," & bytNew & _
                                                ")"
        End If
        
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        gstrSQL = "Insert Into 体检人员档案_干保(登记id,病人id,任务包号,人员序号,单位编码,单位名称,在职情况,任职级别) Values ("
        gstrSQL = gstrSQL & lngKey & ","
        gstrSQL = gstrSQL & lng病人id & ","
        gstrSQL = gstrSQL & "'" & rsTask("taskcode").Value & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("taskseq").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("b0110").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("b0105").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("a6405").Value) & "',"
        gstrSQL = gstrSQL & "'" & NVL(rs("a0704").Value) & "')"
        gcnOracle.Execute gstrSQL
        
        rs.MoveNext
    Loop
    
    '4.项目
    frmWait.WaitInfo = "正在接受体检项目..."
    
    gstrSQL = "Select Distinct a.asmname,b.unioncode From htaskasm a,htaskasmunion b " & _
                "Where a.taskcode=b.taskcode and a.asmcode=b.asmcode and a.asmseq=b.asmseq and a.taskcode='" & rsTask("taskcode").Value & "'"
                
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open gstrSQL, gcnAccess, adOpenStatic
    If rs.BOF Then GoTo over
    
    Dim lngItemKey As Long
    
    intCount1 = 0
    lngTotal = rs.RecordCount
    Do While Not rs.EOF
        
        intCount1 = intCount1 + 1
        
        frmWait.WaitProgress = Format(100 * intCount1 / lngTotal, "0.00")
        
        gstrSQL = "Select b.ID,b.类别,0 As 基本价格,0 As 体检价格,0 As 执行科室id,0 As 采集方式id,0 As 采集科室id,'' As 检验标本,'' As 价格清单 " & _
                    "From 诊疗项目目录_干保 a,诊疗项目目录 b " & _
                    "Where a.诊疗项目id=b.ID and a.干保编码='" & rs("unioncode").Value & "'"
        
        Set rsItem = New ADODB.Recordset
        rsItem.Open gstrSQL, gcnOracle, adOpenStatic, adLockOptimistic
        
        If rsItem.BOF = False Then
        
            Call FinishFillItem(rsItem, Val(GetSetting("ZLSOFT", "公共全局\干保接口", "体检部门", 0)), "接受任务")
            
            lngItemKey = GetNextId("体检项目清单")
            
            gstrSQL = "ZL_体检项目清单_INSERT(" & lngKey & "," & _
                                            "'" & rs("asmname").Value & "'," & _
                                            rsItem("ID").Value & "," & _
                                            "NULL," & _
                                            Val(rsItem("基本价格").Value) & "," & _
                                            Val(rsItem("体检价格").Value) & "," & _
                                            rsItem("执行科室id").Value & "," & _
                                            IIf(rsItem("采集方式id") = "", "NULL", rsItem("采集方式id")) & "," & _
                                            IIf(rsItem("采集科室id") = "", "NULL", rsItem("采集科室id")) & ",'" & _
                                            rsItem("检验标本").Value & "'," & _
                                            "NULL," & _
                                            "NULL,NULL, 1,'" & _
                                            NVL(rsItem("价格清单").Value, "") & "')"
            
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
            
'            gstrSQL = "Insert Into 体检项目清单_干保(清单id,任务包号,组合编码,项目编码,项目分支,项目方法,组合科室) Values ("
'            gstrSQL = gstrSQL & lngItemKey & ","
'            gstrSQL = gstrSQL & "'" & rsTask("taskcode").Value & "',"
'            gstrSQL = gstrSQL & "'" & NVL(rs("a6405").Value) & "',"
'            gstrSQL = gstrSQL & "'" & NVL(rs("a0704").Value) & "',"
'            gstrSQL = gstrSQL & "'" & NVL(rs("asmcode").Value) & "',"
'            gstrSQL = gstrSQL & "'" & NVL(rs("asmseq").Value) & "',"
'            gstrSQL = gstrSQL & "'" & NVL(rs("asmname").Value) & "')"
'
'            gcnOracle.Execute gstrSQL
            
        End If
        
        rs.MoveNext
    Loop
    
    gstrSQL = "ZL_体检登记记录_STATE(" & lngKey & ",2)"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    frmWait.CloseWait
    
    gcnOracle.CommitTrans
    
    AcceptPackage = True
    
    Exit Function
    
over:
    frmWait.CloseWait
    gcnOracle.RollbackTrans
    
    If gcnAccess.State = adStateOpen Then gcnAccess.Close
    
    Exit Function
    
errHand:
    Dim strError As String
    
    strError = Err.Description
    
    frmWait.CloseWait
    
    gcnOracle.RollbackTrans
    
    ShowSimpleMsg strError
    If gcnAccess.State = adStateOpen Then gcnAccess.Close
    
'    Resume
End Function

Private Function FinishFillItem(ByRef rsItem As ADODB.Recordset, ByVal mlngDept As Long, Optional ByVal strTitle As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strKeys As String
    Dim sglSum As Single
    Dim lngExecDept As Long
    Dim strTmp As String
    Dim lngLoop As Long
    Dim strCombList As String
    Dim lngKey As Long
        
    On Error GoTo errHand
    
    lngKey = rsItem("ID").Value
    
    '执行科室id
    gstrSQL = GetPublicSQL(SQL.诊疗执行科室)
    If gstrSQL <> "" Then
        Set rs = OpenSQLRecord(gstrSQL, strTitle, lngKey, mlngDept, UserInfo.部门ID, "%%")
        If rs.BOF = False Then
            rsItem("执行科室id").Value = NVL(rs("ID").Value)
        End If
    End If
        
    If rsItem("类别").Value = "C" Then
        '采集方式id
        gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A,诊疗用法用量 B WHERE A.ID=B.用法id AND A.类别='E' AND A.操作类型='6' AND B.项目ID=[1]"
        Set rs = OpenSQLRecord(gstrSQL, strTitle, lngKey)
        If rs.BOF = False Then
            rsItem("采集方式id").Value = NVL(rs("ID").Value)
        Else
            gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A WHERE A.类别='E' AND A.操作类型='6'"
            Set rs = OpenSQLRecord(gstrSQL, strTitle)
            If rs.BOF = False Then
                rsItem("采集方式id").Value = NVL(rs("ID").Value)
            End If
        End If
            
        '采集科室id
        gstrSQL = GetPublicSQL(SQL.诊疗执行科室)
        Set rs = OpenSQLRecord(gstrSQL, strTitle, Val(rsItem("采集方式id").Value), mlngDept, UserInfo.部门ID, "%%")
        If rs.BOF = False Then
            rsItem("采集科室id").Value = NVL(rs("ID").Value)
        End If
        
        
        '检验标本
        gstrSQL = "SELECT 1 FROM 诊疗项目目录 WHERE 组合项目=1 AND ID=[1]"
        Set rs = OpenSQLRecord(gstrSQL, strTitle, lngKey)
        If rs.BOF = False Then
            '是组合项目
            
            gstrSQL = "SELECT DISTINCT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                    "WHERE C.ID<>[1] AND nvl(C.组合项目,0)=0 " & _
                        "AND B.报告项目id=A.项目id and rownum<2"
                        
            gstrSQL = gstrSQL & "AND B.诊疗项目id IN (SELECT C.ID " & _
                         "FROM 检验报告项目 A," & _
                              "(SELECT 报告项目id FROM 检验报告项目 WHERE 诊疗项目id = [1]) B," & _
                              "诊疗项目目录 C,诊治所见项目 D,检验项目 E,检验报告项目 F " & _
                        "WHERE A.报告项目id = B.报告项目id AND A.诊疗项目id <> [1] AND " & _
                              "nvl(C.组合项目,0) = 0 AND A.诊疗项目id = C.ID AND C.ID=F.诊疗项目id AND F.报告项目id=D.ID AND D.ID=E.诊治项目id)  and rownum<2 "
                                      
        Else
            gstrSQL = "SELECT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                    "WHERE C.ID=[1] AND nvl(C.组合项目,0)=0 AND B.诊疗项目id=[1] and B.报告项目id=A.项目id  and rownum<2"
        End If
        
        Set rs = OpenSQLRecord(gstrSQL, strTitle, lngKey)
        If rs.BOF = False Then
            rsItem("检验标本").Value = rs("名称").Value
        Else
            
            '没有对应时，读取所有标本类型
            gstrSQL = "SELECT 名称 FROM 诊疗检验标本 A where rownum<2"
            Set rs = OpenSQLRecord(gstrSQL, strTitle)
            If rs.BOF = False Then
                rsItem("检验标本").Value = rs("名称").Value
            End If
            
        End If
    End If
    
    '价格
        
    strKeys = rsItem("ID").Value & "'" & rsItem("采集方式id").Value & "'0"
    
    gstrSQL = GetPublicSQL(SQL.体检项目价表, strKeys)
    Set rs = OpenSQLRecord(gstrSQL, strTitle)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            
            sglSum = sglSum + NVL(rs("收费数量"), 0) * NVL(rs("现价"), 0)
                        
            '收费执行科室
            If InStr("4,5,6,7", rs("类别").Value) > 0 Then
                gstrSQL = GetPublicSQL(SQL.药品执行科室)
                Set rsTmp = OpenSQLRecord(gstrSQL, strTitle, rs("类别").Value)
            Else
                gstrSQL = GetPublicSQL(SQL.收费执行科室)
                Set rsTmp = OpenSQLRecord(gstrSQL, strTitle, lngKey, mlngDept, UserInfo.部门ID, "%%")
            End If
            
            If rsTmp.BOF = False Then
                lngExecDept = NVL(rsTmp("ID").Value)
            Else
                lngExecDept = rsItem("执行科室id").Value
            End If
            
            strTmp = strTmp & ";" & NVL(rs("ID")) & ":" & NVL(rs("收费数量")) & ":" & NVL(rs("现价")) & ":" & NVL(rs("现价")) & ":" & lngExecDept & ":" & NVL(rs("计价性质"))
            
            
            rs.MoveNext
        Loop
    End If
    
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    rsItem("价格清单").Value = strTmp
    rsItem("基本价格").Value = sglSum
    rsItem("体检价格").Value = sglSum
    
    FinishFillItem = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description
'    Resume
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时清除参数出错
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = ActualLen(varValue)
            If intMax = 0 Or intMax < 10 Then intMax = 10
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = ActualLen(varValue(lngLeft))
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next

    '执行返回记录集
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
    End If
    cmdData.CommandText = strSQL
    
    Set OpenSQLRecord = cmdData.Execute
    
End Function

Public Function ImportData(frmMain As Object, _
                            ByVal strFile As String, _
                            ByVal lng诊疗分类id As Long, _
                            ByVal str诊疗分类编码 As String, _
                            ByVal lng诊治分类id As Long, _
                            ByVal str诊治分类编码 As String, _
                            Optional ByVal lng检验执行科室 As Long) As Boolean
    
    '------------------------------------------------------------------------------------------------------------------
    '功能:导入卫生局提供的体检项目，并进行对码等
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim strSvrCode As String
    Dim lng诊疗项目id As Long
    Dim lng诊治项目id As Long
    Dim lng组合项目id As Long
    Dim lngNo As Long
    Dim lngNo2 As Long
    Dim lngTotal As Long
    Dim lngElementID As Long
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim lngCount As Long
    Dim byt类型 As Byte
    
    ReDim Preserve strSQL(1 To 1)
    
    On Error GoTo errHand
    
    If ConnectAccess(strFile) = False Then Exit Function
    
    gcnOracle.BeginTrans
    
    frmWait.OpenWait frmMain, "转换体检项目"
    
    frmWait.WaitInfo = "正在删除原有数据..."
    
    '检验对应
    gstrSQL = "Delete From 检验报告项目 Where 诊疗项目id In (Select ID From 诊疗项目目录 Where 分类id=" & lng诊疗分类id & ")"
    gcnOracle.Execute gstrSQL
    
    '检查对应
    gstrSQL = "Delete From 病历元素目录 Where 类型=-1 AND 编码='00000'"
    gcnOracle.Execute gstrSQL
    
    gstrSQL = "Delete From 诊治所见项目_干保"
    gcnOracle.Execute gstrSQL
    
    '检查的
    gstrSQL = "Delete From 诊治所见项目 Where 分类id=" & lng诊治分类id
    gcnOracle.Execute gstrSQL
            
    '检验的
    gstrSQL = "Delete From 诊治所见项目 Where 分类id IS NULL AND 编码 In (Select 编码 From 诊疗项目目录 Where 类别='C' AND 分类id=" & lng诊疗分类id & ")"
    gcnOracle.Execute gstrSQL
    
    
    gstrSQL = "Delete From 诊疗项目目录_干保"
    gcnOracle.Execute gstrSQL
    
    '所有
    gstrSQL = "Delete From 诊疗项目目录 Where 分类id=" & lng诊疗分类id
    gcnOracle.Execute gstrSQL
    
        
    frmWait.WaitInfo = "正在打开数据..."
            
    gstrSQL = "Select * From 体检项目枚举值 Order By 体检项目编,序"
    If rsData.State = adStateOpen Then rsData.Close
    rsData.CursorLocation = adUseClient
    rsData.Open gstrSQL, gcnAccess, adOpenStatic
    
    gstrSQL = "Select * From 体检项目_所有 Where 组合编码 Not In ('R80') Order By 组合编码,项目编码"
    
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open gstrSQL, gcnAccess, adOpenStatic
    If rs.BOF = False Then
        
        frmWait.WaitInfo = "正在转换数据..."
        frmWait.ShowProgress = True
        
        lngTotal = rs.RecordCount
        lngLoop = 0
        
        lngElementID = GetNextId("病历元素目录")
        gstrSQL = "ZL_病历元素_INSERT(-1," & lngElementID & ",'00000','体检检查对应','','宋体,9',1,Null,'00001')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        Do While Not rs.EOF
            
            lngLoop = lngLoop + 1
            frmWait.WaitProgress = Format(100 * lngLoop / lngTotal, "0.00")
            
            If strSvrCode <> rs("组合编码").Value Then
                
                lngCount = lngCount + 1
                
                strSvrCode = rs("组合编码").Value
                
                lng组合项目id = GetNextId("诊疗项目目录")
                                
                gstrSQL = "zl_诊疗项目_Insert('"
                gstrSQL = gstrSQL & IIf(UCase(Left(rs("组合编码").Value, 1)) = "R", "C", "D") & "',"
                gstrSQL = gstrSQL & lng诊疗分类id & ","
                gstrSQL = gstrSQL & lng组合项目id & ",'"
                gstrSQL = gstrSQL & (str诊疗分类编码 & Format(lngCount, "0000")) & "','"
                gstrSQL = gstrSQL & rs("组合名称").Value & "','"
                gstrSQL = gstrSQL & zlGetSymbol(rs("组合名称").Value, 0) & "','"                    '别名拼音_IN 诊疗项目别名.简码%TYPE := NULL,
                gstrSQL = gstrSQL & zlGetSymbol(rs("组合名称").Value, 1) & "',"                    '别名五笔_IN 诊疗项目别名.简码%TYPE := NULL,
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,'"
                gstrSQL = gstrSQL & "其他',"
                gstrSQL = gstrSQL & "1,"
                gstrSQL = gstrSQL & "1,"
                gstrSQL = gstrSQL & "3,'"
                gstrSQL = gstrSQL & rs("单位").Value & "',"
                gstrSQL = gstrSQL & "0,"
                gstrSQL = gstrSQL & "0,"
                gstrSQL = gstrSQL & "1,"
                gstrSQL = gstrSQL & "1,"
                gstrSQL = gstrSQL & "'',"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "4,"
                If UCase(Left(rs("组合编码").Value, 1)) = "R" Then
                    gstrSQL = gstrSQL & IIf(lng检验执行科室 = 0, "NULL", lng检验执行科室) & ","         '门诊执行科室
                Else
                    gstrSQL = gstrSQL & "NULL,"             '门诊执行科室
                End If
                gstrSQL = gstrSQL & "NULL,"             '住院执行科室
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "0)"
                
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                
                gstrSQL = "Insert Into 诊疗项目目录_干保(诊疗项目id,干保编码,干保名称,组合科室) Values ("
                gstrSQL = gstrSQL & lng组合项目id & ","
                gstrSQL = gstrSQL & "'" & rs("组合编码").Value & "',"
                gstrSQL = gstrSQL & "'" & rs("组合名称").Value & "',"
                gstrSQL = gstrSQL & "'" & rs("科室编码").Value & "')"
                
                gcnOracle.Execute gstrSQL
                
                lngNo = 0
                lngNo2 = 0
            End If
                        
            If NVL(rs("上限").Value) <> "" Or NVL(rs("下限").Value) <> "" Then
                byt类型 = 0
                '数字
            Else
                '文本
                byt类型 = 1
            End If
            
            If UCase(Left(rs("组合编码").Value, 1)) = "R" Then
                
                '检验项目
                lngCount = lngCount + 1
                
                lng诊疗项目id = GetNextId("诊疗项目目录")
                
                gstrSQL = "Select a.诊治项目id From 诊治所见项目_干保 a Where a.干保编码=[1]"
                Set rsTmp = OpenSQLRecord(gstrSQL, "转换数据", rs("项目编码").Value)
                If rsTmp.BOF Then
                    gstrSQL = "zl_诊疗项目_Insert('"
                    gstrSQL = gstrSQL & "C',"                                                       '类别
                    gstrSQL = gstrSQL & lng诊疗分类id & ","                                         '分类ID
                    gstrSQL = gstrSQL & lng诊疗项目id & ",'"
                    gstrSQL = gstrSQL & (str诊疗分类编码 & Format(lngCount, "0000")) & "','"
                    gstrSQL = gstrSQL & rs("项目名称").Value & "','"                                '名称
                    gstrSQL = gstrSQL & zlGetSymbol(rs("项目名称").Value, 0) & "','"
                    gstrSQL = gstrSQL & zlGetSymbol(rs("项目名称").Value, 1) & "',"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,'"
                    gstrSQL = gstrSQL & "其他',"
                    gstrSQL = gstrSQL & "1,"
                    gstrSQL = gstrSQL & "0,"
                    gstrSQL = gstrSQL & "3,'"
                    gstrSQL = gstrSQL & rs("单位").Value & "',"
                    gstrSQL = gstrSQL & "0,"
                    gstrSQL = gstrSQL & "0,"
                    gstrSQL = gstrSQL & "1,"
                    gstrSQL = gstrSQL & "0,"                '组合项目
                    gstrSQL = gstrSQL & "'',"               '标本部位
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "4,"                '执行科室
                    gstrSQL = gstrSQL & IIf(lng检验执行科室 = 0, "NULL", lng检验执行科室) & ","         '门诊执行科室
                    gstrSQL = gstrSQL & "NULL,"             '住院执行科室
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "0)"
                    
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    
                    lng诊治项目id = GetNextId("诊治所见项目")
    
                    gstrSQL = "ZL_所见项目_INSERT("
                    gstrSQL = gstrSQL & lng诊治项目id & ","
                    gstrSQL = gstrSQL & "NULL,'"
                    gstrSQL = gstrSQL & (str诊疗分类编码 & Format(lngCount, "0000")) & "','"
                    gstrSQL = gstrSQL & rs("项目名称").Value & "',"                             '中文名
                    gstrSQL = gstrSQL & "NULL,"                                                 '英文名
                    gstrSQL = gstrSQL & byt类型 & ","                                                    '类型
                    gstrSQL = gstrSQL & "50,"
                    gstrSQL = gstrSQL & "0,'"
                    gstrSQL = gstrSQL & rs("单位").Value & "',"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "0,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL)"
    
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    
                    gstrSQL = "ZL_检验项目_UPDATE("
                    gstrSQL = gstrSQL & lng诊治项目id & ","
                    gstrSQL = gstrSQL & "NULL,"                     '缩写
                    gstrSQL = gstrSQL & "NULL,"                     '报告代号
                    gstrSQL = gstrSQL & "1,"                        '项目类别
                    gstrSQL = gstrSQL & IIf(byt类型 = 0, 1, 2) & ",'"                   '结果类型
                    gstrSQL = gstrSQL & rs("单位").Value & "',"     '单位
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & "NULL)"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    
                    gstrSQL = "ZL_检验报告项目_UPDATE(" & lng诊疗项目id & ",'^" & lng诊治项目id & "')"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                                        
                    gstrSQL = "Insert Into 诊治所见项目_干保(诊治项目id,干保编码,干保名称,项目分支,项目方法) VALUES ("
                    gstrSQL = gstrSQL & lng诊治项目id & ","
                    gstrSQL = gstrSQL & "'" & rs("项目编码").Value & "',"
                    gstrSQL = gstrSQL & "'" & rs("项目名称").Value & "',"
                    gstrSQL = gstrSQL & "'',"
                    gstrSQL = gstrSQL & "'')"
                    gcnOracle.Execute gstrSQL
                    
                    gstrSQL = "ZL_检验项目取值_DELETE(" & lng诊治项目id & ")"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    
                    '查找可选数据,有可选数据
                    rsData.Filter = ""
                    rsData.Filter = "体检项目编='" & NVL(rs("项目编码").Value) & "'"
                    If rsData.RecordCount > 0 Then
                        rsData.MoveFirst
                        Do While Not rsData.EOF
                            
                            gstrSQL = "ZL_检验项目取值_INSERT(" & lng诊治项目id & ",'" & NVL(rsData("序").Value) & "','" & NVL(rsData("枚举值").Value) & "',0)"
                            gcnOracle.Execute gstrSQL, , adCmdStoredProc
                            
                            rsData.MoveNext
                        Loop
                    End If
                    
                    '检查上限下限
                    If byt类型 = 0 Then
                        
                        gstrSQL = "ZL_检验项目参考_DELETE(" & lng诊治项目id & ")"
                        gcnOracle.Execute gstrSQL, , adCmdStoredProc
                        
                        gstrSQL = "ZL_检验项目参考_INSERT(" & lng诊治项目id & ",'',0,NULL,NULL,NULL," & Val(NVL(rs("上限").Value)) & "," & Val(NVL(rs("下限").Value)) & ",'')"
                        gcnOracle.Execute gstrSQL, , adCmdStoredProc
                        
                    End If
                Else
                    lng诊治项目id = rsTmp("诊治项目id").Value
                End If
                
                lngNo = lngNo + 1
                gstrSQL = "insert into 检验报告项目(诊疗项目ID,检验标本,报告项目ID,排列序号) values (" & lng组合项目id & ",NULL," & lng诊治项目id & "," & lngNo & ")"
                gcnOracle.Execute gstrSQL
                
            Else
                
                '填写所见项目
                
                lngCount = lngCount + 1
                
                gstrSQL = "Select a.诊治项目id From 诊治所见项目_干保 a Where a.干保编码=[1]"
                Set rsTmp = OpenSQLRecord(gstrSQL, "转换数据", rs("项目编码").Value)
                If rsTmp.BOF Then
                    
                    strTmp = ""
                    
                    '查找可选数据,有可选数据，就定义为下拉选择框;否则定为文本输入框
                    If byt类型 = 1 Then
                        rsData.Filter = ""
                        rsData.Filter = "体检项目编='" & NVL(rs("项目编码").Value) & "'"
                        If rsData.RecordCount > 0 Then
                            rsData.MoveFirst
                            Do While Not rsData.EOF
                                strTmp = strTmp & ";" & rsData("枚举值").Value
                                rsData.MoveNext
                            Loop
                            If strTmp <> "" Then strTmp = Mid(strTmp, 2)
                        End If
                    Else
                        strTmp = Val(NVL(rs("下限").Value)) & ";" & Val(NVL(rs("上限").Value))
                    End If
                    
                    lng诊治项目id = GetNextId("诊治所见项目")
                    
                    gstrSQL = "ZL_所见项目_INSERT("
                    gstrSQL = gstrSQL & lng诊治项目id & ","
                    gstrSQL = gstrSQL & lng诊治分类id & ",'"
                    gstrSQL = gstrSQL & (str诊治分类编码 & Format(lngCount, "0000")) & "','"
                    gstrSQL = gstrSQL & rs("项目名称").Value & "',"                             '中文名
                    gstrSQL = gstrSQL & "NULL,"                                                 '英文名
                    gstrSQL = gstrSQL & byt类型 & ","                                           '类型
                    gstrSQL = gstrSQL & "50,"
                    gstrSQL = gstrSQL & "0,'"
                    gstrSQL = gstrSQL & rs("单位").Value & "',"
                    gstrSQL = gstrSQL & "NULL,"
                    gstrSQL = gstrSQL & IIf(strTmp <> "" And byt类型 = 1, "2", "0") & ","                       '表示法
                    gstrSQL = gstrSQL & "NULL,'"
                    gstrSQL = gstrSQL & strTmp & "',"                                           '数值域
                    gstrSQL = gstrSQL & "NULL,"                                                 '初始值
                    gstrSQL = gstrSQL & "NULL,"                                                 '文字表述
                    gstrSQL = gstrSQL & "NULL)"                                                 '空值文字
    
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
                    gstrSQL = "Insert Into 诊治所见项目_干保(诊治项目id,干保编码,干保名称,项目分支,项目方法) VALUES ("
                    gstrSQL = gstrSQL & lng诊治项目id & ","
                    gstrSQL = gstrSQL & "'" & rs("项目编码").Value & "',"
                    gstrSQL = gstrSQL & "'" & rs("项目名称").Value & "',"
                    gstrSQL = gstrSQL & "'',"
                    gstrSQL = gstrSQL & "'')"
                    gcnOracle.Execute gstrSQL
                Else
                    lng诊治项目id = rsTmp("诊治项目id").Value
                End If
                
                lngNo2 = lngNo2 + 1
  
                gstrSQL = "ZL_所见单_SAVE("
                gstrSQL = gstrSQL & lngElementID & ","
                gstrSQL = gstrSQL & lngNo2 & ","
                gstrSQL = gstrSQL & "'2',"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & lng组合项目id & ","
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & lng诊治项目id & ","
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "'" & NVL(rs("单位").Value) & "',"                  '单位
                gstrSQL = gstrSQL & "NULL)"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                
            End If
            
            rs.MoveNext
        Loop
            
    End If
    
    frmWait.ShowProgress = False
    frmWait.WaitInfo = "正在保存数据..."
    
        
    gcnOracle.CommitTrans
    
    frmWait.CloseWait
    
    ImportData = True
    
    Exit Function
    
errHand:
    Dim strError As String
    
    strError = Err.Description
    frmWait.CloseWait
    gcnOracle.RollbackTrans
    ShowSimpleMsg strError
    
'    Resume
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode('" & strInput & "') from dual"
    Else
        strSQL = "select zlSpellcode('" & strInput & "') from dual"
    End If
    On Error GoTo errHand
    With rsTmp
        If .State = adStateOpen Then .Close
        
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

errHand:
    
    zlGetSymbol = "-"
End Function




