Attribute VB_Name = "mdlVsGridProperty"
Option Explicit
Public Const GRD_GOTFOCUS_COLORSEL = &H8000000D '16772055 '    '进入控件时,选择显示颜色
Public Const GRD_LOSTFOCUS_COLORSEL = &HE0E0E0  '&H80000010  '离开焦点时,选择的显示颜色
Public Enum mTextType
    m文本式 = 0
    m数字式 = 1
    m金额式 = 2
    m负金额式 = 3
End Enum
Public Function GetVsGridBoolColVal(ByVal vsGrid As VSFlexGrid, lngRow As Long, lngCol As Long) As Boolean
    '------------------------------------------------------------------------------
    '功能:获取bool列的值
    '返回:是该单元格为true,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/28
    '------------------------------------------------------------------------------
    
    GetVsGridBoolColVal = grid.BoolVal(vsGrid, lngRow, lngCol)
    
End Function

Public Sub VsFlxGridCheckKeyPress(ByVal objCtl As Object, Row As Long, Col As Long, KeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '功能:只能输入数字和回车及退格
    '参数:
    '   objctl:Vsgrid8.0控件
    '   Keyascii:
    '           Keyascii:8 (退格)
    '   Row-当前行
    '   Col-当前列
    '   TextType:(0-文本式;1-数字式;2-金额式)
    '返回:一个KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Call grid.CheckKeyPress(objCtl, Row, Col, KeyAscii, TextType)
    
End Sub


Public Function zl_VsGridAfterSort(ByVal vsGrid As VSFlexGrid, ByVal intCol As Integer, ByVal intOrder As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:排序后调用此事件(主要是处理行的背景色)
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-28 11:26:52
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        .Redraw = flexRDNone
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = .BackColor
        .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = 16772055
        .Redraw = flexRDBuffered
    End With
    zl_VsGridAfterSort = True
End Function

Public Sub zlVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng主例 As Long = -1, Optional lng尾列 As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef lngRow As Long = -1)
    '-----------------------------------------------------------------------------------------------------------
    '功能:移动单元格的列
    '入参:blnEdit-当前正处于编辑状态,允许新增行
    '     lng主例-主列,如果<0,则主列为0列,否则为指定的列
    '     lng尾列-尾列,如果<0,则主列为.cols-1,否则为指定的列
    '出参:lngRow-如果存在插入行,则返回被插入的行号,否则返回-1
    '返回:
    '编制:刘兴洪
    '日期:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------
    Dim lngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    
    'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
    If lng主例 <> -1 Then
        lngCol = lng主例
    Else
        lngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If lngCol = -1 Then lngCol = 0
    lngLastCol = IIf(lng尾列 < 0, vsGrid.Cols - 1, lng尾列)
    lngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = lngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        lngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                arrSplit = Split(.ColData(i) & "||", "||")
                If .ColHidden(i) Or Val(arrSplit(1)) >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = lngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    lngRow = .Row
                                End If
                            End If
                            .Col = lngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.Width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
Errhand:
End Sub
Public Function zlVsInsertIntoRow(ByVal vsGrid As VSFlexGrid, ByVal lngRow As Long, Optional blnBefor As Boolean = False, _
    Optional blnMoveNewRow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '功能:插入行
    '参数:vsGrid-插入行的网格格件
    '     lngRow-当前行
    '     blnBefor-在lngrow之间或之后.true:之间,false-之后
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intCol As Integer
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        If blnBefor Then
            .AddItem "", lngRow
            For intCol = 0 To .Cols - 1
                .Cell(flexcpBackColor, .Rows - 1, intCol, .Rows - 1, intCol) = .Cell(flexcpBackColor, 1, intCol, 1, intCol)
            Next
        Else
            .AddItem "", lngRow + 1
            For intCol = 0 To .Cols - 1
                .Cell(flexcpBackColor, .Rows - 1, intCol, .Rows - 1, intCol) = .Cell(flexcpBackColor, 1, intCol, 1, intCol)
            Next
        End If
        If blnMoveNewRow = True Then
            If blnBefor Then '
                .Row = lngRow
            Else
                .Row = lngRow + 1
            End If
        End If
    End With
    zlVsInsertIntoRow = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
'*********************************************************************************************************************
'**处理网格控件
Public Sub zl_VsGridGotFocus(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：进入网格控件时选择的颜色
    '入参：CustomColor-自定颜色
    '编制：刘兴洪
    '日期：2010-03-23 10:52:23
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    '进入控件
    With vsGrid
         If CustomColor <> -1 Then
             .FocusRect = flexFocusSolid
             .HighLight = flexHighlightNever
             If .Row >= .FixedRows Then
                If .Rows - 1 > .FixedRows Then  '清除选择颜色
                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
                End If
                 .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
             End If
         Else
            .FocusRect = flexFocusSolid 'IIf(vsGrid.Editable = flexEDNone, flexFocusNone, flexFocusSolid)
            .HighLight = flexHighlightNever
            .BackColorSel = GRD_GOTFOCUS_COLORSEL
        End If
    End With
    Call zl_VsGridRowChange(vsGrid, vsGrid.Row, vsGrid.Row, 0, 0)
End Sub
Public Sub zl_VsGridLOSTFOCUS(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1, Optional ForeColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
   '功能：离开网格控件时选择的颜色
    '入参：CustomColor-是否用自定义颜色来设置(BackColor)的方式来进行)
    '编制：刘兴洪
    '日期：2010-03-23 11:03:05
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With vsGrid
        If CustomColor <> -1 Then
            If .Row >= .FixedRows Then
                .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
            End If
        Else
            .SelectionMode = flexSelectionByRow
            .FocusRect = IIf(vsGrid.Editable = flexEDNone, flexFocusHeavy, flexFocusSolid)
            If ForeColor = -1 Then .HighLight = flexHighlightAlways
            .BackColorSel = GRD_LOSTFOCUS_COLORSEL
        End If
        If ForeColor <> -1 Then
            .Cell(flexcpForeColor, .Row, .FixedCols, .Row, .Cols - 1) = ForeColor
        End If
        .ForeColorSel = .ForeColor
    End With
End Sub
Public Sub zl_VsGridRowChange(ByVal vsGrid As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngNewRow As Long, _
    ByVal lngoldCol As Long, ByVal lngNewCol As Long, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：行列改变时,设置相关的颜色
    '入参：CustomColor-自定义颜色
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-03-23 11:22:38
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    '行改变时
    Err = 0: On Error Resume Next
    If lngOldRow = lngNewRow Then
        vsGrid.Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, vsGrid.Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
        Exit Sub
    End If
    With vsGrid
        .Cell(flexcpBackColor, lngOldRow, vsGrid.FixedCols, lngOldRow, .Cols - 1) = .BackColor
        .Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, .Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
    End With
End Sub

'排序处理
Public Sub zl_VsGridBeforeSort(ByVal vsGrid As VSFlexGrid, ByRef Col As Long, ByRef Order As Integer, Optional strSpaceRowNotCheckCol As String = "")
    '-----------------------------------------------------------------------------------------------------------
    '功能:处理排序(排序时,不包含空白行)
    '入参:strSpaceRowNotCheckCol-不检查空行中的哪些列(列1,列2...)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-07-25 11:38:23
    '-----------------------------------------------------------------------------------------------------------
    Dim lngStartRow As Long, lngEndRow As Long, lngStartCol As Long, lngEndCol As Long
    Dim lngRow As Long, lngCol As Long
    Dim blnAllowSelect As Boolean, blnAllowBigSel As Boolean
    Dim lngOldBackColor As Long

    If vsGrid.ExplorerBar > &H1000& Then Exit Sub
    '保存当前的选择区域
    vsGrid.GetSelection lngStartRow, lngStartCol, lngEndRow, lngEndCol
    vsGrid.Redraw = flexRDNone
    blnAllowBigSel = vsGrid.AllowBigSelection: blnAllowSelect = vsGrid.AllowSelection
    
    '不排序空白行
    With vsGrid
        For lngRow = .Rows - 1 To .FixedRows Step -1
            For lngCol = 0 To .Cols - 1
               If InStr(1, "," & strSpaceRowNotCheckCol & ",", "," & lngCol & ",") > 0 Then
               Else
                    If Trim(.TextMatrix(lngRow, lngCol)) <> "" Then GoTo GoNext:
               End If
            Next
        Next
GoNext:
        If lngRow > .FixedRows Then
            
             .Select .FixedRows, Col, lngRow, Col
            .Sort = Order
        End If
        ' 恢复以前选择的区域
        .Select lngStartRow, lngStartCol, lngEndRow, lngEndCol
            
        .Redraw = flexRDDirect
    End With
    Order = 0
End Sub


Public Function zl_vsGrid_Para_Save(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional bln强制保存 As Boolean = False, Optional blnHaveParaPrivs As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '功能:保存vsFlex的宽度到注册表
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     strKey-主建
    '返回:保存成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    If blnSaveToDataBase = False Then
        zl_vsGrid_Para_Save = True
        If bln强制保存 = False Then
            If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
        End If
    End If
    zl_vsGrid_Para_Save = False
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    If blnSaveToDataBase Then
        zlDatabase.SetPara strKey, strCol, glngSys, lngModule, blnHaveParaPrivs
    Else
        Call SaveRegInFor(g私有模块, strCaption, strKey, strCol)
    End If
    zl_vsGrid_Para_Save = True
End Function

Public Function zl_vsGrid_Para_Restore(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String, _
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
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, arrtemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    If blnSaveToDataBase = False Then
        '只有在本地注册表中才会处理个性化设置
        zl_vsGrid_Para_Restore = True
        If bln强制恢复保存 = False Then
            If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g私有模块, strCaption, strKey, strParaValue)
    Else
        strParaValue = zlDatabase.GetPara(strKey, glngSys, lngModule)
    End If
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    Err = 0: On Error GoTo Errhand:
    arrReg = Split(strParaValue, "|")
    If vsGrid.Cols <> UBound(arrReg) + 1 Then Exit Function
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrtemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrtemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(arrtemp(1))
                If Val(arrtemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
Errhand:
End Function

Public Function zl_vsGrid_GetCols_Property(ByVal vsGrid As VSFlexGrid) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取列头宽度
    '参数:vsGrid-对应的网络控件
    '返回:返回列头信息,格式为:列主键,列宽,列隐藏|列主键,列宽,列隐藏|....
    '编制:刘兴洪
    '日期:2014-10-09 12:08:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    zl_vsGrid_GetCols_Property = strCol
End Function

Public Sub zl_vsGrid_RestoreCols_Property(ByVal vsGrid As VSFlexGrid, ByVal strColsInfor As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置列宽
    '参数:vsGrid-对应的网络控件
    '     strColsInfor-列信息:列主键,列宽,列隐藏|列主键,列宽,列隐藏|....
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-10-09 12:34:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCols As Integer, arrReg As Variant, arrtemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    If strColsInfor = "" Then Exit Sub
    Err = 0: On Error GoTo Errhand:
    arrReg = Split(strColsInfor, "|")
    
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrtemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrtemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(arrtemp(1))
                If Val(arrtemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    Exit Sub
Errhand:
End Sub

