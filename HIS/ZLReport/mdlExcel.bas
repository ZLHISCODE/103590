Attribute VB_Name = "mdlExcel"
Option Explicit
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
'表附加项目描述
Public gobjTitle As RPTItem '报表标题
Public gcolUpItem As RPTItems  '表上项集合,以行顺序存放,一个元素为一行
Public gcolDownItem As RPTItems '表下项集合,以行顺序存放,一个元素为一行c

Public gobjExcel As Object 'Excel对象
Public gobjHead As Object '要输出的任意表的表头
Public gobjBody As Object '要输出的任意表或分类表的表体
Public gblnIsWps As Boolean     '刘兴宏:2007/04/20加入,主要确定是否向wps中传数据

Private W_Excel As Integer

Private Const H_Excel = 20
Private Const xlNormal = -4143

Private lngGridW As Long
Private sngPer As Single, lngTotal As Long
Private frmParent As Object


Public Function HaveExcel() As Boolean
'功能：判断系统是否安装了Excel
'说明：同时初始化Excel对象
    On Error Resume Next
    Err.Clear
    gblnIsWps = False
    Set gobjExcel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then Err.Clear: GoTo GoWps:
    HaveExcel = True
    Exit Function
GoWps:
   '刘兴宏:2007/4/20
    '以WPS为准
    Err = 0: On Error GoTo ErrHand:
    Set gobjExcel = CreateObject("ET.Application") '打开一个WPS中的ET程序
    gblnIsWps = True
    HaveExcel = True
    Exit Function
ErrHand:
    HaveExcel = False
End Function

Public Function isExporting() As Boolean
'功能：根据Excel对象判断是否已经有一个报表在正在输出
    isExporting = Not (gobjExcel Is Nothing)
End Function

Private Function SpecSort(str As String) As String
'功能：对形如"X1,ID1|X2,ID2|..."的字符串按X排序,并返回相同格式的字符串
    Dim arrStr() As String, strTmp As String
    Dim i As Long, j As Long
    
    arrStr = Split(str, "|")
    
    For i = 0 To UBound(arrStr) - 1
        For j = i + 1 To UBound(arrStr)
            If arrStr(j) < arrStr(i) Then
                strTmp = arrStr(i)
                arrStr(i) = arrStr(j)
                arrStr(j) = strTmp
            End If
        Next
    Next
    For i = 0 To UBound(arrStr)
        SpecSort = SpecSort & "|" & arrStr(i)
    Next
    SpecSort = Mid(SpecSort, 2)
End Function

Public Function Col2Excel(ByVal lngCol As Long) As String
'功能：将数字列号转换为EXCEL中的表示方法
    Dim lng1 As Long, lng2 As Long, lng3 As Long
    
    If lngCol <= 26 Then
        Col2Excel = Chr(Asc("A") + lngCol - 1)
    ElseIf lngCol <= 26 ^ 2 + 26 Then
        lng1 = lngCol \ 26
        lng2 = lngCol Mod 26
        If lng2 = 0 Then lng1 = lng1 - 1: lng2 = 26
        Col2Excel = Chr(Asc("A") + lng1 - 1) & Chr(Asc("A") + lng2 - 1)
    ElseIf lngCol <= 26 ^ 3 + 26 ^ 2 + 26 Then
        'A-Z=0-25
        lngCol = lngCol - (26 ^ 2 + 26) - 1
        
        lng1 = Int(lngCol / 26 ^ 2)
        lng2 = Int((lngCol Mod 26 ^ 2) / 26) 'lng2 = Int((lngCol - lng1 * 26 ^ 2) / 26)
        lng3 = (lngCol Mod 26 ^ 2) Mod 26 'lng3 = lngCol - lng1 * 26 ^ 2 - lng2 * 26

        Col2Excel = Chr(Asc("A") + lng1) & Chr(Asc("A") + lng2) & Chr(Asc("A") + lng3)
    End If
End Function

Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Private Function GetShowCol(lngCol As Long, objGrid As Object) As Long
'功能：根据表格列号获取可见列号
    Dim i As Long, j As Long
    
    For i = 0 To lngCol - 1
       If objGrid.ColWidth(i) = 0 Then j = j + 1
    Next
    GetShowCol = lngCol - j
End Function

Private Function GetstrWidth(strFontName As String, curFontSize As Currency, str As String) As Long
'功能：根据字体设置返回一个字符串宽度
    Dim objFrm As frmFlash
    Set objFrm = New frmFlash
    objFrm.Font.name = strFontName
    objFrm.Font.Size = curFontSize
    GetstrWidth = objFrm.TextWidth(str)
    Unload objFrm
    Set objFrm = Nothing
End Function

Private Function GetstrHeight(strFontName As String, curFontSize As Currency, str As String) As Long
'功能：根据字体设置返回一个字符串高度
    Dim objFrm As frmFlash
    Set objFrm = New frmFlash
    objFrm.Font.name = strFontName
    objFrm.Font.Size = curFontSize
    GetstrHeight = objFrm.TextHeight(str)
    Unload objFrm
    Set objFrm = Nothing
End Function

Private Function GrdAlignment(objGrid As Object) As Long
'功能：将FlexGrid的对齐方式转换为EXCEL中的对齐方式
'参数：objGrid=FlexGrid对象
    Dim Alignment As Integer
        
    '获取对齐属性：
    If objGrid.CellAlignment <> 0 Then
        Alignment = objGrid.CellAlignment           '参照单元
    Else
        If objGrid.Col < objGrid.FixedCols Or objGrid.Row < objGrid.FixedRows Then
            Alignment = objGrid.FixedAlignment(objGrid.Col) '参照固定单元
        Else
            Alignment = objGrid.ColAlignment(objGrid.Col)   '参照列
        End If
    End If
    Select Case Alignment
    Case 0, 1, 2        '左对齐
        GrdAlignment = -4131
    Case 3, 4, 5        '居中
        GrdAlignment = -4108
    Case 6, 7, 8        '右对齐
        GrdAlignment = -4152
    Case 9
        If IsNumeric(Trim(objGrid.Text)) Then
            If IsNumberEx(objGrid.Text) Then
                GrdAlignment = -4152
            Else
                GrdAlignment = -4131
            End If
        Else
            GrdAlignment = -4131
        End If
    Case Else
        GrdAlignment = -4131
    End Select
End Function

Private Sub ExcelMerge(objGrid As Object, lngRow As Long, ByVal LeftCol As Long, ByVal RightCol As Long, ByVal TopRow As Long, ByVal BottomRow As Long)
'功能：按照表格的合并格式输出到Excel中
'参数：objGrid=表格，可为表头或表体
    Dim i As Long, j As Long
    Dim lngY As Long, strTmp As String
    Dim lngTmp As Long, lngRowUp As Long
    Dim blnMerge As Boolean
    
    If objGrid.MergeCells = 0 Then Exit Sub
    
    '合并单元格
    '横向合并
    lngY = lngRow
    For i = TopRow To BottomRow - 1
        If objGrid.MergeRow(i) Then
            For j = LeftCol To RightCol - 2
                If objGrid.TextMatrix(i, j) <> "" Then '空白区域就不合并
                    strTmp = Col2Excel(GetShowCol(j, objGrid) + 1) & Trim(str(lngY))
                    If Not gobjExcel.Range(strTmp).MergeCells Then   '已合并了
                        blnMerge = False
                        For lngTmp = j + 1 To RightCol - 1
                            '不知第一个单元相同就退出
                            If objGrid.TextMatrix(i, j) <> objGrid.TextMatrix(i, lngTmp) Then Exit For
                            If objGrid.MergeCells = 3 Or objGrid.MergeCells = 4 Then  '有列限制
                                lngRowUp = i - 1
                                Do While lngRowUp >= TopRow
                                    '上面没合并就退出
                                    If objGrid.TextMatrix(lngRowUp, j) <> objGrid.TextMatrix(lngRowUp, lngTmp) Then Exit For
                                    lngRowUp = lngRowUp - 1
                                Loop
                            End If
                            blnMerge = True
                        Next
                        If blnMerge Then
                            strTmp = Col2Excel(GetShowCol(j, objGrid) + 1) & Trim(str(lngY)) & ":" & Col2Excel(GetShowCol(lngTmp, objGrid)) & Trim(str(lngY))
                            gobjExcel.Range(strTmp).MergeCells = True
                            j = lngTmp - 1 '跳过已合并的列
                        End If
                    End If
                End If
            Next
         End If
        lngY = lngY + 1
        
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    Next
    
    '纵向合并
    lngY = lngRow
    For j = LeftCol To RightCol - 1
        If objGrid.MergeCol(j) Then
            For i = TopRow To BottomRow - 2
                If objGrid.TextMatrix(i, j) <> "" Then '空白区域就不合并
                    strTmp = Col2Excel(GetShowCol(j, objGrid) + 1) & Trim(str(lngY + i - TopRow))
                    If Not gobjExcel.Range(strTmp).MergeCells Then   '已合并了
                        blnMerge = False
                        For lngTmp = i + 1 To BottomRow - 1
                            '不知第一个单元相同就退出
                            If objGrid.TextMatrix(i, j) <> objGrid.TextMatrix(lngTmp, j) Then Exit For
                            If objGrid.MergeCells = 2 Or objGrid.MergeCells = 4 Then  '有行限制
                                lngRowUp = j - 1
                                Do While lngRowUp >= LeftCol
                                    '左面没合并就退出
                                    If objGrid.TextMatrix(i, lngRowUp) <> objGrid.TextMatrix(lngTmp, lngRowUp) Then Exit For
                                    lngRowUp = lngRowUp - 1
                                Loop
                            End If
                            blnMerge = True
                        Next
                        If blnMerge Then
                            strTmp = Col2Excel(GetShowCol(j, objGrid) + 1) & Trim(str(lngY + i - TopRow)) & ":" & Col2Excel(GetShowCol(j, objGrid) + 1) & Trim(str(lngY + lngTmp - 1 - TopRow))
                            gobjExcel.Range(strTmp).MergeCells = True
                            i = lngTmp - 1 '跳过已合并的列
                        End If
                    End If
                End If
            Next
         End If
    
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    Next
End Sub

Private Function ConverItem(strItem As String, strFontName As String, curFontSize As Currency) As String
'功能：处理表上、下项目
    Dim i As Long, strTmp As String, lngLen As Long
    
    If strItem = "" Then Exit Function
    If InStr(strItem, "|") = 0 Then ConverItem = strItem: Exit Function
    
    lngLen = GetstrWidth(strFontName, curFontSize, Replace(strItem, "|", ""))
    lngLen = (lngGridW - lngLen) * (0.1188 / 15)
    If lngLen < 0 Then lngLen = 0
        
    For i = 0 To UBound(Split(strItem, "|"))
        If i = 0 Then
            strTmp = strTmp & Split(strItem, "|")(i)
        Else
            strTmp = strTmp & String(lngLen / UBound(Split(strItem, "|")), " ") & Split(strItem, "|")(i)
        End If
    Next
    ConverItem = strTmp
End Function

Private Function GridToTXT(arrFormat As Variant) As String
'功能：将报表内容输出到一个临时文本文件中
'参数：In/Out列数据格式设置
'返回：所产生的临时文件
    Dim strFile As String, strPath As String, intFile As Integer
    Dim tmpItem As RPTItem, strText As String
    Dim i As Long, j As Long
    
    '产生临时文件
    strPath = Space(256): strFile = Space(256)
    Call GetTempPath(256, strPath)
    strPath = Left(strPath, InStr(strPath, Chr(0)) - 1)
    
    Call GetTempFileName(strPath, "Excel", 0, strFile)
    strFile = Left(strFile, InStr(strFile, Chr(0)) - 1)
    
    intFile = FreeFile()
    Open strFile For Binary Access Write As intFile
    
    '标题输出
    Put intFile, , Replace(gobjTitle.内容, vbCrLf, "")
    Put intFile, , vbCrLf
    
    sngPer = sngPer + 1
    Call ShowFlash("正在输出到 Excel：处理数据 ...", sngPer / lngTotal, frmParent, True)
    
    '表上项目输出
    For Each tmpItem In gcolUpItem
        Put intFile, , ConverItem(tmpItem.内容, tmpItem.字体, tmpItem.字号)
        Put intFile, , vbCrLf
        
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理数据 ...", sngPer / lngTotal, frmParent, True)
    Next
   
   '表头输出
   If gobjHead Is Nothing Then
        For i = 0 To gobjBody.FixedRows - 1
            If gobjBody.RowHeight(i) <> 0 Then
                For j = 0 To gobjBody.Cols - 1
                    If gobjBody.ColWidth(j) <> 0 Then
                        Put intFile, , Replace(gobjBody.TextMatrix(i, j), "<换行分隔符>", "")
                        Put intFile, , vbTab
                    End If
                Next
                Put intFile, , vbCrLf
            End If
        
            sngPer = sngPer + 1
            Call ShowFlash("正在输出到 Excel：处理数据 ...", sngPer / lngTotal, frmParent, True)
        Next
    Else
        For i = 0 To gobjHead.FixedRows - 1
            If gobjHead.RowHeight(i) <> 0 Then
                For j = 0 To gobjHead.Cols - 1
                    If gobjHead.ColWidth(j) <> 0 Then
                        Put intFile, , Replace(gobjHead.TextMatrix(i, j), "<换行分隔符>", "")
                        Put intFile, , vbTab
                    End If
                Next
                Put intFile, , vbCrLf
            End If
            
            sngPer = sngPer + 1
            Call ShowFlash("正在输出到 Excel：处理数据 ...", sngPer / lngTotal, frmParent, True)
        Next
    End If
         
     '网格内容输出
    For i = gobjBody.FixedRows To gobjBody.Rows - 1
        If gobjBody.RowHeight(i) <> 0 Then
            For j = 0 To gobjBody.Cols - 1
                If gobjBody.ColWidth(j) <> 0 Then
                    strText = Replace(gobjBody.TextMatrix(i, j), "<换行分隔符>", "")
                    
                     '特殊格式的数字
                    If IsNumeric(strText) Then
                        If Len(strText) > 15 Then
                            arrFormat(j, 1) = 2 '太长的数字强行处理为文本格式
                        ElseIf IsNumberEx(strText) = False Then
                            arrFormat(j, 1) = 2
                        End If
                    End If
                    
                    '写入文件
                    Put intFile, , strText
                    Put intFile, , vbTab
                    
                    '侦测类型：只要有一个不是数字或长度大于15，则为文本
                    strText = Trim(strText)
                    If strText <> "" And (Not IsNumeric(strText) Or Len(strText) > 15) Then
                        arrFormat(j, 1) = 2
                    End If
                End If
            Next
            Put intFile, , vbCrLf
        End If
        
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理数据 ...", sngPer / lngTotal, frmParent, True)
    Next
    
    '表下项目输出
    For Each tmpItem In gcolDownItem
        Put intFile, , ConverItem(tmpItem.内容, tmpItem.字体, tmpItem.字号)
        Put intFile, , vbCrLf
        
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理数据 ...", sngPer / lngTotal, frmParent, True)
    Next
    
    Close intFile
    GridToTXT = strFile
End Function

Public Function ExportExcel(ByVal frmMain As Object, Optional ByVal strExcelFile As String) As Boolean
'功能：导出报表内容到Excel
'参数：本模块定义的公共参数
'      strExcelFile=指定的输出文件
    Dim i As Long, j As Long, strTmp As String
    Dim lngZeroCol As Long, strFile As String
    Dim arrFormat As Variant
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    On Error GoTo errH
    
    sngPer = 0
    lngTotal = gobjBody.Cols + gobjBody.Rows * 2 + gcolUpItem.count * 2 + gcolDownItem.count * 2 + 26
    If Not gobjHead Is Nothing Then lngTotal = lngTotal + gobjHead.FixedRows * 2
    Set frmParent = frmMain
    '得出1磅多少个缇
    W_Excel = 120
    
    Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    
    gobjExcel.Workbooks.Add
    
    '设置列宽
    lngGridW = 0
    For i = 0 To gobjBody.Cols - 1
        If gobjBody.ColWidth(i) <> 0 Then
            lngGridW = lngGridW + gobjBody.ColWidth(i)
            If gobjBody.ColWidth(i) / W_Excel > 0 Then
                On Error Resume Next
                gobjExcel.Columns(Col2Excel(i - lngZeroCol + 1) & ":" & Col2Excel(i - lngZeroCol + 1)).ColumnWidth = gobjBody.ColWidth(i) / W_Excel
                If Err.Number <> 0 Then
                    If Err.Number = 13 Then
                        Err.Clear: On Error GoTo 0
                        Set gobjExcel = Nothing: Call ShowFlash
                        MsgBox "输出表格列内容太多，在本机资源情况下无法输出。", vbInformation, App.ProductName
                        Exit Function
                    Else
                        GoTo errH
                    End If
                Else
                    On Error GoTo errH
                End If
            End If
        Else
            lngZeroCol = lngZeroCol + 1
        End If
    Next
    
    sngPer = sngPer + 1
    Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    
    '设置列数据格式、对齐
    lngZeroCol = 0
    For i = 0 To gobjBody.Cols - 1
        gobjBody.Col = i
        If gobjBody.ColWidth(i) = 0 Then
            lngZeroCol = lngZeroCol + 1
        Else
            For j = gobjBody.FixedRows To gobjBody.Rows - 1
                gobjBody.Row = j
                If Trim(gobjBody.Text) <> "" Then Exit For '不是空白就有格式
            Next
            strTmp = Col2Excel(i - lngZeroCol + 1)
            strTmp = strTmp & ":" & strTmp
            With gobjExcel.Columns(strTmp)
                If IsNumeric(gobjBody.Text) _
                    And Not Len(gobjBody.Text) > 15 _
                    And IsNumberEx(gobjBody.Text) Then
                    'And Not (Left(gobjBody.Text, 1) = "0" And InStr(gobjBody.Text, ".") = 0) Then
                    j = InStr(gobjBody.Text, ".")
                    If j = 0 Then
                        .NumberFormatLocal = "0_ "
                    Else
                        .NumberFormatLocal = "0." & String(Len(Mid(gobjBody.Text, j + 1)), "0") & "_ "
                    End If
                ElseIf IsDate(Replace(gobjBody.Text, "○", "字")) Then
                    If InStr(gobjBody.Text, ":") > 0 Or InStr(gobjBody.Text, "分") > 0 Then
                        If InStr(gobjBody.Text, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-MM-dd hh:mm:ss"
                        Else
                            .NumberFormatLocal = "yyyy""年""MM""月""dd""日"" hh""时""mm""分""ss""秒"""
                        End If
                    Else
                        If InStr(gobjBody.Text, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-MM-dd"
                        Else
                            .NumberFormatLocal = "yyyy""年""MM""月""dd""日"""
                        End If
                    End If
                Else
                    .NumberFormatLocal = "@"
                End If
                .HorizontalAlignment = GrdAlignment(gobjBody)
                .VerticalAlignment = -4108
            End With
        End If
    Next

    sngPer = sngPer + 1
    Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    
    '输出表头及表体的合并格式
    If gobjHead Is Nothing Then
        'intRow参数(Excel起始行号)加2的原因：MSHGrid的0行号对应Excel的1行号,+1;标题占一行,+1。
        Call ExcelMerge(gobjBody, gcolUpItem.count + 2, 0, gobjBody.Cols, 0, gobjBody.FixedRows)
        Call ExcelMerge(gobjBody, gcolUpItem.count + gobjBody.FixedRows + 2, 0, gobjBody.Cols, gobjBody.FixedRows, gobjBody.Rows)
    Else
        Call ExcelMerge(gobjHead, gcolUpItem.count + 2, 0, gobjHead.Cols, 0, gobjHead.FixedRows)
        Call ExcelMerge(gobjBody, gcolUpItem.count + gobjBody.FixedRows + gobjHead.FixedRows + 2, 0, gobjBody.Cols, gobjBody.FixedRows, gobjBody.Rows)
    End If
    
    '表格内部样式设计
    If gobjHead Is Nothing Then
        i = gcolUpItem.count + gobjBody.Rows + 1
    Else
        i = gcolUpItem.count + gobjHead.FixedRows + gobjBody.Rows + 1
    End If
    strTmp = "A" & Trim(str(gcolUpItem.count + 2)) & ":" & Col2Excel(gobjBody.Cols - lngZeroCol) & Trim(str(i))
    With gobjExcel.Range(strTmp)
        '字体
        .Font.name = gobjBody.Font.name
        .Font.Size = gobjBody.Font.Size
        .Font.Bold = gobjBody.Font.Bold
        .Font.Italic = gobjBody.Font.Italic
        If gblnIsWps Then
            '刘兴宏:2007/4/20
            .Font.Underline = IIF(gobjBody.Font.Underline, 2, -4142)
        Else
            .Font.Underline = gobjBody.Font.Underline
        End If
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
        
        '行高
        .RowHeight = gobjBody.RowHeight(gobjBody.FixedRows) / H_Excel
        
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
        
        '背景色
        .Interior.Pattern = -4142 'xlPatternNone
        .Interior.Color = CDbl(gobjBody.BackColor)
        
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
        
        '前景色
        .Font.Color = CDbl(gobjBody.ForeColor)
        
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
        
        '边线
        .Borders(7).LineStyle = 1 'xlContinuous
        .Borders(8).LineStyle = 1
        .Borders(9).LineStyle = 1
        .Borders(10).LineStyle = 1
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
        
        '边线色
        .Borders(7).Color = CDbl(gobjBody.GridColor)
        .Borders(8).Color = CDbl(gobjBody.GridColor)
        .Borders(9).Color = CDbl(gobjBody.GridColor)
        .Borders(10).Color = CDbl(gobjBody.GridColor)
        
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
        
        If gobjBody.Cols - lngZeroCol > 1 Then
            .Borders(11).LineStyle = 1
            .Borders(11).Color = CDbl(gobjBody.GridColor)
        End If
        If i <> gcolUpItem.count + 2 Then
            .Borders(12).LineStyle = 1
            .Borders(12).Color = CDbl(gobjBody.GridColor)
        End If
        '刘兴宏:2007/04/20
        If gblnIsWps Then
            .Borders.Weight = 2
        End If
    End With
    
    sngPer = sngPer + 1
    Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    
    '处理表格上面一行的下边线,因为第一行如果行高为0且纵向合并,则上边线不可见
    strTmp = "A" & gcolUpItem.count + 1 & ":" & Col2Excel(gobjBody.Cols - lngZeroCol) & gcolUpItem.count + 1
    With gobjExcel.Range(strTmp)
        .Borders(9).LineStyle = 1
        .Borders(9).Color = CDbl(gobjBody.GridColor)
        '刘兴宏:2007/04/20
        If gblnIsWps Then
            .Borders(9).Weight = 2
        End If
    End With
    
    sngPer = sngPer + 1
    Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    
    '表头格式
    If gobjHead Is Nothing Then
        strTmp = "A" & gcolUpItem.count + 2 & ":" & Col2Excel(gobjBody.Cols - lngZeroCol) & gobjBody.FixedRows + gcolUpItem.count + 1
    Else
        strTmp = "A" & gcolUpItem.count + 2 & ":" & Col2Excel(gobjBody.Cols - lngZeroCol) & gobjHead.FixedRows + gcolUpItem.count + 1
    End If
    With gobjExcel.Range(strTmp)
        .NumberFormatLocal = "@"
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        If Not gobjHead Is Nothing Then .WrapText = True
    End With
    
    sngPer = sngPer + 1
    Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    
    '标题格式
    With gobjExcel.Range("A1:" & Col2Excel(gobjBody.Cols - lngZeroCol) & "1")
        .Font.name = gobjTitle.字体
        .Font.Size = gobjTitle.字号
        .Font.Bold = gobjTitle.粗体
        .Font.Italic = gobjTitle.斜体
        If gblnIsWps Then
            '刘兴宏:2007/4/20
            .Font.Underline = IIF(gobjTitle.下线, 2, -4142)
        Else
            .Font.Underline = gobjTitle.下线
        End If
        .Font.Color = CDbl(gobjTitle.前景)
        Select Case gobjTitle.对齐
            Case 0
                .HorizontalAlignment = -4131 'xlLeft
            Case 1
                .HorizontalAlignment = -4108 'xlCenter
            Case 2
                .HorizontalAlignment = -4152 'xlRight
        End Select
        
        .VerticalAlignment = -4108 'xlCenter
        
        If gobjTitle.边框 Then
            .Borders(7).LineStyle = 1
            .Borders(8).LineStyle = 1
            .Borders(9).LineStyle = 1
            .Borders(10).LineStyle = 1
        End If
        If gobjTitle.背景 <> &HFFFFFF Then
            .Interior.Pattern = -4142 'xlPatternNone
            .Interior.Color = CDbl(gobjTitle.背景)
        End If
        .RowHeight = GetstrHeight(gobjTitle.字体, gobjTitle.字号, gobjTitle.内容) * 1.5 / H_Excel
        .MergeCells = True
    End With

    sngPer = sngPer + 1
    Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    
    '表上项目格式
    j = 2
    For i = 1 To gcolUpItem.count
        With gobjExcel.Range("A" & Trim(str(j)) & ":" & Col2Excel(gobjBody.Cols - lngZeroCol) & Trim(str(j)))
            .Font.name = gcolUpItem(i).字体
            .Font.Size = gcolUpItem(i).字号
            .Font.Bold = gcolUpItem(i).粗体
            .Font.Italic = gcolUpItem(i).斜体
            If gblnIsWps Then
                '刘兴宏:2007/4/20
                .Font.Underline = IIF(gcolUpItem(i).下线, 2, -4142)
            Else
                .Font.Underline = gcolUpItem(i).下线
            End If
            .Font.Color = CDbl(gcolUpItem(i).前景)
            
            .VerticalAlignment = -4108 'xlCenter
            If InStr(gcolUpItem(i).内容, "|") = 0 Then
                 '只有一项时根据位置对齐
                If gcolUpItem(i).X + gcolUpItem(i).W / 2 < gobjBody.Left + gobjBody.Width / 3 Then
                    .HorizontalAlignment = -4131 'xlLeft
                ElseIf gcolUpItem(i).X + gcolUpItem(i).W / 2 > gobjBody.Left + gobjBody.Width * (2 / 3) Then
                    .HorizontalAlignment = -4152 'xlRight
                Else
                    .HorizontalAlignment = -4108 'xlCenter
                End If
            Else
                .HorizontalAlignment = -4108 'xlCenter
            End If
            
            If gcolUpItem(i).边框 Then
                .Borders(7).LineStyle = 1
                .Borders(8).LineStyle = 1
                .Borders(9).LineStyle = 1
                .Borders(10).LineStyle = 1
            End If
            If gcolUpItem(i).背景 <> &HFFFFFF Then
                .Interior.Pattern = -4142 'xlPatternNone
                .Interior.Color = CDbl(gcolUpItem(i).背景)
            End If
            .RowHeight = GetstrHeight(gcolUpItem(i).字体, gcolUpItem(i).字号, gcolUpItem(i).内容) * 1.5 / H_Excel
            '.NumberFormat = "G/通用格式"
            .NumberFormatLocal = "G/通用格式"
            .MergeCells = True
        End With
        j = j + 1
        
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    Next

    '表下项目格式
    j = gcolUpItem.count + gobjBody.Rows + 2
    If Not gobjHead Is Nothing Then j = j + gobjHead.FixedRows
    For i = 1 To gcolDownItem.count
        With gobjExcel.Range("A" & Trim(str(j)) & ":" & Col2Excel(gobjBody.Cols - lngZeroCol) & Trim(str(j)))
            .Font.name = gcolDownItem(i).字体
            .Font.Size = gcolDownItem(i).字号
            .Font.Bold = gcolDownItem(i).粗体
            .Font.Italic = gcolDownItem(i).斜体
            If gblnIsWps Then
                '刘兴宏:2007/4/20
                .Font.Underline = IIF(gcolDownItem(i).下线, 2, -4142)
            Else
                .Font.Underline = gcolDownItem(i).下线
            End If
            .Font.Color = CDbl(gcolDownItem(i).前景)
            
            .VerticalAlignment = -4108 'xlCenter
            If InStr(gcolDownItem(i).内容, "|") = 0 Then
                 '只有一项时根据位置对齐
                If gcolDownItem(i).X + gcolDownItem(i).W / 2 < gobjBody.Left + gobjBody.Width / 3 Then
                    .HorizontalAlignment = -4131 'xlLeft
                ElseIf gcolDownItem(i).X + gcolDownItem(i).W / 2 > gobjBody.Left + gobjBody.Width * (2 / 3) Then
                    .HorizontalAlignment = -4152 ' xlRight
                Else
                    .HorizontalAlignment = -4108 'xlCenter
                End If
            Else
                .HorizontalAlignment = -4108 'xlCenter
            End If
            
            If gcolDownItem(i).边框 Then
                .Borders(7).LineStyle = 1
                .Borders(8).LineStyle = 1
                .Borders(9).LineStyle = 1
                .Borders(10).LineStyle = 1
            End If
            If gcolDownItem(i).背景 <> &HFFFFFF Then
                .Interior.Pattern = -4142 'xlPatternNone
                .Interior.Color = CDbl(gcolDownItem(i).背景)
            End If
            .RowHeight = GetstrHeight(gcolDownItem(i).字体, gcolDownItem(i).字号, gcolDownItem(i).内容) * 1.5 / H_Excel
            '.NumberFormat = "G/通用格式"
            .NumberFormatLocal = "G/通用格式"
            .MergeCells = True
        End With
        j = j + 1
    
        sngPer = sngPer + 1
        Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
    Next
    
    '最后处理行高为0的行
    If gobjHead Is Nothing Then
        For i = 0 To gobjBody.Rows - 1
            If gobjBody.RowHeight(i) = 0 Then gobjExcel.Rows(gcolUpItem.count + 2 + i).Delete
        Next
    Else
        For i = 0 To gobjHead.FixedRows - 1
            If gobjHead.RowHeight(i) = 0 Then gobjExcel.Rows(gcolUpItem.count + 2 + i).Delete
        Next
        For i = 0 To gobjBody.Rows - 1
            If gobjBody.RowHeight(i) = 0 Then gobjExcel.Rows(gcolUpItem.count + 2 + i + gobjHead.FixedRows).Delete
        Next
    End If
    
    sngPer = sngPer + 1
    Call ShowFlash("正在输出到 Excel：处理格式 ...", sngPer / lngTotal, frmParent, True)
        
    '定义各数据列数据格式
    ReDim arrFormat(gobjBody.Cols - 1, 1) As Integer
    For i = 0 To gobjBody.Cols - 1
        arrFormat(i, 0) = i + 1 '列号
        arrFormat(i, 1) = 1 '列数据类型：xlGeneralFormat = 1,xlTextFormat = 2
    Next
    '产生文本文件,同时侦测列数据类型
    strFile = GridToTXT(arrFormat)
    
    If gblnIsWps Then
        '刘兴宏:2007/4/20
        gobjExcel.Workbooks.Add
        gobjExcel.Cells.Select
        gobjExcel.Selection.NumberFormatLocal = "@"
        gobjExcel.Range("A1").Select
    
        sngPer = sngPer + 5
        Call ShowFlash("正在输出到 Excel：处理数据 ...", sngPer / lngTotal, frmParent, True)
        
        gobjExcel.Windows(2).Activate
        Set objText = objFile.OpenTextFile(strFile)
        Clipboard.Clear
        Clipboard.SetText objText.ReadAll
        Call gobjExcel.Sheets(1).Paste
        Clipboard.Clear
        objText.Close
        
        gobjExcel.Windows(1).Activate
        gobjExcel.Cells.Select
        Clipboard.Clear
        gobjExcel.Selection.Copy
        gobjExcel.Windows(2).Activate
        gobjExcel.Cells.Select
        gobjExcel.Selection.PasteSpecial Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False
    Else
        '导入文本文件
        Call gobjExcel.Workbooks.OpenText(strFile, , 1, 1, 1, False, True, False, False, False, False, , arrFormat)
        sngPer = sngPer + 5
        Call ShowFlash("正在输出到 Excel：处理数据 ...", sngPer / lngTotal, frmParent, True)
    
        gobjExcel.Windows(2).Activate
        gobjExcel.Cells.Select
        gobjExcel.Selection.Copy
        gobjExcel.Windows(2).Activate
        gobjExcel.Cells.Select
        gobjExcel.Selection.PasteSpecial Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False
    End If
    
    sngPer = sngPer + 1
    Call ShowFlash("正在输出到 Excel：处理数据 ...", sngPer / lngTotal, frmParent, True)
    Clipboard.Clear
    
    If gblnIsWps Then
      gobjExcel.Windows(1).Close False
    Else
      gobjExcel.Windows(2).Close False
    End If
    Call ShowFlash("正在输出到 Excel：完成 ...", 1, frmParent, True)
    
    gobjExcel.Range("A1").Select
    gobjExcel.Caption = App.Title & " - " & gobjTitle.内容
    
    '另存为Excel文件
    If strExcelFile <> "" Then
        If objFile.FileExists(strExcelFile) Then
            Call objFile.DeleteFile(strExcelFile, True)
        End If
        gobjExcel.DisplayAlerts = False
        gobjExcel.ActiveWorkbook.SaveAs strExcelFile, xlNormal
    Else
        gobjExcel.UserControl = True
        gobjExcel.Visible = True
    End If
    
    Set gobjExcel = Nothing
    
    Call ShowFlash
    
    ExportExcel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Call ShowFlash
    Set gobjExcel = Nothing
End Function

Public Sub MakeAppend(frmSource As Object, objPaper As Object)
'功能：从报表标签集合中产生标题、表上项、表下项
'参数：如果没有标题，则取报表名称作为缺省标题
    Dim strIDs As String, lngTmp As Long, strTmp As String
    Dim lngStep As Long, lngYStart As Long, lngBtn As Long
    Dim i As Long, j As Long, objLbl As Object
    
    Set gobjTitle = New RPTItem
    Set gcolUpItem = New RPTItems
    Set gcolDownItem = New RPTItems
    
    If frmSource.lbl.UBound = 0 Then
        '缺省标题
        gobjTitle.内容 = frmSource.mobjReport.名称
        gobjTitle.字体 = "楷体_GB2312"
        gobjTitle.字号 = 18
        gobjTitle.对齐 = 1
        gobjTitle.背景 = vbWhite
        gobjTitle.前景 = 0
        gobjTitle.H = 400
        Exit Sub
    End If
    
    gobjTitle.Y = 32767 '较大值用于最上位置比较
    
    '-------------------------------------------------------------------------
    '标题:表格之上,第一个
    If gobjHead Is Nothing Then
        lngBtn = gobjBody.Top
    Else
        lngBtn = gobjHead.Top
    End If
    
    For Each objLbl In frmSource.lbl
        If objLbl.Index > 0 And objLbl.Container Is objPaper And objLbl.Top + objLbl.Height < lngBtn _
            And objLbl.Top < gobjTitle.Y And Trim(objLbl.Caption) <> "" Then
            With objLbl
                gobjTitle.id = .Index
                gobjTitle.内容 = .Caption
                gobjTitle.X = .Left
                gobjTitle.Y = .Top
                gobjTitle.W = .Width
                gobjTitle.H = .Height
                gobjTitle.粗体 = .FontBold
                gobjTitle.斜体 = .FontItalic
                gobjTitle.下线 = .FontUnderline
                gobjTitle.边框 = (.BorderStyle = 1)
                gobjTitle.字体 = .FontName
                gobjTitle.字号 = .FontSize
                gobjTitle.背景 = .BackColor
                gobjTitle.前景 = .ForeColor
                gobjTitle.前景 = .ForeColor
            End With
        End If
    Next
    If gobjTitle.id <> 0 Then
        strIDs = strIDs & "," & gobjTitle.id
        '根据标题的中心位置判断对齐方式
        lngTmp = gobjTitle.X + gobjTitle.W / 2
        If lngTmp <= gobjBody.Left + gobjBody.Width / 3 Then
            gobjTitle.对齐 = 0
        ElseIf lngTmp >= gobjBody.Left + gobjBody.Width * (2 / 3) Then
            gobjTitle.对齐 = 2
        Else
            gobjTitle.对齐 = 1
        End If
    Else
        '缺省标题
        gobjTitle.内容 = frmSource.mobjReport.名称
        gobjTitle.字体 = "楷体_GB2312"
        gobjTitle.字号 = 18
        gobjTitle.对齐 = 1
        gobjTitle.背景 = vbWhite
        gobjTitle.前景 = 0
        gobjTitle.H = 400
    End If
    
    '-------------------------------------------------------------------------
    '表上项目
    '确定起始点:lngYstart
    '确定间隔:lngStep
    '以最小标签高度为间隔
    lngYStart = 32767
    lngStep = 32767
    If gobjHead Is Nothing Then
        lngBtn = gobjBody.Top
    Else
        lngBtn = gobjHead.Top
    End If
    For Each objLbl In frmSource.lbl
        If objLbl.Index > 0 And objLbl.Container Is objPaper And objLbl.Top + objLbl.Height <= lngBtn And InStr(strIDs & ",", "," & objLbl.Index & ",") = 0 Then
            If objLbl.Top < lngYStart Then lngYStart = objLbl.Top
            If objLbl.Height < lngStep Then lngStep = objLbl.Height
        End If
    Next
    If lngStep = 32767 Then lngStep = 180 '最小间隔为一个9号字符高度
    
    i = lngYStart
    Do While i + lngStep <= lngBtn
        '当前分隔行内组织一行表上项目
        strTmp = ""
        For Each objLbl In frmSource.lbl
            If objLbl.Index > 0 And objLbl.Container Is objPaper And objLbl.Top >= i And objLbl.Top < i + lngStep _
                And InStr(strIDs & ",", "," & objLbl.Index & ",") = 0 And Trim(objLbl.Caption) <> "" Then
                strIDs = strIDs & "," & objLbl.Index
                strTmp = strTmp & "|" & Format(objLbl.Left, "00000") & "," & objLbl.Index
            End If
        Next
        If strTmp <> "" Then
            strTmp = Mid(strTmp, 2)
            '从左到右排序
            strTmp = SpecSort(strTmp)
            '该行格式以第一个标签为准,内容为所有之和
            '每行对齐方式在输出时根据标签数量及坐标来决定
            With frmSource.lbl(CInt(Split(Split(strTmp, "|")(0), ",")(1)))
                gcolUpItem.Add .Index, 0, "标签", 0, 2, 0, "", 0, .Caption, "", .Left, .Top, .Width, .Height _
                    , 0, 0, True, .FontName, .FontSize, .FontBold, .FontUnderline, .FontItalic, 0, .ForeColor _
                    , .BackColor, (.BorderStyle = 1), 0, "", "", "", False, False, , , , , , "_" & .Index
                lngTmp = .Index
            End With
            For j = 1 To UBound(Split(strTmp, "|"))
                gcolUpItem("_" & lngTmp).内容 = gcolUpItem("_" & lngTmp).内容 & "|" & _
                    frmSource.lbl(CInt(Split(Split(strTmp, "|")(j), ",")(1))).Caption
            Next
        End If
        
        '求下一行起始点
        lngYStart = 32767
        For Each objLbl In frmSource.lbl
            If objLbl.Index > 0 And objLbl.Container Is objPaper And objLbl.Top + objLbl.Height <= lngBtn _
                And objLbl.Top >= i + lngStep And InStr(strIDs & ",", "," & objLbl.Index & ",") = 0 Then
                If objLbl.Top < lngYStart Then lngYStart = objLbl.Top
            End If
        Next
        i = lngYStart
    Loop
    
    '-------------------------------------------------------------------------
    '表下项目
    '确定起始点:lngYstart
    '确定间隔:lngStep
    '以最小标签高度为间隔
    lngYStart = 32767
    For Each objLbl In frmSource.lbl
        If objLbl.Index > 0 And objLbl.Container Is objPaper And objLbl.Top >= gobjBody.Top + gobjBody.Height And InStr(strIDs & ",", "," & objLbl.Index & ",") = 0 Then
            If objLbl.Top < lngYStart Then lngYStart = objLbl.Top
        End If
    Next
    
    lngBtn = lngYStart
    For Each objLbl In frmSource.lbl
        If objLbl.Index > 0 And objLbl.Container Is objPaper And objLbl.Top >= lngYStart And InStr(strIDs & ",", "," & objLbl.Index & ",") = 0 Then
            If objLbl.Top > lngBtn Then lngBtn = objLbl.Top
        End If
    Next
    
    lngStep = 32767
    For Each objLbl In frmSource.lbl
        If objLbl.Index > 0 And objLbl.Container Is objPaper And objLbl.Top >= lngYStart And objLbl.Top <= lngBtn And InStr(strIDs & ",", "," & objLbl.Index & ",") = 0 Then
            If objLbl.Height < lngStep Then lngStep = objLbl.Height
        End If
    Next
    If lngStep = 32767 Then lngStep = 180 '最小间隔为一个9号字符高度
    
    i = lngYStart
    Do While i <= lngBtn And lngBtn <> 32767
        '当前分隔行内组织一行表上项目
        strTmp = ""
        For Each objLbl In frmSource.lbl
            If objLbl.Index > 0 And objLbl.Container Is objPaper And objLbl.Top >= i And objLbl.Top < i + lngStep And InStr(strIDs & ",", "," & objLbl.Index & ",") = 0 And Trim(objLbl.Caption) <> "" Then
                strIDs = strIDs & "," & objLbl.Index
                strTmp = strTmp & "|" & Format(objLbl.Left, "00000") & "," & objLbl.Index
            End If
        Next
        If strTmp <> "" Then
            strTmp = Mid(strTmp, 2)
            '从左到右排序
            strTmp = SpecSort(strTmp)
            '该行格式以第一个标签为准,内容为所有之和
            '每行对齐方式在输出时根据标签数量及坐标来决定
            With frmSource.lbl(CInt(Split(Split(strTmp, "|")(0), ",")(1)))
                gcolDownItem.Add .Index, 0, "标签", 0, 2, 0, "", 0, .Caption, "", .Left, .Top, .Width, .Height _
                    , 0, 0, True, .FontName, .FontSize, .FontBold, .FontUnderline, .FontItalic, 0, .ForeColor _
                    , .BackColor, (.BorderStyle = 1), 0, "", "", "", False, False, , , , , , "_" & .Index
                lngTmp = .Index
            End With
            For j = 1 To UBound(Split(strTmp, "|"))
                gcolDownItem("_" & lngTmp).内容 = gcolDownItem("_" & lngTmp).内容 & "|" & _
                    frmSource.lbl(CInt(Split(Split(strTmp, "|")(j), ",")(1))).Caption
            Next
        End If
        
        '求下一行起始点
        lngYStart = 32767
        For Each objLbl In frmSource.lbl
            If objLbl.Index > 0 And objLbl.Container Is objPaper And objLbl.Top <= lngBtn And objLbl.Top >= i + lngStep And InStr(strIDs & ",", "," & objLbl.Index & ",") = 0 Then
                If objLbl.Top < lngYStart Then lngYStart = objLbl.Top
            End If
        Next
        i = lngYStart
    Loop
End Sub

Private Function IsNumberEx(ByVal strText As String) As Boolean
'功能：
'参数：
'返回：

    strText = Trim(strText)
    If strText Like "0[0-9]*" Then
        IsNumberEx = False
    Else
        IsNumberEx = True
    End If
End Function

