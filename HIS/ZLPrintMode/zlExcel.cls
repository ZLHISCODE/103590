VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlExcel"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'EXCEL的VBA中的一些常数
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const xlCenter As Integer = -4108
Private Const xlEdgeLeft As Integer = 7
Private Const xlEdgeTop As Integer = 8
Private Const xlEdgeBottom As Integer = 9
Private Const xlEdgeRight As Integer = 10
Private Const xlContinuous As Integer = 1
Private Const xlInsideVertical As Integer = 11
Private Const xlInsideHorizontal As Integer = 12
Private Const MSTR_VSF As String = "VSFLEXGRID"

Public ExcelApp As Object        '成品时要改为Object对象
Public frmTempExcel As Object

Private mlngWideTable As Long
Private mintChar As Integer            '一个字符的宽度
Private mintProgress As Integer
Private mstrTempFile As String

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'功能： 本类用于完成向EXCEEL输出数据
'
'zlExcelFile                进行输出的中转站

'OutTextFile                输出表格内容到文本文件中
'File1Grd                   单msFlexGrid对象的输出到文本文件
'File2Grd                   两个msFlexGrid对象的输出到文本文件
'FileGrds                   多个msFlexGrid对象的输出到文本文件
'FileDbGrd                  单DBGrid对象的输出到文本文件
'FileFlxDb                  DBGrid和fsFlexGrid组合对象的输出到文本文件
'FileLvw                    listview对象的输出到文本文件

'Excel1Grd                  单msFlexGrid对象的输出到Excel
'Excel2Grd                  两个msFlexGrid对象的输出到Excel
'ExcelGrds                  多个msFlexGrid对象的输出到Excel
'ExcelDbGrd                 单DBGrid对象的输出到Excel
'ExcelFlxDb                 DBGrid和fsFlexGrid组合对象的输出到Excel
'ExcelLvw                   listview对象的输出到Excel

'strRow                     求出一行表上或表下项目在EXCEL中输出时的内容
'RowInttoStr                将数字列号转换为EXCEL中的表示方法
'chgAlignment               将VB的对齐方式转换为EXCEL中的对齐方式
'GrdAlignment               将FlexGrid的对齐方式转换为EXCEL中的对齐方式
'SetProgress                显示进度
'

Public Sub zlExcelFile()
     '------------------------------------------------
    '功能： 进行输出的中转站
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
'    On Error GoTo errHandle
    
    If gblnIsWps Then
        '刘兴宏:2007/4/20
        '以WPS为准
        Set ExcelApp = CreateObject("ET.Application") '打开一个WPS中的ET程序
    Else
        '增加到EXCEL中
        Set ExcelApp = CreateObject("Excel.Application") '打开一个EXCEL程序
    End If
    
    SetProgress 1
    
    '增加一张格式临时表
    ExcelApp.Workbooks.Add
    gfrmTemp.Font.Name = ExcelApp.StandardFont
    gfrmTemp.Font.Size = CCur(ExcelApp.StandardFontSize)
    
    '得出1磅多少个缇
    mintChar = CInt(gfrmTemp.TextWidth("0"))   '得到一个表上、表下项目的字符宽度

    gfrmTemp.Font.Name = "宋体"
    gfrmTemp.Font.Size = 9
    SetProgress 2
    
    Call setHeaderFooter
    SetProgress 3
    
    If TypeOf gobjSend Is zlPrintLvw Then ExcelLvw
    If TypeOf gobjSend Is zlPrint1Grd Then Excel1Grd
    If TypeOf gobjSend Is zlPrint2Grd Then Excel2Grd
    If TypeOf gobjSend Is zlPrintGrds Then ExcelGrds
    If TypeOf gobjSend Is zlPrintDbGrd Then ExcelDbGrd
    If TypeOf gobjSend Is zlPrintFlxDB Then ExcelFlxDb
    SetProgress 50
    
    '先输出到文本文件中
    Call OutTextFile '同时侦测列数据类型
    
    If gblnIsWps Then
        '刘兴宏:2007/4/20
        ExcelApp.Workbooks.Add
        ExcelApp.Cells.Select
        ExcelApp.Selection.NumberFormatLocal = "@"
        ExcelApp.Range("A1").Select
        SetProgress 95
        
        '粘贴格式
        Clipboard.Clear
        
        ExcelApp.Windows(2).Activate
        Set objText = objFile.OpenTextFile(mstrTempFile)
        Clipboard.Clear
        Clipboard.SetText objText.ReadAll
        Call ExcelApp.Sheets(1).Paste
        
        ExcelApp.DisplayAlerts = False
        Call ExcelApp.Sheets(1).Cells.Replace(What:="<换行分隔符>", Replacement:=vbCrLf, LookAt:=2, MatchCase:=False, SearchOrder:=1, SearchFormat:=False, ReplaceFormat:=False)
        ExcelApp.DisplayAlerts = True
        
        Clipboard.Clear
        
        With ExcelApp
            .Windows(1).Activate
            .Cells.Select
            .Selection.Copy
            .Windows(2).Activate
            .Cells.Select
            .Selection.PasteSpecial Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False
        End With
        SetProgress 96
    Else
        Call ExcelApp.Workbooks.OpenText(mstrTempFile, , 1, 1, 1, False, True, False, False, False, False, , arrFormat)
        SetProgress 95
        '粘贴格式
        With ExcelApp
            .Windows(2).Activate
            .Cells.Select
            .Selection.Copy
            .Windows(2).Activate
            .Cells.Select
            .Selection.PasteSpecial Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False
            
            '输出到文本文件时将vbcrlf替换成了"<换行分隔符>"（否则会导致以OpenText方式打开时作为新行）
            .DisplayAlerts = False
            .Cells.Replace What:="<换行分隔符>", Replacement:=vbCrLf, LookAt:=2, MatchCase:=False, SearchOrder:=1, SearchFormat:=False, ReplaceFormat:=False
            .DisplayAlerts = True
        End With
        SetProgress 96
    End If
    
    Clipboard.Clear
    SetProgress 97
    If gblnIsWps Then
        '刘兴宏:2007/4/20:关闭第一个窗口
        ExcelApp.Windows(1).Close False
    Else
        ExcelApp.Windows(2).Close False
    End If
    SetProgress 98
    
    ExcelApp.Range("A1").Select
    SetProgress 99
    ExcelApp.Visible = True
    SetProgress 100
    Unload frmTempExcel
    Set frmTempExcel = Nothing
    Exit Sub
    
errHandle:
    MsgBox "输出到Excel时出现严重错误，输出被迫中止。" & vbCrLf & Err.Description, vbCritical, gstrSysName
    Unload frmTempExcel
    Set ExcelApp = Nothing
    Set frmTempExcel = Nothing
End Sub

Private Function OutTextFile() As Boolean
    '------------------------------------------------
    '功能：先把要输出的内容生成一个文本文件
    '------------------------------------------------
    Dim strFile As String, strPath As String
    Dim intFileNum As Integer
    '产生临时文件
    strPath = Space(256): strFile = Space(256)
    GetTempPath 256, strPath
    strPath = Left$(strPath, InStr(strPath, Chr(0)) - 1)
    
    GetTempFileName strPath, "excel", 0, strFile
    mstrTempFile = Left$(strFile, InStr(strFile, Chr(0)) - 1)
    '打开文件准备输出
    intFileNum = FreeFile()
    Open mstrTempFile For Binary Access Write As intFileNum
    
    SetProgress 51
    Select Case TypeName(gobjSend)
        Case "zlPrintLvw"
            FileLvw intFileNum
        Case "zlPrint1Grd"
            File1Grd intFileNum
        Case "zlPrint2Grd"
            File2Grd intFileNum
        Case "zlPrintGrds"
            FileGrds intFileNum
        Case "zlPrintDbGrd"
            FileDbGrd intFileNum
        Case "zlPrintFlxDB"
            FileFlxDb intFileNum
    End Select
    Close intFileNum
    SetProgress 94
End Function

Private Sub setHeaderFooter()
    '------------------------------------------------
    '功能： 产生页眉与页脚
    '------------------------------------------------
    Dim strLeft As String, strMiddle As String, strRight As String
    Dim intPos As Long
    Dim intPos1 As Long
    Dim strHeader As String, strFooter As String
    
    On Error Resume Next
    strHeader = ConvHFtoExcel(gobjSend.Header)
    intPos = InStr(strHeader, ";")
    intPos1 = intPos + 1
    strLeft = Mid(strHeader, 1, intPos - 1)
    intPos = InStr(intPos1, strHeader, ";")
    strMiddle = Mid(strHeader, intPos1, intPos - intPos1)
    intPos1 = intPos + 1
    strRight = Mid(strHeader, intPos1)

    With ExcelApp.ActiveSheet.PageSetup
        .LeftHeader = strLeft
        .CenterHeader = strMiddle
        .RightHeader = strRight
    End With
    
    strFooter = ConvHFtoExcel(gobjSend.Footer)
    intPos = InStr(strFooter, ";")
    intPos1 = intPos + 1
    strLeft = Mid(strFooter, 1, intPos - 1)
    intPos = InStr(intPos1, strFooter, ";")
    strMiddle = Mid(strFooter, intPos1, intPos - intPos1)
    intPos1 = intPos + 1
    strRight = Mid(strFooter, intPos1)

    With ExcelApp.ActiveSheet.PageSetup
        .LeftFooter = strLeft
        .CenterFooter = strMiddle
        .RightFooter = strRight
    End With

End Sub

Private Function ConvHFtoExcel(ByVal strSource As String) As String
    '------------------------------------------------
    '功能：将页眉与页脚转换成实际打印的内容
    '参数：strSource    页眉与页脚
    '返回：实际打印的内容
    '------------------------------------------------
    Dim strTemp As String
    
    strTemp = Replace(strSource, "[页码]", "&P")
    strTemp = Replace(strTemp, "[页数]", "&N")
    strTemp = Replace(strTemp, "[时间]", "&T")
    strTemp = Replace(strTemp, "[日期]", "&D")
    
    Dim strKeyValue As String       '键值
    Dim lngKey As Long
    Dim lngKeySize As Long
    Dim strRegPath As String
    If IsWindows95 Then
        strRegPath = "Software\MicroSoft\Windows\CurrentVersion"
    Else
        strRegPath = "Software\MicroSoft\Windows NT\CurrentVersion"
    End If
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, strRegPath, 0, KEY_READ, lngKey) = 0 Then
        strKeyValue = Space(256)
        lngKeySize = 256
        If RegQueryValueEx(lngKey, "RegisteredOrganization", 0, 1, strKeyValue, lngKeySize) = 0 Then
            gstrUnitName = StripTerminator(strKeyValue)
        End If
        strKeyValue = Space(256)
        lngKeySize = 256
        If RegQueryValueEx(lngKey, "RegisteredOwner", 0, 1, strKeyValue, lngKeySize) = 0 Then
            gstrUserName = StripTerminator(strKeyValue)
        End If
    End If
    RegCloseKey lngKey

    strTemp = Replace(strTemp, "[用户名]", gstrUserName)
    strTemp = Replace(strTemp, "[单位名]", gstrUnitName)
    ConvHFtoExcel = strTemp
End Function

Private Sub ExcelGrid(GrdTemp As Object, intRow As Long, ByVal LeftCol As Long, ByVal RightCol As Long, ByVal TopRow As Long, ByVal BottomRow As Long, intBegin As Long, intPrecent As Integer)
'功能：根据指定区域的表格设置对应Excel区域的格式（包括合并单元格）
'参数：GrdTemp   表格控件
'      intRow    Excel中的起始行
'      LeftCol   表格中的起始列
'      RightCol  表格中的终止列
'      TopRow    表格中的起始行
'      BottomRow 表格中的终止行
'      intBegin      起始进度值
'      intPrecent    本过程占用的进度值

    Dim intY As Long
    Dim strTemp As String
    Dim iCount As Long, intColcnt As Long
    Dim intBlank前 As Long, intBlank后 As Long '合并单元格前后的隐藏列
    Dim iTemp As Long, intRowUp As Long
    Dim blnMerge As Boolean
    Dim bln表头合并 As Boolean
    
    
    If GrdTemp.MergeCells = 0 Then Exit Sub
    
    If GrdTemp.MergeCells = 5 Then
        '功能：表头并
        If TopRow = 0 And BottomRow = GrdTemp.FixedRows Then
            bln表头合并 = True
        End If
    End If
    
    '合并单元格
    '横向合并
    intY = intRow
    For iCount = TopRow To BottomRow - 1
        If GrdTemp.MergeRow(iCount) = True Or bln表头合并 Then
            intBlank前 = 0
            For intColcnt = LeftCol To RightCol - 2
                If GrdTemp.ColWidth(intColcnt) <> 0 Then
                    '只有非隐藏列才得到处理
                    If GrdTemp.TextMatrix(iCount, intColcnt) <> "" Then '空白区域就不合并
                        strTemp = RowInttoStr(intColcnt + 1 - intBlank前) & Trim(str(intY))
                        ExcelApp.Range(strTemp).Select
                        If Not ExcelApp.Selection.MergeCells Then   '已合并了
                            '初始化循环内部使用的变量
                            blnMerge = False
                            intBlank后 = 0
                            For iTemp = intColcnt + 1 To RightCol - 1
                                If GrdTemp.ColWidth(iTemp) <> 0 Then
                                    '仍然要判断列宽，隐藏列不显示
                                    '不知第一个单元相同就退出
                                    If GrdTemp.TextMatrix(iCount, intColcnt) <> GrdTemp.TextMatrix(iCount, iTemp) Then Exit For
                                    If GrdTemp.MergeCells = 3 Or GrdTemp.MergeCells = 4 Then  '有列限制
                                        intRowUp = iCount - 1
                                        Do While intRowUp >= TopRow
                                            '上面没合并就退出
                                            If GrdTemp.TextMatrix(intRowUp, intColcnt) <> GrdTemp.TextMatrix(intRowUp, iTemp) Then Exit For
                                            intRowUp = intRowUp - 1
                                        Loop
                                    End If
                                    
                                    blnMerge = True
                                Else
                                    intBlank后 = intBlank后 + 1
                                End If
                            Next
                            If blnMerge = True Then
                                strTemp = RowInttoStr(intColcnt + 1 - intBlank前) & Trim(str(intY)) & ":" & RowInttoStr(iTemp - intBlank前 - intBlank后) & Trim(str(intY))
                                ExcelApp.Range(strTemp).Select
                                ExcelApp.Selection.MergeCells = True
                                intColcnt = iTemp - 1 '跳过已合并的列
                            End If
                        End If
                    End If
                Else
                    intBlank前 = intBlank前 + 1
                End If
            Next
         End If
        intY = intY + 1
        SetProgress intBegin + (iCount - TopRow) / (BottomRow - TopRow) * intPrecent / 2
    Next
    '纵向合并
    intY = intRow
    intBlank前 = 0
    For intColcnt = LeftCol To RightCol - 1
        If GrdTemp.ColWidth(intColcnt) <> 0 Then
            '只处理不隐藏的列
            If GrdTemp.MergeCol(intColcnt) = True Or bln表头合并 Then
                For iCount = TopRow To BottomRow - 2
                    If GrdTemp.TextMatrix(iCount, intColcnt) <> "" Then '空白区域就不合并
                        strTemp = RowInttoStr(intColcnt + 1 - intBlank前) & Trim(str(intY + iCount))
                        ExcelApp.Range(strTemp).Select
                        If ExcelApp.Selection.MergeCells = False Then  '已合并了
                            blnMerge = False
                            For iTemp = iCount + 1 To BottomRow - 1
                                '不知第一个单元相同就退出
                                If GrdTemp.TextMatrix(iCount, intColcnt) <> GrdTemp.TextMatrix(iTemp, intColcnt) Then Exit For
                                If GrdTemp.MergeCells = 2 Or GrdTemp.MergeCells = 4 Then  '有行限制
                                    intRowUp = intColcnt - 1
                                    Do While intRowUp >= LeftCol
                                        '左面没合并就退出
                                        If GrdTemp.TextMatrix(iCount, intRowUp) <> GrdTemp.TextMatrix(iTemp, intRowUp) Then Exit For
                                        intRowUp = intRowUp - 1
                                    Loop
                                End If
                                
                                blnMerge = True
                            Next
                            If blnMerge = True Then
                                strTemp = RowInttoStr(intColcnt + 1 - intBlank前) & Trim(str(intY + iCount - TopRow)) & ":" & RowInttoStr(intColcnt + 1 - intBlank前) & Trim(str(intY + iTemp - 1 - TopRow))
                                ExcelApp.Range(strTemp).Select
                                ExcelApp.Selection.MergeCells = True
                                iCount = iTemp - 1 '跳过已合并的列
                            End If
                        End If
                    End If
                Next
             End If
        Else
            intBlank前 = intBlank前 + 1
        End If
        SetProgress intBegin + intPrecent / 2 + (intColcnt - LeftCol) / (RightCol - LeftCol) * intPrecent / 2
    Next
End Sub

Private Function Excel1Grd()
    '------------------------------------------------
    '功能： 单msFlexGrid或VSFlexGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim intY As Long
    Dim GrdTemp As Object
    Dim strTemp As String, strText As String
    Dim iCount As Long, intColcnt As Long
    Dim intBlankCol As Long '空白列的数量
    Dim lngPos As Long
    
    Set GrdTemp = gobjSend.Body
    mlngWideTable = 0
    For intColcnt = 0 To GrdTemp.Cols - 1
        '判断列宽和列隐藏
        If GrdTemp.ColWidth(intColcnt) <= 0 Or ColHidden(GrdTemp, intColcnt) Then
            intBlankCol = intBlankCol + 1
        Else
            mlngWideTable = mlngWideTable + GrdTemp.ColWidth(intColcnt)
            ExcelApp.Columns(RowInttoStr(intColcnt - intBlankCol + 1) & ":" & RowInttoStr(intColcnt - intBlankCol + 1)).Select
            If GrdTemp.ColWidth(intColcnt) / mintChar > 0 Then ExcelApp.Selection.ColumnWidth = GrdTemp.ColWidth(intColcnt) / mintChar
        End If
        SetProgress 3 + (intColcnt / GrdTemp.Cols) * 3
    Next
    
    '网格内容格式
    intBlankCol = 0
    For intColcnt = 0 To GrdTemp.Cols - 1
        If GrdTemp.ColWidth(intColcnt) <= 0 Or ColHidden(GrdTemp, intColcnt) Then
            intBlankCol = intBlankCol + 1
        Else
            strText = ""
            '判断列的格式
            For iCount = GrdTemp.FixedRows To GrdTemp.Rows - 1
                strText = Trim(GrdTemp.TextMatrix(iCount, intColcnt))
                '不是空白就有格式
                If strText <> "" Then Exit For
            Next
            strTemp = RowInttoStr(intColcnt - intBlankCol + 1)
            strTemp = strTemp & ":" & strTemp
            ExcelApp.Columns(strTemp).Select
            With ExcelApp.Selection
                If IsNumeric(strText) Then
                    If Len(strText) <= 15 Then
                        lngPos = InStr(strText, ".")
                        If lngPos = 0 Then
                            .NumberFormatLocal = "0_ "
                        Else
                            .NumberFormatLocal = "0." & String(Len(Mid(strText, lngPos + 1)), "0") & "_ "
                        End If
                    Else
                        .NumberFormatLocal = "@"
                    End If
                ElseIf IsDate(strText) Then
                    If InStr(strText, ":") > 0 Or InStr(strText, "分") > 0 Then
                        If InStr(strText, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd hh:mm:ss"
                        Else
                            .NumberFormatLocal = "yyyy""年""mm""月""dd""日"" hh""时""mm""分""ss""秒"""
                        End If
                    Else
                        If InStr(strText, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd"
                        Else
                            .NumberFormatLocal = "yyyy""年""mm""月""dd""日"""
                        End If
                    End If
                Else
                    .NumberFormatLocal = "@"
                End If
                
                '控件判断
                If iCount >= GrdTemp.Rows Then iCount = GrdTemp.Rows - 1
                
                If UCase(TypeName(GrdTemp)) = MSTR_VSF Then
                    .Font.Name = GrdTemp.Cell(flexcpFontName, iCount, intColcnt)
                    .Font.Size = GrdTemp.Cell(flexcpFontSize, iCount, intColcnt)
                Else
                    'MSHFlexGrid需要定位到单元格CellFontName、CellFontName才有效
                    GrdTemp.Row = iCount
                    GrdTemp.Col = intColcnt
                    .Font.Name = GrdTemp.CellFontName
                    .Font.Size = GrdTemp.CellFontSize
                End If
                
                If ExcelApp.Selection.MergeCells = True Then
                    .HorizontalAlignment = -4108
                Else
                    .HorizontalAlignment = GrdAlignment(GrdTemp, intColcnt, iCount)
                End If
                .VerticalAlignment = -4108
            End With
        End If
        SetProgress 6 + (intColcnt / GrdTemp.Cols) * 10
    Next
    
    ExcelGrid GrdTemp, gobjSend.UnderAppRows.Count + 2, 0, GrdTemp.Cols, 0, GrdTemp.FixedRows, 16, 9
    ExcelGrid GrdTemp, gobjSend.UnderAppRows.Count + GrdTemp.FixedRows + 2, 0, GrdTemp.Cols, GrdTemp.FixedRows, GrdTemp.Rows, 25, 10
    
    '设置网格线
    intY = gobjSend.UnderAppRows.Count + GrdTemp.Rows + 1
    ExcelApp.Range("A" & Trim(str(gobjSend.UnderAppRows.Count + 2)) & ":" & RowInttoStr(GrdTemp.Cols - intBlankCol) & Trim(str(intY))).Select
    With ExcelApp.Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        If GrdTemp.Cols - intBlankCol > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
        If intY <> gobjSend.UnderAppRows.Count + 2 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        '2007/04/20:刘兴宏加入,主要是导到wps时,是虚线,需要将其设置线的精细程度
        If gblnIsWps Then
            .Borders.Weight = 2
        End If
        
    End With
    SetProgress 36
    
    '网头格式
    intBlankCol = 0
    For intColcnt = 0 To GrdTemp.Cols - 1
        If GrdTemp.ColWidth(intColcnt) <= 0 Or ColHidden(GrdTemp, intColcnt) Then
            intBlankCol = intBlankCol + 1
        Else
            For iCount = 0 To GrdTemp.FixedRows - 1
                strTemp = RowInttoStr(intColcnt - intBlankCol + 1) & Trim(CStr(iCount + gobjSend.UnderAppRows.Count + 2))
                ExcelApp.Range(strTemp).Select
                With ExcelApp.Selection
                    .NumberFormatLocal = "@"
                    If UCase(TypeName(GrdTemp)) = MSTR_VSF Then
                        .Font.Name = GrdTemp.Cell(flexcpFontName, iCount, intColcnt)
                        .Font.Size = GrdTemp.Cell(flexcpFontSize, iCount, intColcnt)
                    Else
                        GrdTemp.Row = iCount
                        GrdTemp.Col = intColcnt
                        .Font.Name = GrdTemp.CellFontName
                        .Font.Size = GrdTemp.CellFontSize
                    End If
                    .HorizontalAlignment = -4108
                    .VerticalAlignment = -4108
                End With
            Next
        End If
        SetProgress 37 + (intColcnt / GrdTemp.Cols) * 10
    Next
    
    '标题的格式
    ExcelApp.Range("A1:" & RowInttoStr(GrdTemp.Cols - intBlankCol) & "1").Select
    With ExcelApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = -4108
        .MergeCells = True
        .Font.Name = gobjSend.Title.Font.Name
        .Font.Size = gobjSend.Title.Font.Size
        .RowHeight = GetstrHeight(.Font.Name, .Font.Size, "刘") * 2 / 20
    End With
    SetProgress 48
    '表上项目输出
    intY = 2
    For iCount = 1 To gobjSend.UnderAppRows.Count
        ExcelApp.Range("A" & Trim(str(intY)) & ":" & RowInttoStr(GrdTemp.Cols - intBlankCol) & Trim(str(intY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
        End With
        intY = intY + 1
    Next
    SetProgress 49
    '表下项目输出
    intY = gobjSend.UnderAppRows.Count + GrdTemp.Rows + 2
    For iCount = 1 To gobjSend.BelowAppRows.Count
        ExcelApp.Range("A" & Trim(str(intY)) & ":" & RowInttoStr(GrdTemp.Cols - intBlankCol) & Trim(str(intY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
            
        End With
        intY = intY + 1
    Next
End Function

Private Function Excel2Grd()
    '------------------------------------------------
    '功能： 两个msFlexGrid或VSFlexGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim intY As Long
    Dim GrdHead As Object
    Dim GrdBody As Object
    Dim strTemp As String, strText As String
    Dim iCount As Long, intColcnt As Long, lngPos As Long
    Dim intBlankCol As Long '空白列的数量
    
    Set GrdHead = gobjSend.BodyHead
    Set GrdBody = gobjSend.BodyGrid
    
    mlngWideTable = 0
    For intColcnt = 0 To GrdHead.Cols - 1
        If GrdHead.ColWidth(intColcnt) <= 0 Or ColHidden(GrdHead, intColcnt) Then
            intBlankCol = intBlankCol + 1
        Else
            mlngWideTable = mlngWideTable + GrdHead.ColWidth(intColcnt)
            ExcelApp.Columns(RowInttoStr(intColcnt - intBlankCol + 1) & ":" & RowInttoStr(intColcnt - intBlankCol + 1)).Select
            If GrdHead.ColWidth(intColcnt) / mintChar > 0 Then ExcelApp.Selection.ColumnWidth = GrdHead.ColWidth(intColcnt) / mintChar
        End If
        SetProgress 3 + (intColcnt / GrdHead.Cols) * 3
    Next
    
     '网格内容格式
    intBlankCol = 0
    For intColcnt = 0 To GrdBody.Cols - 1
        If GrdBody.ColWidth(intColcnt) <= 0 Or ColHidden(GrdBody, intColcnt) Then
            intBlankCol = intBlankCol + 1
        Else
            strText = ""
            For iCount = 0 To GrdBody.Rows - 1
                strText = Trim(GrdBody.TextMatrix(iCount, intColcnt))
                '不是空白就有格式
                If strText <> "" Then Exit For
            Next
            strTemp = RowInttoStr(intColcnt - intBlankCol + 1)
            strTemp = strTemp & ":" & strTemp
            ExcelApp.Columns(strTemp).Select
            With ExcelApp.Selection
                If IsNumeric(strText) Then
                    If Len(strText) <= 15 Then
                        lngPos = InStr(strText, ".")
                        If lngPos = 0 Then
                            .NumberFormatLocal = "0_ "
                        Else
                            .NumberFormatLocal = "0." & String(Len(Mid(strText, lngPos + 1)), "0") & "_ "
                        End If
                    Else
                        .NumberFormatLocal = "@"
                    End If
                ElseIf IsDate(strText) Then
                    If InStr(strText, ":") > 0 Or InStr(strText, "分") > 0 Then
                        If InStr(strText, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd hh:mm:ss"
                        Else
                            .NumberFormatLocal = "yyyy""年""mm""月""dd""日"" hh""时""mm""分""ss""秒"""
                        End If
                    Else
                        If InStr(strText, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd"
                        Else
                            .NumberFormatLocal = "yyyy""年""mm""月""dd""日"""
                        End If
                    End If
                Else
                    .NumberFormatLocal = "@"
                End If
                
                If iCount >= GrdBody.Rows Then iCount = GrdBody.Rows - 1
                
                If UCase(TypeName(GrdBody)) = MSTR_VSF Then
                    .Font.Name = GrdBody.Cell(flexcpFontName, iCount, intColcnt)
                    .Font.Size = GrdBody.Cell(flexcpFontSize, iCount, intColcnt)
                Else
                    GrdBody.Row = iCount
                    GrdBody.Col = intColcnt
                    .Font.Name = GrdBody.CellFontName
                    .Font.Size = GrdBody.CellFontSize
                End If
                
                If ExcelApp.Selection.MergeCells = True Then
                    .HorizontalAlignment = -4108
                Else
                    .HorizontalAlignment = GrdAlignment(GrdBody, intColcnt, iCount)
                End If
                .VerticalAlignment = -4108
            End With
        End If
        SetProgress 6 + (intColcnt / GrdBody.Cols) * 10
    Next
    
    ExcelGrid GrdHead, gobjSend.UnderAppRows.Count + 2, 0, GrdHead.Cols, 0, GrdHead.FixedRows, 16, 9
    ExcelGrid GrdBody, gobjSend.UnderAppRows.Count + GrdHead.FixedRows + 2, 0, GrdHead.Cols, 0, GrdBody.Rows, 25, 10
    
    '设置网格线
    intY = gobjSend.UnderAppRows.Count + GrdHead.FixedRows + GrdBody.Rows + 1
    ExcelApp.Range("A" & Trim(str(gobjSend.UnderAppRows.Count + 2)) & ":" & RowInttoStr(GrdBody.Cols - intBlankCol) & Trim(str(intY))).Select
    With ExcelApp.Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        If GrdBody.Cols > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
        If intY <> gobjSend.UnderAppRows.Count + 2 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        '2007/04/20:刘兴宏加入,主要是导到wps时,是虚线,需要将其设置线的精细程度
        If gblnIsWps Then
            .Borders.Weight = 2
        End If
    End With
    SetProgress 36
    
    '网头格式
    intBlankCol = 0
    For intColcnt = 0 To GrdHead.Cols - 1
        If GrdHead.ColWidth(intColcnt) <= 0 Or ColHidden(GrdHead, intColcnt) Then
            intBlankCol = intBlankCol + 1
        Else
            For iCount = 0 To GrdHead.FixedRows - 1
                strTemp = RowInttoStr(intColcnt - intBlankCol + 1) & Trim(CStr(iCount + gobjSend.UnderAppRows.Count + 2))
                ExcelApp.Range(strTemp).Select
                With ExcelApp.Selection
                    .NumberFormatLocal = "@"
                    If UCase(TypeName(GrdHead)) = MSTR_VSF Then
                        .Font.Name = GrdHead.Cell(flexcpFontName, iCount, intColcnt)
                        .Font.Size = GrdHead.Cell(flexcpFontSize, iCount, intColcnt)
                    Else
                        GrdHead.Row = iCount
                        GrdHead.Col = intColcnt
                        .Font.Name = GrdHead.CellFontName
                        .Font.Size = GrdHead.CellFontSize
                    End If
                    .HorizontalAlignment = -4108
                    .VerticalAlignment = -4108
                End With
            Next
        End If
        SetProgress 37 + (intColcnt / GrdHead.Cols) * 10
    Next
  
    '标题的格式
    ExcelApp.Range("A1:" & RowInttoStr(GrdHead.Cols - intBlankCol) & "1").Select
    With ExcelApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = -4108
        .MergeCells = True
        .Font.Name = gobjSend.Title.Font.Name
        .Font.Size = gobjSend.Title.Font.Size
        .RowHeight = GetstrHeight(.Font.Name, .Font.Size, "刘") * 2 / 20
        
    End With
    SetProgress 48
    '表上项目输出
    intY = 2
    For iCount = 1 To gobjSend.UnderAppRows.Count
        ExcelApp.Range("A" & Trim(str(intY)) & ":" & RowInttoStr(GrdHead.Cols - intBlankCol) & Trim(str(intY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
        End With
        intY = intY + 1
    Next
    SetProgress 49
    '表下项目输出
    intY = gobjSend.UnderAppRows.Count + GrdHead.FixedRows + GrdBody.Rows + 2
    For iCount = 1 To gobjSend.BelowAppRows.Count
        ExcelApp.Range("A" & Trim(str(intY)) & ":" & RowInttoStr(GrdHead.Cols - intBlankCol) & Trim(str(intY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
        End With
        intY = intY + 1
    Next

End Function

Private Function ExcelGrds()
    '------------------------------------------------
    '功能： 多个msFlexGrid或VSFlexGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim intY As Long
    Dim GrdTemp As Object
    Dim strTemp As String, strText As String
    Dim iCount As Long, intColcnt As Long
    Dim intCount As Long, intSum As Long
    Dim blnRedraw As Boolean
    Dim intBlankCol As Long '空白列的数量
    Dim lngPos As Long

    Set GrdTemp = gobjSend.Grds(1)
    mlngWideTable = 0
    For intColcnt = 0 To GrdTemp.Cols - 1
        If GrdTemp.ColWidth(0) <= 0 Or ColHidden(GrdTemp, intColcnt) Then
            intBlankCol = intBlankCol + 1
        Else
            mlngWideTable = mlngWideTable + GrdTemp.ColWidth(intColcnt)
            ExcelApp.Columns(RowInttoStr(intColcnt - intBlankCol + 1) & ":" & RowInttoStr(intColcnt - intBlankCol + 1)).Select
            If GrdTemp.ColWidth(intColcnt) / mintChar > 0 Then ExcelApp.Selection.ColumnWidth = GrdTemp.ColWidth(intColcnt) / mintChar
        End If
        SetProgress 3 + (intColcnt / GrdTemp.Cols) * 3
    Next
    
    '网格内容输出
    intY = gobjSend.UnderAppRows.Count + 2 'intY指向每个表的表头
    intSum = gobjSend.Grds.Count
    For Each GrdTemp In gobjSend.Grds
        blnRedraw = GrdTemp.Redraw
        
         '网格内容格式
        For intColcnt = 0 To GrdTemp.Cols - 1
            For iCount = GrdTemp.FixedRows To GrdTemp.Rows - 1
                strText = Trim(GrdTemp.TextMatrix(iCount, intColcnt))
                '不是空白就有格式
                If strText <> "" Then Exit For
            Next
            strTemp = RowInttoStr(intColcnt + 1)
            strTemp = strTemp & Trim(str(intY + GrdTemp.FixedRows)) & ":" & strTemp & Trim(str(intY + GrdTemp.Rows))
            ExcelApp.Range(strTemp).Select
            With ExcelApp.Selection
                If IsNumeric(strText) Then
                    If Len(strText) <= 15 Then
                        lngPos = InStr(strText, ".")
                        If lngPos = 0 Then
                            .NumberFormatLocal = "0_ "
                        Else
                            .NumberFormatLocal = "0." & String(Len(Mid(strText, lngPos + 1)), "0") & "_ "
                        End If
                    Else
                        .NumberFormatLocal = "@"
                    End If
                ElseIf IsDate(strText) Then
                    If InStr(strText, ":") > 0 Or InStr(strText, "分") > 0 Then
                        If InStr(strText, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd hh:mm:ss"
                        Else
                            .NumberFormatLocal = "yyyy""年""mm""月""dd""日"" hh""时""mm""分""ss""秒"""
                        End If
                    Else
                        If InStr(strText, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd"
                        Else
                            .NumberFormatLocal = "yyyy""年""mm""月""dd""日"""
                        End If
                    End If
                Else
                    .NumberFormatLocal = "@"
                End If
                
                If iCount >= GrdTemp.Rows Then iCount = GrdTemp.Rows - 1
                
                If UCase(TypeName(GrdTemp)) = MSTR_VSF Then
                    .Font.Name = GrdTemp.Cell(flexcpFontName, iCount, intColcnt)
                    .Font.Size = GrdTemp.Cell(flexcpFontSize, iCount, intColcnt)
                Else
                    GrdTemp.Row = iCount
                    GrdTemp.Col = intColcnt
                    .Font.Name = GrdTemp.CellFontName
                    .Font.Size = GrdTemp.CellFontSize
                End If
                
                If ExcelApp.Selection.MergeCells = True Then
                    .HorizontalAlignment = -4108
                Else
                    .HorizontalAlignment = GrdAlignment(GrdTemp, intColcnt, iCount)
                End If
                .VerticalAlignment = -4108
            End With
        Next
        
        ExcelGrid GrdTemp, intY, 0, GrdTemp.Cols, 0, GrdTemp.FixedRows, 6 + intCount / intSum * 42, 42 / intSum * 0.4
        ExcelGrid GrdTemp, intY + GrdTemp.FixedRows, 0, GrdTemp.Cols, GrdTemp.FixedRows, GrdTemp.Rows, 6 + intCount / intSum * 42 + 42 / intSum * 0.4, 42 / intSum * 0.4
        
        '网头格式
        For intColcnt = 0 To GrdTemp.Cols - 1
            For iCount = 0 To GrdTemp.FixedRows - 1
                strTemp = RowInttoStr(intColcnt + 1) & Trim(CStr(iCount + intY))
                ExcelApp.Range(strTemp).Select
                With ExcelApp.Selection
                    .NumberFormatLocal = "@"
                    If UCase(TypeName(GrdTemp)) = MSTR_VSF Then
                        .Font.Name = GrdTemp.Cell(flexcpFontName, iCount, intColcnt)
                        .Font.Size = GrdTemp.Cell(flexcpFontSize, iCount, intColcnt)
                    Else
                        GrdTemp.Row = iCount
                        GrdTemp.Col = intColcnt
                        .Font.Name = GrdTemp.CellFontName
                        .Font.Size = GrdTemp.CellFontSize
                    End If
                    .HorizontalAlignment = -4108
                    .VerticalAlignment = -4108
                End With
            Next
            SetProgress 6 + intCount / intSum * 42 + 42 / intSum * 0.9 + (intColcnt / GrdTemp.Cols) * 42 / intSum * 0.1
        Next
        
        '表格线
        ExcelApp.Range("A" & Trim(str(intY)) & ":" & RowInttoStr(GrdTemp.Cols) & Trim(str(intY + GrdTemp.Rows - 1))).Select
        With ExcelApp.Selection
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            If GrdTemp.Cols > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
            If GrdTemp.Rows > 1 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            '2007/04/20:刘兴宏加入,主要是导到wps时,是虚线,需要将其设置线的精细程度
            If gblnIsWps Then
                .Borders.Weight = 2
            End If
        End With
        intY = intY + GrdTemp.Rows + 1
        
        GrdTemp.Redraw = blnRedraw
        intCount = intCount + 1
    Next
    Set GrdTemp = gobjSend.Grds(1)
    '标题的格式
    ExcelApp.Range("A1:" & RowInttoStr(GrdTemp.Cols - intBlankCol) & "1").Select
    With ExcelApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = -4108
        .MergeCells = True
        .Font.Name = gobjSend.Title.Font.Name
        .Font.Size = gobjSend.Title.Font.Size
        .RowHeight = GetstrHeight(.Font.Name, .Font.Size, "刘") * 2 / 20
    End With
    SetProgress 48
    '表上项目输出
    For iCount = 1 To gobjSend.UnderAppRows.Count
        ExcelApp.Range("A" & Trim(str(iCount + 1)) & ":" & RowInttoStr(GrdTemp.Cols - intBlankCol) & Trim(str(iCount + 1))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
        End With
    Next
    SetProgress 49
    '表下项目输出
    intY = intY - 1
    For iCount = 1 To gobjSend.BelowAppRows.Count
        ExcelApp.Range("A" & Trim(str(intY)) & ":" & RowInttoStr(GrdTemp.Cols - intBlankCol) & Trim(str(intY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
        End With
        intY = intY + 1
    Next
End Function

Private Function ExcelDbGrd()
    '------------------------------------------------
    '功能： 单DBGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim objDbGrd As Object, objSource As Object
    Dim lngY As Long
    Dim strTemp As String
    Dim intColcnt As Long
    Dim intCol As Long
    Dim objCol As Object
    Dim intColumnSub As Long
    
    Set objDbGrd = gobjSend.BodyGrid
    Set objSource = gobjSend.DataSource
    
    Dim lngPos As Long '保存记录位置
    lngPos = objDbGrd.Row
    
    mlngWideTable = 0
    intColcnt = 0
    For Each objCol In objDbGrd.Columns
        If objCol.Visible Then
            mlngWideTable = mlngWideTable + objCol.Width
            intColcnt = intColcnt + 1
            ExcelApp.Columns(RowInttoStr(intColcnt) & ":" & RowInttoStr(intColcnt)).Select
            ExcelApp.Selection.ColumnWidth = objCol.Width / mintChar
        End If
        SetProgress 3 + (intColcnt / objDbGrd.Columns.Count) * 3
    Next
   
   '设置网格线
    lngY = gobjSend.UnderAppRows.Count + objSource.RecordCount + IIf(objDbGrd.ColumnHeaders, 1, 0) + 1
    ExcelApp.Range("A" & Trim(str(gobjSend.UnderAppRows.Count + 2)) & ":" & RowInttoStr(intColcnt) & Trim(str(lngY))).Select
    With ExcelApp.Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        If intColcnt > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
        If lngY <> gobjSend.UnderAppRows.Count + 2 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        '2007/04/20:刘兴宏加入,主要是导到wps时,是虚线,需要将其设置线的精细程度
        If gblnIsWps Then
            .Borders.Weight = 2
        End If
    End With
    SetProgress 6
  
    '表头输出
    lngY = gobjSend.UnderAppRows.Count + 2
    If objDbGrd.ColumnHeaders Then
        For intCol = 0 To objDbGrd.Columns.Count - 1
            Set objCol = objDbGrd.Columns(intCol)
            If objCol.Visible Then
                intColumnSub = intColumnSub + 1
                strTemp = RowInttoStr(intColumnSub) & Trim(str(lngY))
                ExcelApp.Range(strTemp).Select
                With ExcelApp.Selection
                    .Font.Name = objDbGrd.HeadFont.Name
                    .Font.Size = objDbGrd.HeadFont.Size
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = -4108
                End With
            End If
        Next
        lngY = lngY + 1
    End If
    intColumnSub = 0
    Dim intColTemp As Long, strCol As String
    Dim iCount As Long
     '网格内容输出
    For intCol = 0 To objDbGrd.Columns.Count - 1
        Set objCol = objDbGrd.Columns(intCol)
        If objCol.Visible Then
            strCol = ""
            For iCount = 0 To objSource.RecordCount - 1
                strCol = Trim(objCol.CellText(objDbGrd.GetBookmark(iCount - lngPos)))
                If strCol <> "" Then Exit For
            Next
            
            intColTemp = intColTemp + 1
            strTemp = RowInttoStr(intColTemp)
            strTemp = strTemp & ":" & strTemp
            ExcelApp.Range(strTemp).Select
            With ExcelApp.Selection
                If IsNumeric(strCol) And Len(strCol) <= 15 Then
                    iCount = InStr(strCol, ".")
                    If iCount = 0 Then
                        .NumberFormatLocal = "0_ "
                    Else
                        .NumberFormatLocal = "0." & String(Len(Mid(strCol, iCount + 1)), "0") & "_ "
                    End If
                ElseIf IsDate(strCol) Then
                    If InStr(strCol, ":") > 0 Or InStr(strCol, "分") > 0 Then
                        If InStr(strCol, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd hh:mm:ss"
                        Else
                            .NumberFormatLocal = "yyyy""年""mm""月""dd""日"" hh""时""mm""分""ss""秒"""
                        End If
                    Else
                        If InStr(strCol, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd"
                        Else
                            .NumberFormatLocal = "yyyy""年""mm""月""dd""日"""
                        End If
                    End If
                Else
                    .NumberFormatLocal = "@"
                End If
                .Font.Name = objDbGrd.Font.Name
                .Font.Size = objDbGrd.Font.Size
                .HorizontalAlignment = chgAlignment(objCol.Alignment)
                .VerticalAlignment = -4108
            End With
        End If
        SetProgress 6 + (intCol / objDbGrd.Columns.Count) * 34
    Next
    '标题的格式
    ExcelApp.Range("A1:" & RowInttoStr(intColcnt) & "1").Select
    With ExcelApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = -4108
        .MergeCells = True
        .Font.Name = gobjSend.Title.Font.Name
        .Font.Size = gobjSend.Title.Font.Size
        .RowHeight = GetstrHeight(.Font.Name, .Font.Size, "刘") * 2 / 20
        
    End With
    SetProgress 43
    '表上项目输出
    lngY = 2
    For iCount = 1 To gobjSend.UnderAppRows.Count
        ExcelApp.Range("A" & Trim(str(lngY)) & ":" & RowInttoStr(intColcnt) & Trim(str(lngY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
            
        End With
        lngY = lngY + 1
    Next
    SetProgress 46
    '表下项目输出
    lngY = gobjSend.UnderAppRows.Count + objSource.RecordCount + IIf(objDbGrd.ColumnHeaders, 1, 0) + 2
    For iCount = 1 To gobjSend.BelowAppRows.Count
        ExcelApp.Range("A" & Trim(str(lngY)) & ":" & RowInttoStr(intColcnt) & Trim(str(lngY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
        End With
        lngY = lngY + 1
    Next
    SetProgress 49
End Function
Private Function ExcelFlxDb()
    '------------------------------------------------
    '功能： DBGrid和fsFlexGrid组合对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim GrdHead As Object
    Dim objDbGrd As Object, objSource As Object
    Dim objCol As Object
    
    Dim lngY As Long
    Dim strTemp As String
    Dim iCount As Long, intColcnt As Long
    Dim intColumnSub As Long
    
    Set GrdHead = gobjSend.BodyHead
    Set objDbGrd = gobjSend.BodyGrid
    Set objSource = gobjSend.DataSource
    
    Dim lngPos As Long '保存记录位置
    lngPos = objDbGrd.Row
    
    mlngWideTable = 0
    For Each objCol In objDbGrd.Columns
        If objCol.Visible Then
            mlngWideTable = mlngWideTable + objCol.Width
            intColcnt = intColcnt + 1
            ExcelApp.Columns(RowInttoStr(intColcnt) & ":" & RowInttoStr(intColcnt)).Select
            ExcelApp.Selection.ColumnWidth = objCol.Width / mintChar
            SetProgress 3 + (intColcnt / GrdHead.Cols) * 3
        End If
    Next
    
    Dim intColTemp As Long
    Dim strCol As String
     '网格内容输出
    intColTemp = 0
    lngY = gobjSend.UnderAppRows.Count + GrdHead.FixedRows + 2
    For intColcnt = 0 To objDbGrd.Columns.Count - 1
        Set objCol = objDbGrd.Columns(intColcnt)
        If objCol.Visible Then
            strCol = ""
            For iCount = 0 To objSource.RecordCount - 1
                strCol = Trim(objCol.CellText(objDbGrd.GetBookmark(iCount - lngPos)))
                If strCol <> "" Then Exit For
            Next
            intColTemp = intColTemp + 1
            strTemp = RowInttoStr(intColTemp)
            strTemp = strTemp & Trim(str(lngY)) & ":" & strTemp & Trim(str(lngY + objSource.RecordCount - 1))
            ExcelApp.Range(strTemp).Select
            With ExcelApp.Selection
                If IsNumeric(strCol) And Len(strCol) <= 15 Then
                    iCount = InStr(strCol, ".")
                    If iCount = 0 Then
                        .NumberFormatLocal = "0_ "
                    Else
                        .NumberFormatLocal = "0." & String(Len(Mid(strCol, iCount + 1)), "0") & "_ "
                    End If
                ElseIf IsDate(strCol) Then
                    If InStr(strCol, ":") > 0 Or InStr(strCol, "分") > 0 Then
                        If InStr(strCol, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd hh:mm:ss"
                        Else
                            .NumberFormatLocal = "yyyy""年""mm""月""dd""日"" hh""时""mm""分""ss""秒"""
                        End If
                    Else
                        If InStr(strCol, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd"
                        Else
                            .NumberFormatLocal = "yyyy""年""mm""月""dd""日"""
                        End If
                    End If
                Else
                    .NumberFormatLocal = "@"
                End If
                .Font.Name = objDbGrd.Font.Name
                .Font.Size = objDbGrd.Font.Size
                .HorizontalAlignment = chgAlignment(objCol.Alignment)
                .VerticalAlignment = -4108
            End With
            intColumnSub = intColumnSub + 1
        End If
        SetProgress 6 + (intColumnSub / GrdHead.Cols) * 10
    Next
    
    ExcelGrid GrdHead, gobjSend.UnderAppRows.Count + 2, 0, GrdHead.Cols, 0, GrdHead.FixedRows, 16, 9
    
    '设置网格线
    lngY = gobjSend.UnderAppRows.Count + GrdHead.FixedRows + objSource.RecordCount + 1
    ExcelApp.Range("A" & Trim(str(gobjSend.UnderAppRows.Count + 2)) & ":" & RowInttoStr(objDbGrd.Columns.Count) & Trim(str(lngY))).Select
    With ExcelApp.Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        If objDbGrd.Columns.Count > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
        If lngY <> gobjSend.UnderAppRows.Count + 2 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        '2007/04/20:刘兴宏加入,主要是导到wps时,是虚线,需要将其设置线的精细程度
        If gblnIsWps Then
            .Borders.Weight = 2
        End If
    End With
    SetProgress 26
    
    '表头输出
    lngY = gobjSend.UnderAppRows.Count + 2
    For iCount = 0 To GrdHead.FixedRows - 1
        GrdHead.Row = iCount
        For intColcnt = 0 To GrdHead.Cols - 1
            GrdHead.Col = intColcnt
            strTemp = RowInttoStr(intColcnt + 1) & Trim(str(lngY))
            ExcelApp.Range(strTemp).Select
            With ExcelApp.Selection
                .Font.Name = GrdHead.CellFontName
                .Font.Size = GrdHead.CellFontSize
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = -4108
            End With
        Next
        SetProgress 26 + iCount / GrdHead.FixedRows * 4
        lngY = lngY + 1
    Next
    
    '表格头部格式
    For intColcnt = 0 To GrdHead.Cols - 1
        GrdHead.Col = intColcnt
        For iCount = 0 To GrdHead.FixedRows - 1
            GrdHead.Row = iCount
            strTemp = RowInttoStr(intColcnt + 1) & Trim(CStr(iCount + gobjSend.UnderAppRows.Count + 2))
            ExcelApp.Range(strTemp).Select
            With ExcelApp.Selection
                .NumberFormatLocal = "@"
                .Font.Name = GrdHead.CellFontName
                .Font.Size = GrdHead.CellFontSize
                .HorizontalAlignment = -4108
                .VerticalAlignment = -4108
            End With
        Next
        SetProgress 30 + (intColcnt / GrdHead.Cols) * 17
    Next
    '标题的格式
    ExcelApp.Range("A1:" & RowInttoStr(intColcnt) & "1").Select
    With ExcelApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = -4108
        .MergeCells = True
        .Font.Name = gobjSend.Title.Font.Name
        .Font.Size = gobjSend.Title.Font.Size
        .RowHeight = GetstrHeight(.Font.Name, .Font.Size, "刘") * 2 / 20
        
    End With
    SetProgress 48
    '表上项目输出
    lngY = 2
    For iCount = 1 To gobjSend.UnderAppRows.Count
        ExcelApp.Range("A" & Trim(str(lngY)) & ":" & RowInttoStr(intColcnt) & Trim(str(lngY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
        End With
        lngY = lngY + 1
    Next
    SetProgress 49
    '表下项目输出
    
    lngY = gobjSend.UnderAppRows.Count + GrdHead.FixedRows + objSource.RecordCount + 2
    For iCount = 1 To gobjSend.BelowAppRows.Count
        ExcelApp.Range("A" & Trim(str(lngY)) & ":" & RowInttoStr(intColcnt) & Trim(str(lngY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
        End With
        lngY = lngY + 1
    Next
End Function


Private Function ExcelLvw()
    '------------------------------------------------
    '功能： listview对象的输出
    '参数：
    '------------------------------------------------
    '----------------------------------------------------
    '   变量设置
    '----------------------------------------------------
    Dim lvwTemp As Object
    Dim lvwHeader As Object
    Dim lvwListitem As Object
    Dim lngWidth As Long
    Dim strTemp As String
    Dim intY As Long
    Dim intCol As Long
    Dim i As Long
    Dim intColumnSum As Long
    
    Set lvwTemp = gobjSend.Body.objData
    mlngWideTable = 0
    
    For intY = 1 To lvwTemp.ColumnHeaders.Count
        For i = 1 To lvwTemp.ColumnHeaders.Count
            Set lvwHeader = lvwTemp.ColumnHeaders(i)
            If lvwHeader.Position = intY Then Exit For
        Next
        mlngWideTable = mlngWideTable + lvwHeader.Width
        ExcelApp.Columns(RowInttoStr(intY) & ":" & RowInttoStr(intY)).Select
        ExcelApp.Selection.ColumnWidth = lvwHeader.Width / mintChar
        intColumnSum = intColumnSum + 1
        SetProgress 3 + (intY / lvwTemp.ColumnHeaders.Count) * 3
    Next
    
    '设置网格线
    intY = IIf(gobjSend.UnderAppItems.Count > 0, 1, 0) + lvwTemp.ListItems.Count + 2
    ExcelApp.Range("A" & Trim(str(IIf(gobjSend.UnderAppItems.Count > 0, 1, 0) + 2)) & ":" & _
        RowInttoStr(lvwTemp.ColumnHeaders.Count) & Trim(str(intY))).Select
    With ExcelApp.Selection
            .Font.Name = gobjSend.Body.Font.Name
            .Font.Size = gobjSend.Body.Font.Size
            
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            If lvwTemp.ColumnHeaders.Count > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
            If intY <> IIf(gobjSend.UnderAppItems.Count > 0, 1, 0) + 2 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            '2007/04/20:刘兴宏加入,主要是导到wps时,是虚线,需要将其设置线的精细程度
            If gblnIsWps Then
                .Borders.Weight = 2
            End If
    End With
    SetProgress 6
    
    'listitems输出
    Dim strCol As String
    For intCol = 1 To lvwTemp.ColumnHeaders.Count
        For i = 1 To lvwTemp.ColumnHeaders.Count
            Set lvwHeader = lvwTemp.ColumnHeaders(i)
            If lvwHeader.Position = intCol Then Exit For
        Next
        
        For Each lvwListitem In lvwTemp.ListItems
            If lvwHeader.SubItemIndex = 0 Then
                strCol = lvwListitem.Text
            Else
                strCol = lvwListitem.SubItems(lvwHeader.SubItemIndex)
            End If
            '不是空白就有格式
            If Trim(strCol) <> "" Then Exit For
        Next
        strTemp = RowInttoStr(intCol)
        strTemp = strTemp & ":" & strTemp
        ExcelApp.Columns(strTemp).Select
        With ExcelApp.Selection
            If IsNumeric(strCol) And Len(strCol) <= 15 Then
                i = InStr(strCol, ".")
                If i = 0 Then
                    .NumberFormatLocal = "0_ "
                Else
                    .NumberFormatLocal = "0." & String(Len(Mid(strCol, i + 1)), "0") & "_ "
                End If
            ElseIf IsDate(strCol) Then
                If InStr(strCol, ":") > 0 Or InStr(strCol, "分") > 0 Then
                    If InStr(strCol, "-") > 0 Then
                        .NumberFormatLocal = "yyyy-mm-dd hh:mm:ss"
                    Else
                        .NumberFormatLocal = "yyyy""年""mm""月""dd""日"" hh""时""mm""分""ss""秒"""
                    End If
                Else
                    If InStr(strCol, "-") > 0 Then
                        .NumberFormatLocal = "yyyy-mm-dd"
                    Else
                        .NumberFormatLocal = "yyyy""年""mm""月""dd""日"""
                    End If
                End If
            Else
                .NumberFormatLocal = "@"
            End If
            .HorizontalAlignment = chgAlignment(lvwHeader.Alignment)
            .VerticalAlignment = -4108
        End With
        SetProgress 6 + (intCol / lvwTemp.ColumnHeaders.Count) * 34
    Next 'intCol
    '标题的格式
    ExcelApp.Range("A1:" & RowInttoStr(lvwTemp.ColumnHeaders.Count) & "1").Select
    With ExcelApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = -4108
        .MergeCells = True
        .Font.Name = gobjSend.Title.Font.Name
        .Font.Size = gobjSend.Title.Font.Size
        .RowHeight = GetstrHeight(.Font.Name, .Font.Size, "刘") * 2 / 20
    End With
    SetProgress 43
    '表上项目输出
    If gobjSend.UnderAppItems.Count > 0 Then
        ExcelApp.Range("A2:" & RowInttoStr(lvwTemp.ColumnHeaders.Count) & "2").Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
            
        End With
    End If
    SetProgress 46
    '表下项目输出
    intY = IIf(gobjSend.UnderAppItems.Count > 0, 1, 0) + lvwTemp.ListItems.Count + 3
    If gobjSend.BelowAppItems.Count > 0 Then
        ExcelApp.Range("A" & Trim(str(intY)) & ":" & RowInttoStr(lvwTemp.ColumnHeaders.Count) & Trim(str(intY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
        End With
    End If
    SetProgress 49
End Function

Private Function File1Grd(ByVal intFileNumber As Integer)
    '------------------------------------------------
    '功能： 单msFlexGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim GrdTemp As Object
    Dim iCount As Long, intColcnt As Long, intSum As Long, i As Long
    Dim strText As String
    Dim AppRowTemp As zlTabAppRow
    
    Set GrdTemp = gobjSend.Body
  
    '标题输出
    Put intFileNumber, , Replace(Replace(Replace(gobjSend.Title.Text, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
    Put intFileNumber, , vbCrLf
    SetProgress 52
    
    '表上项目输出
    For Each AppRowTemp In gobjSend.UnderAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    SetProgress 53
    
    '表头输出
    intSum = GrdTemp.FixedRows
    For iCount = 0 To GrdTemp.FixedRows - 1
        For intColcnt = 0 To GrdTemp.Cols - 1
            If GrdTemp.ColWidth(intColcnt) <= 0 Or mdlCommon.ColHidden(GrdTemp, intColcnt) Then
                '...
            Else
                strText = GrdTemp.TextMatrix(iCount, intColcnt)
                Put intFileNumber, , Replace(Replace(Replace(strText, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
                Put intFileNumber, , vbTab
            End If
        Next
        Put intFileNumber, , vbCrLf
        SetProgress 53 + (iCount / intSum * 7)
    Next
    
    'Excel输出预设格式
    iCount = 0
    For i = 0 To GrdTemp.Cols - 1
        If GrdTemp.ColWidth(i) <= 0 Or mdlCommon.ColHidden(GrdTemp, i) Then
            '...
        Else
            iCount = iCount + 1
        End If
    Next
    If iCount <> 0 Then
        ReDim arrFormat(iCount - 1, 1) As Integer
        For i = 0 To iCount - 1
            arrFormat(i, 0) = i + 1 '列号
            arrFormat(i, 1) = 1 '列数据类型：xlGeneralFormat = 1,xlTextFormat = 2
        Next
    End If
    
    '网格内容输出
    intSum = GrdTemp.Rows - GrdTemp.FixedRows
    For iCount = GrdTemp.FixedRows To GrdTemp.Rows - 1
        i = 0
        For intColcnt = 0 To GrdTemp.Cols - 1
            If GrdTemp.ColWidth(intColcnt) <= 0 Or mdlCommon.ColHidden(GrdTemp, intColcnt) Then
                '...
            Else
                strText = GrdTemp.TextMatrix(iCount, intColcnt)
                If IsNumeric(Trim(strText)) And Len(Trim(strText)) <= 15 Then
                    strText = Trim(strText)
                End If
                strText = Replace(Replace(Replace(strText, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
                
               '带空格的数字强行处理为文本格式
                If IsNumeric(strText) And Left(strText, 2) = "  " Then
                    arrFormat(i, 1) = 2
                End If
            
                Put intFileNumber, , strText
                Put intFileNumber, , vbTab
                
                '侦测类型:只要有一个不是数字,或长度大于15,则为文本
                strText = Trim(strText)
                If strText <> "" And (Not IsNumeric(strText) Or Len(strText) > 15) _
                    And InStr(",合计,小计,", "," & Replace(strText, " ", "") & ",") = 0 Then
                    arrFormat(i, 1) = 2
                End If
                i = i + 1
            End If
        Next
        Put intFileNumber, , vbCrLf
        SetProgress 60 + (iCount / intSum * 34)
    Next
    
    '表下项目输出
    For Each AppRowTemp In gobjSend.BelowAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    SetProgress 94
End Function

Private Function File2Grd(ByVal intFileNumber As Integer)
    '------------------------------------------------
    '功能： 两个msFlexGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim GrdHead As Object, GrdBody As Object
    Dim iCount As Long, intColcnt As Long, intSum As Long, i As Long
    Dim strText As String
    Dim AppRowTemp As zlTabAppRow
    
    Set GrdHead = gobjSend.BodyHead
    Set GrdBody = gobjSend.BodyGrid
  
    '标题输出
    Put intFileNumber, , Replace(Replace(Replace(gobjSend.Title.Text, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
    Put intFileNumber, , vbCrLf
    SetProgress 52
    
    '表上项目输出
    For Each AppRowTemp In gobjSend.UnderAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    SetProgress 53
    
    '表头输出
    intSum = GrdHead.FixedRows
    For iCount = 0 To GrdHead.FixedRows - 1
        For intColcnt = 0 To GrdHead.Cols - 1
            If GrdHead.ColWidth(intColcnt) <= 0 Or mdlCommon.ColHidden(GrdHead, intColcnt) Then
                '...
            Else
                strText = GrdHead.TextMatrix(iCount, intColcnt)
                Put intFileNumber, , Replace(Replace(Replace(strText, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
                Put intFileNumber, , vbTab
            End If
        Next
        Put intFileNumber, , vbCrLf
        SetProgress 53 + (iCount / intSum * 7)
    Next
    
    'Excel输出预设格式
    iCount = 0
    For i = 0 To GrdBody.Cols - 1
        If GrdBody.ColWidth(i) <= 0 Or mdlCommon.ColHidden(GrdBody, i) Then
            '...
        Else
            iCount = iCount + 1
        End If
    Next
    If iCount <> 0 Then
        ReDim arrFormat(iCount - 1, 1) As Integer
        For i = 0 To iCount - 1
            arrFormat(i, 0) = i + 1 '列号
            arrFormat(i, 1) = 1 '列数据类型：xlGeneralFormat = 1,xlTextFormat = 2
        Next
    End If
    
    '网格内容输出
    intSum = GrdBody.Rows - GrdBody.FixedRows
    For iCount = 0 To GrdBody.Rows - 1
        i = 0
        For intColcnt = 0 To GrdBody.Cols - 1
            If GrdBody.ColWidth(intColcnt) <= 0 Or mdlCommon.ColHidden(GrdBody, intColcnt) Then
                '...
            Else
                strText = GrdBody.TextMatrix(iCount, intColcnt)
                If IsNumeric(Trim(strText)) And Len(Trim(strText)) <= 15 Then
                    strText = Trim(strText)
                End If
                strText = Replace(Replace(Replace(strText, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
                '带空格的数字强行处理为文本格式
                If IsNumeric(strText) And Left(strText, 2) = "  " Then
                    arrFormat(i, 1) = 2
                End If
                
                Put intFileNumber, , strText
                Put intFileNumber, , vbTab
                
                '侦测类型:只要有一个不是数字,或长度大于15,则为文本
                strText = Trim(strText)
                If strText <> "" And (Not IsNumeric(strText) Or Len(strText) > 15) Then
                    arrFormat(i, 1) = 2
                End If
                i = i + 1
            End If
        Next
        Put intFileNumber, , vbCrLf
        SetProgress 60 + (iCount / intSum * 34)
    Next
    
    '表下项目输出
    For Each AppRowTemp In gobjSend.BelowAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    SetProgress 94
End Function

Private Function FileGrds(ByVal intFileNumber As Integer)
    '------------------------------------------------
    '功能： 多个msFlexGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim GrdTemp As Object
    Dim iCount As Long, intColcnt As Long, intGrid As Long
    Dim blnRedraw As Boolean
    Dim strText As String
    Dim intSum As Long, intCount As Long, intMax As Long, i As Long
    Dim AppRowTemp As zlTabAppRow
    
    Set GrdTemp = gobjSend.Grds(1)
  
    '标题输出
    Put intFileNumber, , Replace(Replace(Replace(gobjSend.Title.Text, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
    Put intFileNumber, , vbCrLf
    SetProgress 52
    
    '表上项目输出
    For Each AppRowTemp In gobjSend.UnderAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    SetProgress 53
    
    'Excel输出预设格式
    intMax = 0
    For Each GrdTemp In gobjSend.Grds
        iCount = 0
        For i = 0 To GrdTemp.Cols - 1
            If GrdTemp.ColWidth(i) <= 0 Or mdlCommon.ColHidden(GrdTemp, i) Then
                '...
            Else
                iCount = iCount + 1
            End If
        Next
        If iCount > intMax Then intMax = iCount
    Next
    iCount = intMax
    If iCount <> 0 Then
        ReDim arrFormat(iCount - 1, 1) As Integer
        For i = 0 To iCount - 1
            arrFormat(i, 0) = i + 1 '列号
            arrFormat(i, 1) = 1 '列数据类型：xlGeneralFormat = 1,xlTextFormat = 2
        Next
    End If
    
    '网格内容输出
    intCount = gobjSend.Grds.Count
    For Each GrdTemp In gobjSend.Grds
        intSum = GrdTemp.Rows
        blnRedraw = GrdTemp.Redraw
        intGrid = intGrid + 1
        For iCount = 0 To GrdTemp.Rows - 1
            i = 0
            For intColcnt = 0 To GrdTemp.Cols - 1
                If GrdTemp.ColWidth(intColcnt) <= 0 Or mdlCommon.ColHidden(GrdTemp, intColcnt) Then
                    '...
                Else
                    strText = GrdTemp.TextMatrix(iCount, intColcnt)
                    If IsNumeric(Trim(strText)) And Len(Trim(strText)) <= 15 Then
                        strText = Trim(strText)
                    End If
                    strText = Replace(Replace(Replace(strText, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
                    '带空格的数字强行处理为文本格式
                    If IsNumeric(strText) And Left(strText, 2) = "  " Then
                        arrFormat(i, 1) = 2
                    End If
            
                    Put intFileNumber, , strText
                    Put intFileNumber, , vbTab
                    
                    '侦测类型:只要有一个不是数字,或长度大于15,则为文本
                    strText = Trim(strText)
                    If strText <> "" And (Not IsNumeric(strText) Or Len(strText) > 15) Then
                        arrFormat(i, 1) = 2
                    End If
                    i = i + 1
                End If
            Next
            Put intFileNumber, , vbCrLf
            SetProgress 53 + (iCount / intSum) * (intGrid / intCount) * 41
        Next
        If intGrid < gobjSend.Grds.Count Then
            Put intFileNumber, , vbCrLf
        End If
        
        GrdTemp.Redraw = blnRedraw
    Next
    
    '表下项目输出
    For Each AppRowTemp In gobjSend.BelowAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    SetProgress 94
End Function

Private Function FileDbGrd(ByVal intFileNumber As Long)
    '------------------------------------------------
    '功能： 单DBGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim objDbGrd As Object, objSource As Object, objCol As Object
    Dim intColcnt As Long, intCol As Long
    Dim strTemp As String
    Dim intSum As Long, iCount As Long, i As Long
    
    Set objDbGrd = gobjSend.BodyGrid
    Set objSource = gobjSend.DataSource
    
    Dim lngPos As Long '保存记录位置
    
    '标题输出
    Put intFileNumber, , Replace(Replace(Replace(gobjSend.Title.Text, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
    Put intFileNumber, , vbCrLf
    SetProgress 52
    
    '表上项目输出
    Dim AppRowTemp As zlTabAppRow
    For Each AppRowTemp In gobjSend.UnderAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    SetProgress 53
    
    '表头输出
    intSum = objDbGrd.Columns.Count
    If objDbGrd.ColumnHeaders Then
        For intCol = 0 To objDbGrd.Columns.Count - 1
            Set objCol = objDbGrd.Columns(intCol)
            If objCol.Visible Then
                Put intFileNumber, , Replace(Replace(Replace(objCol.Caption, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
                Put intFileNumber, , vbTab
            End If
            SetProgress 53 + (intCol / intSum * 2)
        Next
        Put intFileNumber, , vbCrLf
    End If
    
    'Excel输出预设格式
    iCount = 0
    For i = 0 To objDbGrd.Columns.Count - 1
        If objDbGrd.Columns(i).Visible Then iCount = iCount + 1
    Next
    If iCount <> 0 Then
        ReDim arrFormat(iCount - 1, 1) As Integer
        For i = 0 To iCount - 1
            arrFormat(i, 0) = i + 1 '列号
            arrFormat(i, 1) = 1 '列数据类型：xlGeneralFormat = 1,xlTextFormat = 2
        Next
    End If
    
    '网格内容输出
    lngPos = objDbGrd.Row
    intSum = objSource.RecordCount
    For iCount = 0 To intSum - 1
        i = 0
        For intCol = 0 To objDbGrd.Columns.Count - 1
            Set objCol = objDbGrd.Columns(intCol)
            If objCol.Visible Then
                strTemp = objCol.CellText(objDbGrd.GetBookmark(iCount - lngPos))
                 
                '带空格的数字强行处理为文本格式
                If IsNumeric(strTemp) And Left(strTemp, 2) = "  " Then
                    arrFormat(i, 1) = 2
                End If
            
                Put intFileNumber, , strTemp
                Put intFileNumber, , vbTab
                
                '侦测类型:只要有一个不是数字,或长度大于15,则为文本
                strTemp = Trim(strTemp)
                If strTemp <> "" And (Not IsNumeric(strTemp) Or Len(strTemp) > 15) Then
                    arrFormat(i, 1) = 2
                End If
                i = i + 1
            End If
        Next
        Put intFileNumber, , vbCrLf
        SetProgress 55 + (iCount / intSum * 39)
    Next
    
    '表下项目输出
    For Each AppRowTemp In gobjSend.BelowAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    SetProgress 94
End Function

Private Function FileFlxDb(ByVal intFileNumber As Integer)
    '------------------------------------------------
    '功能： DBGrid和fsFlexGrid组合对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    Dim GrdHead As Object, objDbGrd As Object, objSource As Object, objCol As Object
    Dim strText As String
    Dim iCount As Long, intColcnt As Long, intSum As Long, i As Long
    Dim AppRowTemp As zlTabAppRow
    Dim blnIsVSFlex As Boolean
    Dim lngPos As Long '保存记录位置
    
    Set GrdHead = gobjSend.BodyHead
    Set objDbGrd = gobjSend.BodyGrid
    Set objSource = gobjSend.DataSource
    
    '标题输出
    Put intFileNumber, , Replace(Replace(Replace(gobjSend.Title.Text, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
    Put intFileNumber, , vbCrLf
    SetProgress 52
    
    '表上项目输出
    For Each AppRowTemp In gobjSend.UnderAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    SetProgress 53
    
    '表头输出
    blnIsVSFlex = UCase(TypeName(GrdHead)) = MSTR_VSF Or UCase(TypeName(GrdHead)) = UCase("MSHFlexGrid")
    intSum = GrdHead.FixedRows
    For iCount = 0 To GrdHead.FixedRows - 1
        If blnIsVSFlex = False Then GrdHead.Row = iCount
        For intColcnt = 0 To GrdHead.Cols - 1
            If GrdHead.ColWidth(intColcnt) <> 0 Then
                If blnIsVSFlex Then
                    strText = GrdHead.TextMatrix(iCount, intColcnt)
                Else
                    GrdHead.Col = intColcnt
                    strText = GrdHead.Text
                End If
                Put intFileNumber, , Replace(Replace(Replace(strText, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
                Put intFileNumber, , vbTab
            End If
        Next
        SetProgress 53 + (iCount / intSum * 2)
        Put intFileNumber, , vbCrLf
    Next
    
    'Excel输出预设格式
    iCount = 0
    For i = 0 To objDbGrd.Columns.Count - 1
        If objDbGrd.Columns(i).Visible Then iCount = iCount + 1
    Next
    If iCount <> 0 Then
        ReDim arrFormat(iCount - 1, 1) As Integer
        For i = 0 To iCount - 1
            arrFormat(i, 0) = i + 1 '列号
            arrFormat(i, 1) = 1 '列数据类型：xlGeneralFormat = 1,xlTextFormat = 2
        Next
    End If
    
    '网格内容输出
    lngPos = objDbGrd.Row
    intSum = objSource.RecordCount
    For iCount = 0 To intSum - 1
        i = 0
        For intColcnt = 0 To objDbGrd.Columns.Count - 1
            Set objCol = objDbGrd.Columns(intColcnt)
            If objCol.Visible Then
                strText = Replace(Replace(Replace(objCol.CellText(objDbGrd.GetBookmark(iCount - lngPos)), vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
                 
                '带空格的数字强行处理为文本格式
                If IsNumeric(strText) And Left(strText, 2) = "  " Then
                    arrFormat(i, 1) = 2
                End If
                
                Put intFileNumber, , strText
                Put intFileNumber, , vbTab
                
                '侦测类型:只要有一个不是数字,或长度大于15,则为文本
                strText = Trim(strText)
                If strText <> "" And (Not IsNumeric(strText) Or Len(strText) > 15) Then
                    arrFormat(i, 1) = 2
                End If
                i = i + 1
            End If
        Next
        Put intFileNumber, , vbCrLf
        SetProgress 55 + (iCount / intSum * 39)
    Next
    
    '表下项目输出
    For Each AppRowTemp In gobjSend.BelowAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    SetProgress 94
End Function

Private Function FileLvw(ByVal intFileNumber As Integer)
    '------------------------------------------------
    '功能： listview对象的输出
    '参数：
    '------------------------------------------------
    Dim lvwTemp As Object
    Dim lvwHeader As Object
    Dim lvwListitem As Object
    Dim intCol As Long, i As Long
    Dim intSum As Long, intCount As Long
    Dim strText As String
    
    Set lvwTemp = gobjSend.Body.objData
    '标题输出
    Put intFileNumber, , Replace(Replace(Replace(gobjSend.Title.Text, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
    Put intFileNumber, , vbCrLf
    SetProgress 52
    '表上项目输出
    If gobjSend.UnderAppItems.Count > 0 Then
        Put intFileNumber, , Replace(Replace(Replace(strRow(gobjSend.UnderAppItems), vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
        Put intFileNumber, , vbCrLf
    End If
    SetProgress 53
    
    '表头输出
    For intCol = 1 To lvwTemp.ColumnHeaders.Count
        For i = 1 To lvwTemp.ColumnHeaders.Count
            Set lvwHeader = lvwTemp.ColumnHeaders(i)
            If lvwHeader.Position = intCol Then Exit For
        Next
        Put intFileNumber, , Replace(Replace(Replace(lvwHeader.Text, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
        Put intFileNumber, , vbTab
    Next
    Put intFileNumber, , vbCrLf
    SetProgress 55
    
    'Excel输出预设格式
    ReDim arrFormat(lvwTemp.ColumnHeaders.Count - 1, 1) As Integer
    For i = 0 To lvwTemp.ColumnHeaders.Count - 1
        arrFormat(i, 0) = i + 1 '列号
        arrFormat(i, 1) = 1 '列数据类型：xlGeneralFormat = 1,xlTextFormat = 2
    Next
    
    'ListItems输出
    intSum = lvwTemp.ListItems.Count
    For Each lvwListitem In lvwTemp.ListItems
        For intCol = 1 To lvwTemp.ColumnHeaders.Count
            For i = 1 To lvwTemp.ColumnHeaders.Count
                Set lvwHeader = lvwTemp.ColumnHeaders(i)
                If lvwHeader.Position = intCol Then Exit For
            Next
            If lvwHeader.SubItemIndex = 0 Then
                strText = Replace(Replace(Replace(lvwListitem.Text, vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
            Else
                strText = Replace(Replace(Replace(lvwListitem.SubItems(lvwHeader.SubItemIndex), vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
            End If
            
            '带空格的数字强行处理为文本格式
            If IsNumeric(strText) And Left(strText, 2) = "  " Then
                arrFormat(lvwHeader.SubItemIndex, 1) = 2
            End If
            
            Put intFileNumber, , strText
            Put intFileNumber, , vbTab
            
            '侦测类型:只要有一个不是数字,或长度大于15,则为文本
            strText = Trim(strText)
            If strText <> "" And (Not IsNumeric(strText) Or Len(strText) > 15) Then
                arrFormat(lvwHeader.SubItemIndex, 1) = 2
            End If
        Next
        Put intFileNumber, , vbCrLf
        intCount = intCount + 1
        SetProgress 55 + (intCount / intSum * 39)
    Next
    
    '表下项目输出
    If gobjSend.BelowAppItems.Count > 0 Then
        Put intFileNumber, , Replace(Replace(Replace(strRow(gobjSend.BelowAppItems), vbCrLf, "<换行分隔符>"), vbLf, "<换行分隔符>"), vbCr, "<换行分隔符>")
        Put intFileNumber, , vbCrLf
    End If
    SetProgress 94
End Function

Private Function strRow(zlrow As zlTabAppRow) As String
    '------------------------------------------------
    '功能：求出一行表上或表下项目在EXCEL中输出时的内容
    '参数：
    '------------------------------------------------

    Dim varTempstr As Variant
    Dim intCurrentWidth As Long
    Dim strBlankSum As String
    
    gfrmTemp.Font.Name = gobjSend.AppFont.Name
    gfrmTemp.Font.Size = CCur(gobjSend.AppFont.Size)
    
    For Each varTempstr In zlrow
        intCurrentWidth = intCurrentWidth + CInt(gfrmTemp.TextWidth(varTempstr.Text)) '此时得到的是有效字符长度
    Next
    intCurrentWidth = (mlngWideTable - intCurrentWidth) / CInt(gfrmTemp.TextWidth("0"))  '此时得到的是空格数
    If intCurrentWidth < 0 Then intCurrentWidth = 0
    strBlankSum = String(intCurrentWidth, " ")
    
    gfrmTemp.Font.Name = "宋体"
    gfrmTemp.Font.Size = 9
    
    For Each varTempstr In zlrow
        If zlrow.Count = 1 Then
            strRow = strRow & varTempstr.Text
        Else
            strRow = strRow & varTempstr.Text & String(intCurrentWidth / (zlrow.Count - 1), " ")
        End If
    Next
    strRow = RTrim(strRow)
End Function

Private Function RowInttoStr(ByVal intRow As Long) As String
    '------------------------------------------------
    '功能：将数字列号转换为EXCEL中的表示方法
    '参数：
    '------------------------------------------------

    Dim int1 As Long, int2 As Long
    If intRow < 27 Then
        RowInttoStr = Chr(Asc("A") + intRow - 1)
        Exit Function
    End If
    
    int1 = intRow \ 26
    int2 = intRow Mod 26
    If int2 = 0 Then
        int1 = int1 - 1
        int2 = 26
    End If
    RowInttoStr = Chr(Asc("A") + int1 - 1) & Chr(Asc("A") + int2 - 1)
End Function
Private Function chgAlignment(ByVal GrdAlignment As Integer) As Long
    '------------------------------------------------
    '功能：将VB的对齐方式转换为EXCEL中的对齐方式
    '参数：GrdAlignment     FlexGrid的对齐方式
    '返回：EXCEL中的对齐方式
    '------------------------------------------------

    Select Case GrdAlignment
         Case 0       '左对齐
            chgAlignment = -4131
         Case 2        '居中
            chgAlignment = -4108
         Case 1        '右对齐
            chgAlignment = -4152
        Case Else
            chgAlignment = -4131
    End Select
End Function

Private Function GrdAlignment(objGrid As Object, ByVal lngCol As Long, ByVal lngRow As Long) As Long
    '------------------------------------------------
    '功能：将FlexGrid的对齐方式转换为EXCEL中的对齐方式
    '参数：objGrid      FlexGrid对象
    '返回：EXCEL中的对齐方式
    '------------------------------------------------
    Dim Alignment As Integer
    Dim blnVSF As Boolean
        
    If UCase(TypeName(objGrid)) = UCase("MSHFlexGrid") Then
        objGrid.Row = lngRow
        objGrid.Col = lngCol
        blnVSF = False
    Else
        blnVSF = True
    End If
        
    '【容错处理】检查有无定位到单元格，如果没有使用CellAlignment属性会出错
    If AvailableAlignment(objGrid) = False Then
        Alignment = flexAlignLeftTop    '缺省
    Else
        '获取对齐属性：
        If objGrid.CellAlignment <> 0 And objGrid.CellAlignment <> 9 Then
            Alignment = objGrid.CellAlignment                       '参照单元格
        Else
            If lngCol < objGrid.FixedCols Or lngRow < objGrid.FixedRows Then
                If blnVSF = False Then
                    Alignment = objGrid.ColAlignmentFixed(lngCol)   '参照固定单元格
                Else
                    Alignment = objGrid.ColAlignment(lngCol)
                End If
            Else
                Alignment = objGrid.ColAlignment(lngCol)            '参照列
            End If
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
        If lngRow >= objGrid.Rows Then lngRow = objGrid.Rows - 1
        If lngCol >= objGrid.Cols Then lngCol = objGrid.Cols - 1
        If IsNumeric(Trim(objGrid.TextMatrix(lngRow, lngCol))) Then
            GrdAlignment = -4152
        Else
            GrdAlignment = -4131
        End If
    Case Else
        GrdAlignment = -4131
    End Select
End Function

Private Sub SetProgress(ByVal i As Integer)
    '------------------------------------------------
    '功能：显示进度
    '参数：i        进度值
    '返回：
    '------------------------------------------------

    mintProgress = i
    If mintProgress > 100 Then mintProgress = 100
    frmTempExcel.lblnum = Trim(str(mintProgress))
    frmTempExcel.ProgressBar1.Value = mintProgress
    frmTempExcel.Refresh
End Sub

Private Function GetstrHeight(strFontName As String, curFontSize As Currency, str As String) As Long
'功能：根据字体设置返回一个字符串高度
    Dim objFrm As frmAbout
    
    Set objFrm = New frmAbout
    objFrm.Font.Name = strFontName
    objFrm.Font.Size = curFontSize
    GetstrHeight = objFrm.TextHeight(str)
    Unload objFrm
    Set objFrm = Nothing
End Function

Private Function AvailableAlignment(ByVal objGrid As Object) As Boolean
    Dim i As Integer
    
    AvailableAlignment = False
    If objGrid Is Nothing Then Exit Function
    
    i = -1
    On Error Resume Next
    i = objGrid.CellAlignment
    AvailableAlignment = i >= 0
    On Error GoTo 0
End Function


