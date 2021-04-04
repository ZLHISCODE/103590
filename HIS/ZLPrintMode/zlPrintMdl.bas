Attribute VB_Name = "zlPrintMdl"
Option Explicit

Public Const conLineWide As Integer = 30        '横线所占宽度(单位为缇)占两条线宽度
Public Const conLineHigh As Integer = 30        '竖线所占高度(单位为缇)占两条线高度
Public Const conRatemmToTwip As Single = 56.6857142857143      '毫米与缇的比率
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public gblnIsWps As Boolean      '刘兴宏加入:判断是否向WPS中传数据

Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90

Public Const conSize1 = "信笺， 8 1/2 x 11 英寸"
Public Const conSize2 = "+A611 小型信笺， 8 1/2 x 11 英寸"
Public Const conSize3 = "小型报， 11 x 17 英寸"
Public Const conSize4 = "分类帐， 17 x 11 英寸"
Public Const conSize5 = "法律文件， 8 1/2 x 14 英寸"
Public Const conSize6 = "声明书，5 1/2 x 8 1/2 英寸"
Public Const conSize7 = "行政文件，7 1/2 x 10 1/2 英寸"
Public Const conSize8 = "A3, 297 x 420 毫米"
Public Const conSize9 = "A4, 210 x 297 毫米"
Public Const conSize10 = "A4小号， 210 x 297 毫米"
Public Const conSize11 = "A5, 148 x 210 毫米"
Public Const conSize12 = "B4, 250 x 354 毫米"
Public Const conSize13 = "B5, 182 x 257 毫米"
Public Const conSize14 = "对开本， 8 1/2 x 13 英寸"
Public Const conSize15 = "四开本， 215 x 275 毫米"
Public Const conSize16 = "10 x 14 英寸"
Public Const conSize17 = "11 x 17 英寸"
Public Const conSize18 = "便条，8 1/2 x 11 英寸"
Public Const conSize19 = "#9 信封， 3 7/8 x 8 7/8 英寸"
Public Const conSize20 = "#10 信封， 4 1/8 x 9 1/2 英寸"
Public Const conSize21 = "#11 信封， 4 1/2 x 10 3/8 英寸"
Public Const conSize22 = "#12 信封， 4 1/2 x 11 英寸"
Public Const conSize23 = "#14 信封， 5 x 11 1/2 英寸"
Public Const conSize24 = "C 尺寸工作单"
Public Const conSize25 = "D 尺寸工作单"
Public Const conSize26 = "E 尺寸工作单"
Public Const conSize27 = "DL 型信封， 110 x 220 毫米"
Public Const conSize28 = "C5 型信封， 162 x 229 毫米"
Public Const conSize29 = "C3 型信封， 324 x 458 毫米"
Public Const conSize30 = "C4 型信封， 229 x 324 毫米"
Public Const conSize31 = "C6 型信封， 114 x 162 毫米"
Public Const conSize32 = "C65 型信封，114 x 229 毫米"
Public Const conSize33 = "B4 型信封， 250 x 353 毫米"
Public Const conSize34 = "B5 型信封，176 x 250 毫米"
Public Const conSize35 = "B6 型信封， 176 x 125 毫米"
Public Const conSize36 = "信封， 110 x 230 毫米"
Public Const conSize37 = "信封大王， 3 7/8 x 7 1/2 英寸"
Public Const conSize38 = "信封， 3 5/8 x 6 1/2 英寸"
Public Const conSize39 = "U.S. 标准复写簿， 14 7/8 x 11 英寸"
Public Const conSize40 = "德国标准复写簿， 8 1/2 x 12 英寸"
Public Const conSize41 = "德国法律复写簿， 8 1/2 x 13 英寸"

Public Const conBin1 = "上层纸盒进纸"
Public Const conBin2 = "下层纸盒进纸"
Public Const conBin3 = "中间纸盒进纸"
Public Const conBin4 = "等待手动插入每页纸"
Public Const conBin5 = "信封进纸器进纸"
Public Const conBin6 = "信封进纸器进纸；但要等待手动插入"
Public Const conBin7 = "当前缺省纸盒进纸"
Public Const conBin8 = "拖拉进纸器进纸"
Public Const conBin9 = "小型进纸器进纸"
Public Const conBin10 = "大型纸盒进纸"
Public Const conBin11 = "大容量进纸器进纸"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PaperName                  根据当前打印机的设置，获取纸张名称
'PaperSource                根据当前打印机的设置，获取送纸方式描述
'zlPutPrinterSet            向系统注册表中保存打印缺省设置
'PrintLvw                   listview对象的输出
'Print1Grd                  单MSFlexGrid对象的输出
'Print2Grd                  两个MSFlexGrid对象的输出
'PrintGrds                  多msFlexGrid对象的输出
'PrintDBGrd                 单DBGrid对象的输出
'PrintFlxDB                 DBGrid和fsFlexGrid组合对象的输出
'GridCellPrint              分析并打印网格的一个单元
'PrintCell                  按指定坐标打印一个数据单元,并将当前坐标移动到单元右上角位置
'HaveExcel                  判断本机上装有EXCEL没有
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function PaperName() As String
    '------------------------------------------------
    '功能： 根据当前打印机的设置，获取纸张名称
    '参数：
    '返回： 纸张名称
    '------------------------------------------------
    Dim mSize As Integer
    Err = 0
    On Error GoTo errHand
    
    If Printer.PaperSize = 256 Then
        PaperName = "用户自定义，" _
            & Printer.Width / 56.6857142857143 & "x" _
            & Printer.Height / 56.6857142857143 & "毫米"
        Exit Function
    End If
    If Printer.PaperSize >= 1 And Printer.PaperSize <= 41 Then
        mSize = Printer.PaperSize
        PaperName = IIf(Printer.Orientation = 1, "纵向", "横向") & Space(2) _
            & Switch( _
            mSize = 1, conSize1, mSize = 2, conSize2, mSize = 3, conSize3, mSize = 4, conSize4, mSize = 5, conSize5, _
            mSize = 6, conSize6, mSize = 7, conSize7, mSize = 8, conSize8, mSize = 9, conSize9, mSize = 10, conSize10, _
            mSize = 11, conSize11, mSize = 12, conSize12, mSize = 13, conSize13, mSize = 14, conSize14, mSize = 15, conSize15, _
            mSize = 16, conSize16, mSize = 17, conSize17, mSize = 18, conSize18, mSize = 19, conSize19, mSize = 20, conSize20, _
            mSize = 21, conSize21, mSize = 22, conSize22, mSize = 23, conSize23, mSize = 24, conSize24, mSize = 25, conSize25, _
            mSize = 26, conSize26, mSize = 27, conSize27, mSize = 28, conSize28, mSize = 29, conSize29, mSize = 30, conSize30, _
            mSize = 31, conSize31, mSize = 32, conSize32, mSize = 33, conSize33, mSize = 34, conSize34, mSize = 35, conSize35, _
            mSize = 36, conSize36, mSize = 37, conSize37, mSize = 38, conSize38, mSize = 39, conSize39, mSize = 40, conSize40, _
            mSize = 41, conSize41)
        Exit Function
    End If
errHand:
    PaperName = "不可测的纸张"

End Function

Public Function PaperSource() As String
    '------------------------------------------------
    '功能： 根据当前打印机的设置，获取送纸方式描述
    '参数：
    '返回： 送纸方式字符串
    '------------------------------------------------
    Dim mBin As Integer
    
    Err = 0
    On Error GoTo errHand
    
    If Printer.PaperBin = 14 Then
        PaperSource = "附加的卡式纸盒进纸"
        Exit Function
    End If
    If Printer.PaperBin >= 1 And Printer.PaperBin <= 11 Then
        PaperSource = Switch( _
            mBin = 1, conBin1, mBin = 2, conBin2, mBin = 3, conBin3, mBin = 4, conBin4, mBin = 5, conBin5, _
            mBin = 6, conBin6, mBin = 7, conBin7, mBin = 8, conBin8, mBin = 9, conBin9, mBin = 10, conBin10, _
            mBin = 11, conBin11)
        Exit Function
    End If
errHand:
    PaperSource = "不可测的进纸方式"

End Function

Public Function zlPutPrinterSet() As Boolean
    '------------------------------------------------
    '功能：向系统注册表中保存打印缺省设置
    '------------------------------------------------
    If Printers.Count = 0 Then
        zlPutPrinterSet = False
        Exit Function
    End If
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperSize", Printer.PaperSize
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperBin", Printer.PaperBin
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Orientation", Printer.Orientation
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Width", Printer.Width
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Height", Printer.Height
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PageHead", gstrHeader
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PageFoot", gstrFooter
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PageUp", gsngUp         '上边距 '以毫米为单位
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PageDown", gsngDown     '下边距   '以毫米为单位
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PageLeft", gsngLeft     '左边距 '以毫米为单位
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PageRight", gsngRight   '右边边距   '以毫米为单位

    zlPutPrinterSet = True
End Function

Public Function PrintLvw(ByVal PageRow As Long, ByVal PageCol As Long) As Boolean
    '------------------------------------------------
    '功能： listview对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------
    
    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    '   变量设置
    '----------------------------------------------------
    
    Dim intColcnt   As Integer   '列计数器
    Dim iCount      As Long   '自由计数器
    Dim i As Long
    Dim sngHeight   As Single
    Dim CellsForward As New Collection   '向前搜索产生已经打印的单元
    Dim objLvw As Object
    Dim objLitem As ListItem
    Dim objHeader As ColumnHeader
    
    Set objLvw = gobjSend.Body
    With gobjOutTo
        .FontName = objLvw.Font.Name
        .FontSize = objLvw.Font.Size * gsngScale
        .FontBold = objLvw.Font.Bold
        .FontItalic = objLvw.Font.Italic
        sngHeight = .TextHeight("A") * gobjSend.RowSpaceRate * 1.5
        '表头输出
        For iCount = 1 To gintFixRow
            .CurrentX = gsngLeft * conRatemmToTwip
            For intColcnt = 1 To gintFixCol
                For i = 1 To objLvw.objData.ColumnHeaders.Count
                    Set objHeader = objLvw.objData.ColumnHeaders(i)
                    If objHeader.Position = intColcnt Then Exit For
                Next
                PrintCell objHeader.Text, .CurrentX, .CurrentY, _
                    objHeader.Width, sngHeight, objHeader.Alignment, , gobjSend.GridColor, _
                    RGB(195, 195, 195), String(4, Trim(CStr(gobjSend.GridLines)))

            Next
            For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                For i = 1 To objLvw.objData.ColumnHeaders.Count
                    Set objHeader = objLvw.objData.ColumnHeaders(i)
                    If objHeader.Position = intColcnt Then Exit For
                Next
                PrintCell objHeader.Text, .CurrentX, .CurrentY, _
                    objHeader.Width, sngHeight, objHeader.Alignment, , gobjSend.GridColor, _
                    RGB(195, 195, 195), String(4, Trim(CStr(gobjSend.GridLines)))
            Next
            
            .CurrentY = .CurrentY + sngHeight
        Next
        sngHeight = (.TextHeight("A") + 2 * conLineHigh) * gobjSend.RowSpaceRate
        '表格内容输出
        Dim strTemp As String
        For iCount = gintRow(1, PageRow) To gintRow(2, PageRow)
            .CurrentX = gsngLeft * conRatemmToTwip
            For intColcnt = 1 To gintFixCol
                For i = 1 To objLvw.objData.ColumnHeaders.Count
                    Set objHeader = objLvw.objData.ColumnHeaders(i)
                    If objHeader.Position = intColcnt Then Exit For
                Next
                If iCount = 1 Then
                    strTemp = objHeader.Text
                Else
                    If objHeader.Index = 1 Then
                        strTemp = objLvw.objData.ListItems(iCount - 1).Text
                    Else
                        strTemp = objLvw.objData.ListItems(iCount - 1).SubItems(objHeader.Index - 1)
                    End If
                End If
                PrintCell strTemp, .CurrentX, .CurrentY, _
                    objHeader.Width, sngHeight, objHeader.Alignment, , gobjSend.GridColor, _
                    RGB(195, 195, 195), String(4, Trim(CStr(gobjSend.GridLines)))
            Next
            
            For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                For i = 1 To objLvw.objData.ColumnHeaders.Count
                    Set objHeader = objLvw.objData.ColumnHeaders(i)
                    If objHeader.Position = intColcnt Then Exit For
                Next
                If iCount = 1 Then
                    strTemp = objHeader.Text
                Else
                    If objHeader.Index = 1 Then
                        strTemp = objLvw.objData.ListItems(iCount - 1).Text
                    Else
                        strTemp = objLvw.objData.ListItems(iCount - 1).SubItems(objHeader.Index - 1)
                    End If
                End If
                PrintCell strTemp, .CurrentX, .CurrentY, _
                    objHeader.Width, sngHeight, objHeader.Alignment, , gobjSend.GridColor, _
                    RGB(195, 195, 195), String(4, Trim(CStr(gobjSend.GridLines)))
            Next
            .CurrentY = .CurrentY + sngHeight
        Next
    End With
    
    PrintLvw = True
    Exit Function

errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    PrintLvw = False

End Function


Public Function Print1Grd(ByVal PageRow As Long, ByVal PageCol As Long) As Boolean
    '------------------------------------------------
    '功能： 单MSFlexGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------

    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    '   变量设置
    '----------------------------------------------------
    
    Dim intColcnt   As Integer   '列计数器
    Dim iCount      As Long   '自由计数器
    Dim CellsForward As New Collection   '向前搜索产生已经打印的单元
        
    With gobjOutTo
                
        '表头输出
        If gobjSend.PageCols > 0 Then
            '临床路径表单使用
            For iCount = gintFixedRow(1, PageRow) To gintFixedRow(2, PageRow)
                .CurrentX = gsngLeft * conRatemmToTwip
                For intColcnt = 1 To gintFixCol
                    GridCellPrint gobjSend.Body, iCount - 1, intColcnt - 1, CellsForward
                Next
                For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                    GridCellPrint gobjSend.Body, iCount - 1, intColcnt - 1, CellsForward, , gintCol(2, PageCol)
                Next
                
                .CurrentY = .CurrentY + gobjSend.Body.RowHeight(iCount - 1)
            Next
        Else
            For iCount = 1 To gintFixRow
                .CurrentX = gsngLeft * conRatemmToTwip
                For intColcnt = 1 To gintFixCol
                    GridCellPrint gobjSend.Body, iCount - 1, intColcnt - 1, CellsForward
                Next
                For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                    GridCellPrint gobjSend.Body, iCount - 1, intColcnt - 1, CellsForward, , gintCol(2, PageCol)
                Next
                .CurrentY = .CurrentY + gobjSend.Body.RowHeight(iCount - 1)
            Next
        End If
        
        '表格内容输出
        For iCount = gintRow(1, PageRow) To gintRow(2, PageRow)
            .CurrentX = gsngLeft * conRatemmToTwip
            For intColcnt = 1 To gintFixCol
                GridCellPrint gobjSend.Body, iCount - 1, intColcnt - 1, CellsForward, gintRow(2, PageRow)
            Next
            
            For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                    GridCellPrint gobjSend.Body, iCount - 1, intColcnt - 1, CellsForward, gintRow(2, PageRow), gintCol(2, PageCol)
            Next
            .CurrentY = .CurrentY + gobjSend.Body.RowHeight(iCount - 1)
        Next
    End With
    
    Print1Grd = True
    Exit Function

errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    Print1Grd = False

End Function


Public Function Print2Grd(ByVal PageRow As Long, ByVal PageCol As Long) As Boolean
    '------------------------------------------------
    '功能： 两个MSFlexGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------

    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    '   变量设置
    '----------------------------------------------------
    
    Dim intColcnt   As Integer   '列计数器
    Dim iCount      As Long   '自由计数器
    Dim CellsForward As New Collection   '向前搜索产生已经打印的单元
        
    With gobjOutTo
                
        '表头输出
        For iCount = 1 To gintFixRow
            .CurrentX = gsngLeft * conRatemmToTwip
            For intColcnt = 1 To gintFixCol
                GridCellPrint gobjSend.BodyHead, iCount - 1, intColcnt - 1, CellsForward
            Next
            For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                GridCellPrint gobjSend.BodyHead, iCount - 1, intColcnt - 1, CellsForward, , gintCol(2, PageCol)
            Next
            
            .CurrentY = .CurrentY + gobjSend.BodyHead.RowHeight(iCount - 1)
        Next
        
        '表格内容输出
        For iCount = gintRow(1, PageRow) To gintRow(2, PageRow)
            .CurrentX = gsngLeft * conRatemmToTwip
            For intColcnt = 1 To gintFixCol
                If iCount <= gobjSend.BodyHead.FixedRows Then
                    GridCellPrint gobjSend.BodyHead, iCount - 1, intColcnt - 1, CellsForward, gintRow(2, PageRow)
                Else
                    GridCellPrint gobjSend.BodyGrid, iCount - gobjSend.BodyHead.FixedRows - 1, intColcnt - 1, CellsForward, gintRow(2, PageRow), gintCol(2, PageCol)
                End If
            Next
            
            For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                If iCount <= gobjSend.BodyHead.FixedRows Then
                    GridCellPrint gobjSend.BodyHead, iCount - 1, intColcnt - 1, CellsForward, gintRow(2, PageRow), gintCol(2, PageCol)
                Else
                    GridCellPrint gobjSend.BodyGrid, iCount - gobjSend.BodyHead.FixedRows - 1, intColcnt - 1, CellsForward, gintRow(2, PageRow), gintCol(2, PageCol)
                End If
            Next
            If iCount <= gobjSend.BodyHead.FixedRows Then
                .CurrentY = .CurrentY + gobjSend.BodyHead.RowHeight(iCount - 1)
            Else
                .CurrentY = .CurrentY + gobjSend.BodyGrid.RowHeight(iCount - gobjSend.BodyHead.FixedRows - 1)
            End If
        Next
    End With
    
    Print2Grd = True
    Exit Function

errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    Print2Grd = False

End Function

Public Function PrintGrds(ByVal PageRow As Long, ByVal PageCol As Long) As Boolean
    '------------------------------------------------
    '功能： 多msFlexGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------

    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    '   变量设置
    '----------------------------------------------------
    
    Dim intColcnt   As Integer   '列计数器
    Dim iCount      As Long   '自由计数器
    Dim objGrid As Object
    Dim CellsForward As Collection  '向前搜索产生已经打印的单元
    Set CellsForward = New Collection
    
    Dim intRows As Long, intCols As Long
    With gobjOutTo
    For Each objGrid In gobjSend.Grds
        
        '表格内容输出
        intRows = objGrid.Rows
        intCols = objGrid.Cols
        For iCount = 1 To intRows
            .CurrentX = gsngLeft * conRatemmToTwip
            For intColcnt = 1 To intCols
                GridCellPrint objGrid, iCount - 1, intColcnt - 1, CellsForward, intRows, intCols
            Next
            .CurrentY = .CurrentY + objGrid.RowHeight(iCount - 1)
        Next
        .CurrentY = .CurrentY + 300
    Next
    End With
    
    PrintGrds = True
    Exit Function

errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    PrintGrds = False

End Function



Public Function PrintDBGrd(ByVal PageRow As Long, ByVal PageCol As Long) As Boolean
    '------------------------------------------------
    '功能： 单DBGrid对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------

    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    '   变量设置
    '----------------------------------------------------
    
    Dim intColcnt   As Integer   '列计数器
    Dim iCount      As Long   '自由计数器
    Dim iTemp As Long
    Dim CellsForward As New Collection   '向前搜索产生已经打印的单元
    
    Dim objDbGrd As Object, objCol As Object, objSource As Object
    Set objDbGrd = gobjSend.BodyGrid
    Set objSource = gobjSend.DataSource
    
    Dim lngGrdColor As Long
    Dim sngHighFont As Single
    Dim sngHighHead As Single
    
    Select Case objDbGrd.RowDividerStyle
    Case 0
        lngGrdColor = objDbGrd.BackColor
    Case 1
        lngGrdColor = RGB(0, 0, 0)
    Case 2
        lngGrdColor = RGB(80, 80, 80)
    Case 5
        lngGrdColor = objDbGrd.ForeColor
    Case Else
        lngGrdColor = RGB(110, 110, 110)
    End Select
    
    With gobjOutTo
              
        If objDbGrd.ColumnHeaders Then
            .FontName = objDbGrd.HeadFont.Name
            .FontSize = objDbGrd.HeadFont.Size * gsngScale
            .FontBold = objDbGrd.HeadFont.Bold
            .FontItalic = objDbGrd.HeadFont.Italic
            sngHighFont = .TextHeight("赵") + conLineHigh
        End If
        sngHighHead = sngHighFont * objDbGrd.HeadLines
        '表头输出
        For iCount = 1 To gintFixRow
            .CurrentX = gsngLeft * conRatemmToTwip
            intColcnt = 0
            For iTemp = 1 To gintTotalCol
                Set objCol = objDbGrd.Columns(iTemp - 1)
                If objCol.Visible Then
                    intColcnt = intColcnt + 1
                    If intColcnt <= gintFixCol Or (intColcnt >= gintCol(1, PageCol) And intColcnt <= gintCol(2, PageCol)) Then
                        PrintCell objCol.Caption, .CurrentX, .CurrentY, objCol.Width, sngHighHead, _
                            objCol.Alignment, , RGB(128, 128, 128), RGB(220, 220, 220), "1111"
                    End If
                End If
            Next
            .CurrentY = .CurrentY + sngHighHead
        Next
        
        '表格内容输出
        Dim iRowInTb As Long
        Dim j As Long
        iRowInTb = objSource.AbsolutePosition
        For iCount = gintRow(1, PageRow) To gintRow(2, PageRow)
            .CurrentX = gsngLeft * conRatemmToTwip
            If iCount = 1 Then
                intColcnt = 0
                For iTemp = 1 To gintTotalCol
                    Set objCol = objDbGrd.Columns(iTemp - 1)
                    If objCol.Visible Then
                        intColcnt = intColcnt + 1
                        If intColcnt <= gintFixCol Or (intColcnt >= gintCol(1, PageCol) And intColcnt <= gintCol(2, PageCol)) Then
                            PrintCell objCol.Caption, .CurrentX, .CurrentY, objCol.Width, sngHighHead, _
                                objCol.Alignment, , RGB(128, 128, 128), RGB(220, 220, 220), "1111"
                        End If
                    End If
                Next
                .CurrentY = .CurrentY + sngHighHead
            Else
                
                '求出书签的位置
                j = iCount - iRowInTb - 1
                intColcnt = 0
                For iTemp = 1 To gintTotalCol
                    Set objCol = objDbGrd.Columns(iTemp - 1)
                    If objCol.Visible Then
                        intColcnt = intColcnt + 1
                        If intColcnt <= gintFixCol Or (intColcnt >= gintCol(1, PageCol) And intColcnt <= gintCol(2, PageCol)) Then
                            PrintCell objCol.CellText(objDbGrd.GetBookmark(j)), .CurrentX, .CurrentY, objCol.Width, CSng(objDbGrd.RowHeight), _
                                objCol.Alignment, objDbGrd.ForeColor, lngGrdColor, objDbGrd.BackColor, "1111"
                        End If
                    End If
                Next
                .CurrentY = .CurrentY + CSng(objDbGrd.RowHeight)
            End If
        Next
    End With
    
    PrintDBGrd = True
    Exit Function

errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    PrintDBGrd = False

End Function

Public Function PrintFlxDB(ByVal PageRow As Long, ByVal PageCol As Long) As Boolean
    '------------------------------------------------
    '功能： DBGrid和fsFlexGrid组合对象的输出
    '参数：
    '返回： 成功返回true ；错误返回false
    '------------------------------------------------


    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    '   变量设置
    '----------------------------------------------------
    
    Dim intColcnt   As Integer   '列计数器
    Dim iCount      As Long   '自由计数器
    Dim iTemp As Long
    Dim CellsForward As New Collection   '向前搜索产生已经打印的单元
        
    Dim objDbGrd As Object, objCol As Object, objSource As Object
    Set objDbGrd = gobjSend.BodyGrid
    Set objSource = gobjSend.DataSource
    
    Dim lngGrdColor As Long
    Dim sngHighFont As Single
    Dim sngHighHead As Single
    
    Select Case objDbGrd.RowDividerStyle
    Case 0
        lngGrdColor = objDbGrd.BackColor
    Case 1
        lngGrdColor = RGB(0, 0, 0)
    Case 2
        lngGrdColor = RGB(80, 80, 80)
    Case 5
        lngGrdColor = objDbGrd.ForeColor
    Case Else
        lngGrdColor = RGB(110, 110, 110)
    End Select
    
    With gobjOutTo
                
        '表头输出
        For iCount = 1 To gintFixRow
            .CurrentX = gsngLeft * conRatemmToTwip
            For intColcnt = 1 To gintFixCol
                GridCellPrint gobjSend.BodyHead, iCount - 1, intColcnt - 1, CellsForward
            Next
            For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                GridCellPrint gobjSend.BodyHead, iCount - 1, intColcnt - 1, CellsForward, , gintCol(2, PageCol)
            Next
            
            .CurrentY = .CurrentY + gobjSend.BodyHead.RowHeight(iCount - 1)
        Next
        
        
        '表格内容输出
        Dim iRowInTb As Long
        Dim j As Long
        iRowInTb = objSource.AbsolutePosition
        For iCount = gintRow(1, PageRow) To gintRow(2, PageRow)
            .CurrentX = gsngLeft * conRatemmToTwip
            If iCount <= gobjSend.BodyHead.FixedRows Then
                For intColcnt = 1 To gintFixCol
                    GridCellPrint gobjSend.BodyHead, iCount - 1, intColcnt - 1, CellsForward
                Next
                For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                    GridCellPrint gobjSend.BodyHead, iCount - 1, intColcnt - 1, CellsForward
                Next
                .CurrentY = .CurrentY + gobjSend.BodyHead.RowHeight(iCount - 1)
            Else
                '求出书签的位置
                j = iCount - iRowInTb - gobjSend.BodyHead.FixedRows
                intColcnt = 0
                For iTemp = 1 To gintTotalCol
                    Set objCol = objDbGrd.Columns(iTemp - 1)
                    If objCol.Visible Then
                        intColcnt = intColcnt + 1
                        If intColcnt <= gintFixCol Or (intColcnt >= gintCol(1, PageCol) And intColcnt <= gintCol(2, PageCol)) Then
                            PrintCell objCol.CellText(objDbGrd.GetBookmark(j)), .CurrentX, .CurrentY, objCol.Width, CSng(objDbGrd.RowHeight), _
                                objCol.Alignment, objDbGrd.ForeColor, lngGrdColor, objDbGrd.BackColor, "1111"
                        End If
                    End If
                Next
                .CurrentY = .CurrentY + CSng(objDbGrd.RowHeight)
            End If
        Next
        
    End With


    PrintFlxDB = True
    Exit Function

errHand:
    MsgBox "系统出现不可预知的错误" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    PrintFlxDB = False

End Function

Public Sub GridCellPrint(objGrid As Object, Row As Long, Col As Long, _
    AcrossCells As Collection, Optional MaxRow, Optional MaxCol)
    '------------------------------------------------
    '功能： 分析并打印网格的一个单元
    '参数：
    '   objGrid:需要输出的MSFlexGrid对象
    '   Row:行号
    '   Col:列号
    '   AcrossCells:应忽略的单元集合，他们已经作为合并单元提前打印
    '返回：
    '------------------------------------------------
    Dim iCount As Long
    For iCount = 1 To AcrossCells.Count
        If Trim(CStr(Row)) & "," & Trim(CStr(Col)) = AcrossCells.Item(iCount) Then
            gobjOutTo.CurrentX = gobjOutTo.CurrentX + objGrid.ColWidth(Col)
            AcrossCells.Remove iCount
            Exit Sub
        End If
    Next
    
    '对应于单元的变量：
    Dim Text As String
    Dim X As Long, Y As Long
    Dim Wide As Long
    Dim High As Long
    Dim Alignment As Byte
    Dim PortraitAlignment As Byte
    Dim ForeColor As Long
    Dim GridColor As Long
    Dim FillColor As Long
    Dim LineStyle As String
    Dim FontName
    Dim FontSize
    Dim FontBold
    Dim FontItalic
    
    Dim iRow As Long, iCol As Long
    Dim bln表头合并 As Boolean
    
    If IsMissing(MaxRow) Then MaxRow = gintFixRow
    If IsMissing(MaxCol) Then MaxCol = gintFixCol
    
    objGrid.Row = Row
    objGrid.Col = Col
    
    '忽略隐藏的列
    If objGrid.ColWidth(Col) <= 0 Then
        Exit Sub
    ElseIf UCase(TypeName(objGrid)) = UCase("VSFlexGrid") Then
        'ColHidden方法VSFlexGrid才有
        If objGrid.ColHidden(Col) Then
            Exit Sub
        End If
    End If
        
    '获取对齐属性：
    If objGrid.CellAlignment <> 0 And objGrid.CellAlignment <> 9 Then
        Alignment = objGrid.CellAlignment           '参照单元
    Else
        If Col < objGrid.FixedCols Or Row < objGrid.FixedRows Or objGrid.Rowdata(Row) = UCase("FIXEDROW") Then
            Alignment = objGrid.FixedAlignment(Col) '参照固定单元
        Else
            Alignment = objGrid.ColAlignment(Col)   '参照列
        End If
    End If
    Select Case Alignment
    Case 1, 4, 7, 9
        PortraitAlignment = 2       '中
    Case 2, 5, 8
        PortraitAlignment = 1       '下
    Case 0, 3, 6
        PortraitAlignment = 0       '上
    End Select
    Select Case Alignment
    Case 0, 1, 2        '左对齐
        Alignment = 0
    Case 3, 4, 5        '居中
        Alignment = 2
    Case 6, 7, 8        '右对齐
        Alignment = 1
    Case 9
        If IsNumeric(Trim(objGrid.Text)) Then
            Alignment = 1
        Else
            Alignment = 0
        End If
    Case Else
            Alignment = 0
    End Select
    
    '获取背景色：
    If CLng(objGrid.CellBackColor) <> 0 Then
        FillColor = objGrid.CellBackColor
    Else
        If Col < objGrid.FixedCols Or Row < objGrid.FixedRows Or objGrid.Rowdata(Row) = UCase("FIXEDROW") Then
            FillColor = objGrid.BackColorFixed
        Else
            FillColor = objGrid.BackColor
        End If
    End If
    
    '获取前景色：
    If CLng(objGrid.CellForeColor) <> 0 Then
        ForeColor = objGrid.CellForeColor
    Else
        If Col < objGrid.FixedCols Or Row < objGrid.FixedRows Or objGrid.Rowdata(Row) = UCase("FIXEDROW") Then
            ForeColor = objGrid.ForeColorFixed
        Else
            ForeColor = objGrid.ForeColor
        End If
    End If
    
    '网格线颜色：
    If Col < objGrid.FixedCols Or Row < objGrid.FixedRows Or objGrid.Rowdata(Row) = UCase("FIXEDROW") Then
        GridColor = IIf(objGrid.GridLinesFixed = 1, objGrid.GridColorFixed, 0)   '参照固定线
    Else
        GridColor = IIf(objGrid.GridLines = 1, objGrid.GridColor, 0)             '参照标准线
    End If
    
    '网格线宽度：
    If Col < objGrid.FixedCols Or Row < objGrid.FixedRows Or objGrid.Rowdata(Row) = UCase("FIXEDROW") Then
        LineStyle = IIf(objGrid.GridLinesFixed = 0, "0000", "1111")         '参照固定线
    Else
        LineStyle = IIf(objGrid.GridLines = 0, "0000", "1111")              '参照标准线
    End If
    
    '向前搜索合并单元，获取正确高度、宽度：
    Text = objGrid.Text
    High = objGrid.RowHeight(Row)
    Wide = objGrid.ColWidth(Col)
    If Text <> "" And objGrid.MergeCells <> 0 Then
        If objGrid.MergeCells = 5 Then
            '功能：表头并
            If objGrid.FixedRows > Row Then
                bln表头合并 = True
            End If
        End If
        If objGrid.MergeRow(Row) Or bln表头合并 Then
            For iCol = Col + 1 To IIf(Col < objGrid.FixedCols, objGrid.FixedCols, objGrid.Cols) - 1
                If iCol > MaxCol - 1 Then Exit For
                objGrid.Col = iCol
                If Text = objGrid.Text Then
                    If objGrid.MergeCells = 3 Or objGrid.MergeCells = 4 Then
                        iCount = Row - 1
                        Do While iCount >= 0
                            If objGrid.TextMatrix(iCount, Col) <> objGrid.TextMatrix(iCount, iCol) Then Exit For
                            iCount = iCount - 1
                        Loop
                    End If
                    Wide = Wide + objGrid.ColWidth(iCol)
                    AcrossCells.Add Trim(CStr(Row)) & "," & Trim(CStr(iCol))
                Else
                    Exit For
                End If
            Next
        End If
        
        objGrid.Row = Row
        objGrid.Col = Col
        If objGrid.MergeCol(Col) Or bln表头合并 Then
            For iRow = Row + 1 To IIf(Row < objGrid.FixedRows, objGrid.FixedRows, objGrid.Rows) - 1
                If iRow > MaxRow - 1 And objGrid.Rowdata(Row) <> UCase("FIXEDROW") Then Exit For
                objGrid.Row = iRow
                If Text = objGrid.Text Then
                    If objGrid.MergeCells = 2 Or objGrid.MergeCells = 4 Then
                        iCount = Col - 1
                        Do While iCount >= 0
                            If objGrid.TextMatrix(Row, iCount) <> objGrid.TextMatrix(iRow, iCount) Then Exit For
                            iCount = iCount - 1
                        Loop
                    End If
                    High = High + objGrid.RowHeight(iRow)
                    AcrossCells.Add Trim(CStr(iRow)) & "," & Trim(CStr(Col))
                Else
                    Exit For
                End If
            Next
        End If
        objGrid.Row = Row
        objGrid.Col = Col
    End If
    '单元输出：
    Dim CurrentX As Long
    CurrentX = gobjOutTo.CurrentX
    PrintCell Text, gobjOutTo.CurrentX, gobjOutTo.CurrentY, Wide, High, Alignment, _
        ForeColor, GridColor, FillColor, LineStyle, _
        objGrid.CellFontName, objGrid.CellFontSize * gsngScale, _
        objGrid.CellFontBold, objGrid.CellFontItalic, PortraitAlignment
    gobjOutTo.CurrentX = CurrentX + objGrid.ColWidth(Col)

End Sub
Public Function CellTextRows(ByVal strText As String, ByVal Wide, ByVal High) As Variant
    '----------------------------------------------------------------------------------
    '功能:将文本按行转换成数组返回
    '参数:strText-单元格式
    '     Wide -宽度
    '     hight-高度
    '返回:单印的单元格的数组
    '编制:刘兴宏
    '日期:2007/09/04
    '----------------------------------------------------------------------------------
    Dim arrPrintRow()  As String
    Dim arrPrintText As Variant
    Dim i As Long, intRow As Integer
    Dim intAllRow As Integer, strPrintText As String, strRest As String
    Dim j As Long
    Dim strTmp As String
    
    If InStr(1, strText, vbCrLf) > 0 Then
        arrPrintText = Split(Trim(strText), vbCrLf)
    Else
        arrPrintText = Split(Trim(strText), Chr(13))
    End If
    
    j = 0
    For i = 0 To UBound(arrPrintText)
            strPrintText = arrPrintText(i)
            
            If Wide - conLineWide < gobjOutTo.TextWidth("1") Then    '小于一个字符
                intAllRow = 1
            Else
'                If gobjOutTo.TextWidth(strPrintText) Mod (Wide - conLineWide) = 0 Then
'                    intAllRow = gobjOutTo.TextWidth(strPrintText) \ (Wide - conLineWide)
'                Else
'                    intAllRow = gobjOutTo.TextWidth(strPrintText) \ (Wide - conLineWide) + 1
'                End If
                
                '计算多行内容的行数，2008-04-08 By FrChen
                strTmp = ""
                intAllRow = 0
                For intRow = 1 To Len(strPrintText)
                    If gobjOutTo.TextWidth(strTmp & Mid(strPrintText, intRow, 1)) > (Wide - conLineWide) Then
                        intAllRow = intAllRow + 1
                        strTmp = Mid(strPrintText, intRow, 1)
                    Else
                        strTmp = strTmp & Mid(strPrintText, intRow, 1)
                    End If
                Next
                If strTmp <> "" Then intAllRow = intAllRow + 1
                
            End If
            
            For intRow = intAllRow To 1 Step -1
                If High >= gobjOutTo.TextHeight(strPrintText) * intRow Then
                    Exit For
                End If
            Next
            intAllRow = intRow
            strRest = strPrintText
            For intRow = 0 To intAllRow - 1
                Do While gobjOutTo.TextWidth(strPrintText) > Wide - conLineWide
                    If Len(Trim(strPrintText)) <= 1 Then Exit Do
                    strPrintText = Left(strPrintText, Len(strPrintText) - 1)
                Loop
                strRest = Mid(strRest, Len(strPrintText) + 1)
                If intAllRow = 1 Then
                    If Len(strRest) = 1 Then
                       strPrintText = strPrintText & strRest
                    End If
                End If
                ReDim Preserve arrPrintRow(j)
                If j = 0 Then
                    arrPrintRow(j) = strPrintText
                Else
                    arrPrintRow(j) = strPrintText
                End If
                j = j + 1
                strPrintText = strRest
            Next
    Next
    intAllRow = UBound(arrPrintRow) + 1
    For intRow = intAllRow To 1 Step -1
        If High >= gobjOutTo.TextHeight(arrPrintRow(0)) * intRow Then
            Exit For
        End If
    Next
    ReDim Preserve arrPrintRow(intRow)
    CellTextRows = arrPrintRow
End Function

Public Sub PrintCell(ByVal Text As String, _
    ByVal X As Single, ByVal Y As Single, _
    Optional ByVal Wide, _
    Optional ByVal High, _
    Optional Alignment As Byte = 0, _
    Optional ForeColor As Long = 0, _
    Optional GridColor As Long = 0, _
    Optional FillColor As Long = 0, _
    Optional LineStyle As String = "1111", _
    Optional FontName, Optional FontSize, _
    Optional FontBold, Optional FontItalic, _
    Optional PortraitAlignment As Byte = 2)
    '------------------------------------------------
    '功能： 按指定坐标打印一个数据单元,并将当前坐标移动到单元右上角位置
    '参数：
    '   Text:    输出的字符串,其中不包含回车或换行符
    '   X:       左上角X坐标
    '   Y:       左上角Y坐标
    '   Wide:    输出宽度
    '   High:    输出高度
    '   Alignment:    对齐模式，0-左对齐(缺省),1-右对齐,2-居中
    '   PortraitAlignment:纵向对齐模式，0-居上;1-居下,2-居中
    '   ForeColor前景色,缺省为黑色
    '   GridColor边线色,缺省为黑色
    '   FillColor填充色,缺省为设备背景色,由于系统采用了黑色的色码，所以将不允许填充黑色
    '   LineStyle:依序分别为上左右下的线条宽度
    '           0-无线，1-9依序加粗，1为缺省
    '   FontName,FontSize,FontBold,FontItalic:字体属性
    '返回：
    '------------------------------------------------
    Dim aryString() As String       '回车分割的字符串
    Dim lngOldForeColor As Long     '输出设备缺省前景色
    Dim intRow As Long, intAllRow As Long
    Dim strRest As String, sngYMove As Single
    Dim oldFontName, oldFontSize, oldFontBold, oldFontItalic
    Dim strTmp As String
    Dim intLineWidth As Integer, intLineWidthOld As Integer
    Dim lngDC As Long, lngPPI As Long, lngDPI As Long
    
    lngOldForeColor = gobjOutTo.ForeColor
    
    On Error Resume Next
    With gobjOutTo
        If Not IsMissing(FontName) Then
            oldFontName = gobjOutTo.FontName
            .FontName = FontName
        End If
        If Not IsMissing(FontSize) Then
            .FontSize = FontSize
            oldFontSize = gobjOutTo.FontSize
        End If
        If Not IsMissing(FontBold) Then
            .FontBold = FontBold
            oldFontBold = gobjOutTo.FontBold
        End If
        If Not IsMissing(FontItalic) Then
            .FontItalic = FontItalic
            oldFontItalic = gobjOutTo.FontItalic
        End If
    End With
    
    If IsMissing(Wide) Then Wide = gobjOutTo.TextWidth(Text) + 2 * conLineWide
    If IsMissing(High) Then High = gobjOutTo.TextHeight(Text) + 2 * conLineHigh
'    Wide = CLng(Wide)
'    High = CLng(High)
    If Wide * High = 0 Then Exit Sub
    
    If UCase(TypeName(LineStyle)) <> "STRING" Then LineStyle = CStr(LineStyle)
    If Len(LineStyle) < 4 Then
        LineStyle = Left(LineStyle & "1111", 4)
    End If
    
    '获取DPI、PPI
    lngDPI = GetDeviceCaps(Printer.hDC, LOGPIXELSY)
    lngDC = GetDC(0)
    lngPPI = GetDeviceCaps(lngDC, LOGPIXELSY)
    ReleaseDC 0, lngDC
    
    '------------------------------------------
    '   边线打印
    '------------------------------------------
    intLineWidthOld = gobjOutTo.DrawWidth
    If Val(Mid(LineStyle, 1, 1)) <> 0 Then
        intLineWidth = Val(Mid(LineStyle, 1, 1))
        If TypeOf gobjOutTo Is Printer Then
            '通过PPI换算DPI的等宽值，如果lngPPI为0，默认96像素/英寸
            intLineWidth = intLineWidth * lngDPI / IIf(lngPPI = 0, 96, lngPPI)
        End If
        gobjOutTo.DrawWidth = intLineWidth
        gobjOutTo.Line (X, Y)-(X + Wide, Y), GridColor
    End If
    
    If Val(Mid(LineStyle, 2, 1)) <> 0 Then
        intLineWidth = Val(Mid(LineStyle, 2, 1))
        If TypeOf gobjOutTo Is Printer Then
            '通过PPI换算DPI的等宽值，如果lngPPI为0，默认96像素/英寸
            intLineWidth = intLineWidth * lngDPI / IIf(lngPPI = 0, 96, lngPPI)
        End If
        gobjOutTo.DrawWidth = intLineWidth
        gobjOutTo.Line (X, Y)-(X, Y + High), GridColor
    End If
    
    If Val(Mid(LineStyle, 3, 1)) <> 0 Then
        intLineWidth = Val(Mid(LineStyle, 3, 1))
        If TypeOf gobjOutTo Is Printer Then
            '通过PPI换算DPI的等宽值，如果lngPPI为0，默认96像素/英寸
            intLineWidth = intLineWidth * lngDPI / IIf(lngPPI = 0, 96, lngPPI)
        End If
        gobjOutTo.DrawWidth = intLineWidth
        gobjOutTo.Line (X + Wide, Y)-(X + Wide, Y + High), GridColor
    End If
    
    If Val(Mid(LineStyle, 4, 1)) <> 0 Then
        intLineWidth = Val(Mid(LineStyle, 4, 1))
        If TypeOf gobjOutTo Is Printer Then
            '通过PPI换算DPI的等宽值，如果lngPPI为0，默认96像素/英寸
            intLineWidth = intLineWidth * lngDPI / IIf(lngPPI = 0, 96, lngPPI)
        End If
        gobjOutTo.DrawWidth = intLineWidth
        gobjOutTo.Line (X, Y + High)-(X + Wide, Y + High), GridColor
    End If
    
    If Wide > conLineWide And High > conLineHigh Then
        '------------------------------------------
        '   底色填充
        '------------------------------------------
'        If FillColor <> 0 Then
'            Printer.FillStyle = 1
'            gobjOutTo.Line (X + conLineWide / 2, Y + conLineHigh / 2)- _
'                (X + Wide - conLineWide / 2, Y + High - conLineHigh / 2), _
'                FillColor, BF
'        End If
        
        '------------------------------------------
        '   文字打印
        '------------------------------------------
        gobjOutTo.ForeColor = ForeColor
    
        If InStr(1, Text, vbCrLf) = 0 And InStr(1, Text, Chr(13)) = 0 Then
            If Wide - conLineWide < gobjOutTo.TextWidth("1") Then    '小于一个字符
                intAllRow = 1
            Else
'                If gobjOutTo.TextWidth(Text) Mod (Wide - conLineWide) = 0 Then
'                    intAllRow = gobjOutTo.TextWidth(Text) \ (Wide - conLineWide)
'                Else
'                    intAllRow = gobjOutTo.TextWidth(Text) \ (Wide - conLineWide) + 1
'                End If

                '计算多行内容的行数，2008-04-08 By FrChen
                strTmp = ""
                intAllRow = 0
                For intRow = 1 To Len(Text)
                    If gobjOutTo.TextWidth(strTmp & Mid(Text, intRow, 1)) > (Wide - conLineWide) Then
                        intAllRow = intAllRow + 1
                        strTmp = Mid(Text, intRow, 1)
                    Else
                        strTmp = strTmp & Mid(Text, intRow, 1)
                    End If
                Next
                If strTmp <> "" Then intAllRow = intAllRow + 1
                
            End If
            For intRow = intAllRow To 1 Step -1
                If High >= gobjOutTo.TextHeight(Text) * intRow Then
                    Exit For
                End If
            Next
            intAllRow = intRow
            
            Select Case PortraitAlignment
            Case 0
                sngYMove = conLineHigh                                                          '居上
            Case 1
                sngYMove = (High - conLineHigh - gobjOutTo.TextHeight(Text) * intAllRow)        '居下
            Case Else
                sngYMove = (High - conLineHigh - gobjOutTo.TextHeight(Text) * intAllRow) / 2    '居中
            End Select
            If sngYMove < 0 Then sngYMove = conLineHigh
            
            strRest = Text
            For intRow = 0 To intAllRow - 1
                Do While gobjOutTo.TextWidth(Text) > Wide - conLineWide
                    If Len(Trim(Text)) <= 1 Then Exit Do
                    Text = Left(Text, Len(Text) - 1)
                Loop
                strRest = Mid(strRest, Len(Text) + 1)
                Select Case Alignment
                Case 2
                    gobjOutTo.CurrentX = X + (Wide - gobjOutTo.TextWidth(Text)) / 2             '居中
                Case 1
                    gobjOutTo.CurrentX = X - conLineWide / 2 + Wide - gobjOutTo.TextWidth(Text) '居右
                Case Else
                    gobjOutTo.CurrentX = X + conLineWide / 2                                    '居左
                End Select
                gobjOutTo.CurrentY = Y + conLineHigh / 2 + sngYMove + intRow * gobjOutTo.TextHeight(Text)
                
                If intAllRow = 1 Then
                    If Len(strRest) = 1 Then
                        gobjOutTo.Print Text & strRest
                    Else
                        gobjOutTo.Print Text
                    End If
                Else
                    gobjOutTo.Print Text
                End If
                Text = strRest
            Next
        Else
            '刘兴宏:有回车答,也应该自动索减列宽
            aryString = CellTextRows(Text, Wide, High)
        
'            If InStr(1, Text, vbCrLf) > 0 Then
'                aryString = Split(Trim(Text), vbCrLf)
'            Else
'                aryString = Split(Trim(Text), Chr(13))
'            End If

            intAllRow = UBound(aryString)
            sngYMove = (High - conLineHigh - gobjOutTo.TextHeight("ZYL") * intAllRow) / 2
            
            strRest = Text
            For intRow = 0 To intAllRow
                strRest = aryString(intRow)
                Select Case Alignment
                Case 2
                    Dim blnLR As Boolean
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        blnLR = Not blnLR
                        strRest = IIf(blnLR, Left(strRest, Len(strRest) - 1), Right(strRest, Len(strRest) - 1))
                    Loop
                    gobjOutTo.CurrentX = X + (Wide - gobjOutTo.TextWidth(strRest)) / 2
                Case 1
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        strRest = Right(strRest, Len(strRest) - 1)
                    Loop
                    gobjOutTo.CurrentX = X - conLineWide / 2 + Wide - gobjOutTo.TextWidth(strRest)
                Case Else
                    Do While Wide < gobjOutTo.TextWidth(strRest)
                        strRest = Left(strRest, Len(strRest) - 1)
                    Loop
                    gobjOutTo.CurrentX = X + conLineWide / 2
                End Select
                
                gobjOutTo.CurrentY = Y + conLineHigh / 2 + sngYMove + intRow * gobjOutTo.TextHeight(strRest)
                If gobjOutTo.CurrentY + gobjOutTo.TextHeight(strRest) > Y + High Then Exit For
                If gobjOutTo.CurrentY >= Y Then gobjOutTo.Print strRest
            
            Next
        End If
    End If
    gobjOutTo.CurrentX = X + Wide
    gobjOutTo.CurrentY = Y
    gobjOutTo.DrawStyle = 0
    gobjOutTo.DrawWidth = intLineWidthOld
    gobjOutTo.ForeColor = lngOldForeColor

    If Not IsMissing(FontName) Then gobjOutTo.FontName = oldFontName
    If Not IsMissing(FontSize) Then gobjOutTo.FontSize = oldFontSize
    If Not IsMissing(FontBold) Then gobjOutTo.FontBold = oldFontBold
    If Not IsMissing(FontItalic) Then gobjOutTo.FontItalic = oldFontItalic

End Sub

Public Function HaveExcel() As Boolean
    '------------------------------------------------
    '功能：判断本机上装有EXCEL没有
    '参数：
    '返回：有则返回True
    '------------------------------------------------

    On Error GoTo errHandle1
    Dim objTemp  As Object
    gblnIsWps = False
    Set objTemp = CreateObject("Excel.Application") '打开一个EXCEL程序
    Set objTemp = Nothing
    HaveExcel = True
    Exit Function

errHandle1:

    '刘兴宏:2007/4/20
    '以WPS为准
    Err = 0: On Error GoTo errHand:
    Set objTemp = CreateObject("ET.Application") '打开一个WPS中的ET程序
    Set objTemp = Nothing
    HaveExcel = True
    gblnIsWps = True
    Exit Function
errHand:
    Set objTemp = Nothing
    HaveExcel = False

End Function

