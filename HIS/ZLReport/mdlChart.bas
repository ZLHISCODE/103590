Attribute VB_Name = "mdlChart"
Option Explicit
Public Enum ChartType
    柱形图 = 51 'xlColumnClustered
    三维柱形图 = -4100 'xl3DColumn
    条形图 = 57 'xlBarClustered
    折线图 = 4 'xlLine
    饼图 = 5 'xlPie
    散点图 = -4169 'xlXYScatter
    气泡图 = 15 'xlBubble
    面积图 = 1 'xlArea
    圆环图 = -4120 'xlDoughnut
    雷达图 = -4151 'xlRadar
    曲面图 = 83 'xlSurface
    股价图 = 88 'xlStockHLC
    圆柱图 = 92 'xlCylinderColClustered
    圆锥图 = 99 'xlConeColClustered
    棱锥图 = 106 'xlPyramidColClustered
End Enum
Private Const xlEdgeLeft As Long = 7
Private Const xlEdgeTop As Long = 8
Private Const xlEdgeBottom As Long = 9
Private Const xlEdgeRight As Long = 10
Private Const xlInsideVertical As Long = 11
Private Const xlInsideHorizontal As Long = 12
Private Const xlContinuous As Long = 1
Private Const xlCenter As Long = -4108
Private Const xlLocationAsNewSheet As Long = 1
Private Const xlRows As Long = 1
Private Const xlColumns As Long = 2
Private Const xlColumnClustered As Long = 51

Public Function ExcelChart(frmMain As Object, objBody As Object, objHead As Object, ByVal lngMode As Long, ByVal strTitle As String, Optional ByVal chtType As ChartType = 柱形图) As Boolean
'功能：利用Excel显示图表
'参数：frmMain   进度窗口显示的所有者
'      objBody   数据显示的主体
'      objHead   对于清册表，还有表头
'      lngMode   报表类型。1-汇总表  2-清册表
'      strTitle  标题
'      chtType   图表的风格
    Dim objApplication As Object
    Dim LngRows As Long, lngCols As Long            '行的总数，列的总数
    Dim lngFixedRows As Long, lngFixedCols As Long  '标题行的总数，标题列的总数
    Dim strFile As String, strPath As String, intFileNum As Integer '与临时文件有关的变量
    
    
    Dim lngRow As Long, lngCol As Long            '行计数器，列计数器
    Dim lngExcelCount As Long, lngCount As Long   '计数器
    Dim lngTemp As Long
    Dim strTemp As String
    
    '创建Excel对象
    On Error Resume Next
    Set objApplication = CreateObject("Excel.Application")
    If Err <> 0 Then
        Err.Clear
        MsgBox "打开Excel程序失败。" & vbCrLf & "可能是程序未正确安装，或系统资源不足。", vbExclamation, App.Title
        Exit Function
    End If
    objApplication.Workbooks.Add
    Call ShowFlash("正在为创建Excel图表做准备……", 0.01, frmMain, True)
    
    '保存控件的属性
    objBody.Redraw = False
    If Not objHead Is Nothing Then objHead.Redraw = False
    
    On Error GoTo errHandle
    
    '产生临时文件
    strPath = Space(256): strFile = Space(256)
    GetTempPath 256, strPath
    strPath = Left$(strPath, InStr(strPath, Chr(0)) - 1)
    
    GetTempFileName strPath, "excel", 0, strFile
    strFile = Left$(strFile, InStr(strFile, Chr(0)) - 1)
    '打开文件准备输出
    intFileNum = FreeFile()
    Open strFile For Binary Access Write As intFileNum
    Call ShowFlash("正在为创建Excel图表做准备……", 0.02, frmMain, True)
    
    '输出标题
    Put intFileNum, , Replace(strTitle, vbCrLf, "") & vbCrLf
    
    
    If lngMode = 1 Then
        '汇总表的处理,除了输出以外,还有合并单元格
        
        '首先处理固定行的合并
        For lngRow = 0 To objBody.FixedRows - 1
            If objBody.RowHeight(lngRow) <> 0 Then
                lngFixedRows = lngFixedRows + 1
                lngExcelCount = 0
                lngCol = 0
                Do Until lngCol > objBody.Cols - 1
                    If objBody.ColWidth(lngCol) <> 0 And objBody.ColData(lngCol) = 0 Then
                        lngExcelCount = lngExcelCount + 1
                        lngCount = lngExcelCount    '保存第一个合并单元的列数(Excel)
                        lngTemp = lngCol            '保存第一个合并单元的列数(Grid)
                        
                        lngCol = lngCol + 1
                        Do Until lngCol > objBody.Cols - 1
                            If objBody.ColWidth(lngCol) <> 0 And objBody.ColData(lngCol) = 0 Then
                                If objBody.TextMatrix(lngRow, lngTemp) = objBody.TextMatrix(lngRow, lngCol) Then
                                    '相同合并,并检查下一行
                                    lngExcelCount = lngExcelCount + 1
                                    lngCol = lngCol + 1
                                Else
                                    Exit Do
                                End If
                            Else
                                lngCol = lngCol + 1
                            End If
                        Loop
                        If lngExcelCount <> lngCount Then
                            '进行合并
                            strTemp = Col2Excel(lngCount) & Trim(str(lngFixedRows + 1)) & ":" & Col2Excel(lngExcelCount) & Trim(str(lngFixedRows + 1))
                            objApplication.Range(strTemp).Select
                            objApplication.Selection.MergeCells = True
                            
                        End If
                    Else
                        lngCol = lngCol + 1
                    End If
                Loop
                
            End If
            Call ShowFlash("正在为创建Excel图表做准备……", 0.01 + lngRow / objBody.FixedRows * 0.05, frmMain, True)
        Next
        '首先处理固定列的合并
        For lngCol = 0 To objBody.FixedCols - 1
            If objBody.ColWidth(lngCol) <> 0 Then
                lngFixedCols = lngFixedCols + 1
                lngExcelCount = 0
                lngRow = 0
                Do Until lngRow > objBody.Rows - 1
                    If objBody.RowHeight(lngRow) <> 0 And objBody.RowData(lngRow) = 0 Then
                        lngExcelCount = lngExcelCount + 1
                        lngCount = lngExcelCount    '保存第一个合并单元的行数(Excel)
                        lngTemp = lngRow            '保存第一个合并单元的行数(Grid)
                        
                        lngRow = lngRow + 1
                        Do Until lngRow > objBody.Rows - 1
                            If objBody.RowHeight(lngRow) <> 0 And objBody.RowData(lngRow) = 0 Then
                                If objBody.TextMatrix(lngTemp, lngCol) = objBody.TextMatrix(lngRow, lngCol) Then
                                    '相同合并,并检查下一行
                                    lngExcelCount = lngExcelCount + 1
                                    lngRow = lngRow + 1
                                Else
                                    Exit Do
                                End If
                            Else
                                lngRow = lngRow + 1
                            End If
                        Loop
                        If lngExcelCount <> lngCount Then
                            '进行合并
                            strTemp = Col2Excel(lngFixedCols) & Trim(str(lngCount + 1)) & ":" & Col2Excel(lngFixedCols) & Trim(str(lngExcelCount + 1))
                            objApplication.Range(strTemp).Select
                            objApplication.Selection.MergeCells = True
                            
                        End If
                    Else
                        lngRow = lngRow + 1
                    End If
                Loop
                
            End If
            Call ShowFlash("正在为创建Excel图表做准备……", 0.05 + lngCol / objBody.FixedCols * 0.05, frmMain, True)
        Next
        
        '显示主体网络的内容
        For lngCol = 0 To objBody.Cols - 1
            If objBody.ColWidth(lngCol) <> 0 And objBody.ColData(lngCol) = 0 Then
                lngCols = lngCols + 1
            End If
        Next
        For lngRow = 0 To objBody.Rows - 1
            If objBody.RowHeight(lngRow) <> 0 And objBody.RowData(lngRow) = 0 Then
                LngRows = LngRows + 1
                For lngCol = 0 To objBody.Cols - 1
                    If objBody.ColWidth(lngCol) <> 0 And objBody.ColData(lngCol) = 0 Then
                        If lngRow >= objBody.FixedRows And lngCol >= objBody.FixedCols Then
                            '汇总表表体数据部份输出时去掉空格(否则可能影响图形生成)
                            Put intFileNum, , Replace(Trim(objBody.TextMatrix(lngRow, lngCol)), vbCrLf, "") & vbTab
                        Else
                            Put intFileNum, , Replace(objBody.TextMatrix(lngRow, lngCol), vbCrLf, "") & vbTab
                        End If
                    End If
                Next
                Put intFileNum, , vbCrLf
             End If
             Call ShowFlash("正在为创建Excel图表做准备……", 0.1 + lngRow / objBody.Rows * 0.8, frmMain, True)   '这一过程占有80%
        Next
    Else
        '清册表的处理
        '清册表的表头要单独显示
        If Not (objHead Is Nothing) Then
            LngRows = 1
            lngFixedRows = 1
            
            lngTemp = objHead.FixedRows - 1 '只要固定行中的最后一行
            For lngCount = 0 To objHead.Cols - 1
                If objHead.ColWidth(lngCount) <> 0 Then
                    lngCols = lngCols + 1
                    
                    Put intFileNum, , Replace(objHead.TextMatrix(lngTemp, lngCount), vbCrLf, "") & vbTab
                End If
            Next
            Put intFileNum, , vbCrLf
            Call ShowFlash("正在为创建Excel图表做准备……", 0.1, frmMain, True)
        End If
        '显示主体网络的内容
        lngFixedCols = 1
        For lngRow = 0 To objBody.Rows - 1
            If objBody.RowHeight(lngRow) <> 0 Then
                LngRows = LngRows + 1
                For lngCol = 0 To objBody.Cols - 1
                    If objBody.ColWidth(lngCol) <> 0 Then
                        Put intFileNum, , Replace(objBody.TextMatrix(lngRow, lngCol), vbCrLf, "") & vbTab
                    End If
                Next
                Put intFileNum, , vbCrLf
             End If
             Call ShowFlash("正在为创建Excel图表做准备……", 0.1 + lngRow / objBody.Rows * 0.8, frmMain, True)   '这一过程占有80%
        Next
    End If
    Close #intFileNum
    '合并标题行
    strTemp = "A1:" & Col2Excel(lngCols) & "1"
    objApplication.Range(strTemp).Select
    With objApplication.Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 22
    End With
    '显示网络线
        '固定行
    If lngFixedRows > 0 Then
        strTemp = "A2:" & Col2Excel(lngCols) & Trim(str(lngFixedRows + 1))
        objApplication.Range(strTemp).Select
        With objApplication.Selection
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            If lngCols > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
            If lngFixedRows > 1 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
    Call ShowFlash("正在为创建Excel图表做准备……", 0.91, frmMain, True)
    '固定行
    If lngFixedCols > 0 Then
        strTemp = "A2:" & Col2Excel(lngFixedCols) & Trim(str(LngRows + 1))
        objApplication.Range(strTemp).Select
        With objApplication.Selection
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            If lngFixedCols > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
            If LngRows > 1 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
    Call ShowFlash("正在为创建Excel图表做准备……", 0.92, frmMain, True)
    '整个表
    strTemp = "A2:" & Col2Excel(lngCols) & Trim(str(LngRows + 1))
    objApplication.Range(strTemp).Select
    With objApplication.Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
    Call ShowFlash("正在为创建Excel图表做准备……", 0.93, frmMain, True)
    
    '导入数据
    objApplication.Workbooks.OpenText strFile, , 1, 1, 1, False, True, False, False, False, False, False, Array(1, 1)
    Call ShowFlash("正在为创建Excel图表做准备……", 0.94, frmMain, True)
    With objApplication
        .Windows(2).Activate
        .Cells.Select
        .Selection.Copy
        .Windows(2).Activate
        .Cells.Select
        .Selection.PasteSpecial Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False
    End With
    Call ShowFlash("正在为创建Excel图表做准备……", 0.95, frmMain, True)
    Clipboard.Clear
    objApplication.Windows(2).Close False
    objApplication.ActiveSheet.Range("A1").Select
    Call ShowFlash("正在为创建Excel图表做准备……", 0.96, frmMain, True)
    strPath = objApplication.ActiveSheet.Name
    '显示图表
    objApplication.Charts.Add
    With objApplication.ActiveChart
        strTemp = "A2:" & Col2Excel(lngCols) & Trim(str(LngRows + 1))
        If LngRows - lngFixedRows = 1 Then
            '系列产生在行
            .SetSourceData Source:=objApplication.Sheets(strPath).Range(strTemp), PlotBy:=xlRows
        Else
            .SetSourceData Source:=objApplication.Sheets(strPath).Range(strTemp), PlotBy:=xlColumns
        End If
        .ChartType = chtType
        
        On Error Resume Next
        lngCol = 1
        If LngRows - lngFixedRows > 1 Then
            '系列产生在列
            For lngCount = lngFixedCols + 1 To lngCols
                If .SeriesCollection.Count < lngCol Then .SeriesCollection.NewSeries
                .SeriesCollection(lngCol).Values = objApplication.Sheets(strPath).Range(Col2Excel(lngCount) & Trim(str(lngFixedRows + 2)) & ":" & Col2Excel(lngCount) & Trim(str(LngRows + 1)))
                If lngFixedRows > 0 Then
                    .SeriesCollection(lngCol).Name = objApplication.Sheets(strPath).Range(Col2Excel(lngCount) & "2:" & Col2Excel(lngCount) & Trim(str(lngFixedRows + 1)))
                End If
                If Err <> 0 Then
                    Err.Clear
                End If
                lngCol = lngCol + 1
                Call ShowFlash("正在为创建Excel图表做准备……", 0.96 + lngCount / lngCols * 0.04, frmMain, True)
            Next
        End If
        .Location Where:=xlLocationAsNewSheet, Name:="图表"
        .HasDataTable = False
        .HasTitle = True
        .ChartTitle.Text = strTitle
        .ChartTitle.Font.Size = 20
        .Legend.Left = .PlotArea.Left + .PlotArea.Width
        .PlotArea.Select
        objApplication.Selection.Interior.ColorIndex = 2
        objApplication.Selection.Width = Screen.Width / 20 * (2 / 3)
        objApplication.Selection.Height = Screen.Height / 20 / 2
        Dim lng(1 To 4) As Single
        lng(1) = objApplication.Selection.Left: lng(2) = objApplication.Selection.Width
        lng(3) = objApplication.Selection.Top: lng(4) = objApplication.Selection.Height
        
        .Legend.Select
        objApplication.Selection.Left = lng(1) + lng(2) + 5
        objApplication.Selection.Top = (lng(3) + lng(4) - objApplication.Selection.Height) / 2

    End With
    
    '结束
    Call ShowFlash
    '恢复控件的属性
    objBody.Redraw = True
    If Not objHead Is Nothing Then objHead.Redraw = True
    ExcelChart = True
    objApplication.UserControl = True
    objApplication.Visible = True
    Set objApplication = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call ShowFlash
    '恢复控件的属性
    objBody.Redraw = True
    If Not objHead Is Nothing Then objHead.Redraw = True
    objApplication.DisplayAlerts = False
    objApplication.Quit
    Set objApplication = Nothing
End Function
