Attribute VB_Name = "mdlChart"
Option Explicit
Public Enum ChartType
    ����ͼ = 51 'xlColumnClustered
    ��ά����ͼ = -4100 'xl3DColumn
    ����ͼ = 57 'xlBarClustered
    ����ͼ = 4 'xlLine
    ��ͼ = 5 'xlPie
    ɢ��ͼ = -4169 'xlXYScatter
    ����ͼ = 15 'xlBubble
    ���ͼ = 1 'xlArea
    Բ��ͼ = -4120 'xlDoughnut
    �״�ͼ = -4151 'xlRadar
    ����ͼ = 83 'xlSurface
    �ɼ�ͼ = 88 'xlStockHLC
    Բ��ͼ = 92 'xlCylinderColClustered
    Բ׶ͼ = 99 'xlConeColClustered
    ��׶ͼ = 106 'xlPyramidColClustered
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

Public Function ExcelChart(frmMain As Object, objBody As Object, objHead As Object, ByVal lngMode As Long, ByVal strTitle As String, Optional ByVal chtType As ChartType = ����ͼ) As Boolean
'���ܣ�����Excel��ʾͼ��
'������frmMain   ���ȴ�����ʾ��������
'      objBody   ������ʾ������
'      objHead   �����������б�ͷ
'      lngMode   �������͡�1-���ܱ�  2-����
'      strTitle  ����
'      chtType   ͼ��ķ��
    Dim objApplication As Object
    Dim LngRows As Long, lngCols As Long            '�е��������е�����
    Dim lngFixedRows As Long, lngFixedCols As Long  '�����е������������е�����
    Dim strFile As String, strPath As String, intFileNum As Integer '����ʱ�ļ��йصı���
    
    
    Dim lngRow As Long, lngCol As Long            '�м��������м�����
    Dim lngExcelCount As Long, lngCount As Long   '������
    Dim lngTemp As Long
    Dim strTemp As String
    
    '����Excel����
    On Error Resume Next
    Set objApplication = CreateObject("Excel.Application")
    If Err <> 0 Then
        Err.Clear
        MsgBox "��Excel����ʧ�ܡ�" & vbCrLf & "�����ǳ���δ��ȷ��װ����ϵͳ��Դ���㡣", vbExclamation, App.Title
        Exit Function
    End If
    objApplication.Workbooks.Add
    Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.01, frmMain, True)
    
    '����ؼ�������
    objBody.Redraw = False
    If Not objHead Is Nothing Then objHead.Redraw = False
    
    On Error GoTo errHandle
    
    '������ʱ�ļ�
    strPath = Space(256): strFile = Space(256)
    GetTempPath 256, strPath
    strPath = Left$(strPath, InStr(strPath, Chr(0)) - 1)
    
    GetTempFileName strPath, "excel", 0, strFile
    strFile = Left$(strFile, InStr(strFile, Chr(0)) - 1)
    '���ļ�׼�����
    intFileNum = FreeFile()
    Open strFile For Binary Access Write As intFileNum
    Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.02, frmMain, True)
    
    '�������
    Put intFileNum, , Replace(strTitle, vbCrLf, "") & vbCrLf
    
    
    If lngMode = 1 Then
        '���ܱ�Ĵ���,�����������,���кϲ���Ԫ��
        
        '���ȴ���̶��еĺϲ�
        For lngRow = 0 To objBody.FixedRows - 1
            If objBody.RowHeight(lngRow) <> 0 Then
                lngFixedRows = lngFixedRows + 1
                lngExcelCount = 0
                lngCol = 0
                Do Until lngCol > objBody.Cols - 1
                    If objBody.ColWidth(lngCol) <> 0 And objBody.ColData(lngCol) = 0 Then
                        lngExcelCount = lngExcelCount + 1
                        lngCount = lngExcelCount    '�����һ���ϲ���Ԫ������(Excel)
                        lngTemp = lngCol            '�����һ���ϲ���Ԫ������(Grid)
                        
                        lngCol = lngCol + 1
                        Do Until lngCol > objBody.Cols - 1
                            If objBody.ColWidth(lngCol) <> 0 And objBody.ColData(lngCol) = 0 Then
                                If objBody.TextMatrix(lngRow, lngTemp) = objBody.TextMatrix(lngRow, lngCol) Then
                                    '��ͬ�ϲ�,�������һ��
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
                            '���кϲ�
                            strTemp = Col2Excel(lngCount) & Trim(str(lngFixedRows + 1)) & ":" & Col2Excel(lngExcelCount) & Trim(str(lngFixedRows + 1))
                            objApplication.Range(strTemp).Select
                            objApplication.Selection.MergeCells = True
                            
                        End If
                    Else
                        lngCol = lngCol + 1
                    End If
                Loop
                
            End If
            Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.01 + lngRow / objBody.FixedRows * 0.05, frmMain, True)
        Next
        '���ȴ���̶��еĺϲ�
        For lngCol = 0 To objBody.FixedCols - 1
            If objBody.ColWidth(lngCol) <> 0 Then
                lngFixedCols = lngFixedCols + 1
                lngExcelCount = 0
                lngRow = 0
                Do Until lngRow > objBody.Rows - 1
                    If objBody.RowHeight(lngRow) <> 0 And objBody.RowData(lngRow) = 0 Then
                        lngExcelCount = lngExcelCount + 1
                        lngCount = lngExcelCount    '�����һ���ϲ���Ԫ������(Excel)
                        lngTemp = lngRow            '�����һ���ϲ���Ԫ������(Grid)
                        
                        lngRow = lngRow + 1
                        Do Until lngRow > objBody.Rows - 1
                            If objBody.RowHeight(lngRow) <> 0 And objBody.RowData(lngRow) = 0 Then
                                If objBody.TextMatrix(lngTemp, lngCol) = objBody.TextMatrix(lngRow, lngCol) Then
                                    '��ͬ�ϲ�,�������һ��
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
                            '���кϲ�
                            strTemp = Col2Excel(lngFixedCols) & Trim(str(lngCount + 1)) & ":" & Col2Excel(lngFixedCols) & Trim(str(lngExcelCount + 1))
                            objApplication.Range(strTemp).Select
                            objApplication.Selection.MergeCells = True
                            
                        End If
                    Else
                        lngRow = lngRow + 1
                    End If
                Loop
                
            End If
            Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.05 + lngCol / objBody.FixedCols * 0.05, frmMain, True)
        Next
        
        '��ʾ�������������
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
                            '���ܱ�������ݲ������ʱȥ���ո�(�������Ӱ��ͼ������)
                            Put intFileNum, , Replace(Trim(objBody.TextMatrix(lngRow, lngCol)), vbCrLf, "") & vbTab
                        Else
                            Put intFileNum, , Replace(objBody.TextMatrix(lngRow, lngCol), vbCrLf, "") & vbTab
                        End If
                    End If
                Next
                Put intFileNum, , vbCrLf
             End If
             Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.1 + lngRow / objBody.Rows * 0.8, frmMain, True)   '��һ����ռ��80%
        Next
    Else
        '����Ĵ���
        '����ı�ͷҪ������ʾ
        If Not (objHead Is Nothing) Then
            LngRows = 1
            lngFixedRows = 1
            
            lngTemp = objHead.FixedRows - 1 'ֻҪ�̶����е����һ��
            For lngCount = 0 To objHead.Cols - 1
                If objHead.ColWidth(lngCount) <> 0 Then
                    lngCols = lngCols + 1
                    
                    Put intFileNum, , Replace(objHead.TextMatrix(lngTemp, lngCount), vbCrLf, "") & vbTab
                End If
            Next
            Put intFileNum, , vbCrLf
            Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.1, frmMain, True)
        End If
        '��ʾ�������������
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
             Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.1 + lngRow / objBody.Rows * 0.8, frmMain, True)   '��һ����ռ��80%
        Next
    End If
    Close #intFileNum
    '�ϲ�������
    strTemp = "A1:" & Col2Excel(lngCols) & "1"
    objApplication.Range(strTemp).Select
    With objApplication.Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 22
    End With
    '��ʾ������
        '�̶���
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
    Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.91, frmMain, True)
    '�̶���
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
    Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.92, frmMain, True)
    '������
    strTemp = "A2:" & Col2Excel(lngCols) & Trim(str(LngRows + 1))
    objApplication.Range(strTemp).Select
    With objApplication.Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
    Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.93, frmMain, True)
    
    '��������
    objApplication.Workbooks.OpenText strFile, , 1, 1, 1, False, True, False, False, False, False, False, Array(1, 1)
    Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.94, frmMain, True)
    With objApplication
        .Windows(2).Activate
        .Cells.Select
        .Selection.Copy
        .Windows(2).Activate
        .Cells.Select
        .Selection.PasteSpecial Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False
    End With
    Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.95, frmMain, True)
    Clipboard.Clear
    objApplication.Windows(2).Close False
    objApplication.ActiveSheet.Range("A1").Select
    Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.96, frmMain, True)
    strPath = objApplication.ActiveSheet.Name
    '��ʾͼ��
    objApplication.Charts.Add
    With objApplication.ActiveChart
        strTemp = "A2:" & Col2Excel(lngCols) & Trim(str(LngRows + 1))
        If LngRows - lngFixedRows = 1 Then
            'ϵ�в�������
            .SetSourceData Source:=objApplication.Sheets(strPath).Range(strTemp), PlotBy:=xlRows
        Else
            .SetSourceData Source:=objApplication.Sheets(strPath).Range(strTemp), PlotBy:=xlColumns
        End If
        .ChartType = chtType
        
        On Error Resume Next
        lngCol = 1
        If LngRows - lngFixedRows > 1 Then
            'ϵ�в�������
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
                Call ShowFlash("����Ϊ����Excelͼ����׼������", 0.96 + lngCount / lngCols * 0.04, frmMain, True)
            Next
        End If
        .Location Where:=xlLocationAsNewSheet, Name:="ͼ��"
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
    
    '����
    Call ShowFlash
    '�ָ��ؼ�������
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
    '�ָ��ؼ�������
    objBody.Redraw = True
    If Not objHead Is Nothing Then objHead.Redraw = True
    objApplication.DisplayAlerts = False
    objApplication.Quit
    Set objApplication = Nothing
End Function
