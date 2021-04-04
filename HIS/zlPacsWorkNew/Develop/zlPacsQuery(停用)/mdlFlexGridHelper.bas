Attribute VB_Name = "mdlFlexGridHelper"
Option Explicit

Public Sub GridInit(strColName As String, objGrid As VSFlexGrid)
    '初始化配置列表
    Dim i As Integer
    Dim lngCount As Long
    Dim arrData() As String
    
    arrData = Split(strColName, "|")
    lngCount = UBound(arrData) + 1
    
    With objGrid
'        .ColWidthMin = 1500
    
        .Cols = lngCount
        .Rows = 1
        .FixedRows = 1
'        .RowHeightMin = 480
'        .Cell(flexcpAlignment, 0, 0, 0, lngCount - 1) = flexAlignCenterCenter

        '最后一列自动填充满列表
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .AutoResize = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionByRow
        .AllowSelection = False
        
        For i = 0 To lngCount - 1
            .TextMatrix(0, i) = arrData(i)
        Next

        If .Rows > 1 Then .RowSel = 1
    End With
    
End Sub

Public Function NewRow(objGrid As VSFlexGrid) As Long
'新增数据行
    NewRow = -1
    
    objGrid.Rows = objGrid.Rows + 1
    objGrid.RowSel = objGrid.Rows - 1
    
    Call LocateRow(objGrid, objGrid.RowSel)
    NewRow = objGrid.RowSel
End Function

Public Sub LocateRow(objGrid As VSFlexGrid, Optional ByVal lngRowIndex As Long = -1)
'定位指定行，默认定位为最后一位
    Dim lngRow As Long
    Dim iCol As Long
    
    If objGrid.Rows <= 1 Then Exit Sub
    
    lngRow = lngRowIndex
    If lngRow < 0 Then
        lngRow = objGrid.Rows - 1
    End If
    
    '取得第一个未隐藏的列
    For iCol = 0 To objGrid.Cols - 1
        If Not objGrid.ColHidden(iCol) Then Exit For
    Next iCol
    
    Call objGrid.Select(lngRow, iCol)
    Call objGrid.ShowCell(lngRow, iCol)
End Sub

Public Sub MoveUp(objGrid As VSFlexGrid)
'上移一行
    Dim strRowData As Variant
    Dim strRowText As Variant
    Dim varRowPic  As Variant
    Dim strRowDataChange As Variant
    Dim varBackColor As Variant
    
    Dim lngRow As Long
    Dim i As Long
    Dim lngUpRow As Long
    
    lngRow = objGrid.Row
    If lngRow <= 1 Then Exit Sub

    lngUpRow = lngRow - 1
    
    Do While lngUpRow > 0
        If objGrid.RowHidden(lngUpRow) Then
            lngUpRow = lngUpRow - 1
        Else
            Exit Do
        End If
    Loop
    
    If objGrid.RowHidden(lngUpRow) Then Exit Sub

    For i = 0 To objGrid.Cols - 1
        
        strRowText = objGrid.TextMatrix(lngUpRow, i)
        strRowData = objGrid.Cell(flexcpData, lngUpRow, i)
        Set varRowPic = objGrid.Cell(flexcpPicture, lngUpRow, i)
        varBackColor = objGrid.Cell(flexcpBackColor, lngUpRow, i)
        
        objGrid.TextMatrix(lngUpRow, i) = objGrid.TextMatrix(lngRow, i)
        objGrid.Cell(flexcpData, lngUpRow, i) = objGrid.Cell(flexcpData, lngRow, i)
        objGrid.Cell(flexcpPicture, lngUpRow, i) = objGrid.Cell(flexcpPicture, lngRow, i)
        objGrid.Cell(flexcpBackColor, lngUpRow, i) = objGrid.Cell(flexcpBackColor, lngRow, i)
        
        objGrid.TextMatrix(lngRow, i) = strRowText
        objGrid.Cell(flexcpData, lngRow, i) = strRowData
        objGrid.Cell(flexcpPicture, lngRow, i) = varRowPic
        objGrid.Cell(flexcpBackColor, lngRow, i) = varBackColor
        
    Next i
    
    strRowDataChange = objGrid.RowData(lngUpRow)
    objGrid.RowData(lngUpRow) = objGrid.RowData(lngRow)
    objGrid.RowData(lngRow) = strRowDataChange
    
    Call objGrid.Select(lngUpRow, 0)
End Sub

Public Sub MoveDown(objGrid As VSFlexGrid)
'下移一行
    Dim strRowData As Variant
    Dim strRowText As Variant
    Dim varRowPic  As Variant
    Dim strRowDataChange As Variant
    Dim varBackColor As Variant
    
    Dim lngRow As Long
    Dim i As Long
    Dim lngDownRow As Long
    
    lngRow = objGrid.Row
    If lngRow = 0 Then Exit Sub
    If lngRow >= objGrid.Rows - 1 Then Exit Sub
    
    lngDownRow = lngRow + 1
    
    Do While lngDownRow < objGrid.Rows - 1
        If objGrid.RowHidden(lngDownRow) Then
            lngDownRow = lngDownRow + 1
        Else
            Exit Do
        End If
    Loop

    If objGrid.RowHidden(lngDownRow) Then Exit Sub
    
    For i = 0 To objGrid.Cols - 1
        
        strRowText = objGrid.TextMatrix(lngDownRow, i)
        strRowData = objGrid.Cell(flexcpData, lngDownRow, i)
        Set varRowPic = objGrid.Cell(flexcpPicture, lngDownRow, i)
        varBackColor = objGrid.Cell(flexcpBackColor, lngDownRow, i)
        
        objGrid.TextMatrix(lngDownRow, i) = objGrid.TextMatrix(lngRow, i)
        objGrid.Cell(flexcpData, lngDownRow, i) = objGrid.Cell(flexcpData, lngRow, i)
        objGrid.Cell(flexcpPicture, lngDownRow, i) = objGrid.Cell(flexcpPicture, lngRow, i)
        objGrid.Cell(flexcpBackColor, lngDownRow, i) = objGrid.Cell(flexcpBackColor, lngRow, i)
        
        objGrid.TextMatrix(lngRow, i) = strRowText
        objGrid.Cell(flexcpData, lngRow, i) = strRowData
        objGrid.Cell(flexcpPicture, lngRow, i) = varRowPic
        objGrid.Cell(flexcpBackColor, lngRow, i) = varBackColor
    Next i

    strRowDataChange = objGrid.RowData(lngDownRow)
    objGrid.RowData(lngDownRow) = objGrid.RowData(lngRow)
    objGrid.RowData(lngRow) = strRowDataChange
    
    Call objGrid.Select(lngDownRow, 0)
End Sub

Public Function CheckRepet(objGrid As VSFlexGrid, lngCol As Long) As Boolean
'检查参数是否重复
    Dim i As Long
    Dim j As Long
    Dim lngRow As Long
    Dim lngStartRow As Long
    
    CheckRepet = False

    For i = 1 To objGrid.Rows - 2
        If Len(objGrid.TextMatrix(i, lngCol)) > 0 And Not objGrid.RowHidden(i) Then
            lngStartRow = i + 1
            Do
                lngRow = objGrid.FindRow(objGrid.TextMatrix(i, lngCol), lngStartRow, 0, False)
                
                If lngRow > 0 Then
                    If Not objGrid.RowHidden(lngRow) Then
                        CheckRepet = True
                        Exit Function
                    Else
                        If lngStartRow < objGrid.Rows - 1 Then
                            lngStartRow = lngRow + 1
                        Else
                            Exit Do
                        End If
                    End If
                Else
                    Exit Do
                End If
            Loop
'            For j = i + 1 To objGrid.Rows - 1
'                If Len(objGrid.TextMatrix(j, lngCol)) > 0 And Not objGrid.RowHidden(j) Then
'                    If UCase(Trim(objGrid.TextMatrix(j, lngCol))) = UCase(Trim(objGrid.TextMatrix(i, lngCol))) Then
'                        CheckRepet = True
'                        Exit Function
'                    End If
'                End If
'            Next
        End If
    Next
    
End Function


Public Function IsSelectionRow(objGrid As VSFlexGrid) As Boolean
    IsSelectionRow = False

    If objGrid.Rows <= 1 Then Exit Function
    If objGrid.RowSel <= 0 Or objGrid.RowSel >= objGrid.Rows Then Exit Function
    If objGrid.RowHidden(objGrid.RowSel) = True Then Exit Function
    
    IsSelectionRow = True
End Function
