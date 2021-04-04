Attribute VB_Name = "mEPRTable"
Option Explicit


Public Sub SaveTable(theTable As TTF160Ctl.F1Book, cTable As cEPRTable)
    '保存 F1Book 控件内容到 cEPRTable 中（不包括Pictures和Elements）
    Dim strMerge As String, bLocked As Boolean, bHide As Boolean
    Dim i As Long, j As Long, k As Long
    Dim OldRow As Long, OldCol As Long

    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    Dim lngID As Long, T As Variant

    With theTable
        '保存报表属性
        cTable.FixedRows = .FixedRows
        cTable.FixedCols = .FixedCols
        cTable.Rows = .MaxRow
        cTable.Cols = .MaxCol

        OldRow = .Row: OldCol = .Col
        iStartRow = .SelStartRow: iEndRow = .SelEndRow
        iStartCol = .SelStartCol: iEndCol = .SelEndCol

        .SetSelection iStartRow, iStartCol, iStartRow, iStartCol

        Dim NewCells As New cEPRCells
        For i = 1 To .MaxRow
            For j = 1 To .MaxCol
                .SetActiveCell i, j
                NewCells.Add i, j
                Set NewCells.LastCell.CellFormat = .GetCellFormat   '获取单元格格式
                
                NewCells.LastCell.Text = .EntryRC(i, j)
                
                If .SelStartRow <> .SelEndRow Or .SelStartCol <> .SelEndCol Then
                    strMerge = Format(.SelStartRow, "0000") & Format(.SelStartCol, "0000") & Format(.SelEndRow, "0000") & Format(.SelEndCol, "0000")
                Else
                    strMerge = ""
                End If
                NewCells.LastCell.MergeNo = strMerge
                NewCells.LastCell.Width = .ColWidthTwips(j)
                NewCells.LastCell.Height = .RowHeight(i)
                
                T = Split(.GetCellFormat.ValidationText, "_")
                If UBound(T) >= 1 Then
                    If T(0) = "P" Then
                        lngID = Val(T(1))
                        NewCells.LastCell.CellType = cprCTEPicture
                    ElseIf T(0) = "E" Then
                        lngID = Val(T(1))
                        NewCells.LastCell.CellType = cprCTEElement
                    Else
                        NewCells.LastCell.CellType = cprCTEText
                    End If
                End If
                
                .GetProtection bLocked, bHide
                NewCells.LastCell.Locked = bLocked
            Next j
        Next i
        Set cTable.Cells = NewCells.Clone   '单元格的保存

        .SetSelection iStartRow, iStartCol, iEndRow, iEndCol
        .Row = OldRow: .Col = OldCol
    End With
End Sub

Public Sub ReadTable(cTable As cEPRTable, theTable As TTF160Ctl.F1Book)
    '读取 cEPRTable 内容到 F1Book 控件中（不包括Pictures和Elements）
    Dim strMerge As String, bLocked As Boolean
    Dim R1 As Long, C1 As Long, R2 As Long, C2 As Long
    Dim i As Long, j As Long, k As Long
    Dim CellFmt As TTF160Ctl.F1CellFormat

    With theTable
        .FixedRows = cTable.FixedRows
        .FixedCols = cTable.FixedCols
        .MaxRow = cTable.Rows
        .MaxCol = cTable.Cols
                
        '处理普通文本和诊治要素
        For k = 1 To cTable.Cells.Count     '循环恢复每一单元格内容和格式
            i = cTable.Cells(k).Row
            j = cTable.Cells(k).Col
            .SetActiveCell i, j             '设置活动单元格
            If cTable.Cells(k).CellType = cprCTEElement Then    '显示诊治要素，同时暂存诊治要素结果值到 Elements中。
                Elements.AddEPRElement cTable.Elements(cTable.Cells(k).ElementKey).Clone
                Set CellFmt = .GetCellFormat
                CellFmt.ValidationText = "E_" & Elements.LastElement.Key       '保存当前单元格的图片Key。
                .SetCellFormat CellFmt
                bLocked = True
                .EntryRC(i, j) = Elements.LastElement.结果值      '诊治要素只保存其结果值。
            Else
                bLocked = cTable.Cells(k).Locked
                .EntryRC(i, j) = cTable.Cells(k).Text               '纯文本
            End If

            .SetProtection bLocked, False   '设置单元格是否Locked

            .ColWidthTwips(j) = cTable.Cells(k).Width       '恢复单元格高、宽
            .RowHeight(i) = cTable.Cells(k).Height

            .SetCellFormat cTable.Cells(k).CellFormat       '恢复单元格格式

            strMerge = cTable.Cells(k).MergeNo              '恢复单元格的合并
            If strMerge <> "" Then
                R1 = Val(Left(strMerge, 4))
                C1 = Val(Mid(strMerge, 5, 4))
                R2 = Val(Mid(strMerge, 9, 4))
                C2 = Val(Mid(strMerge, 13))
                .SetSelection R1, C1, R2, C2
                Set CellFmt = .GetCellFormat
                CellFmt.MergeCells = True
                .SetCellFormat CellFmt
            End If
        Next
        .Row = 1: .Col = 1
    End With
End Sub

