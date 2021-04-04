Attribute VB_Name = "mEPRTable"
Option Explicit


Public Sub SaveTable(theTable As TTF160Ctl.F1Book, cTable As cEPRTable)
    '���� F1Book �ؼ����ݵ� cEPRTable �У�������Pictures��Elements��
    Dim strMerge As String, bLocked As Boolean, bHide As Boolean
    Dim i As Long, j As Long, k As Long
    Dim OldRow As Long, OldCol As Long

    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    Dim lngID As Long, T As Variant

    With theTable
        '���汨������
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
                Set NewCells.LastCell.CellFormat = .GetCellFormat   '��ȡ��Ԫ���ʽ
                
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
        Set cTable.Cells = NewCells.Clone   '��Ԫ��ı���

        .SetSelection iStartRow, iStartCol, iEndRow, iEndCol
        .Row = OldRow: .Col = OldCol
    End With
End Sub

Public Sub ReadTable(cTable As cEPRTable, theTable As TTF160Ctl.F1Book)
    '��ȡ cEPRTable ���ݵ� F1Book �ؼ��У�������Pictures��Elements��
    Dim strMerge As String, bLocked As Boolean
    Dim R1 As Long, C1 As Long, R2 As Long, C2 As Long
    Dim i As Long, j As Long, k As Long
    Dim CellFmt As TTF160Ctl.F1CellFormat

    With theTable
        .FixedRows = cTable.FixedRows
        .FixedCols = cTable.FixedCols
        .MaxRow = cTable.Rows
        .MaxCol = cTable.Cols
                
        '������ͨ�ı�������Ҫ��
        For k = 1 To cTable.Cells.Count     'ѭ���ָ�ÿһ��Ԫ�����ݺ͸�ʽ
            i = cTable.Cells(k).Row
            j = cTable.Cells(k).Col
            .SetActiveCell i, j             '���û��Ԫ��
            If cTable.Cells(k).CellType = cprCTEElement Then    '��ʾ����Ҫ�أ�ͬʱ�ݴ�����Ҫ�ؽ��ֵ�� Elements�С�
                Elements.AddEPRElement cTable.Elements(cTable.Cells(k).ElementKey).Clone
                Set CellFmt = .GetCellFormat
                CellFmt.ValidationText = "E_" & Elements.LastElement.Key       '���浱ǰ��Ԫ���ͼƬKey��
                .SetCellFormat CellFmt
                bLocked = True
                .EntryRC(i, j) = Elements.LastElement.���ֵ      '����Ҫ��ֻ��������ֵ��
            Else
                bLocked = cTable.Cells(k).Locked
                .EntryRC(i, j) = cTable.Cells(k).Text               '���ı�
            End If

            .SetProtection bLocked, False   '���õ�Ԫ���Ƿ�Locked

            .ColWidthTwips(j) = cTable.Cells(k).Width       '�ָ���Ԫ��ߡ���
            .RowHeight(i) = cTable.Cells(k).Height

            .SetCellFormat cTable.Cells(k).CellFormat       '�ָ���Ԫ���ʽ

            strMerge = cTable.Cells(k).MergeNo              '�ָ���Ԫ��ĺϲ�
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

