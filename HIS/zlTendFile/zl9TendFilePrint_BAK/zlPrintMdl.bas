Attribute VB_Name = "zlPrintMdl"
Option Explicit

Public Const conLineWide As Integer = 30        '������ռ���(��λΪ�)ռ�����߿��
Public Const conLineHigh As Integer = 30        '������ռ�߶�(��λΪ�)ռ�����߸߶�
Public Const conRatemmToTwip As Single = 56.6857142857143      '������羵ı���
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public gblnIsWps As Boolean      '���˺����:�ж��Ƿ���WPS�д�����

Public Const conSize1 = "�ż㣬 8 1/2 x 11 Ӣ��"
Public Const conSize2 = "+A611 С���ż㣬 8 1/2 x 11 Ӣ��"
Public Const conSize3 = "С�ͱ��� 11 x 17 Ӣ��"
Public Const conSize4 = "�����ʣ� 17 x 11 Ӣ��"
Public Const conSize5 = "�����ļ��� 8 1/2 x 14 Ӣ��"
Public Const conSize6 = "�����飬5 1/2 x 8 1/2 Ӣ��"
Public Const conSize7 = "�����ļ���7 1/2 x 10 1/2 Ӣ��"
Public Const conSize8 = "A3, 297 x 420 ����"
Public Const conSize9 = "A4, 210 x 297 ����"
Public Const conSize10 = "A4С�ţ� 210 x 297 ����"
Public Const conSize11 = "A5, 148 x 210 ����"
Public Const conSize12 = "B4, 250 x 354 ����"
Public Const conSize13 = "B5, 182 x 257 ����"
Public Const conSize14 = "�Կ����� 8 1/2 x 13 Ӣ��"
Public Const conSize15 = "�Ŀ����� 215 x 275 ����"
Public Const conSize16 = "10 x 14 Ӣ��"
Public Const conSize17 = "11 x 17 Ӣ��"
Public Const conSize18 = "������8 1/2 x 11 Ӣ��"
Public Const conSize19 = "#9 �ŷ⣬ 3 7/8 x 8 7/8 Ӣ��"
Public Const conSize20 = "#10 �ŷ⣬ 4 1/8 x 9 1/2 Ӣ��"
Public Const conSize21 = "#11 �ŷ⣬ 4 1/2 x 10 3/8 Ӣ��"
Public Const conSize22 = "#12 �ŷ⣬ 4 1/2 x 11 Ӣ��"
Public Const conSize23 = "#14 �ŷ⣬ 5 x 11 1/2 Ӣ��"
Public Const conSize24 = "C �ߴ繤����"
Public Const conSize25 = "D �ߴ繤����"
Public Const conSize26 = "E �ߴ繤����"
Public Const conSize27 = "DL ���ŷ⣬ 110 x 220 ����"
Public Const conSize28 = "C5 ���ŷ⣬ 162 x 229 ����"
Public Const conSize29 = "C3 ���ŷ⣬ 324 x 458 ����"
Public Const conSize30 = "C4 ���ŷ⣬ 229 x 324 ����"
Public Const conSize31 = "C6 ���ŷ⣬ 114 x 162 ����"
Public Const conSize32 = "C65 ���ŷ⣬114 x 229 ����"
Public Const conSize33 = "B4 ���ŷ⣬ 250 x 353 ����"
Public Const conSize34 = "B5 ���ŷ⣬176 x 250 ����"
Public Const conSize35 = "B6 ���ŷ⣬ 176 x 125 ����"
Public Const conSize36 = "�ŷ⣬ 110 x 230 ����"
Public Const conSize37 = "�ŷ������ 3 7/8 x 7 1/2 Ӣ��"
Public Const conSize38 = "�ŷ⣬ 3 5/8 x 6 1/2 Ӣ��"
Public Const conSize39 = "U.S. ��׼��д���� 14 7/8 x 11 Ӣ��"
Public Const conSize40 = "�¹���׼��д���� 8 1/2 x 12 Ӣ��"
Public Const conSize41 = "�¹����ɸ�д���� 8 1/2 x 13 Ӣ��"

Public Const conBin1 = "�ϲ�ֽ�н�ֽ"
Public Const conBin2 = "�²�ֽ�н�ֽ"
Public Const conBin3 = "�м�ֽ�н�ֽ"
Public Const conBin4 = "�ȴ��ֶ�����ÿҳֽ"
Public Const conBin5 = "�ŷ��ֽ����ֽ"
Public Const conBin6 = "�ŷ��ֽ����ֽ����Ҫ�ȴ��ֶ�����"
Public Const conBin7 = "��ǰȱʡֽ�н�ֽ"
Public Const conBin8 = "������ֽ����ֽ"
Public Const conBin9 = "С�ͽ�ֽ����ֽ"
Public Const conBin10 = "����ֽ�н�ֽ"
Public Const conBin11 = "��������ֽ����ֽ"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PaperName                  ���ݵ�ǰ��ӡ�������ã���ȡֽ������
'PaperSource                ���ݵ�ǰ��ӡ�������ã���ȡ��ֽ��ʽ����
'zlPutPrinterSet            ��ϵͳע����б����ӡȱʡ����
'PrintLvw                   listview��������
'PrintTends                  ��MSFlexGrid��������
'Print2Grd                  ����MSFlexGrid��������
'PrintGrds                  ��msFlexGrid��������
'PrintDBGrd                 ��DBGrid��������
'PrintFlxDB                 DBGrid��fsFlexGrid��϶�������
'GridCellPrint              ��������ӡ�����һ����Ԫ
'PrintCell                  ��ָ�������ӡһ�����ݵ�Ԫ,������ǰ�����ƶ�����Ԫ���Ͻ�λ��
'HaveExcel                  �жϱ�����װ��EXCELû��
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function PaperName() As String
    '------------------------------------------------
    '���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡֽ������
    '������
    '���أ� ֽ������
    '------------------------------------------------
    Dim mSize As Integer
    Err = 0
    On Error GoTo errHand
    
    If Printer.PaperSize = 256 Then
        PaperName = "�û��Զ��壬" _
            & Printer.Width / 56.6857142857143 & "x" _
            & Printer.Height / 56.6857142857143 & "����"
        Exit Function
    End If
    If Printer.PaperSize >= 1 And Printer.PaperSize <= 41 Then
        mSize = Printer.PaperSize
        PaperName = IIf(Printer.Orientation = 1, "����", "����") & Space(2) _
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
    PaperName = "���ɲ��ֽ��"

End Function

Public Function PaperSource() As String
    '------------------------------------------------
    '���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡ��ֽ��ʽ����
    '������
    '���أ� ��ֽ��ʽ�ַ���
    '------------------------------------------------
    Dim mBin As Integer
    
    Err = 0
    On Error GoTo errHand
    
    If Printer.PaperBin = 14 Then
        PaperSource = "���ӵĿ�ʽֽ�н�ֽ"
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
    PaperSource = "���ɲ�Ľ�ֽ��ʽ"

End Function

Public Function zlPutPrinterSet() As Boolean
    '------------------------------------------------
    '���ܣ���ϵͳע����б����ӡȱʡ����
    '------------------------------------------------
    If Printers.Count = 0 Then
        zlPutPrinterSet = False
        Exit Function
    End If
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperSize", Printer.PaperSize
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperBin", Printer.PaperBin
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Orientation", Printer.Orientation
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Width", Printer.Width
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Height", Printer.Height
    zlPutPrinterSet = True
End Function


Public Function PrintTends(ByVal PageRow As Long, ByVal PageCol As Long) As Boolean
    '------------------------------------------------
    '���ܣ� ��MSFlexGrid��������
    '������
    '���أ� �ɹ�����true �����󷵻�false
    '------------------------------------------------

    Err = 0
    On Error GoTo errHand
    '----------------------------------------------------
    '   ��������
    '----------------------------------------------------
    
    Dim intColcnt   As Integer   '�м�����
    Dim iCount      As Long   '���ɼ�����
    Dim CellsForward As New Collection   '��ǰ���������Ѿ���ӡ�ĵ�Ԫ
    
    With gobjOutTo
        gintCol(2, 1) = gobjSend.Body.Cols
        gintRow(2, 1) = gobjSend.Body.Rows
        '��ͷ���
        For iCount = 1 To gintFixRow
            If Not gobjSend.Body.RowHidden(iCount - 1) Then
                .CurrentX = gsngLeft * conRatemmToTwip
                For intColcnt = 1 To gintFixCol
                    If Not gobjSend.Body.ColHidden(intColcnt - 1) Then
                        GridCellPrint gobjSend.Body, iCount - 1, intColcnt - 1, CellsForward
                    End If
                Next
                For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                    If Not gobjSend.Body.ColHidden(intColcnt - 1) Then
                        GridCellPrint gobjSend.Body, iCount - 1, intColcnt - 1, CellsForward, , gintCol(2, PageCol)
                    End If
                Next
                .CurrentY = .CurrentY + gobjSend.Body.RowHeightMin
            End If
        Next
        
        '����������
        For iCount = gintRow(1, 1) To gintRow(2, 1)
            If Not gobjSend.Body.RowHidden(iCount - 1) Then
                .CurrentX = gsngLeft * conRatemmToTwip
                If iCount > glngPrintRow Or Not gblnPrintMode Or gintPrintState > 1 Then
                    For intColcnt = 1 To gintFixCol
                        If Not gobjSend.Body.ColHidden(intColcnt - 1) Then
                            GridCellPrint gobjSend.Body, iCount - 1, intColcnt - 1, CellsForward, gintRow(2, PageRow)
                        End If
                    Next
                    For intColcnt = gintCol(1, PageCol) To gintCol(2, PageCol)
                        If Not gobjSend.Body.ColHidden(intColcnt - 1) Then
                            GridCellPrint gobjSend.Body, iCount - 1, intColcnt - 1, CellsForward, gintRow(2, PageRow), gintCol(2, PageCol)
                        End If
                    Next
                End If
                .CurrentY = .CurrentY + gobjSend.Body.RowHeightMin
            End If
        Next
    End With
    
    PrintTends = True
    Exit Function

errHand:
    MsgBox "ϵͳ���ֲ���Ԥ֪�Ĵ���" & vbCrLf & Err.Description, vbCritical + vbOKOnly, gstrSysName
    PrintTends = False

End Function

Public Sub GridCellPrint(objGrid As Object, ROW As Long, COL As Long, _
    AcrossCells As Collection, Optional MaxRow, Optional MaxCol)
    '------------------------------------------------
    '���ܣ� ��������ӡ�����һ����Ԫ
    '������
    '   objGrid:��Ҫ�����MSFlexGrid����
    '   Row:�к�
    '   Col:�к�
    '   AcrossCells:Ӧ���Եĵ�Ԫ���ϣ������Ѿ���Ϊ�ϲ���Ԫ��ǰ��ӡ
    '���أ�
    '------------------------------------------------
    Dim iCount As Long
    For iCount = 1 To AcrossCells.Count
        If Trim(CStr(ROW)) & "," & Trim(CStr(COL)) = AcrossCells.Item(iCount) Then
            gobjOutTo.CurrentX = gobjOutTo.CurrentX + objGrid.ColWidth(COL)
            AcrossCells.Remove iCount
            Exit Sub
        End If
    Next
    
    '��Ӧ�ڵ�Ԫ�ı�����
    Dim Text As String
    Dim x As Long, Y As Long
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
    
    If IsMissing(MaxRow) Then MaxRow = gintFixRow
    If IsMissing(MaxCol) Then MaxCol = gintFixCol
    objGrid.ROW = ROW
    objGrid.COL = COL
    If objGrid.RowHidden(ROW) Then Exit Sub
    If objGrid.ColHidden(COL) Then Exit Sub
        
    '��ȡ�������ԣ�
    If COL < objGrid.FixedCols Or ROW < objGrid.FixedRows Then
        Alignment = objGrid.FixedAlignment(COL) '���չ̶���Ԫ
    Else
        Alignment = objGrid.ColAlignment(COL)   '������
    End If
    If ROW < objGrid.FixedRows Then Alignment = 4
    If Alignment = 11 Then Alignment = 7
    
    Select Case Alignment
    Case 1, 4, 7, 9
        PortraitAlignment = 2       '��
    Case 2, 5, 8
        PortraitAlignment = 1       '��
    Case 0, 3, 6
        PortraitAlignment = 0       '��
    End Select
    Select Case Alignment
    Case 0, 1, 2        '�����
        Alignment = 0
    Case 3, 4, 5        '����
        Alignment = 2
    Case 6, 7, 8        '�Ҷ���
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
    
    '��ȡ����ɫ��
    If CLng(objGrid.CellBackColor) <> 0 Then
        FillColor = objGrid.CellBackColor
    Else
        If COL < objGrid.FixedCols Or ROW < objGrid.FixedRows Then
            FillColor = objGrid.BackColorFixed
        Else
            FillColor = objGrid.BackColor
        End If
    End If
    
    '��ȡǰ��ɫ��
    If CLng(objGrid.CellForeColor) <> 0 Then
        ForeColor = objGrid.CellForeColor
    Else
        If COL < objGrid.FixedCols Or ROW < objGrid.FixedRows Then
            ForeColor = objGrid.ForeColorFixed
        Else
            ForeColor = objGrid.ForeColor
        End If
    End If
    
    '��������ɫ��
    If COL < objGrid.FixedCols Or ROW < objGrid.FixedRows Then
        GridColor = IIf(objGrid.GridLinesFixed = 1, objGrid.GridColorFixed, 0)   '���չ̶���
    Else
        GridColor = IIf(objGrid.GridLines = 1, objGrid.GridColor, 0)             '���ձ�׼��
    End If
    
    '�����߿�ȣ�
    If COL < objGrid.FixedCols Or ROW < objGrid.FixedRows Then
        LineStyle = IIf(objGrid.GridLinesFixed = 0, "0000", "1111")         '���չ̶���
    Else
        LineStyle = IIf(objGrid.GridLines = 0, "0000", "1111")              '���ձ�׼��
    End If
    
    '��ǰ�����ϲ���Ԫ����ȡ��ȷ�߶ȡ���ȣ�
    Text = objGrid.Text
    High = objGrid.RowHeightMin
    Wide = objGrid.ColWidth(COL)
    If Text <> "" And objGrid.MergeCells <> 0 Then
        If objGrid.MergeRow(ROW) Then
            For iCol = COL + 1 To IIf(COL < objGrid.FixedCols, objGrid.FixedCols, objGrid.Cols) - 1
                If iCol > MaxCol - 1 Then Exit For
                objGrid.COL = iCol
                If Text = objGrid.Text Then
                    If objGrid.MergeCells = 3 Or objGrid.MergeCells = 4 Then
                        iCount = ROW - 1
                        Do While iCount >= 0
                            If objGrid.TextMatrix(iCount, COL) <> objGrid.TextMatrix(iCount, iCol) Then Exit For
                            iCount = iCount - 1
                        Loop
                    End If
                    Wide = Wide + objGrid.ColWidth(iCol)
                    AcrossCells.Add Trim(CStr(ROW)) & "," & Trim(CStr(iCol))
                Else
                    Exit For
                End If
            Next
        End If
        
        objGrid.ROW = ROW
        objGrid.COL = COL
        If objGrid.MergeCol(COL) And ROW < objGrid.FixedRows Then
            For iRow = ROW + 1 To IIf(ROW < objGrid.FixedRows, objGrid.FixedRows, objGrid.Rows) - 1
                If iRow > MaxRow - 1 Then Exit For
                objGrid.ROW = iRow
                If Text = objGrid.Text Then
                    If objGrid.MergeCells = 2 Or objGrid.MergeCells = 4 Then
                        iCount = COL - 1
                        Do While iCount >= 0
                            If objGrid.TextMatrix(ROW, iCount) <> objGrid.TextMatrix(iRow, iCount) Then Exit For
                            iCount = iCount - 1
                        Loop
                    End If
                    High = High + objGrid.RowHeightMin
                    AcrossCells.Add Trim(CStr(iRow)) & "," & Trim(CStr(COL))
                Else
                    Exit For
                End If
            Next
        End If
        objGrid.ROW = ROW
        objGrid.COL = COL
    End If
    '��Ԫ�����
    Dim CurrentX As Long
    Dim bytCollType As Byte
    CurrentX = gobjOutTo.CurrentX
    
    '����������˫����
    bytCollType = Val(objGrid.TextMatrix(ROW, objGrid.Cols - 4))
    If bytCollType = 2 Then
        If InStr(1, "|" & frmTendFileReader.GetCollectCols & ";", "|" & COL - (2 + objGrid.FixedCols - 1) & ";") = 0 Then
            bytCollType = 0
        End If
    End If
    
    PrintCell Text, gobjOutTo.CurrentX, gobjOutTo.CurrentY, Wide, High, Alignment, _
        ForeColor, GridColor, FillColor, LineStyle, _
        objGrid.CellFontName, objGrid.CellFontSize * gsngScale, _
        objGrid.CellFontBold, objGrid.CellFontItalic, PortraitAlignment, _
        IIf(InStr(1, gstr�Խ���, "," & COL - 1 & ",") <> 0, 1, 0), bytCollType, _
        IIf(ROW < glngPrintRow, 1, Val(objGrid.TextMatrix(ROW, objGrid.Cols - 2)))
    gobjOutTo.CurrentX = CurrentX + objGrid.ColWidth(COL)

End Sub

Public Function CellTextRows(ByVal strText As String, ByVal Wide, ByVal High) As Variant
    '----------------------------------------------------------------------------------
    '����:���ı�����ת�������鷵��
    '����:strText-��Ԫ��ʽ
    '     Wide -���
    '     hight-�߶�
    '����:��ӡ�ĵ�Ԫ�������
    '����:���˺�
    '����:2007/09/04
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
            
            If Wide - conLineWide < gobjOutTo.TextWidth("1") Then    'С��һ���ַ�
                intAllRow = 1
            Else
'                If gobjOutTo.TextWidth(strPrintText) Mod (Wide - conLineWide) = 0 Then
'                    intAllRow = gobjOutTo.TextWidth(strPrintText) \ (Wide - conLineWide)
'                Else
'                    intAllRow = gobjOutTo.TextWidth(strPrintText) \ (Wide - conLineWide) + 1
'                End If
                
                '����������ݵ�������2008-04-08 By FrChen
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
    ByVal x As Single, ByVal Y As Single, _
    Optional ByVal Wide, _
    Optional ByVal High, _
    Optional Alignment As Byte = 0, _
    Optional ForeColor As Long = 0, _
    Optional GridColor As Long = 0, _
    Optional FillColor As Long = 0, _
    Optional LineStyle As String = "1111", _
    Optional FontName, Optional FontSize, _
    Optional FontBold, Optional FontItalic, _
    Optional PortraitAlignment As Byte = 2, _
    Optional Catercorner As Byte = 0, _
    Optional CollectType As Byte = 0, _
    Optional PrintedPage As Long = 0)
    '------------------------------------------------
    '���ܣ� ��ָ�������ӡһ�����ݵ�Ԫ,������ǰ�����ƶ�����Ԫ���Ͻ�λ��
    '������
    '   Text:    ������ַ���,���в������س����з�
    '   X:       ���Ͻ�X����
    '   Y:       ���Ͻ�Y����
    '   Wide:    ������
    '   High:    ����߶�
    '   Alignment:    ����ģʽ��0-�����(ȱʡ),1-�Ҷ���,2-����
    '   PortraitAlignment:�������ģʽ��0-����;1-����,2-����
    '   ForeColorǰ��ɫ,ȱʡΪ��ɫ
    '   GridColor����ɫ,ȱʡΪ��ɫ
    '   FillColor���ɫ,ȱʡΪ�豸����ɫ,����ϵͳ�����˺�ɫ��ɫ�룬���Խ�����������ɫ
    '   LineStyle:����ֱ�Ϊ�������µ��������
    '           0-���ߣ�1-9����Ӵ֣�1Ϊȱʡ
    '   FontName,FontSize,FontBold,FontItalic:��������
    '   Catercorner: 0-�޶Խ���;1-�Խ���
    '   CollectType: 0-������;1-������������
    '   PrintedPage: �����ʾ�Ѵ�ӡ,�����Ի�ɫ����
    '���أ�
    '------------------------------------------------
    Dim aryString() As String       '�س��ָ���ַ���
    Dim lngOldForeColor As Long     '����豸ȱʡǰ��ɫ
    Dim intRow As Long, intAllRow As Long
    Dim strRest As String, sngYMove As Single
    Dim oldFontName, oldFontSize, oldFontBold, oldFontItalic
    Dim strTmp As String
    
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
    
    If Not (gblnPrintMode And PrintedPage > 0 And glngPage = 1) Or gintPrintState > 1 Then
        If UCase(TypeName(LineStyle)) <> "STRING" Then LineStyle = CStr(LineStyle)
        If Len(LineStyle) < 4 Then
            LineStyle = Left(LineStyle & "1111", 4)
        End If
        
        '------------------------------------------
        '   ���ߴ�ӡ
        '------------------------------------------
        If Mid(LineStyle, 1, 1) <> 0 Then
            gobjOutTo.DrawWidth = Mid(LineStyle, 1, 1)
            gobjOutTo.Line (x, Y)-(x + Wide, Y), IIf(PrintedPage = 0 Or gintPrintState > 1, GridColor, ForeColor)
        End If
        
        If Mid(LineStyle, 2, 1) <> 0 Then
            gobjOutTo.DrawWidth = Mid(LineStyle, 2, 1)
            gobjOutTo.Line (x, Y)-(x, Y + High), IIf(PrintedPage = 0 Or gintPrintState > 1, GridColor, ForeColor)
        End If
        
        If Mid(LineStyle, 3, 1) <> 0 Then
            gobjOutTo.DrawWidth = Mid(LineStyle, 3, 1)
            gobjOutTo.Line (x + Wide, Y)-(x + Wide, Y + High), IIf(PrintedPage = 0 Or gintPrintState > 1, GridColor, ForeColor)
        End If
        
        If Mid(LineStyle, 4, 1) <> 0 Then
            gobjOutTo.DrawWidth = Mid(LineStyle, 4, 1)
            gobjOutTo.Line (x, Y + High)-(x + Wide, Y + High), IIf(PrintedPage = 0 Or gintPrintState > 1, GridColor, ForeColor)
        End If
        
        If CollectType = 1 Then
            gobjOutTo.DrawWidth = 1
            gobjOutTo.Line (x, Y + 10)-(x + Wide, Y + 10), ForeColor
            gobjOutTo.Line (x, Y + High - 20)-(x + Wide, Y + High - 20), ForeColor
        ElseIf CollectType = 2 Then
            gobjOutTo.DrawWidth = 1
            gobjOutTo.Line (x, Y + High - 40)-(x + Wide, Y + High - 40), IIf(PrintedPage = 0 Or gintPrintState > 1, vbRed, ForeColor)
            gobjOutTo.Line (x, Y + High - 20)-(x + Wide, Y + High - 20), IIf(PrintedPage = 0 Or gintPrintState > 1, vbRed, ForeColor)
        End If
        
        If Catercorner = 1 And InStr(1, Text, "/") <> 0 Then
            '�Խ������������к�/��ֱ�����
            gobjOutTo.DrawWidth = 1
            gobjOutTo.Line (x, Y + High)-(x + Wide, Y), IIf(PrintedPage = 0 Or gintPrintState > 1, GridColor, ForeColor)
            
            gobjOutTo.ForeColor = ForeColor
            '������ݾ��Ͽ�����ʾ
            gobjOutTo.CurrentX = x + conLineWide / 2                                    '����
            gobjOutTo.CurrentY = Y
            gobjOutTo.Print Split(Text, "/")(0)
            
            '�ұ����ݾ��¿�����ʾ
            gobjOutTo.CurrentX = x + Wide - gobjOutTo.TextWidth(Split(Text, "/")(1)) '����
            gobjOutTo.CurrentY = Y + High - gobjOutTo.TextHeight(Text)
            gobjOutTo.Print Split(Text, "/")(1)
        Else
            If Wide > conLineWide And High > conLineHigh Then
                '------------------------------------------
                '   ��ɫ���
                '------------------------------------------
        '        If FillColor <> 0 Then
        '            Printer.FillStyle = 1
        '            gobjOutTo.Line (X + conLineWide / 2, Y + conLineHigh / 2)- _
        '                (X + Wide - conLineWide / 2, Y + High - conLineHigh / 2), _
        '                FillColor, BF
        '        End If
                
                '------------------------------------------
                '   ���ִ�ӡ
                '------------------------------------------
                gobjOutTo.ForeColor = ForeColor
            
    '            If InStr(1, Text, vbCrLf) = 0 And InStr(1, Text, Chr(13)) = 0 Then
    '                If Wide - conLineWide < gobjOutTo.TextWidth("1") Then    'С��һ���ַ�
    '                    intAllRow = 1
    '                Else
    '    '                If gobjOutTo.TextWidth(Text) Mod (Wide - conLineWide) = 0 Then
    '    '                    intAllRow = gobjOutTo.TextWidth(Text) \ (Wide - conLineWide)
    '    '                Else
    '    '                    intAllRow = gobjOutTo.TextWidth(Text) \ (Wide - conLineWide) + 1
    '    '                End If
    '
    '                    '����������ݵ�������2008-04-08 By FrChen
    '                    strTmp = ""
    '                    intAllRow = 0
    '                    For intRow = 1 To Len(Text)
    '                        If gobjOutTo.TextWidth(strTmp & Mid(Text, intRow, 1)) > (Wide - conLineWide) Then
    '                            intAllRow = intAllRow + 1
    '                            strTmp = Mid(Text, intRow, 1)
    '                        Else
    '                            strTmp = strTmp & Mid(Text, intRow, 1)
    '                        End If
    '                    Next
    '                    If strTmp <> "" Then intAllRow = intAllRow + 1
    '
    '                End If
    '                For intRow = intAllRow To 1 Step -1
    '                    If High >= gobjOutTo.TextHeight(Text) * intRow Then
    '                        Exit For
    '                    End If
    '                Next
    '                intAllRow = intRow
    '
    '                Select Case PortraitAlignment
    '                Case 0
    '                    sngYMove = conLineHigh                                                          '����
    '                Case 1
    '                    sngYMove = (High - conLineHigh - gobjOutTo.TextHeight(Text) * intAllRow)        '����
    '                Case Else
    '                    sngYMove = (High - conLineHigh - gobjOutTo.TextHeight(Text) * intAllRow) / 2    '����
    '                End Select
    '                If sngYMove < 0 Then sngYMove = conLineHigh
    '
    '                strRest = Text
    '                For intRow = 0 To intAllRow - 1
    '                    Do While gobjOutTo.TextWidth(Text) > Wide - conLineWide
    '                        If Len(Trim(Text)) <= 1 Then Exit Do
    '                        Text = Left(Text, Len(Text) - 1)
    '                    Loop
    '                    strRest = Mid(strRest, Len(Text) + 1)
    '                    Select Case Alignment
    '                    Case 2
    '                        gobjOutTo.CurrentX = X + (Wide - gobjOutTo.TextWidth(Text)) / 2             '����
    '                    Case 1
    '                        gobjOutTo.CurrentX = X - conLineWide / 2 + Wide - gobjOutTo.TextWidth(Text) '����
    '                    Case Else
    '                        gobjOutTo.CurrentX = X + conLineWide / 2                                    '����
    '                    End Select
    '                    gobjOutTo.CurrentY = Y + conLineHigh / 2 + sngYMove + intRow * gobjOutTo.TextHeight(Text)
    '
    '                    If intAllRow = 1 Then
    '                        If Len(strRest) = 1 Then
    '                            gobjOutTo.Print Text & strRest
    '                        Else
    '                            gobjOutTo.Print Text
    '                        End If
    '                    Else
    '                        gobjOutTo.Print Text
    '                    End If
    '                    Text = strRest
    '                Next
    '            Else
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
                            gobjOutTo.CurrentX = x + (Wide - gobjOutTo.TextWidth(strRest)) / 2
                        Case 1
                            Do While Wide < gobjOutTo.TextWidth(strRest)
                                strRest = Right(strRest, Len(strRest) - 1)
                            Loop
                            gobjOutTo.CurrentX = x - conLineWide / 2 + Wide - gobjOutTo.TextWidth(strRest)
                        Case Else
                            Do While Wide < gobjOutTo.TextWidth(strRest)
                                strRest = Left(strRest, Len(strRest) - 1)
                            Loop
                            gobjOutTo.CurrentX = x + conLineWide / 2
                        End Select
                        
                        gobjOutTo.CurrentY = Y + conLineHigh / 2 + sngYMove + intRow * gobjOutTo.TextHeight(strRest)
                        If gobjOutTo.CurrentY + gobjOutTo.TextHeight(strRest) > Y + High Then Exit For
                        If gobjOutTo.CurrentY >= Y Then gobjOutTo.Print strRest
                    
                    Next '            End If
            End If
        End If
    End If
    gobjOutTo.CurrentX = x + Wide
    gobjOutTo.CurrentY = Y
    gobjOutTo.DrawStyle = 0
    gobjOutTo.DrawWidth = 1
    gobjOutTo.ForeColor = lngOldForeColor

    If Not IsMissing(FontName) Then gobjOutTo.FontName = oldFontName
    If Not IsMissing(FontSize) Then gobjOutTo.FontSize = oldFontSize
    If Not IsMissing(FontBold) Then gobjOutTo.FontBold = oldFontBold
    If Not IsMissing(FontItalic) Then gobjOutTo.FontItalic = oldFontItalic

End Sub

Public Function HaveExcel() As Boolean
    '------------------------------------------------
    '���ܣ��жϱ�����װ��EXCELû��
    '������
    '���أ����򷵻�True
    '------------------------------------------------

    On Error GoTo errHand 'errHandle1
    Dim objTemp  As Object
    gblnIsWps = False
    Set objTemp = CreateObject("Excel.Application") '��һ��EXCEL����
    Set objTemp = Nothing
    HaveExcel = True
    Exit Function

errHandle1:

    '���˺�:2007/4/20
    '��WPSΪ׼
    Err = 0: On Error GoTo errHand:
    Set objTemp = CreateObject("ET.Application") '��һ��WPS�е�ET����
    Set objTemp = Nothing
    HaveExcel = True
    gblnIsWps = True
    Exit Function
errHand:
    Set objTemp = Nothing
    HaveExcel = False

End Function

