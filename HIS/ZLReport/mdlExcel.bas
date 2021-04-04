Attribute VB_Name = "mdlExcel"
Option Explicit
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
'������Ŀ����
Public gobjTitle As RPTItem '�������
Public gcolUpItem As RPTItems  '�������,����˳����,һ��Ԫ��Ϊһ��
Public gcolDownItem As RPTItems '�������,����˳����,һ��Ԫ��Ϊһ��c

Public gobjExcel As Object 'Excel����
Public gobjHead As Object 'Ҫ����������ı�ͷ
Public gobjBody As Object 'Ҫ���������������ı���
Public gblnIsWps As Boolean     '���˺�:2007/04/20����,��Ҫȷ���Ƿ���wps�д�����

Private W_Excel As Integer

Private Const H_Excel = 20
Private Const xlNormal = -4143

Private lngGridW As Long
Private sngPer As Single, lngTotal As Long
Private frmParent As Object


Public Function HaveExcel() As Boolean
'���ܣ��ж�ϵͳ�Ƿ�װ��Excel
'˵����ͬʱ��ʼ��Excel����
    On Error Resume Next
    Err.Clear
    gblnIsWps = False
    Set gobjExcel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then Err.Clear: GoTo GoWps:
    HaveExcel = True
    Exit Function
GoWps:
   '���˺�:2007/4/20
    '��WPSΪ׼
    Err = 0: On Error GoTo ErrHand:
    Set gobjExcel = CreateObject("ET.Application") '��һ��WPS�е�ET����
    gblnIsWps = True
    HaveExcel = True
    Exit Function
ErrHand:
    HaveExcel = False
End Function

Public Function isExporting() As Boolean
'���ܣ�����Excel�����ж��Ƿ��Ѿ���һ���������������
    isExporting = Not (gobjExcel Is Nothing)
End Function

Private Function SpecSort(str As String) As String
'���ܣ�������"X1,ID1|X2,ID2|..."���ַ�����X����,��������ͬ��ʽ���ַ���
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
'���ܣ��������к�ת��ΪEXCEL�еı�ʾ����
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
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Private Function GetShowCol(lngCol As Long, objGrid As Object) As Long
'���ܣ����ݱ���кŻ�ȡ�ɼ��к�
    Dim i As Long, j As Long
    
    For i = 0 To lngCol - 1
       If objGrid.ColWidth(i) = 0 Then j = j + 1
    Next
    GetShowCol = lngCol - j
End Function

Private Function GetstrWidth(strFontName As String, curFontSize As Currency, str As String) As Long
'���ܣ������������÷���һ���ַ������
    Dim objFrm As frmFlash
    Set objFrm = New frmFlash
    objFrm.Font.name = strFontName
    objFrm.Font.Size = curFontSize
    GetstrWidth = objFrm.TextWidth(str)
    Unload objFrm
    Set objFrm = Nothing
End Function

Private Function GetstrHeight(strFontName As String, curFontSize As Currency, str As String) As Long
'���ܣ������������÷���һ���ַ����߶�
    Dim objFrm As frmFlash
    Set objFrm = New frmFlash
    objFrm.Font.name = strFontName
    objFrm.Font.Size = curFontSize
    GetstrHeight = objFrm.TextHeight(str)
    Unload objFrm
    Set objFrm = Nothing
End Function

Private Function GrdAlignment(objGrid As Object) As Long
'���ܣ���FlexGrid�Ķ��뷽ʽת��ΪEXCEL�еĶ��뷽ʽ
'������objGrid=FlexGrid����
    Dim Alignment As Integer
        
    '��ȡ�������ԣ�
    If objGrid.CellAlignment <> 0 Then
        Alignment = objGrid.CellAlignment           '���յ�Ԫ
    Else
        If objGrid.Col < objGrid.FixedCols Or objGrid.Row < objGrid.FixedRows Then
            Alignment = objGrid.FixedAlignment(objGrid.Col) '���չ̶���Ԫ
        Else
            Alignment = objGrid.ColAlignment(objGrid.Col)   '������
        End If
    End If
    Select Case Alignment
    Case 0, 1, 2        '�����
        GrdAlignment = -4131
    Case 3, 4, 5        '����
        GrdAlignment = -4108
    Case 6, 7, 8        '�Ҷ���
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
'���ܣ����ձ��ĺϲ���ʽ�����Excel��
'������objGrid=��񣬿�Ϊ��ͷ�����
    Dim i As Long, j As Long
    Dim lngY As Long, strTmp As String
    Dim lngTmp As Long, lngRowUp As Long
    Dim blnMerge As Boolean
    
    If objGrid.MergeCells = 0 Then Exit Sub
    
    '�ϲ���Ԫ��
    '����ϲ�
    lngY = lngRow
    For i = TopRow To BottomRow - 1
        If objGrid.MergeRow(i) Then
            For j = LeftCol To RightCol - 2
                If objGrid.TextMatrix(i, j) <> "" Then '�հ�����Ͳ��ϲ�
                    strTmp = Col2Excel(GetShowCol(j, objGrid) + 1) & Trim(str(lngY))
                    If Not gobjExcel.Range(strTmp).MergeCells Then   '�Ѻϲ���
                        blnMerge = False
                        For lngTmp = j + 1 To RightCol - 1
                            '��֪��һ����Ԫ��ͬ���˳�
                            If objGrid.TextMatrix(i, j) <> objGrid.TextMatrix(i, lngTmp) Then Exit For
                            If objGrid.MergeCells = 3 Or objGrid.MergeCells = 4 Then  '��������
                                lngRowUp = i - 1
                                Do While lngRowUp >= TopRow
                                    '����û�ϲ����˳�
                                    If objGrid.TextMatrix(lngRowUp, j) <> objGrid.TextMatrix(lngRowUp, lngTmp) Then Exit For
                                    lngRowUp = lngRowUp - 1
                                Loop
                            End If
                            blnMerge = True
                        Next
                        If blnMerge Then
                            strTmp = Col2Excel(GetShowCol(j, objGrid) + 1) & Trim(str(lngY)) & ":" & Col2Excel(GetShowCol(lngTmp, objGrid)) & Trim(str(lngY))
                            gobjExcel.Range(strTmp).MergeCells = True
                            j = lngTmp - 1 '�����Ѻϲ�����
                        End If
                    End If
                End If
            Next
         End If
        lngY = lngY + 1
        
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    Next
    
    '����ϲ�
    lngY = lngRow
    For j = LeftCol To RightCol - 1
        If objGrid.MergeCol(j) Then
            For i = TopRow To BottomRow - 2
                If objGrid.TextMatrix(i, j) <> "" Then '�հ�����Ͳ��ϲ�
                    strTmp = Col2Excel(GetShowCol(j, objGrid) + 1) & Trim(str(lngY + i - TopRow))
                    If Not gobjExcel.Range(strTmp).MergeCells Then   '�Ѻϲ���
                        blnMerge = False
                        For lngTmp = i + 1 To BottomRow - 1
                            '��֪��һ����Ԫ��ͬ���˳�
                            If objGrid.TextMatrix(i, j) <> objGrid.TextMatrix(lngTmp, j) Then Exit For
                            If objGrid.MergeCells = 2 Or objGrid.MergeCells = 4 Then  '��������
                                lngRowUp = j - 1
                                Do While lngRowUp >= LeftCol
                                    '����û�ϲ����˳�
                                    If objGrid.TextMatrix(i, lngRowUp) <> objGrid.TextMatrix(lngTmp, lngRowUp) Then Exit For
                                    lngRowUp = lngRowUp - 1
                                Loop
                            End If
                            blnMerge = True
                        Next
                        If blnMerge Then
                            strTmp = Col2Excel(GetShowCol(j, objGrid) + 1) & Trim(str(lngY + i - TopRow)) & ":" & Col2Excel(GetShowCol(j, objGrid) + 1) & Trim(str(lngY + lngTmp - 1 - TopRow))
                            gobjExcel.Range(strTmp).MergeCells = True
                            i = lngTmp - 1 '�����Ѻϲ�����
                        End If
                    End If
                End If
            Next
         End If
    
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    Next
End Sub

Private Function ConverItem(strItem As String, strFontName As String, curFontSize As Currency) As String
'���ܣ�������ϡ�����Ŀ
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
'���ܣ����������������һ����ʱ�ı��ļ���
'������In/Out�����ݸ�ʽ����
'���أ�����������ʱ�ļ�
    Dim strFile As String, strPath As String, intFile As Integer
    Dim tmpItem As RPTItem, strText As String
    Dim i As Long, j As Long
    
    '������ʱ�ļ�
    strPath = Space(256): strFile = Space(256)
    Call GetTempPath(256, strPath)
    strPath = Left(strPath, InStr(strPath, Chr(0)) - 1)
    
    Call GetTempFileName(strPath, "Excel", 0, strFile)
    strFile = Left(strFile, InStr(strFile, Chr(0)) - 1)
    
    intFile = FreeFile()
    Open strFile For Binary Access Write As intFile
    
    '�������
    Put intFile, , Replace(gobjTitle.����, vbCrLf, "")
    Put intFile, , vbCrLf
    
    sngPer = sngPer + 1
    Call ShowFlash("��������� Excel���������� ...", sngPer / lngTotal, frmParent, True)
    
    '������Ŀ���
    For Each tmpItem In gcolUpItem
        Put intFile, , ConverItem(tmpItem.����, tmpItem.����, tmpItem.�ֺ�)
        Put intFile, , vbCrLf
        
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel���������� ...", sngPer / lngTotal, frmParent, True)
    Next
   
   '��ͷ���
   If gobjHead Is Nothing Then
        For i = 0 To gobjBody.FixedRows - 1
            If gobjBody.RowHeight(i) <> 0 Then
                For j = 0 To gobjBody.Cols - 1
                    If gobjBody.ColWidth(j) <> 0 Then
                        Put intFile, , Replace(gobjBody.TextMatrix(i, j), "<���зָ���>", "")
                        Put intFile, , vbTab
                    End If
                Next
                Put intFile, , vbCrLf
            End If
        
            sngPer = sngPer + 1
            Call ShowFlash("��������� Excel���������� ...", sngPer / lngTotal, frmParent, True)
        Next
    Else
        For i = 0 To gobjHead.FixedRows - 1
            If gobjHead.RowHeight(i) <> 0 Then
                For j = 0 To gobjHead.Cols - 1
                    If gobjHead.ColWidth(j) <> 0 Then
                        Put intFile, , Replace(gobjHead.TextMatrix(i, j), "<���зָ���>", "")
                        Put intFile, , vbTab
                    End If
                Next
                Put intFile, , vbCrLf
            End If
            
            sngPer = sngPer + 1
            Call ShowFlash("��������� Excel���������� ...", sngPer / lngTotal, frmParent, True)
        Next
    End If
         
     '�����������
    For i = gobjBody.FixedRows To gobjBody.Rows - 1
        If gobjBody.RowHeight(i) <> 0 Then
            For j = 0 To gobjBody.Cols - 1
                If gobjBody.ColWidth(j) <> 0 Then
                    strText = Replace(gobjBody.TextMatrix(i, j), "<���зָ���>", "")
                    
                     '�����ʽ������
                    If IsNumeric(strText) Then
                        If Len(strText) > 15 Then
                            arrFormat(j, 1) = 2 '̫��������ǿ�д���Ϊ�ı���ʽ
                        ElseIf IsNumberEx(strText) = False Then
                            arrFormat(j, 1) = 2
                        End If
                    End If
                    
                    'д���ļ�
                    Put intFile, , strText
                    Put intFile, , vbTab
                    
                    '������ͣ�ֻҪ��һ���������ֻ򳤶ȴ���15����Ϊ�ı�
                    strText = Trim(strText)
                    If strText <> "" And (Not IsNumeric(strText) Or Len(strText) > 15) Then
                        arrFormat(j, 1) = 2
                    End If
                End If
            Next
            Put intFile, , vbCrLf
        End If
        
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel���������� ...", sngPer / lngTotal, frmParent, True)
    Next
    
    '������Ŀ���
    For Each tmpItem In gcolDownItem
        Put intFile, , ConverItem(tmpItem.����, tmpItem.����, tmpItem.�ֺ�)
        Put intFile, , vbCrLf
        
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel���������� ...", sngPer / lngTotal, frmParent, True)
    Next
    
    Close intFile
    GridToTXT = strFile
End Function

Public Function ExportExcel(ByVal frmMain As Object, Optional ByVal strExcelFile As String) As Boolean
'���ܣ������������ݵ�Excel
'��������ģ�鶨��Ĺ�������
'      strExcelFile=ָ��������ļ�
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
    '�ó�1�����ٸ��
    W_Excel = 120
    
    Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    
    gobjExcel.Workbooks.Add
    
    '�����п�
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
                        MsgBox "������������̫�࣬�ڱ�����Դ������޷������", vbInformation, App.ProductName
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
    Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    
    '���������ݸ�ʽ������
    lngZeroCol = 0
    For i = 0 To gobjBody.Cols - 1
        gobjBody.Col = i
        If gobjBody.ColWidth(i) = 0 Then
            lngZeroCol = lngZeroCol + 1
        Else
            For j = gobjBody.FixedRows To gobjBody.Rows - 1
                gobjBody.Row = j
                If Trim(gobjBody.Text) <> "" Then Exit For '���ǿհ׾��и�ʽ
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
                ElseIf IsDate(Replace(gobjBody.Text, "��", "��")) Then
                    If InStr(gobjBody.Text, ":") > 0 Or InStr(gobjBody.Text, "��") > 0 Then
                        If InStr(gobjBody.Text, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-MM-dd hh:mm:ss"
                        Else
                            .NumberFormatLocal = "yyyy""��""MM""��""dd""��"" hh""ʱ""mm""��""ss""��"""
                        End If
                    Else
                        If InStr(gobjBody.Text, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-MM-dd"
                        Else
                            .NumberFormatLocal = "yyyy""��""MM""��""dd""��"""
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
    Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    
    '�����ͷ������ĺϲ���ʽ
    If gobjHead Is Nothing Then
        'intRow����(Excel��ʼ�к�)��2��ԭ��MSHGrid��0�кŶ�ӦExcel��1�к�,+1;����ռһ��,+1��
        Call ExcelMerge(gobjBody, gcolUpItem.count + 2, 0, gobjBody.Cols, 0, gobjBody.FixedRows)
        Call ExcelMerge(gobjBody, gcolUpItem.count + gobjBody.FixedRows + 2, 0, gobjBody.Cols, gobjBody.FixedRows, gobjBody.Rows)
    Else
        Call ExcelMerge(gobjHead, gcolUpItem.count + 2, 0, gobjHead.Cols, 0, gobjHead.FixedRows)
        Call ExcelMerge(gobjBody, gcolUpItem.count + gobjBody.FixedRows + gobjHead.FixedRows + 2, 0, gobjBody.Cols, gobjBody.FixedRows, gobjBody.Rows)
    End If
    
    '����ڲ���ʽ���
    If gobjHead Is Nothing Then
        i = gcolUpItem.count + gobjBody.Rows + 1
    Else
        i = gcolUpItem.count + gobjHead.FixedRows + gobjBody.Rows + 1
    End If
    strTmp = "A" & Trim(str(gcolUpItem.count + 2)) & ":" & Col2Excel(gobjBody.Cols - lngZeroCol) & Trim(str(i))
    With gobjExcel.Range(strTmp)
        '����
        .Font.name = gobjBody.Font.name
        .Font.Size = gobjBody.Font.Size
        .Font.Bold = gobjBody.Font.Bold
        .Font.Italic = gobjBody.Font.Italic
        If gblnIsWps Then
            '���˺�:2007/4/20
            .Font.Underline = IIF(gobjBody.Font.Underline, 2, -4142)
        Else
            .Font.Underline = gobjBody.Font.Underline
        End If
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
        
        '�и�
        .RowHeight = gobjBody.RowHeight(gobjBody.FixedRows) / H_Excel
        
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
        
        '����ɫ
        .Interior.Pattern = -4142 'xlPatternNone
        .Interior.Color = CDbl(gobjBody.BackColor)
        
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
        
        'ǰ��ɫ
        .Font.Color = CDbl(gobjBody.ForeColor)
        
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
        
        '����
        .Borders(7).LineStyle = 1 'xlContinuous
        .Borders(8).LineStyle = 1
        .Borders(9).LineStyle = 1
        .Borders(10).LineStyle = 1
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
        
        '����ɫ
        .Borders(7).Color = CDbl(gobjBody.GridColor)
        .Borders(8).Color = CDbl(gobjBody.GridColor)
        .Borders(9).Color = CDbl(gobjBody.GridColor)
        .Borders(10).Color = CDbl(gobjBody.GridColor)
        
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
        
        If gobjBody.Cols - lngZeroCol > 1 Then
            .Borders(11).LineStyle = 1
            .Borders(11).Color = CDbl(gobjBody.GridColor)
        End If
        If i <> gcolUpItem.count + 2 Then
            .Borders(12).LineStyle = 1
            .Borders(12).Color = CDbl(gobjBody.GridColor)
        End If
        '���˺�:2007/04/20
        If gblnIsWps Then
            .Borders.Weight = 2
        End If
    End With
    
    sngPer = sngPer + 1
    Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    
    '����������һ�е��±���,��Ϊ��һ������и�Ϊ0������ϲ�,���ϱ��߲��ɼ�
    strTmp = "A" & gcolUpItem.count + 1 & ":" & Col2Excel(gobjBody.Cols - lngZeroCol) & gcolUpItem.count + 1
    With gobjExcel.Range(strTmp)
        .Borders(9).LineStyle = 1
        .Borders(9).Color = CDbl(gobjBody.GridColor)
        '���˺�:2007/04/20
        If gblnIsWps Then
            .Borders(9).Weight = 2
        End If
    End With
    
    sngPer = sngPer + 1
    Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    
    '��ͷ��ʽ
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
    Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    
    '�����ʽ
    With gobjExcel.Range("A1:" & Col2Excel(gobjBody.Cols - lngZeroCol) & "1")
        .Font.name = gobjTitle.����
        .Font.Size = gobjTitle.�ֺ�
        .Font.Bold = gobjTitle.����
        .Font.Italic = gobjTitle.б��
        If gblnIsWps Then
            '���˺�:2007/4/20
            .Font.Underline = IIF(gobjTitle.����, 2, -4142)
        Else
            .Font.Underline = gobjTitle.����
        End If
        .Font.Color = CDbl(gobjTitle.ǰ��)
        Select Case gobjTitle.����
            Case 0
                .HorizontalAlignment = -4131 'xlLeft
            Case 1
                .HorizontalAlignment = -4108 'xlCenter
            Case 2
                .HorizontalAlignment = -4152 'xlRight
        End Select
        
        .VerticalAlignment = -4108 'xlCenter
        
        If gobjTitle.�߿� Then
            .Borders(7).LineStyle = 1
            .Borders(8).LineStyle = 1
            .Borders(9).LineStyle = 1
            .Borders(10).LineStyle = 1
        End If
        If gobjTitle.���� <> &HFFFFFF Then
            .Interior.Pattern = -4142 'xlPatternNone
            .Interior.Color = CDbl(gobjTitle.����)
        End If
        .RowHeight = GetstrHeight(gobjTitle.����, gobjTitle.�ֺ�, gobjTitle.����) * 1.5 / H_Excel
        .MergeCells = True
    End With

    sngPer = sngPer + 1
    Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    
    '������Ŀ��ʽ
    j = 2
    For i = 1 To gcolUpItem.count
        With gobjExcel.Range("A" & Trim(str(j)) & ":" & Col2Excel(gobjBody.Cols - lngZeroCol) & Trim(str(j)))
            .Font.name = gcolUpItem(i).����
            .Font.Size = gcolUpItem(i).�ֺ�
            .Font.Bold = gcolUpItem(i).����
            .Font.Italic = gcolUpItem(i).б��
            If gblnIsWps Then
                '���˺�:2007/4/20
                .Font.Underline = IIF(gcolUpItem(i).����, 2, -4142)
            Else
                .Font.Underline = gcolUpItem(i).����
            End If
            .Font.Color = CDbl(gcolUpItem(i).ǰ��)
            
            .VerticalAlignment = -4108 'xlCenter
            If InStr(gcolUpItem(i).����, "|") = 0 Then
                 'ֻ��һ��ʱ����λ�ö���
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
            
            If gcolUpItem(i).�߿� Then
                .Borders(7).LineStyle = 1
                .Borders(8).LineStyle = 1
                .Borders(9).LineStyle = 1
                .Borders(10).LineStyle = 1
            End If
            If gcolUpItem(i).���� <> &HFFFFFF Then
                .Interior.Pattern = -4142 'xlPatternNone
                .Interior.Color = CDbl(gcolUpItem(i).����)
            End If
            .RowHeight = GetstrHeight(gcolUpItem(i).����, gcolUpItem(i).�ֺ�, gcolUpItem(i).����) * 1.5 / H_Excel
            '.NumberFormat = "G/ͨ�ø�ʽ"
            .NumberFormatLocal = "G/ͨ�ø�ʽ"
            .MergeCells = True
        End With
        j = j + 1
        
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    Next

    '������Ŀ��ʽ
    j = gcolUpItem.count + gobjBody.Rows + 2
    If Not gobjHead Is Nothing Then j = j + gobjHead.FixedRows
    For i = 1 To gcolDownItem.count
        With gobjExcel.Range("A" & Trim(str(j)) & ":" & Col2Excel(gobjBody.Cols - lngZeroCol) & Trim(str(j)))
            .Font.name = gcolDownItem(i).����
            .Font.Size = gcolDownItem(i).�ֺ�
            .Font.Bold = gcolDownItem(i).����
            .Font.Italic = gcolDownItem(i).б��
            If gblnIsWps Then
                '���˺�:2007/4/20
                .Font.Underline = IIF(gcolDownItem(i).����, 2, -4142)
            Else
                .Font.Underline = gcolDownItem(i).����
            End If
            .Font.Color = CDbl(gcolDownItem(i).ǰ��)
            
            .VerticalAlignment = -4108 'xlCenter
            If InStr(gcolDownItem(i).����, "|") = 0 Then
                 'ֻ��һ��ʱ����λ�ö���
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
            
            If gcolDownItem(i).�߿� Then
                .Borders(7).LineStyle = 1
                .Borders(8).LineStyle = 1
                .Borders(9).LineStyle = 1
                .Borders(10).LineStyle = 1
            End If
            If gcolDownItem(i).���� <> &HFFFFFF Then
                .Interior.Pattern = -4142 'xlPatternNone
                .Interior.Color = CDbl(gcolDownItem(i).����)
            End If
            .RowHeight = GetstrHeight(gcolDownItem(i).����, gcolDownItem(i).�ֺ�, gcolDownItem(i).����) * 1.5 / H_Excel
            '.NumberFormat = "G/ͨ�ø�ʽ"
            .NumberFormatLocal = "G/ͨ�ø�ʽ"
            .MergeCells = True
        End With
        j = j + 1
    
        sngPer = sngPer + 1
        Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
    Next
    
    '������и�Ϊ0����
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
    Call ShowFlash("��������� Excel�������ʽ ...", sngPer / lngTotal, frmParent, True)
        
    '��������������ݸ�ʽ
    ReDim arrFormat(gobjBody.Cols - 1, 1) As Integer
    For i = 0 To gobjBody.Cols - 1
        arrFormat(i, 0) = i + 1 '�к�
        arrFormat(i, 1) = 1 '���������ͣ�xlGeneralFormat = 1,xlTextFormat = 2
    Next
    '�����ı��ļ�,ͬʱ�������������
    strFile = GridToTXT(arrFormat)
    
    If gblnIsWps Then
        '���˺�:2007/4/20
        gobjExcel.Workbooks.Add
        gobjExcel.Cells.Select
        gobjExcel.Selection.NumberFormatLocal = "@"
        gobjExcel.Range("A1").Select
    
        sngPer = sngPer + 5
        Call ShowFlash("��������� Excel���������� ...", sngPer / lngTotal, frmParent, True)
        
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
        '�����ı��ļ�
        Call gobjExcel.Workbooks.OpenText(strFile, , 1, 1, 1, False, True, False, False, False, False, , arrFormat)
        sngPer = sngPer + 5
        Call ShowFlash("��������� Excel���������� ...", sngPer / lngTotal, frmParent, True)
    
        gobjExcel.Windows(2).Activate
        gobjExcel.Cells.Select
        gobjExcel.Selection.Copy
        gobjExcel.Windows(2).Activate
        gobjExcel.Cells.Select
        gobjExcel.Selection.PasteSpecial Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False
    End If
    
    sngPer = sngPer + 1
    Call ShowFlash("��������� Excel���������� ...", sngPer / lngTotal, frmParent, True)
    Clipboard.Clear
    
    If gblnIsWps Then
      gobjExcel.Windows(1).Close False
    Else
      gobjExcel.Windows(2).Close False
    End If
    Call ShowFlash("��������� Excel����� ...", 1, frmParent, True)
    
    gobjExcel.Range("A1").Select
    gobjExcel.Caption = App.Title & " - " & gobjTitle.����
    
    '���ΪExcel�ļ�
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
'���ܣ��ӱ����ǩ�����в������⡢�����������
'���������û�б��⣬��ȡ����������Ϊȱʡ����
    Dim strIDs As String, lngTmp As Long, strTmp As String
    Dim lngStep As Long, lngYStart As Long, lngBtn As Long
    Dim i As Long, j As Long, objLbl As Object
    
    Set gobjTitle = New RPTItem
    Set gcolUpItem = New RPTItems
    Set gcolDownItem = New RPTItems
    
    If frmSource.lbl.UBound = 0 Then
        'ȱʡ����
        gobjTitle.���� = frmSource.mobjReport.����
        gobjTitle.���� = "����_GB2312"
        gobjTitle.�ֺ� = 18
        gobjTitle.���� = 1
        gobjTitle.���� = vbWhite
        gobjTitle.ǰ�� = 0
        gobjTitle.H = 400
        Exit Sub
    End If
    
    gobjTitle.Y = 32767 '�ϴ�ֵ��������λ�ñȽ�
    
    '-------------------------------------------------------------------------
    '����:���֮��,��һ��
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
                gobjTitle.���� = .Caption
                gobjTitle.X = .Left
                gobjTitle.Y = .Top
                gobjTitle.W = .Width
                gobjTitle.H = .Height
                gobjTitle.���� = .FontBold
                gobjTitle.б�� = .FontItalic
                gobjTitle.���� = .FontUnderline
                gobjTitle.�߿� = (.BorderStyle = 1)
                gobjTitle.���� = .FontName
                gobjTitle.�ֺ� = .FontSize
                gobjTitle.���� = .BackColor
                gobjTitle.ǰ�� = .ForeColor
                gobjTitle.ǰ�� = .ForeColor
            End With
        End If
    Next
    If gobjTitle.id <> 0 Then
        strIDs = strIDs & "," & gobjTitle.id
        '���ݱ��������λ���ж϶��뷽ʽ
        lngTmp = gobjTitle.X + gobjTitle.W / 2
        If lngTmp <= gobjBody.Left + gobjBody.Width / 3 Then
            gobjTitle.���� = 0
        ElseIf lngTmp >= gobjBody.Left + gobjBody.Width * (2 / 3) Then
            gobjTitle.���� = 2
        Else
            gobjTitle.���� = 1
        End If
    Else
        'ȱʡ����
        gobjTitle.���� = frmSource.mobjReport.����
        gobjTitle.���� = "����_GB2312"
        gobjTitle.�ֺ� = 18
        gobjTitle.���� = 1
        gobjTitle.���� = vbWhite
        gobjTitle.ǰ�� = 0
        gobjTitle.H = 400
    End If
    
    '-------------------------------------------------------------------------
    '������Ŀ
    'ȷ����ʼ��:lngYstart
    'ȷ�����:lngStep
    '����С��ǩ�߶�Ϊ���
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
    If lngStep = 32767 Then lngStep = 180 '��С���Ϊһ��9���ַ��߶�
    
    i = lngYStart
    Do While i + lngStep <= lngBtn
        '��ǰ�ָ�������֯һ�б�����Ŀ
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
            '����������
            strTmp = SpecSort(strTmp)
            '���и�ʽ�Ե�һ����ǩΪ׼,����Ϊ����֮��
            'ÿ�ж��뷽ʽ�����ʱ���ݱ�ǩ����������������
            With frmSource.lbl(CInt(Split(Split(strTmp, "|")(0), ",")(1)))
                gcolUpItem.Add .Index, 0, "��ǩ", 0, 2, 0, "", 0, .Caption, "", .Left, .Top, .Width, .Height _
                    , 0, 0, True, .FontName, .FontSize, .FontBold, .FontUnderline, .FontItalic, 0, .ForeColor _
                    , .BackColor, (.BorderStyle = 1), 0, "", "", "", False, False, , , , , , "_" & .Index
                lngTmp = .Index
            End With
            For j = 1 To UBound(Split(strTmp, "|"))
                gcolUpItem("_" & lngTmp).���� = gcolUpItem("_" & lngTmp).���� & "|" & _
                    frmSource.lbl(CInt(Split(Split(strTmp, "|")(j), ",")(1))).Caption
            Next
        End If
        
        '����һ����ʼ��
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
    '������Ŀ
    'ȷ����ʼ��:lngYstart
    'ȷ�����:lngStep
    '����С��ǩ�߶�Ϊ���
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
    If lngStep = 32767 Then lngStep = 180 '��С���Ϊһ��9���ַ��߶�
    
    i = lngYStart
    Do While i <= lngBtn And lngBtn <> 32767
        '��ǰ�ָ�������֯һ�б�����Ŀ
        strTmp = ""
        For Each objLbl In frmSource.lbl
            If objLbl.Index > 0 And objLbl.Container Is objPaper And objLbl.Top >= i And objLbl.Top < i + lngStep And InStr(strIDs & ",", "," & objLbl.Index & ",") = 0 And Trim(objLbl.Caption) <> "" Then
                strIDs = strIDs & "," & objLbl.Index
                strTmp = strTmp & "|" & Format(objLbl.Left, "00000") & "," & objLbl.Index
            End If
        Next
        If strTmp <> "" Then
            strTmp = Mid(strTmp, 2)
            '����������
            strTmp = SpecSort(strTmp)
            '���и�ʽ�Ե�һ����ǩΪ׼,����Ϊ����֮��
            'ÿ�ж��뷽ʽ�����ʱ���ݱ�ǩ����������������
            With frmSource.lbl(CInt(Split(Split(strTmp, "|")(0), ",")(1)))
                gcolDownItem.Add .Index, 0, "��ǩ", 0, 2, 0, "", 0, .Caption, "", .Left, .Top, .Width, .Height _
                    , 0, 0, True, .FontName, .FontSize, .FontBold, .FontUnderline, .FontItalic, 0, .ForeColor _
                    , .BackColor, (.BorderStyle = 1), 0, "", "", "", False, False, , , , , , "_" & .Index
                lngTmp = .Index
            End With
            For j = 1 To UBound(Split(strTmp, "|"))
                gcolDownItem("_" & lngTmp).���� = gcolDownItem("_" & lngTmp).���� & "|" & _
                    frmSource.lbl(CInt(Split(Split(strTmp, "|")(j), ",")(1))).Caption
            Next
        End If
        
        '����һ����ʼ��
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
'���ܣ�
'������
'���أ�

    strText = Trim(strText)
    If strText Like "0[0-9]*" Then
        IsNumberEx = False
    Else
        IsNumberEx = True
    End If
End Function

