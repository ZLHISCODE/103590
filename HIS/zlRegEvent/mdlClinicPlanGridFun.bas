Attribute VB_Name = "mdlClinicPlanGridFun"
Option Explicit
Public Enum gPlanGrid_ColIndex '����̶�������
    COL_ͼ��
    COL_����ID
    COL_��ԴID
    COL_����ID
    COL_����
    COL_����
    COL_����
    COL_��Ŀ
    COL_ҽ��
    Col_ҽ��ְ��
    COL_�Ƿ񽨲���
    COL_ԤԼ����
    COL_����Ƶ��
    COL_���տ���״̬
    COL_���ջ���
    COL_�Ű෽ʽ
    COL_��ʼʱ��
    COL_��ֹʱ��
    COL_�Ǽ�ʱ��
    COL_��ʱ����
    COL_�Ƿ����
    COL_�Ƿ��ٴ��Ű�
End Enum

'�����б�̶�����
Public Const gPlanGrid_FixedCols = 22

'���ű�����
Public Enum gPlanGrid_DataStyle
    Data_Templet = 0
    Data_FixedRule = 1
    Data_Plan = 2
    Data_MonthTemplet = 3 '��ģ�壬�³�������ɵ�ģ��
End Enum

Public Sub InitPlanGrid(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    Optional ByVal dtMinDate As Date, Optional ByVal dtMaxDate As Date, _
    Optional ByVal blnPublished As Boolean)
    '���ܣ���ʼ���������ݱ��
    '   vsfGrid - VSF���
    '   bytDataStyle - ��������
    Dim strHead As String, varData As Variant
    Dim strHeadSub As String, varDataSub As Variant
    Dim i As Long, lngCol As Long
    Dim arrDate As Variant, strTemp As String
    Dim dtCurdate As Date, intDays As Integer

    Err = 0: On Error GoTo errHandler
    With vsfGrid
        .Redraw = False
        .Rows = 2
        
        '�̶���
        strHead = " ,4,300|����ID,4,0|��ԴID,4,0|����ID,4,0|����,4,0|����,4,500|����,1,1000|��Ŀ,1,0|ҽ��,1,850|ҽ��ְ��,1,0|" & _
                "����,4,0|ԤԼ����,4,0|����Ƶ��,4,0|���տ���״̬,1,0|���ջ���,4,0|�Ű෽ʽ,4,0|��ʼʱ��,1,0|��ֹʱ��,1,0|" & _
                "�Ǽ�ʱ��,1,0|��ʱ����,4,0|�Ƿ����,4,0|�Ƿ��ٴ��Ű�,4,0"
        strHeadSub = " ,����ID,��ԴID,����ID,����,����,����,��Ŀ,ҽ��,ҽ��ְ��," & _
                "����,ԤԼ����,����Ƶ��,���տ���״̬,���ջ���,�Ű෽ʽ,��ʼʱ��,��ֹʱ��," & _
                "�Ǽ�ʱ��,��ʱ����,�Ƿ����,�Ƿ��ٴ��Ű�"
        varData = Split(strHead, "|")
        varDataSub = Split(strHeadSub, ",")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0): .TextMatrix(1, i) = varDataSub(i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .ColKey(i) = Split(varData(i), ",")(0)
        Next
        .FixedCols = 9: .FixedRows = 2
        '��̬��
        Select Case bytDataStyle
        Case Data_Templet, Data_FixedRule   'ģ��,�̶�����
            strHead = "��һ,1,450|��һ,4,550|��һ,4,550|�ܶ�,1,450|�ܶ�,4,450|�ܶ�,4,550|����,1,450|����,4,550|����,4,550|" & _
                    "����,1,450|����,4,550|����,4,550|����,1,450|����,4,550|����,4,550|����,1,450|����,4,550|����,4,550|" & _
                    "����,1,450|����,4,550|����,4,550"
            strHeadSub = "ʱ��,�޺�,��Լ,ʱ��,�޺�,��Լ,ʱ��,�޺�,��Լ," & _
                    "ʱ��,�޺�,��Լ,ʱ��,�޺�,��Լ,ʱ��,�޺�,��Լ," & _
                    "ʱ��,�޺�,��Լ"
            If bytDataStyle = Data_Templet Then
                strHead = strHead & "|��������,1,1150|��������,1,450|��������,4,550|��������,4,550"
                strHeadSub = strHeadSub & ",������Ŀ,ʱ��,�޺�,��Լ"
            End If
            varData = Split(strHead, "|")
            varDataSub = Split(strHeadSub, ",")
            lngCol = .Cols
            .Cols = .Cols + UBound(varData) + 1
            For i = 0 To UBound(varData)
                .TextMatrix(0, lngCol) = Split(varData(i), ",")(0): .TextMatrix(1, lngCol) = varDataSub(i)
                .Cell(flexcpData, 0, lngCol) = CStr(Split(varData(i), ",")(0))
                .ColAlignment(lngCol) = Split(varData(i), ",")(1)
                .ColWidth(lngCol) = Split(varData(i), ",")(2)
                lngCol = lngCol + 1
            Next
            .FixedAlignment(-1) = flexAlignCenterCenter
            .RowHeight(0) = 420: .RowHeight(1) = 300
            
            .AllowSelection = False
        Case Data_Plan, Data_MonthTemplet '���ż�¼
            intDays = DateDiff("d", dtMinDate, dtMaxDate) + 1 '����
            If intDays < 0 Then intDays = 0
            dtCurdate = dtMinDate
            lngCol = .Cols
            .Cols = .Cols + intDays * 3
            For i = 1 To intDays
                If bytDataStyle = Data_MonthTemplet Then
                    .Cell(flexcpText, 0, lngCol, 0, lngCol + 2) = Day(dtCurdate) & "�� "
                Else
                    strTemp = Decode(bytDataStyle, Data_MonthTemplet, Day(dtCurdate) & "��", Format(dtCurdate, "mm��dd��")) & _
                              Chr(13) & GetWeekName(Weekday(dtCurdate, vbMonday) - 1)
                    .TextMatrix(0, lngCol) = strTemp
                    .TextMatrix(0, lngCol + 1) = strTemp
                    .TextMatrix(0, lngCol + 2) = strTemp
                End If
                .Cell(flexcpData, 0, lngCol, 0, lngCol + 2) = Format(dtCurdate, "yyyy-MM-dd") '����
                .Cell(flexcpText, 1, lngCol, 1, lngCol + 2) = "ʱ��" & vbTab & "�޺�" & vbTab & "��Լ"
                .ColAlignment(lngCol) = 1: .ColAlignment(lngCol + 1) = 4: .ColAlignment(lngCol + 2) = 4
                .ColWidth(lngCol) = 450
                .ColWidth(lngCol + 1) = IIf(bytDataStyle = Data_MonthTemplet Or blnPublished = False, 550, 650)
                .ColWidth(lngCol + 2) = IIf(bytDataStyle = Data_MonthTemplet Or blnPublished = False, 550, 650)
                dtCurdate = DateAdd("d", 1, dtCurdate)
                lngCol = lngCol + 3
            Next
            .FixedAlignment(-1) = flexAlignCenterCenter
            If bytDataStyle = Data_MonthTemplet Then
                .RowHeight(0) = 420: .RowHeight(1) = 300
            Else
                .RowHeight(0) = 500: .RowHeight(1) = 300
            End If
            
            .AllowSelection = blnPublished
        End Select
        
        .AllowBigSelection = False
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .GridLines = flexGridFlat
        .PicturesOver = True '������ͼƬ����
        
        '����������,�����û�ѡ����ʾ��
        For i = 0 To .Cols - 1
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)|������(0-��������,1-��ֹ����,2-��������,�����س���������)
            Select Case i
            Case COL_����ID, COL_��ԴID, COL_����ID, COL_��ʱ����, COL_�Ƿ����
                 .ColData(i) = "-1|1"
            Case COL_����, COL_����, COL_ҽ��
                .ColData(i) = "1|0"
            End Select
        Next
        '�ǹ̶�����ʱ����ʼʱ�����ֹʱ�䲻��ʾ
        If bytDataStyle <> Data_FixedRule Then
            .ColData(COL_��ʼʱ��) = "-1|1": .ColData(COL_��ֹʱ��) = "-1|1": .ColData(COL_�Ǽ�ʱ��) = "-1|1"
        End If

        '�ϲ�����
        .MergeCells = flexMergeRestrictColumns
        .MergeRow(0) = True: .MergeCol(-1) = True
        .Redraw = True
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function LoadPlanDataByRecordset(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    ByVal rsData As ADODB.Recordset, Optional ByVal bytTableMode As Byte, _
    Optional ByRef lngSignalCount As Long, Optional ByVal blnSortLoad As Boolean, _
    Optional ByVal blnPublished As Boolean, _
    Optional ByVal strStartDate As String, Optional ByVal strEndDate As String) As Boolean
    '���ܣ�����Recordset�����������
    '��Σ�
    '   bytTableMode 0-�̶������,1-�³����,2-�ܳ����,3-ģ��
    '   strStartDate,strEndDate ��������ڷ�Χ���ܳ����ʱ����
    '���Σ�
    '   lngSignalCount - ��Դ����
    '˵�������ݱ����ǰ�"����,����,����,��Ŀ,ҽ��"���������˵ģ����������ʾ����ȷ
    Dim i As Long, j As Long, lngCurRow As Long, lngCurCol As Long
    Dim strGroupKey As String '��������"����,����,����,��Ŀ,ҽ��"����
    Dim lngBackColor As Long '����������Ľ���ɫ
    Dim strTemp As String, blnAddRow As Boolean
    Dim lngRowStart As Long, lngRowEnd As Long
    Dim lngoldCol As Long, blnFindCol As Boolean
    Dim lngOldRow As Long, str��Լ�� As String
    Dim strRecordInfo As String
    
    Err = 0: On Error GoTo errHandle
    lngSignalCount = 0
    '��¼��ǰѡ��Ԫ�����ڻָ�ѡ��
    lngOldRow = vsfGrid.Row: lngoldCol = vsfGrid.Col
    '���������
    vsfGrid.Clear 1: vsfGrid.Rows = 2
    
    If rsData Is Nothing Then Exit Function
    If rsData.RecordCount = 0 Then Exit Function
    
    rsData.MoveFirst
    With vsfGrid
        lngCurRow = 2
        strGroupKey = ""
        lngBackColor = G_AlternateColor
        .Redraw = flexRDNone
        Do While Not rsData.EOF
            blnFindCol = False
            '1.�������
            strTemp = Nvl(rsData!����) & "," & Nvl(rsData!����) & "," & Nvl(rsData!�շ���Ŀ) & "," & Nvl(rsData!ҽ������)
            If bytDataStyle = Data_FixedRule Then strTemp = strTemp & "," & Nvl(rsData!����ID)
            If strGroupKey <> strTemp Then
                lngSignalCount = lngSignalCount + 1
                strGroupKey = strTemp
                lngCurCol = gPlanGrid_FixedCols  '�����ж��Ƿ�ȷ������
                lngBackColor = IIf(lngBackColor = vbWindowBackground, G_AlternateColor, vbWindowBackground)
                
                .Rows = .Rows + 1: lngCurRow = .Rows - 1
                .RowData(lngCurRow) = -1 '��ǣ������ж��Ƿ�Ϊ���ؿ���
                
                lngCurRow = lngCurRow + 1
            End If
            '2.�������
            '2.1ȷ����ǰ��
            Select Case bytDataStyle
            Case Data_Templet  'ģ��
                If Nvl(rsData!�Ű����) <> 1 Then '��������
                    '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
                    lngCurCol = .Cols - 4: blnFindCol = True
                Else
                    If Nvl(rsData!������Ŀ) <> "" Then
                        strTemp = Nvl(rsData!������Ŀ)
                        For i = lngCurCol To .Cols - 1 Step 3
                            If strTemp = .Cell(flexcpData, 0, i) Then
                                lngCurCol = i: blnFindCol = True
                                Exit For
                            End If
                        Next
                        'û�ҵ��ٴӿ�ʼ�����ң���Ҫ�ǰ�������Ŀ���������˳��һ��
                        If blnFindCol = False Then
                            For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
                                If strTemp = .Cell(flexcpData, 0, i) Then
                                    lngCurCol = i: blnFindCol = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            Case Data_FixedRule  '�̶�����
                If Nvl(rsData!������Ŀ) <> "" Then
                    strTemp = Nvl(rsData!������Ŀ)
                    For i = lngCurCol To .Cols - 1 Step 3
                        If strTemp = .Cell(flexcpData, 0, i) Then
                            lngCurCol = i: blnFindCol = True
                            Exit For
                        End If
                    Next
                    'û�ҵ��ٴӿ�ʼ�����ң���Ҫ�ǰ�������Ŀ���������˳��һ��
                    If blnFindCol = False Then
                        For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
                            If strTemp = .Cell(flexcpData, 0, i) Then
                                lngCurCol = i: blnFindCol = True
                                Exit For
                            End If
                        Next
                    End If
                End If
            Case Else '���ż�¼
                If Nvl(rsData!��������) = "" Then
                    '��������Ϊ��,ȡ���ŵĿ�ʼʱ��,��Ҫ������ܿ��µ��ܳ�����޳����¼ʱ,����������ʾ��ͬһ��
                    strTemp = Format(Nvl(rsData!��ʼʱ��), "yyyy-mm-dd")
                    lngCurCol = gPlanGrid_FixedCols
                Else
                    strTemp = Format(Nvl(rsData!��������), "yyyy-mm-dd")
                End If
                If IsDate(strTemp) Then
                    For i = lngCurCol To .Cols - 1 Step 3
                        If DateDiff("d", strTemp, .Cell(flexcpData, 0, i)) = 0 Then
                            lngCurCol = i: blnFindCol = True
                            Exit For
                        End If
                    Next
                End If
            End Select
            
            If blnFindCol Then
                '2.2ȷ����ǰ��
                For i = IIf(.Rows - 1 > lngCurRow, lngCurRow, .Rows - 1) To 2 Step -1
                    If .RowData(i) = -1 Or .TextMatrix(i, lngCurCol) <> "" Then  '�����ؿ��л�����������
                        lngCurRow = i + 1: Exit For
                    End If
                Next
            End If
            
            '3.��������
            blnAddRow = False
            If .Rows - 1 < lngCurRow Then
                '�����в�������1��
                .Rows = .Rows + 1: lngCurRow = .Rows - 1
                .RowData(lngCurRow) = lngBackColor '�������ý���ɫ
                
                Select Case bytTableMode
                Case 0 '�̶������
                    If Val(Nvl(rsData!�Ƿ���Ч)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("FixedItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("InvalidFixedItem")
                    End If
                Case 1 '�³����
                    If Val(Nvl(rsData!�Ƿ���Ч)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("MonthItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("InvalidMonthItem")
                    End If
                Case 2 '�ܳ����
                    If Val(Nvl(rsData!�Ƿ���Ч)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("WeekItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("InvalidWeekItem")
                    End If
                Case 3 'ģ��
                    If Nvl(rsData!�Ű෽ʽ) = "�����Ű�" Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("MonthItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("WeekItem")
                    End If
                End Select
                .Cell(flexcpPictureAlignment, lngCurRow, COL_ͼ��) = flexAlignCenterCenter

                If Not (bytDataStyle = Data_Plan And IsDate(strStartDate) And IsDate(strEndDate)) Then
                    .TextMatrix(lngCurRow, COL_����ID) = Nvl(rsData!����ID)
                End If
                .TextMatrix(lngCurRow, COL_��ԴID) = Nvl(rsData!��ԴID)
                .TextMatrix(lngCurRow, COL_����) = Nvl(rsData!����)
                .TextMatrix(lngCurRow, COL_����) = Nvl(rsData!����)
                .TextMatrix(lngCurRow, COL_����) = Nvl(rsData!����)
                .TextMatrix(lngCurRow, COL_��Ŀ) = Nvl(rsData!�շ���Ŀ)
                .TextMatrix(lngCurRow, COL_ҽ��) = Nvl(rsData!��ʶ��) & Nvl(rsData!ҽ������)
                .Cell(flexcpData, lngCurRow, COL_ҽ��) = Nvl(rsData!ҽ������)
                .TextMatrix(lngCurRow, Col_ҽ��ְ��) = Nvl(rsData!ҽ��ְ��)

                .TextMatrix(lngCurRow, COL_�Ƿ񽨲���) = IIf(Val(Nvl(rsData!�Ƿ񽨲���)) = 1, "��", "")
                .TextMatrix(lngCurRow, COL_ԤԼ����) = Nvl(rsData!ԤԼ����)
                .TextMatrix(lngCurRow, COL_����Ƶ��) = Nvl(rsData!����Ƶ��)
                .TextMatrix(lngCurRow, COL_���տ���״̬) = Nvl(rsData!���տ���״̬)
                .TextMatrix(lngCurRow, COL_���ջ���) = IIf(Val(Nvl(rsData!�Ƿ���ջ���)) = 1, "��", "")
                .TextMatrix(lngCurRow, COL_�Ű෽ʽ) = Nvl(rsData!�Ű෽ʽ)
                .TextMatrix(lngCurRow, COL_��ʼʱ��) = Format(Nvl(rsData!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
                .TextMatrix(lngCurRow, COL_��ֹʱ��) = Format(Nvl(rsData!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
                If bytDataStyle = Data_FixedRule Then
                    .TextMatrix(lngCurRow, COL_�Ǽ�ʱ��) = Format(Nvl(rsData!�Ǽ�ʱ��), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngCurRow, COL_��ʱ����) = Val(Nvl(rsData!��ʱ����))
                    .TextMatrix(lngCurRow, COL_�Ƿ����) = Val(Nvl(rsData!�Ƿ����))
                End If
                .TextMatrix(lngCurRow, COL_�Ƿ��ٴ��Ű�) = IIf(Val(Nvl(rsData!�Ƿ��ٴ��Ű�)) = 1, "��", "")
                blnAddRow = True
            End If
                
            If bytDataStyle = Data_Plan And IsDate(strStartDate) And IsDate(strEndDate) Then
                '������ܿ��µ��ܳ�����޳���а���IDΪ��ǰѡ�������к�Դ�İ���ID
                If IsDate(Nvl(rsData!��ʼʱ��)) And IsDate(Nvl(rsData!��ֹʱ��)) Then
                    If DateDiff("d", Nvl(rsData!��ʼʱ��), strStartDate) <= 0 And DateDiff("d", Nvl(rsData!��ֹʱ��), strEndDate) >= 0 Then
                        .TextMatrix(lngCurRow, COL_����ID) = Nvl(rsData!����ID)
                    End If
                End If
            End If
            
            If blnFindCol Then
                '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
                'ԤԼ���Ʒ�ʽ��0-����ԤԼ����;1-�ú����ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
                If Nvl(rsData!�ϰ�ʱ��) <> "" Then
                    str��Լ�� = IIf(Nvl(rsData!ԤԼ���Ʒ�ʽ) = 1, "-", _
                        IIf(Val(Nvl(rsData!��Լ��)) = 0, IIf(Val(Nvl(rsData!�޺���)) = 0, "��", _
                            Val(Nvl(rsData!�޺���))), Val(Nvl(rsData!��Լ��))))
                    Select Case bytDataStyle
                    Case Data_Templet
                        If Nvl(rsData!�Ű����) = 1 Then
                            .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!�ϰ�ʱ��)
                            .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!��¼ID)
                            .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!�޺���)) = 0, "��", Nvl(rsData!�޺���))
                            .TextMatrix(lngCurRow, lngCurCol + 2) = str��Լ��
                        Else
                            .TextMatrix(lngCurRow, lngCurCol) = _
                                IIf(Nvl(rsData!�Ű����) = 4 Or Nvl(rsData!�Ű����) = 5, "��ѭ(" & Val(Nvl(rsData!������Ŀ)) & "��)", Nvl(rsData!������Ŀ))
                            .TextMatrix(lngCurRow, lngCurCol + 1) = Nvl(rsData!�ϰ�ʱ��)
                            .Cell(flexcpData, lngCurRow, lngCurCol + 1) = Nvl(rsData!��¼ID)
                            .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Val(Nvl(rsData!�޺���)) = 0, "��", Nvl(rsData!�޺���))
                            .TextMatrix(lngCurRow, lngCurCol + 3) = str��Լ��
                        End If
                    Case Data_FixedRule
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!�ϰ�ʱ��)
                        .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!��¼ID)
                        .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!�޺���)) = 0, "��", Nvl(rsData!�޺���))
                        .TextMatrix(lngCurRow, lngCurCol + 2) = str��Լ��
                    Case Data_Plan
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!�ϰ�ʱ��)
                        .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!��¼ID)
                        
                        '�޺���flexcpData��ǳ����¼���ͣ���ʽ"�Ƿ���ʱ����|�Ƿ�����|�Ƿ�ͣ��|�Ƿ�����"
                        strRecordInfo = IIf(Val(Nvl(rsData!�Ƿ���ʱ����)) = 1, 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Val(Nvl(rsData!�Ƿ�����)) = 1, 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Nvl(rsData!ͣ�￪ʼʱ��) <> "", 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Nvl(rsData!����ҽ������) <> "", 1, 0)
                        If Nvl(rsData!����ҽ������) <> "" Then '���������ɫ������ʾ����ʾ����ҽ��
                            .TextMatrix(lngCurRow, lngCurCol) = .TextMatrix(lngCurRow, lngCurCol) & vbCrLf & "(" & Nvl(rsData!����ҽ������) & ")"
                        End If
                        .Cell(flexcpData, lngCurRow, lngCurCol + 1) = strRecordInfo
                        'δ�����Ĳ���ʾ�ѹ�������Լ��
                        .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(blnPublished, Nvl(rsData!�ѹ���, "0") & "/", "") & IIf(Nvl(rsData!�޺���) = "", "��", Nvl(rsData!�޺���))
                        
                        .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Nvl(rsData!ԤԼ���Ʒ�ʽ) = 1, "-", IIf(blnPublished, Val(Nvl(rsData!��Լ��)) & "/", "") & str��Լ��)
                        .Cell(flexcpData, lngCurRow, lngCurCol + 2) = Val(Nvl(rsData!����ID)) & "," & Val(Nvl(rsData!����ID))
                    End Select
                ElseIf bytDataStyle = Data_Plan Then
                    .Cell(flexcpData, lngCurRow, lngCurCol + 2) = Val(Nvl(rsData!����ID)) & "," & Val(Nvl(rsData!����ID))
                End If
                If blnAddRow Then lngCurRow = lngCurRow + 1
            End If
            rsData.MoveNext
        Loop
        
        Call SetGridFormat(vsfGrid, bytDataStyle, , , blnSortLoad, strStartDate, strEndDate)
        .Redraw = flexRDBuffered
        
        On Error Resume Next
        If .Rows > .FixedRows And .Cols > .FixedCols Then     'ȱʡ��λ��
            .Row = -1 '��֤��ѡ���в���������Ҳ����RowColChange�¼�
            .Row = IIf(lngOldRow < .FixedRows Or lngOldRow > .Rows - 1, IIf(.Rows > .FixedRows, .FixedRows + 1, .FixedRows), lngOldRow)
            .Col = IIf(lngoldCol = 0 Or lngoldCol > .Cols - 1, .FixedCols, lngoldCol)
            .ShowCell .Row, .Col  '������ʾ��ָ����Ԫ
        End If
    End With
    LoadPlanDataByRecordset = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function RefreshOnePlanData(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    Optional ByVal rsData As ADODB.Recordset, Optional ByVal lngCurRowOld As Long = -1, _
    Optional ByVal blnPublished As Boolean, Optional ByVal bytTableMode As Byte, _
    Optional ByVal strStartDate As String, Optional ByVal strEndDate As String) As Boolean
    'ˢ��ָ���к�Դ����
    '��Σ�
    '   vsfGrid VSFlexGrid������
    '   bytDataStyle ���ű�����
    '   rsData ���ﰲ�ż�¼��,ΪNothing��ʾ�޳����¼
    '   bytTableMode 0-�̶������,1-�³����,2-�ܳ����,3-ģ��
    '   strStartDate,strEndDate ��������ڷ�Χ���ܳ����ʱ����
    Dim lngCurRow As Long, lngCurCol As Long
    Dim lngStartRow As Long, lngEndRow As Long
    Dim i As Long, j As Long
    Dim blnFindCol As Boolean, strTemp As String
    Dim str��Լ�� As String, strRecordInfo As String
    Dim blnFindRow As Boolean, blnRefrashed As Boolean
    Dim blnHaveData As Boolean
    
    Err = 0: On Error GoTo errHandle
    With vsfGrid
'        .Redraw = flexRDNone
        '1.����ú�Դ�ĳ����¼
        lngCurRow = IIf(lngCurRowOld = -1, .Row, lngCurRowOld)
        '108641����ǰ������ǿ��У����ô���
        If .RowData(lngCurRow) = -1 Then RefreshOnePlanData = True: Exit Function
        If GetPlanGroupRange(vsfGrid, lngCurRow, lngStartRow, lngEndRow) = False Then Exit Function
        For i = lngEndRow To lngStartRow + 1 Step -1
            .RemoveItem i  '�Ƴ������У�ֻ����һ��
        Next
        lngCurRow = lngStartRow: lngEndRow = lngStartRow
        lngCurCol = gPlanGrid_FixedCols
        
        .Cell(flexcpText, lngCurRow, gPlanGrid_FixedCols, lngCurRow, .Cols - 1) = ""
        .Cell(flexcpData, lngCurRow, gPlanGrid_FixedCols, lngCurRow, .Cols - 1) = ""
        .Cell(flexcpForeColor, lngCurRow, gPlanGrid_FixedCols, lngCurRow, .Cols - 1) = .ForeColor
        Set .Cell(flexcpPicture, lngCurRow, gPlanGrid_FixedCols, lngCurRow, .Cols - 1) = Nothing
        .TextMatrix(lngCurRow, COL_����ID) = ""
        If bytDataStyle = Data_FixedRule Then
            .TextMatrix(lngCurRow, COL_��ʼʱ��) = ""
            .TextMatrix(lngCurRow, COL_��ֹʱ��) = ""
            .TextMatrix(lngCurRow, COL_�Ǽ�ʱ��) = ""
            .TextMatrix(lngCurRow, COL_��ʱ����) = ""
            .TextMatrix(lngCurRow, COL_�Ƿ����) = ""
        End If
        
        '2.���¼�������
        blnHaveData = True
        If rsData Is Nothing Then blnHaveData = False
        If rsData.RecordCount = 0 Then blnHaveData = False
        If blnHaveData = False Then
            .RemoveItem lngCurRow
            If .RowData(lngCurRow - 1) = -1 And lngCurRow - 1 >= .FixedRows Then
                .RemoveItem lngCurRow - 1
            End If
            RefreshOnePlanData = True
            Exit Function
        End If
        
        Do While Not rsData.EOF
'            lngCurRow = lngStartRow
            blnFindRow = False: blnFindCol = False
            '2.1ȷ����ǰ��
            Select Case bytDataStyle
            Case Data_Templet  'ģ��
                If Nvl(rsData!�Ű����) <> 1 Then '��������
                    '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
                    lngCurCol = .Cols - 4: blnFindCol = True
                Else
                    If Nvl(rsData!������Ŀ) <> "" Then
                        strTemp = Nvl(rsData!������Ŀ)
                        For i = lngCurCol To .Cols - 1 Step 3
                            If strTemp = .Cell(flexcpData, 0, i) Then
                                lngCurCol = i: blnFindCol = True
                                Exit For
                            End If
                        Next
                        'û�ҵ��ٴӿ�ʼ�����ң���Ҫ�ǰ�������Ŀ���������˳��һ��
                        If blnFindCol = False Then
                            For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
                                If strTemp = .Cell(flexcpData, 0, i) Then
                                    lngCurCol = i: blnFindCol = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            Case Data_FixedRule  '�̶�����
                If Nvl(rsData!������Ŀ) <> "" Then
                    strTemp = Nvl(rsData!������Ŀ)
                    For i = lngCurCol To .Cols - 1 Step 3
                        If strTemp = .Cell(flexcpData, 0, i) Then
                            lngCurCol = i: blnFindCol = True
                            Exit For
                        End If
                    Next
                    'û�ҵ��ٴӿ�ʼ�����ң���Ҫ�ǰ�������Ŀ���������˳��һ��
                    If blnFindCol = False Then
                        For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
                            If strTemp = .Cell(flexcpData, 0, i) Then
                                lngCurCol = i: blnFindCol = True
                                Exit For
                            End If
                        Next
                    End If
                End If
            Case Else '���ż�¼
                If Nvl(rsData!��������) = "" Then
                    '��������Ϊ��,ȡ���ŵĿ�ʼʱ��,��Ҫ������ܿ��µ��ܳ�����޳����¼ʱ,����������ʾ��ͬһ��
                    strTemp = Format(Nvl(rsData!��ʼʱ��), "yyyy-mm-dd")
                    lngCurCol = gPlanGrid_FixedCols
                Else
                    strTemp = Format(Nvl(rsData!��������), "yyyy-mm-dd")
                End If
                If IsDate(strTemp) Then
                    For i = lngCurCol To .Cols - 1 Step 3
                        If DateDiff("d", strTemp, .Cell(flexcpData, 0, i)) = 0 Then
                            lngCurCol = i: blnFindCol = True
                            Exit For
                        End If
                    Next
                End If
            End Select
            
            If blnFindCol Then
                '2.2ȷ����ǰ��
                For i = lngEndRow To lngStartRow Step -1
                    If .TextMatrix(i, lngCurCol) <> "" Then   '�����ؿ��л�����������
                        lngCurRow = i + 1: blnFindRow = True
                        Exit For
                    End If
                Next
                If blnFindRow = False Then lngCurRow = lngStartRow
            End If
            
            '2.3��������
            If blnRefrashed = False Then
                blnRefrashed = True
                If Not (bytDataStyle = Data_Plan And IsDate(strStartDate) And IsDate(strEndDate)) Then
                    .TextMatrix(lngCurRow, COL_����ID) = Val(Nvl(rsData!����ID))
                End If
                If bytDataStyle = Data_FixedRule Then
                    .TextMatrix(lngCurRow, COL_��ʼʱ��) = Format(Nvl(rsData!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngCurRow, COL_��ֹʱ��) = Format(Nvl(rsData!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngCurRow, COL_�Ǽ�ʱ��) = Format(Nvl(rsData!�Ǽ�ʱ��), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngCurRow, COL_��ʱ����) = Val(Nvl(rsData!��ʱ����))
                    .TextMatrix(lngCurRow, COL_�Ƿ����) = Val(Nvl(rsData!�Ƿ����))
                End If
                
                '������Դʱ����Ч�����Ч
                Select Case bytTableMode
                Case 0 '�̶������
                    If Val(Nvl(rsData!�Ƿ���Ч)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("FixedItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("InvalidFixedItem")
                    End If
                Case 1 '�³����
                    If Val(Nvl(rsData!�Ƿ���Ч)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("MonthItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("InvalidMonthItem")
                    End If
                Case 2 '�ܳ����
                    If Val(Nvl(rsData!�Ƿ���Ч)) = 1 Then
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("WeekItem")
                    Else
                        Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = GetPlanItemImage("InvalidWeekItem")
                    End If
                End Select
                .Cell(flexcpPictureAlignment, lngCurRow, COL_ͼ��) = flexAlignCenterCenter
            End If
            If lngEndRow < lngCurRow Then
                '�����в�������1��
                lngEndRow = lngEndRow + 1: .AddItem "", lngEndRow
                
                lngCurRow = lngEndRow
                .RowData(lngCurRow) = .RowData(lngCurRow - 1) '�������ý���ɫ
                
                For j = 0 To gPlanGrid_FixedCols - 1
                    .TextMatrix(lngCurRow, j) = .TextMatrix(lngCurRow - 1, j)
                Next
                Set .Cell(flexcpPicture, lngCurRow, COL_ͼ��) = .Cell(flexcpPicture, lngCurRow - 1, COL_ͼ��)
                .Cell(flexcpPictureAlignment, lngCurRow, COL_ͼ��) = flexAlignCenterCenter
            End If
            
            If bytDataStyle = Data_Plan And IsDate(strStartDate) And IsDate(strEndDate) Then
                '������ܿ��µ��ܳ�����޳���а���IDΪ��ǰѡ�������к�Դ�İ���ID
                If IsDate(Nvl(rsData!��ʼʱ��)) And IsDate(Nvl(rsData!��ֹʱ��)) Then
                    If DateDiff("d", Nvl(rsData!��ʼʱ��), strStartDate) <= 0 And DateDiff("d", Nvl(rsData!��ֹʱ��), strEndDate) >= 0 Then
                        .TextMatrix(lngCurRow, COL_����ID) = Nvl(rsData!����ID)
                    End If
                End If
            End If
            
            If blnFindCol Then
                '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
                'ԤԼ���Ʒ�ʽ��0-����ԤԼ����;1-�ú����ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
                If Nvl(rsData!�ϰ�ʱ��) <> "" Then
                    str��Լ�� = IIf(Nvl(rsData!ԤԼ���Ʒ�ʽ) = 1, "-", _
                        IIf(Val(Nvl(rsData!��Լ��)) = 0, IIf(Val(Nvl(rsData!�޺���)) = 0, "��", _
                            Val(Nvl(rsData!�޺���))), Val(Nvl(rsData!��Լ��))))
                    Select Case bytDataStyle
                    Case Data_Templet
                        If Nvl(rsData!�Ű����) = 1 Then
                            .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!�ϰ�ʱ��)
                            .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!��¼ID)
                            .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!�޺���)) = 0, "��", Nvl(rsData!�޺���))
                            .TextMatrix(lngCurRow, lngCurCol + 2) = str��Լ��
                        Else
                            .TextMatrix(lngCurRow, lngCurCol) = _
                                IIf(Nvl(rsData!�Ű����) = 4 Or Nvl(rsData!�Ű����) = 5, "��ѭ(" & Val(Nvl(rsData!������Ŀ)) & "��)", Nvl(rsData!������Ŀ))
                            .TextMatrix(lngCurRow, lngCurCol + 1) = Nvl(rsData!�ϰ�ʱ��)
                            .Cell(flexcpData, lngCurRow, lngCurCol + 1) = Nvl(rsData!��¼ID)
                            .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Val(Nvl(rsData!�޺���)) = 0, "��", Nvl(rsData!�޺���))
                            .TextMatrix(lngCurRow, lngCurCol + 3) = str��Լ��
                        End If
                    Case Data_FixedRule
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!�ϰ�ʱ��)
                        .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!��¼ID)
                        .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!�޺���)) = 0, "��", Nvl(rsData!�޺���))
                        .TextMatrix(lngCurRow, lngCurCol + 2) = str��Լ��
                    Case Data_Plan
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!�ϰ�ʱ��)
                        .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!��¼ID)
                        
                        '�޺���flexcpData��ǳ����¼���ͣ���ʽ"�Ƿ���ʱ����|�Ƿ�����|�Ƿ�ͣ��|�Ƿ�����"
                        strRecordInfo = IIf(Val(Nvl(rsData!�Ƿ���ʱ����)) = 1, 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Val(Nvl(rsData!�Ƿ�����)) = 1, 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Nvl(rsData!ͣ�￪ʼʱ��) <> "", 1, 0)
                        strRecordInfo = strRecordInfo & "|" & IIf(Nvl(rsData!����ҽ������) <> "", 1, 0)
                        If Nvl(rsData!����ҽ������) <> "" Then '���������ɫ������ʾ����ʾ����ҽ��
                            .TextMatrix(lngCurRow, lngCurCol) = .TextMatrix(lngCurRow, lngCurCol) & vbCrLf & "(" & Nvl(rsData!����ҽ������) & ")"
                        End If
                        .Cell(flexcpData, lngCurRow, lngCurCol + 1) = strRecordInfo
                        .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(blnPublished, Nvl(rsData!�ѹ���, "0") & "/", "") & IIf(Nvl(rsData!�޺���) = "", "��", Nvl(rsData!�޺���))
                        
                        .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Nvl(rsData!ԤԼ���Ʒ�ʽ) = 1, "-", IIf(blnPublished, Val(Nvl(rsData!��Լ��)) & "/", "") & str��Լ��)
                        .Cell(flexcpData, lngCurRow, lngCurCol + 2) = Val(Nvl(rsData!����ID)) & "," & Val(Nvl(rsData!����ID))
                    End Select
                ElseIf bytDataStyle = Data_Plan Then
                    .Cell(flexcpData, lngCurRow, lngCurCol + 2) = Val(Nvl(rsData!����ID)) & "," & Val(Nvl(rsData!����ID))
                End If
            End If
            rsData.MoveNext
        Loop
            
        '3.���ø�ʽ
        Call SetGridFormat(vsfGrid, bytDataStyle, lngStartRow, lngEndRow, False, strStartDate, strEndDate)
'        .Redraw = flexRDBuffered
    End With
    RefreshOnePlanData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetGridFormat(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    Optional ByVal lngStartRow As Long = -1, Optional ByVal lngEndRow As Long = -1, _
    Optional ByVal blnSortLoad As Boolean, _
    Optional ByVal strStartDate As String, Optional ByVal strEndDate As String)
    '���õ�Ԫ���ʽ
    '��Σ�
    '   lngStartRow ����ʱ��ȡ.FixedRows
    '   lngEndRow ����ʱ��ȡ.Rows-1
    '   strStartDate,strEndDate ��������ڷ�Χ���ܳ����ʱ����
    Dim i As Long, j As Long
    Dim lngCurRowInGroup As Long, strSpace As String
    Dim intDataType As Integer
    Dim varRecordType As Variant
    Dim lngSortCol As Long, strSort As String
    
    With vsfGrid
        If lngStartRow = -1 Then lngStartRow = .FixedRows
        If lngEndRow = -1 Then lngEndRow = .Rows - 1
        
        '���⴦���Ա��ܹ��ϲ�����
        lngCurRowInGroup = 0 '���������к�
        For i = lngStartRow To lngEndRow
            If .RowData(i) = -1 Then lngCurRowInGroup = 0
            For j = 0 To .Cols - 1
                If .RowData(i) = -1 Then Exit For
                If .TextMatrix(i, j) = "" Then .TextMatrix(i, j) = " " '��ֹ����Ϊ�ղ��ܺϲ�
                If .RowData(i - 1) <> -1 And j >= gPlanGrid_FixedCols Then '�Ƿ�Ϊ��������
                    If (j - gPlanGrid_FixedCols) Mod 3 = 0 Then '"ʱ���"��
                        If .TextMatrix(i, j) = " " Then '�ϲ�����Ŀ���
                            .Cell(flexcpAlignment, i - 1, j, i, j) = flexAlignLeftCenter
                            .TextMatrix(i, j) = .TextMatrix(i - 1, j)
                            .TextMatrix(i, j + 1) = .TextMatrix(i - 1, j + 1)
                            .TextMatrix(i, j + 2) = .TextMatrix(i - 1, j + 2)
                        Else
                            strSpace = Space(lngCurRowInGroup Mod 2) '���ո񣬷�ֹ������ͬ�ϲ�
                            .TextMatrix(i, j + 1) = strSpace & .TextMatrix(i, j + 1) & strSpace
                            .TextMatrix(i, j + 2) = strSpace & .TextMatrix(i, j + 2) & strSpace
                        End If
                        'ģ��������Ű�
                        If bytDataStyle = Data_Templet And j = .Cols - 4 Then
                            If .TextMatrix(i, j) = " " Then '�ϲ�����Ŀ���
                                .TextMatrix(i, j + 3) = .TextMatrix(i - 1, j + 3)
                            Else
                                strSpace = Space(lngCurRowInGroup Mod 2) '���ո񣬷�ֹ������ͬ�ϲ�
                                .TextMatrix(i, j + 1) = LTrim(.TextMatrix(i, j + 1)) 'ȥ����ߵĿո�
                                .TextMatrix(i, j + 3) = strSpace & .TextMatrix(i, j + 3) & strSpace
                            End If
                            j = j + 1
                        End If
                        j = j + 2
                    End If
                End If
            Next
            If .RowData(i) <> -1 Then lngCurRowInGroup = lngCurRowInGroup + 1
        Next
        
        '�б���ɫ
        If .FixedCols <= .Cols - 1 Then
            For i = lngStartRow To lngEndRow
                .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = .RowData(i)
                .RowHeight(i) = 420
                If .RowData(i) = -1 Then .RowHeight(i) = 0 '���������и߶�
            Next
        End If
        
        If bytDataStyle = Data_Plan Then
            '�޺���flexcpData��ǳ����¼���ͣ���ʽ"�Ƿ���ʱ����|�Ƿ�����|�Ƿ�ͣ��|�Ƿ�����"
            For i = lngStartRow To lngEndRow
                For j = gPlanGrid_FixedCols To .Cols - 1 Step 3
                    varRecordType = Split(.Cell(flexcpData, i, j + 1) & "|||", "|")
                    If varRecordType(0) = 1 Then '��ʱ���������ɫ������ʾ
                        .Cell(flexcpForeColor, i, j, i, j + 2) = vbBlue
                    End If
                    If varRecordType(1) = 1 Then '��������ʾ����ͼ��
                        .Cell(flexcpPicture, i, j) = GetLockImage
                        .Cell(flexcpPictureAlignment, i, j) = flexAlignRightBottom
                    End If
                    If varRecordType(2) = 1 Then 'ͣ����ú�ɫ������ʾ
                        .Cell(flexcpBackColor, i, j, i, j + 2) = vbRed
                    End If
                    If varRecordType(3) = 1 Then '���������ɫ������ʾ����ʾ����ҽ��
                        .Cell(flexcpForeColor, i, j, i, j + 2) = vbBlue
                    End If
                Next
            Next
        End If
        
        '����ܳ�������ǵ�ǰ�����������û�ɫ������ʾ
        If bytDataStyle = Data_Plan And IsDate(strStartDate) And IsDate(strEndDate) Then
            For j = gPlanGrid_FixedCols To .Cols - 1 Step 3
                If Not (CDate(.Cell(flexcpData, 0, j)) >= CDate(strStartDate) _
                    And CDate(.Cell(flexcpData, 0, j)) <= CDate(strEndDate)) Then
                    .Cell(flexcpForeColor, 0, j, .Rows - 1, j + 2) = &HC0C0C0
                End If
            Next
        End If
        
        '��������ͼ��
        If blnSortLoad = False Then
            Call SetSortFlexcpData(vsfGrid) '��������ʶ
            vsfGrid.Cell(flexcpData, 1, COL_����) = "ASC" 'ȱʡ��������������
        End If
        
        If .Cell(flexcpData, 1, COL_����) <> "-" Then strSort = .Cell(flexcpData, 1, COL_����): lngSortCol = COL_����
        If .Cell(flexcpData, 1, COL_����) <> "-" Then strSort = .Cell(flexcpData, 1, COL_����): lngSortCol = COL_����
        If .Cell(flexcpData, 1, COL_����) <> "-" Then strSort = .Cell(flexcpData, 1, COL_����): lngSortCol = COL_����
        If .Cell(flexcpData, 1, COL_��Ŀ) <> "-" Then strSort = .Cell(flexcpData, 1, COL_��Ŀ): lngSortCol = COL_��Ŀ
        If .Cell(flexcpData, 1, COL_ҽ��) <> "-" Then strSort = .Cell(flexcpData, 1, COL_ҽ��): lngSortCol = COL_ҽ��
        If lngSortCol = 0 Then lngSortCol = COL_����: strSort = "ASC"
        
        If strSort = "ASC" Then
            Set .Cell(flexcpPicture, 0, lngSortCol, 1, lngSortCol) = GetSortIcon("ASC")
        ElseIf strSort = "DESC" Then
            Set .Cell(flexcpPicture, 0, lngSortCol, 1, lngSortCol) = GetSortIcon("DESC")
        Else
            Set .Cell(flexcpPicture, 0, lngSortCol, 1, lngSortCol) = Nothing
        End If
        .Cell(flexcpPictureAlignment, 0, lngSortCol) = flexAlignCenterBottom
        
        On Error Resume Next
        If .Row < .FixedRows And .Rows > .FixedRows Then .Row = .FixedRows + 1
        If .RowData(.Row) = -1 And .Rows > .Row + 1 Then .Row = .Row + 1
        If .Col < .FixedCols And .Cols > .FixedCols Then .Col = .FixedCols
        Call SetPlanGridRangeColor(vsfGrid, bytDataStyle)
    End With
End Sub

Public Sub ShowHolidayToPlan(vsfGrid As VSFlexGrid, ByVal dtStartDate As Date, ByVal dtEndStart As Date)
    '��ʾ�����ڼ���
    Dim strSQL As String, rsHoliday As ADODB.Recordset
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandler
    '�����ڼ���
    strSQL = "Select ��ʼ����, ��ֹ����, ��������" & vbNewLine & _
            " From �������ձ�" & vbNewLine & _
            " Where ���� = 0 And ��� = To_Number(To_Char([1], 'yyyy'))" & vbNewLine & _
            "       And Not(��ʼ���� > [2] Or ��ֹ���� < [1])"
    Set rsHoliday = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ڼ�������", dtStartDate, dtEndStart)
    If rsHoliday.RecordCount = 0 Then Exit Sub
    
    With vsfGrid
        For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
            If IsDate(.Cell(flexcpData, 0, i)) Then
                rsHoliday.MoveFirst
                Do While Not rsHoliday.EOF
                    If CDate(.Cell(flexcpData, 0, i)) >= CDate(Nvl(rsHoliday!��ʼ����)) _
                        And CDate(.Cell(flexcpData, 0, i)) <= CDate(Nvl(rsHoliday!��ֹ����)) Then
                        .TextMatrix(0, i) = .TextMatrix(0, i) & "(" & Nvl(rsHoliday!��������) & ")"
                        .TextMatrix(0, i + 1) = .TextMatrix(0, i + 1) & "(" & Nvl(rsHoliday!��������) & ")"
                        .TextMatrix(0, i + 2) = .TextMatrix(0, i + 2) & "(" & Nvl(rsHoliday!��������) & ")"
                        Exit Do
                    End If
                    rsHoliday.MoveNext
                Loop
            End If
        Next
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ShowStopVisitPlan(vsfGrid As VSFlexGrid, ByVal dtStartDate As Date, ByVal dtEndStart As Date, _
    Optional ByVal lng��ԴId As Long)
    '��ʾͣ�ﰲ��
    Dim strSQL As String, rsStopVisitPlan As ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo errHandler
    'ͣ�ﰲ��
    strSQL = "Select b.Id As ��Դid, Trunc(a.��ʼʱ��) As ��ʼʱ��, a.��ֹʱ��, a.ͣ��ԭ��" & vbNewLine & _
            " From �ٴ�����ͣ���¼ A, �ٴ������Դ B" & vbNewLine & _
            " Where a.������ = b.ҽ������ And b.ҽ��id Is Not Null" & vbNewLine & _
            "       And a.��¼id Is Null And a.������ Is Not Null And a.ȡ���� Is Null" & vbNewLine & _
            "       And Not (a.��ʼʱ�� > [2] Or a.��ֹʱ�� < [1])" & _
                    IIf(lng��ԴId = 0, "", " And b.ID = [3]")

    Set rsStopVisitPlan = zlDatabase.OpenSQLRecord(strSQL, "��ȡͣ�ﰲ��", dtStartDate, dtEndStart, lng��ԴId)
    If rsStopVisitPlan.RecordCount = 0 Then Exit Sub
    
    rsStopVisitPlan.MoveFirst
    With vsfGrid
        For i = gPlanGrid_FixedCols To .Cols - 1 Step 3
            blnFind = False
            If IsDate(.Cell(flexcpData, 0, i)) Then
                For j = .FixedRows To .Rows - 1
                    If blnFind And lng��ԴId <> 0 And lng��ԴId <> Val(.TextMatrix(j, COL_��ԴID)) Then
                        Exit For
                    End If
                    rsStopVisitPlan.MoveFirst
                    Do While Not rsStopVisitPlan.EOF
                        If CDate(.Cell(flexcpData, 0, i)) >= CDate(Nvl(rsStopVisitPlan!��ʼʱ��)) _
                            And CDate(.Cell(flexcpData, 0, i)) <= CDate(Nvl(rsStopVisitPlan!��ֹʱ��)) _
                            And Val(.TextMatrix(j, COL_��ԴID)) = Val(Nvl(rsStopVisitPlan!��ԴID)) Then
                            blnFind = True
                            If .Cell(flexcpBackColor, j, i) <> vbRed Then
                                .Cell(flexcpForeColor, j, i, j, i + 2) = vbRed
                            End If
                            .Cell(flexcpText, j, i) = Trim(.TextMatrix(j, i)) & IIf(Trim(.TextMatrix(j, i)) = "", "", vbCrLf) & "(" & Nvl(rsStopVisitPlan!ͣ��ԭ��) & ")"
                            Exit Do
                        End If
                        rsStopVisitPlan.MoveNext
                    Loop
                Next
            End If
        Next
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub SetPlanGridSelRange(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle)
    '���ܣ�����ѡ�����з�Χ
    Dim lngRowStart As Long, lngRowEnd As Long '��ʼ�к���ֹ��
    Dim lngColStart As Long, lngColEnd As Long '��ʼ�к���ֹ��
    
    On Error Resume Next
    With vsfGrid
        If Not .Visible Then Exit Sub
        If .Row < .FixedRows Or .RowSel < .FixedRows Then Exit Sub
        If .Col < gPlanGrid_FixedCols And .ColSel < gPlanGrid_FixedCols Then Exit Sub
            
        'ѡ���з�Χ
        lngRowStart = .Row: lngRowEnd = .RowSel
        
        'ѡ���з�Χ
        If .Col >= gPlanGrid_FixedCols And .ColSel < gPlanGrid_FixedCols Then
            '��ʼ��Ϊ�����У�ĩβ��Ϊ�������У���ֻѡ���������
            lngColStart = .ColSel
            lngColEnd = lngColStart
        End If
        
        If .Col < gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
            '��ʼ��Ϊ�������У�ĩβ��Ϊ�����У���ѡ�������а��ŷ�Χ
            lngColStart = GetPlanItemNameCol(.ColSel) 'ȷ��"ʱ���"��
            lngColEnd = lngColStart + 2
        End If
        
        If .Col >= gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
            lngColStart = GetPlanItemNameCol(.Col) 'ȷ��"ʱ���"��
            lngColEnd = GetPlanItemNameCol(.ColSel)
            If lngColStart > lngColEnd Then
                lngColStart = lngColStart + 2
            Else
                lngColEnd = lngColEnd + 2
            End If
        End If
        
        'ģ�����һ�����⴦��
        If bytDataStyle = Data_Templet And lngColStart = .Cols - 4 Then
            lngColEnd = lngColEnd + 1: lngColStart = lngColStart + 1
        End If
        
        '����ѡ��
        .ForeColorSel = .Cell(flexcpForeColor, .RowSel, .ColSel)
        .Select lngRowStart, lngColStart, lngRowEnd, lngColEnd
    End With
End Sub

Public Sub SetPlanGridRangeColor(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    Optional ByVal strOldSelRange As String)
    '���ܣ�����ѡ������ɫ,.RowData�д�����ɫֵ
    'strOldSelRange:��һ��ѡ���������򣬸�ʽ"��ʼ��|������|��ʼ��|������"
    Dim lngRowStart As Long, lngRowEnd As Long '��ʼ�к���ֹ��
    Dim lngColStart As Long, lngColEnd As Long '��ʼ�к���ֹ��
    Dim varRecordType As Variant
    Dim i As Long, j As Long, lngTemp As Long
    Dim varTemp As Variant
    
    On Error Resume Next
    With vsfGrid
        If Not .Visible Then Exit Sub
        '�ָ���ɫ
        If strOldSelRange <> "" Then
            varTemp = Split(strOldSelRange & "|||", "|")
            lngRowStart = varTemp(0): lngRowEnd = varTemp(1)
            lngColStart = varTemp(2): lngColEnd = varTemp(3)
            
            If lngRowStart < .FixedRows Or lngRowEnd < .FixedRows Then
            ElseIf lngColStart < gPlanGrid_FixedCols And lngColEnd < gPlanGrid_FixedCols Then
                .Cell(flexcpBackColor, lngRowStart, lngColStart) = .RowData(lngRowStart)
            Else
                If lngRowStart > lngRowEnd Then lngTemp = lngRowStart: lngRowStart = lngRowEnd: lngRowEnd = lngTemp
                If lngColStart > lngColEnd Then lngTemp = lngColStart: lngColStart = lngColEnd: lngColEnd = lngTemp
                If lngRowStart < .FixedRows Then lngRowStart = .FixedRows
                If lngColStart < .FixedCols Then lngColStart = .FixedCols
                For i = lngRowStart To lngRowEnd
                    .Cell(flexcpBackColor, i, lngColStart, i, lngColEnd) = .RowData(i)
                    For j = lngColStart To lngColEnd
                        '�޺���flexcpData��ǳ����¼���ͣ���ʽ"�Ƿ���ʱ����|�Ƿ�����|�Ƿ�ͣ��|�Ƿ�����"
                        varRecordType = Split(.Cell(flexcpData, i, j + 1) & "|||", "|")
                        If Val(varRecordType(2)) = 1 Then   'ͣ��
                            .Cell(flexcpBackColor, i, j, i, j + 2) = vbRed
                            'ͣ��ʱ�ı���������ɫ�ģ���������лָ�
                            If Val(varRecordType(0)) = 1 Or Val(varRecordType(3)) = 1 Then '��ʱ����/����
                                .Cell(flexcpForeColor, i, j, i, j + 2) = vbBlue
                            Else
                                .Cell(flexcpForeColor, i, j, i, j + 2) = vbBlack
                            End If
                        End If
                        j = j + 2
                    Next
                Next
            End If
        End If
        
        If .Row < .FixedRows Or .RowSel < .FixedRows Then Exit Sub
        If .Col < gPlanGrid_FixedCols And .ColSel < gPlanGrid_FixedCols Then
            lngRowStart = .Row: lngColStart = .Col
            If lngRowStart >= .FixedRows And lngColStart >= .FixedCols Then
                .Cell(flexcpBackColor, lngRowStart, lngColStart) = .BackColorSel
            End If
        Else
            'ѡ���з�Χ
            lngRowStart = .Row: lngRowEnd = .RowSel
            
            'ѡ���з�Χ
            If .Col >= gPlanGrid_FixedCols And .ColSel < gPlanGrid_FixedCols Then
                '��ʼ��Ϊ�����У�ĩβ��Ϊ�������У���ֻѡ���������
                lngColStart = .ColSel
                lngColEnd = lngColStart
                lngRowEnd = lngRowStart
            End If
            
            If .Col < gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
                '��ʼ��Ϊ�������У�ĩβ��Ϊ�����У���ѡ�������а��ŷ�Χ
                lngColStart = GetPlanItemNameCol(.ColSel) 'ȷ��"ʱ���"��
                lngColEnd = lngColStart + 2
            End If
            
            If .Col >= gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
                lngColStart = GetPlanItemNameCol(.Col) 'ȷ��"ʱ���"��
                lngColEnd = GetPlanItemNameCol(.ColSel)
                If lngColStart > lngColEnd Then
                    lngColStart = lngColStart + 2
                Else
                    lngColEnd = lngColEnd + 2
                End If
            End If
            
            'ģ�����һ�����⴦��
            If bytDataStyle = Data_Templet And lngColStart = .Cols - 4 Then
                lngColEnd = lngColEnd + 1: lngColStart = lngColStart + 1
            End If
            
            '����ѡ��
            If lngRowStart <> .Row Or lngColStart <> .Col _
                Or lngRowEnd <> .RowSel Or lngColEnd <> .ColSel Then
                .Select lngRowStart, lngColStart, lngRowEnd, lngColEnd
            End If
            
            If lngRowStart > lngRowEnd Then lngTemp = lngRowStart: lngRowStart = lngRowEnd: lngRowEnd = lngTemp
            If lngColStart > lngColEnd Then lngTemp = lngColStart: lngColStart = lngColEnd: lngColEnd = lngTemp
            If lngRowStart < .FixedRows Then lngRowStart = .FixedRows
            If lngColStart < .FixedCols Then lngColStart = .FixedCols
            For i = lngRowStart To lngRowEnd
                .Cell(flexcpBackColor, i, lngColStart, i, lngColEnd) = .BackColorSel
                For j = lngColStart To lngColEnd
                    '�޺���flexcpData��ǳ����¼���ͣ���ʽ"�Ƿ���ʱ����|�Ƿ�����|�Ƿ�ͣ��|�Ƿ�����"
                    varRecordType = Split(.Cell(flexcpData, i, j + 1) & "|||", "|")
                    If Val(varRecordType(2)) = 1 Then   'ͣ��
                        .Cell(flexcpForeColor, i, j, i, j + 2) = vbRed
                    End If
                    j = j + 2
                Next
            Next
        End If
    End With
End Sub

Public Function GetPlanItemNameCol(ByVal lngCurCol As Long) As Long
    On Error GoTo errHandle
    'ȷ��"ʱ��"�е�������
    GetPlanItemNameCol = lngCurCol - Choose(((lngCurCol - gPlanGrid_FixedCols) Mod 3) + 1, 0, 1, 2)
    Exit Function
errHandle:
    Err.Clear
    GetPlanItemNameCol = 0
End Function

Public Function GetPlanGroupRange(vsfGrid As VSFlexGrid, _
    ByVal lngCurRow As Long, ByRef lngRowStart As Long, ByRef lngRowEnd As Long) As Boolean
    '��ǰ�����������Χ
    Dim i As Integer
    
    With vsfGrid
        lngRowStart = .FixedRows
        For i = lngCurRow To .FixedRows Step -1
            If .RowData(i) = -1 Then lngRowStart = i + 1: Exit For
        Next
        lngRowEnd = .Rows - 1
        For i = lngCurRow + 1 To .Rows - 1
            If .RowData(i) = -1 And i <> .Rows - 1 Then lngRowEnd = i - 1: Exit For
        Next
    End With
    GetPlanGroupRange = True
End Function

Public Function GetPlanSortCircleStr(vsfGrid As VSFlexGrid, ByVal bytDataStyle As gPlanGrid_DataStyle, _
    lngRow As Long, lngCol As Long) As String
    '��ȡ����
    Dim i As Long, strSort As String
    
    On Error GoTo errHandle
    '��������
    Select Case lngCol
    Case COL_����
        strSort = SortCircle(vsfGrid, lngCol, "����") & "�������,"
    Case COL_����
        strSort = SortCircle(vsfGrid, lngCol, "�������")
    Case COL_����
        strSort = SortCircle(vsfGrid, lngCol, "����") & "�������,"
    Case COL_��Ŀ
        strSort = SortCircle(vsfGrid, lngCol, "�շ���Ŀ") & "�������,"
    Case COL_ҽ��
        strSort = SortCircle(vsfGrid, lngCol, "ҽ������") & "�������,"
    End Select

    If strSort <> "" Then
        Call SetSortFlexcpData(vsfGrid, lngCol) '��������ʶ
        
        Select Case bytDataStyle
        Case Data_FixedRule
            strSort = strSort & "��ʼʱ��,��ֹʱ��,������Ŀ,�ϰ�ʱ��"
        Case Data_Templet
            strSort = strSort & "������Ŀ,�ϰ�ʱ��"
        Case Data_Plan
            strSort = strSort & "��������,�ϰ�ʱ��"
        End Select
    End If
    GetPlanSortCircleStr = strSort
    Exit Function
errHandle:
    Err.Clear
End Function

Private Sub SetSortFlexcpData(vsfGrid As VSFlexGrid, Optional ByVal lngSortCol As Long)
    '��������ʶ
    On Error GoTo errHandle
    If lngSortCol <> COL_���� Then vsfGrid.Cell(flexcpData, 1, COL_����) = "-"
    If vsfGrid.Cell(flexcpData, 1, COL_����) = "-" Then Set vsfGrid.Cell(flexcpPicture, 0, COL_����, 1, COL_����) = Nothing
    If lngSortCol <> COL_���� Then vsfGrid.Cell(flexcpData, 1, COL_����) = "-"
    If vsfGrid.Cell(flexcpData, 1, COL_����) = "-" Then Set vsfGrid.Cell(flexcpPicture, 0, COL_����, 1, COL_����) = Nothing
    If lngSortCol <> COL_���� Then vsfGrid.Cell(flexcpData, 1, COL_����) = "-"
    If vsfGrid.Cell(flexcpData, 1, COL_����) = "-" Then Set vsfGrid.Cell(flexcpPicture, 0, COL_����, 1, COL_����) = Nothing
    If lngSortCol <> COL_��Ŀ Then vsfGrid.Cell(flexcpData, 1, COL_��Ŀ) = "-"
    If vsfGrid.Cell(flexcpData, 1, COL_��Ŀ) = "-" Then Set vsfGrid.Cell(flexcpPicture, 0, COL_��Ŀ, 1, COL_��Ŀ) = Nothing
    If lngSortCol <> COL_ҽ�� Then vsfGrid.Cell(flexcpData, 1, COL_ҽ��) = "-"
    If vsfGrid.Cell(flexcpData, 1, COL_ҽ��) = "-" Then Set vsfGrid.Cell(flexcpPicture, 0, COL_ҽ��, 1, COL_ҽ��) = Nothing
    Exit Sub
errHandle:
    Err.Clear
End Sub

Private Function SortCircle(vsfGrid As VSFlexGrid, ByVal lngCol As Long, ByVal strColName As String) As String
    'Cell(flexcpData, 1, lngCol)��¼�˵�ǰ����ʽ��ע�������¼�������ʱ���
    Select Case vsfGrid.Cell(flexcpData, 1, lngCol)
    Case ""
        If lngCol = COL_���� Then '�����г�ʼʱ�������������е�
            vsfGrid.Cell(flexcpData, 1, lngCol) = "DESC"
            SortCircle = strColName & " DESC,"
        Else
            vsfGrid.Cell(flexcpData, 1, lngCol) = "ASC"
            SortCircle = strColName & " Asc,"
        End If
    Case "ASC" '����
        vsfGrid.Cell(flexcpData, 1, lngCol) = "DESC"
        SortCircle = strColName & " Desc,"
    Case "DESC" '����
        If lngCol = COL_���� Then '������Ҫô����Ҫô����
            vsfGrid.Cell(flexcpData, 1, lngCol) = "ASC"
            SortCircle = strColName & " Asc,"
        Else
            vsfGrid.Cell(flexcpData, 1, lngCol) = "-"
            SortCircle = ""
        End If
    Case "-" '������
        vsfGrid.Cell(flexcpData, 1, lngCol) = "ASC"
        SortCircle = strColName & " Asc,"
    End Select
End Function

'�޺���flexcpData��ǳ����¼���ͣ���ʽ"�Ƿ���ʱ����|�Ƿ�����|�Ƿ�ͣ��|�Ƿ�����"
Public Function PlanIsLocked(vsfGrid As VSFlexGrid, _
    Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1) As Boolean
    '�Ƿ�����״̬
    Dim lngCurRow As Long, lngCurCol As Long
    Dim varRecordType As Variant
    
    With vsfGrid
        lngCurRow = IIf(lngRow = -1, .Row, lngRow)
        lngCurCol = IIf(lngCol = -1, GetPlanItemNameCol(.Col), GetPlanItemNameCol(lngCol))
                    
        varRecordType = Split(.Cell(flexcpData, lngCurRow, lngCurCol + 1) & "|||", "|")
        If Val(varRecordType(1)) = 1 Then
            PlanIsLocked = True
        End If
    End With
End Function

Public Function PlanIsStopVisit(vsfGrid As VSFlexGrid) As Boolean
    '�Ƿ�ͣ��״̬
    Dim lngCurRow As Long, lngCurCol As Long
    Dim varRecordType As Variant
    
    lngCurRow = vsfGrid.Row
    lngCurCol = GetPlanItemNameCol(vsfGrid.Col)
    
    varRecordType = Split(vsfGrid.Cell(flexcpData, lngCurRow, lngCurCol + 1) & "|||", "|")
    If Val(varRecordType(2)) = 1 Then
        PlanIsStopVisit = True
    End If
End Function

Public Function PlanIsReplaceDoctor(vsfGrid As VSFlexGrid) As Boolean
    '�Ƿ�������״̬
    Dim lngCurRow As Long, lngCurCol As Long
    Dim varRecordType As Variant
    
    lngCurRow = vsfGrid.Row
    lngCurCol = GetPlanItemNameCol(vsfGrid.Col)
    
    varRecordType = Split(vsfGrid.Cell(flexcpData, lngCurRow, lngCurCol + 1) & "|||", "|")
    If Val(varRecordType(3)) = 1 Then
        PlanIsReplaceDoctor = True
    End If
End Function

Public Function PlanIsSelOne(vsfGrid As VSFlexGrid) As Boolean
    '�Ƿ�ֻѡ����һ��ʱ��
    Dim lngRowStart As Long, lngRowEnd As Long '��ʼ�к���ֹ��
    Dim lngColStart As Long, lngColEnd As Long '��ʼ�к���ֹ��
    
    With vsfGrid
        'ѡ���з�Χ
        lngRowStart = .Row: lngRowEnd = .RowSel
        
        'ѡ���з�Χ
        If .Col >= gPlanGrid_FixedCols And .ColSel < gPlanGrid_FixedCols Then
            '��ʼ��Ϊ�����У�ĩβ��Ϊ�������У���ֻѡ���������
            lngColStart = .ColSel
            lngColEnd = lngColStart
        End If
        
        If .Col < gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
            '��ʼ��Ϊ�������У�ĩβ��Ϊ�����У���ѡ�������а��ŷ�Χ
            lngColStart = GetPlanItemNameCol(.ColSel) 'ȷ��"ʱ���"��
            lngColEnd = lngColStart + 2
        End If
        
        If .Col >= gPlanGrid_FixedCols And .ColSel >= gPlanGrid_FixedCols Then
            lngColStart = GetPlanItemNameCol(.Col) 'ȷ��"ʱ���"��
            lngColEnd = GetPlanItemNameCol(.ColSel)
            If lngColStart > lngColEnd Then
                lngColStart = lngColStart + 2
            Else
                lngColEnd = lngColEnd + 2
            End If
        End If
    End With
    PlanIsSelOne = Not (Abs(lngRowEnd - lngRowStart) > 0 Or Abs(lngColEnd - lngColStart) > 2)
End Function

Public Function SelectedIsNotNull(ByVal vsfGrid As VSFlexGrid) As Boolean
    '�жϵ�ǰѡ��Ԫ���Ƿ��ǿ�ֵ
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Function
    End With
    SelectedIsNotNull = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Is��ֹԤԼ(ByVal vsfGrid As VSFlexGrid) As Boolean
    '�жϵ�ǰѡ�����Ƿ��ֹԤԼ
    On Error GoTo errHandler
    Is��ֹԤԼ = True
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        If Trim(.TextMatrix(.Row, GetPlanItemNameCol(.Col) + 2)) = "-" Then Exit Function
    End With
    Is��ֹԤԼ = False
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsVerified(ByVal vsfGrid As VSFlexGrid) As Boolean
    '�жϵ�ǰѡ�����Ƿ������
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        If Val(.TextMatrix(.Row, COL_�Ƿ����)) = 0 Then Exit Function
    End With
    IsVerified = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsTempPlan(ByVal vsfGrid As VSFlexGrid) As Boolean
    '�жϵ�ǰѡ�����Ƿ���ʱ����
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        If Val(.TextMatrix(.Row, COL_��ʱ����)) = 0 Then Exit Function
    End With
    IsTempPlan = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPlanItemImage(ByVal strKey As String) As IPictureDisp
    '��ȡ���ź�Դ����ͼ��
    '��Σ�
    '   strKey ͼ������:
    '       InvalidFixedItem '��Ч���̶��Ű��Դ
    '       FixedItem '�������̶��Ű��Դ
    '       InvalidMonthItem '��Ч�����Ű��Դ
    '       MonthItem '���������Ű��Դ
    '       InvalidWeekItem '��Ч�����Ű��Դ
    '       WeekItem '���������Ű��Դ
    Set GetPlanItemImage = frmClinicPlanTemp.GetPlanItemImage(strKey)
End Function

Public Function GetSortIcon(ByVal strKey As String) As IPictureDisp
    '��ȡ����ͼ��
    '��Σ�
    '   strKey ͼ������
    '       ASC '����
    '       DESC '����
    Set GetSortIcon = frmClinicPlanTemp.GetSortIcon(strKey)
End Function

Private Function GetLockImage() As IPictureDisp
    '��ȡ����ͼ��
    Set GetLockImage = frmClinicPlanTemp.GetLockPicture
End Function

Public Sub RegistPlan_KeyDown(vsfGrid As VSFlexGrid, KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    On Error Resume Next
    With vsfGrid
        Select Case KeyCode
        Case vbKeyRight '�����ƶ���
            If Shift = vbShiftMask Then
                If .ColSel + 3 >= gPlanGrid_FixedCols Then
                    .ColSel = .ColSel + 3
                    KeyCode = 0 '���μ�ֵ
                Else
                    
                End If
            Else
                If .Col < gPlanGrid_FixedCols Then
                    i = 1
                    Do While .Col + i < .Cols - 1
                        If .ColHidden(.Col + i) Or .ColWidth(.Col + i) = 0 Then
                            i = i + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    .Col = .Col + i
                    KeyCode = 0 '���μ�ֵ
                Else
                    .Col = .Col + 3
                End If
            End If
        Case vbKeyLeft '�����ƶ���
            If Shift = vbShiftMask Then
                If .ColSel - 3 >= gPlanGrid_FixedCols Then
                    .ColSel = .ColSel - 3
                    KeyCode = 0 '���μ�ֵ
                Else
                    
                End If
            End If
        End Select
    End With
End Sub

Public Function GetSelectRange(vsfGrid As VSFlexGrid, ByVal strSelRange As String, _
    ByRef lngRowStart As Long, ByRef lngRowEnd As Long, _
    ByRef lngColStart As Long, ByRef lngColEnd As Long) As Boolean
    '�ֽ��ѡ����������
    '��Σ�
    '   vsfGrid:����ؼ�
    '   strSelRange:��ʽ"��ʼ��|������|��ʼ��|������"
    '���Σ�
    '   lngRowStart ��ʼ��
    '   lngRowEnd ������
    '   lngColStart ��ʼ��
    '   lngColEnd ������
    Dim varTemp As Variant, lngTemp As Long
    
    Err = 0: On Error GoTo errHandler
    If InStr(strSelRange, "|") <= 0 Then Exit Function
    
    varTemp = Split(strSelRange & "|||", "|")
    lngRowStart = varTemp(0): lngRowEnd = varTemp(1)
    lngColStart = varTemp(2): lngColEnd = varTemp(3)
    With vsfGrid
        If lngRowStart < .FixedRows Or lngRowStart > .Rows - 1 Then Exit Function
        If lngRowEnd < .FixedRows Or lngRowEnd > .Rows - 1 Then Exit Function
        If lngColStart < .FixedRows Or lngColStart > .Cols - 1 Then Exit Function
        If lngColEnd < .FixedRows Or lngColEnd > .Cols - 1 Then Exit Function
    End With
    
    If lngRowStart > lngRowEnd Then lngTemp = lngRowStart: lngRowStart = lngRowEnd: lngRowEnd = lngTemp
    If lngColStart > lngColEnd Then lngTemp = lngColStart: lngColStart = lngColEnd: lngColEnd = lngTemp
    GetSelectRange = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function vsGrid_Para_Restore_Plan(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional blnǿ�ƻָ����� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�����ݿ��лָ�����Ŀ�ȵ���Ϣ
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '     blnSaveToDataBase-�Ƿ��������ݿ��б������(����������ݿ��б���,��ǿ�Ʊ���Ϊtrue,��������Ƿ�ʹ�ø��Ի������ȷ��)
    '     blnǿ�ƻָ�����-�����Ƿ񽫱���ע���Ĳ���ֵ,����ǿ�ƻָ�
    '����:�ָ��ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '˵����
    '       ����д�ù�������Ϊ�ٴ����ﰲ�ŵı��Ƚ����⣬ֻ�ָ�ĳЩ�У�ͬʱ����Ҳ�Ƕ�̬�仯��
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, arrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    If blnSaveToDataBase = False Then
        'ֻ���ڱ���ע����вŻᴦ����Ի�����
        vsGrid_Para_Restore_Plan = True
        If blnǿ�ƻָ����� = False Then
            If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g˽��ģ��, strCaption, strKey, strParaValue)
    Else
        strParaValue = zlDatabase.GetPara(strKey, glngSys, lngModule)
    End If
    
    vsGrid_Para_Restore_Plan = False
    If strParaValue = "" Then Exit Function
    'strParaValue:�����ʽ:������,�п�,������|������,�п�,������|...
    Err = 0: On Error GoTo Errhand:
    arrReg = Split(strParaValue, "|")
'    If vsGrid.Cols <> UBound(arrReg) + 1 Then Exit Function
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrTemp(0)
            If strColName <> "" Then
                intTemp = .ColIndex(strColName)
                If intTemp <> -1 Then
                    .ColWidth(intTemp) = Val(arrTemp(1))
                    If Val(arrTemp(2)) = 1 Then
                        .ColHidden(intTemp) = True
                    Else
                        .ColHidden(intTemp) = False
                    End If
                    If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                    .ColPosition(.ColIndex(strColName)) = intCol
                End If
            End If
        Next
    End With
    vsGrid_Para_Restore_Plan = True
    Exit Function
Errhand:
End Function

Public Function VSFlexGridCopyTo(ByVal vsfSource As VSFlexGrid, ByRef vsfNew As VSFlexGrid, _
    Optional ByVal bytMode As Byte) As Boolean
    '����: ��vsfSource�����ݸ��Ƶ�vsfNew�У�������ʾ��ʽ�����ڴ�ӡ\Ԥ��
    '����:
    '     vsfNew-���ƺ�Ķ���
    '     vsfSource-�����ƵĶ���
    '     bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '���أ����Ƴɹ�������True�����򣬷���False
    VSFlexGridCopyTo = frmClinicPlanTemp.VSFlexGridCopyTo(vsfSource, vsfNew, bytMode)
End Function
