Attribute VB_Name = "mdlPubVSFlexGrid"
Option Explicit
Public Type VsfRowCol
    lngRow As Long
    lngCol As Long
End Type

Public Sub vfgSetting(ByVal LngStyle As Long, ByRef objVfg As VSFlexGrid, Optional ByVal strTtile As String, Optional VsfImg As ImageList)
    'lngStyle��0 Ĭ�����ã�ͳһVfg�������
    'strHead��  �����ʽ��
    '           ����1,���,���뷽ʽ;����2,���,���뷽ʽ;.......
    '           ���뷽ʽȡֵ, * ��ʾ����ȡֵ
    '           FlexAlignLeftTop       0   ����
    '           flexAlignLeftCenter    1   ����  *
    '           flexAlignLeftBottom    2   ����
    '           flexAlignCenterTop     3   ����
    '           flexAlignCenterCenter  4   ����  *
    '           flexAlignCenterBottom  5   ����
    '           flexAlignRightTop      6   ����
    '           flexAlignRightCenter   7   ����  *
    '           flexAlignRightBottom   8   ����
    '           flexAlignGeneral       9   ����
    'objVfg:    Ҫ��ʼ���Ŀؼ�
    'VsfImg:    ImageListͼ�꼯�ؼ�����

    Dim arrHead As Variant, i As Long, strHead As String
    If strTtile = "" Then
        strHead = "��1��,900,1;��2��,900,1;��3��,900,1"
    Else
        strHead = strTtile
    End If
    arrHead = Split(strHead, ";")
    
    
    With objVfg
        '1.�߿�
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .GridLines = flexGridFlat
        .GridColorFixed = flexGridFlat
        
        '2.��ɫ
        .BackColor = vbWindowBackground '���ڱ���
        .BackColorAlternate = vbWindowBackground
        .BackColorBkg = vbWindowBackground
        .BackColorFixed = vbButtonFace '��ť����
        .BackColorFrozen = &H0&         '��
        .FloodColor = &HC0&             '��
        .BackColorSel = &HFFEBD7        'ǳ��
        .ForeColor = vbWindowText       '�����ı�
        .ForeColorFixed = vbButtonText  '��ť�ı�
        .ForeColorFrozen = &H0&         '��
        .ForeColorSel = vbWindowText
        
        .GridColor = vbApplicationWorkspace 'Ӧ�ó�������
        .GridColorFixed = vbApplicationWorkspace
        .SheetBorder = vbWindowBackground
        .TreeColor = vbButtonShadow         '��ť��Ӱ
        
        '3.��ʼ������

        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ
            If CheckImgListKey(VsfImg, .TextMatrix(.FixedRows - 1, .FixedCols + i)) = True Then
                .Row = .FixedRows - 1
                .Col = .FixedCols + i
                .CellPicture = VsfImg.ListImages(Split(arrHead(i), ",")(0)).ExtractIcon
                '��ͼ��ʱ����ʾ����
                .TextMatrix(.FixedRows - 1, .FixedCols + i) = ""
            End If
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
            End If
        Next
        
        '�̶������־���
        If .FixedRows > 0 And .Cols > 0 Then
            .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        .RowHeight(0) = 300
        .RowHeightMin = 300
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
        .Redraw = True
        
        
        '4.��������
        .SelectionMode = flexSelectionByRow     '����ѡ��
        .ExplorerBar = flexExNone               '�����������Ӧ�������ƶ��У�����
        .AllowUserResizing = flexResizeColumns  '�ɵ����п�
        .Editable = flexEDNone                  'ֻ��
        
    End With
    
End Sub

Public Function vfgLoadFromRecord(ByRef objVfg As VSFlexGrid, _
                                  ByRef rsTmp As ADODB.Recordset, _
                                  ByRef strErr As String, _
                                  Optional objImgList As ImageList) As Boolean
    '����¼������װ��vfg�ؼ�
    'objVfg : vfg�ؼ�
    'rsTmp  : װ��ؼ��ļ�¼��
    'strErr :��ʾ��Ϣ
    Dim i As Integer, strTitle As String
    On Error GoTo errH
    
    '����
    For i = 0 To rsTmp.Fields.Count - 1
        strTitle = strTitle & ";" & rsTmp.Fields(i).Name & ",0," & flexAlignLeftCenter
    Next
    If strTitle <> "" Then strTitle = Mid(strTitle, 2)
    
    Call vfgSetting(0, objVfg, strTitle, objImgList)
    
    '��������
    With objVfg
        .Rows = .FixedRows + 1
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
        'Set .DataSource = rsTmp ֱ��������Դ����ԭ�����õĸ�ʽ����ȸ�ʽ��ʧ�����ֹ��������
        Do Until rsTmp.EOF
            For i = 0 To rsTmp.Fields.Count - 1
                .TextMatrix(.Rows - 1, i) = CStr("" & rsTmp.Fields(i).Value)
                If Not objImgList Is Nothing Then
                    If CheckImgListKey(objImgList, rsTmp.Fields(i).Name) = True And CheckImgListKey(objImgList, rsTmp.Fields(i).Value & "") = True Then
                        .Row = .Rows - 1
                        .Col = i
                        .CellPicture = objImgList.ListImages(rsTmp.Fields(i).Value).ExtractIcon
                    End If
                End If
            Next
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If .Rows > .FixedRows + 1 Then .Rows = .Rows - 1
    End With
    vfgLoadFromRecord = True
    Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
End Function
Public Function CheckImgListKey(Vfgimg As ImageList, strKey As String) As Boolean
    '����           ���ؼ����Ƿ���ͼ���б��д��ڣ�������ڷ���Ϊ��
    '����
    '               Vfgimg �����ͼ�����
    '               strKey Ҫ��鵱ǰ�����Key�Ƿ����
    '����           �з����棬û�з��ؼ�
    Dim intLoop As Integer
    On Error Resume Next
    If Vfgimg Is Nothing Then Exit Function
    With Vfgimg
        For intLoop = 1 To .ListImages.Count
            If .ListImages(intLoop).Key = strKey Then
                CheckImgListKey = True
                Exit Function
            End If
        Next
    End With
End Function

Public Function vfgFindRowSel(ByRef objVfg As VSFlexGrid, strField As String, FindstrValue, Optional strErr As String) As Long
    '����       ����ָ���ֶκͲ��ҵ�ֵƥ�䣬���ҵ���ѡ��
    '����
    '           objVfg      VSF����
    '           strField    �ֶ�
    '           FindstrValue    ���ҵ�ֵ
    Dim lngLoop As Long
    On Error Resume Next
    vfgFindRowSel = -1
    With objVfg
        For lngLoop = 1 To .Rows - 1
            If .TextMatrix(lngLoop, .ColIndex(strField)) = FindstrValue Then
                .Row = lngLoop
                vfgFindRowSel = lngLoop
                Exit Function
            End If
        Next
    End With
    Exit Function
errH:
    strErr = "������(vfgFindRowSel),������Ϣ:" & Err.Number & " " & Err.Description
End Function
Public Function vfgFindRowSelA(ByRef objVfg As VSFlexGrid, strField As String, FindstrValue, Optional strErr As String) As Long
    '����       ����ָ���ֶκͲ��ҵ�ֵƥ�䣬���ҵ���ѡ��
    '����
    '           objVfg      VSF����
    '           strField    �ֶ�
    '           FindstrValue    ���ҵ�ֵ
    Dim lngLoop As Long
    On Error Resume Next
    vfgFindRowSelA = -1
    With objVfg
        For lngLoop = 1 To .Rows - 1
            If .TextMatrix(lngLoop, .ColIndex(strField)) = FindstrValue Then
'                .Row = lngLoop
                vfgFindRowSelA = lngLoop
                Exit Function
            End If
        Next
    End With
    Exit Function
errH:
    strErr = "������(vfgFindRowSel),������Ϣ:" & Err.Number & " " & Err.Description
End Function
Public Function vfgFindRowCheck(ByRef objVfg As VSFlexGrid, strField As String, FindstrValue As String, Optional lngRow As Long, Optional lngCol As Long) As Boolean
    '����       ����Ƿ��и�����ֵ
    '����
    '           objVfg      VSF����
    '           strField    �ֶ�
    '           FindstrValue    ���ҵ�ֵ
    '����       ������һ����ֵΪ�� ����Ϊ��
    Dim lngLoop As Long
    On Error Resume Next
    With objVfg
        For lngLoop = 1 To .Rows - 1
            If .TextMatrix(lngLoop, .ColIndex(strField)) = FindstrValue Then
                If lngLoop = lngRow And .ColIndex("strField") = lngCol Then
                Else
                    vfgFindRowCheck = True
                End If
                Exit Function
            End If
        Next
    End With
End Function
Public Function VsfColAllSelAllcls(objvsf As VSFlexGrid, intCol As Integer, Optional intSel As Integer, Optional strErr As String) As Boolean
    '����               ȫѡ��ȫ��ѡ���
    '����               intSel 0=����һ�н����ж� 1=ȫ��ѡ�� 2=ȫ����ѡ��
    On Error GoTo errH
    Dim intRow As Integer
    
    With objvsf
        If intSel = 0 Then
            If .Rows = 1 Then Exit Function
            intSel = .Cell(flexcpChecked, 1, intCol, 1, intCol)
            If intSel = 1 Then
                intSel = 2
            Else
                intSel = 1
            End If
        End If
        For intRow = 1 To .Rows - 1
            .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = intSel
        Next
    End With
    VsfColAllSelAllcls = True
    Exit Function
errH:
    strErr = "������(vfgFindRowSel),������Ϣ:" & Err.Number & " " & Err.Description
End Function


