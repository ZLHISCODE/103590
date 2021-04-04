Attribute VB_Name = "mdlPubVSFlexGrid"
Option Explicit
Public Type VsfRowCol
    lngRow As Long
    lngCol As Long
End Type

Public Sub vfgSetting(ByVal LngStyle As Long, ByRef objVfg As VSFlexGrid, Optional ByVal strTtile As String, Optional VsfImg As ImageList, Optional ByVal lngFontSize As Long)
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
        
         '5.����
        If lngFontSize > 0 Then
            .FontSize = lngFontSize
        End If
        
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
          
          '����
1         On Error GoTo vfgLoadFromRecord_Error

2         For i = 0 To rsTmp.Fields.Count - 1
3             strTitle = strTitle & ";" & rsTmp.Fields(i).Name & ",0," & flexAlignLeftCenter
4         Next
5         If strTitle <> "" Then strTitle = Mid(strTitle, 2)
          
6         Call vfgSetting(0, objVfg, strTitle, objImgList)
          
          '��������
7         With objVfg
8             .Rows = .FixedRows + 1
9             .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
              'Set .DataSource = rsTmp ֱ��������Դ����ԭ�����õĸ�ʽ����ȸ�ʽ��ʧ�����ֹ��������
10            Do Until rsTmp.EOF
11                For i = 0 To rsTmp.Fields.Count - 1
12                    .TextMatrix(.Rows - 1, i) = CStr("" & rsTmp.Fields(i).value)
13                    If Not objImgList Is Nothing Then
14                        If CheckImgListKey(objImgList, rsTmp.Fields(i).Name) = True And CheckImgListKey(objImgList, rsTmp.Fields(i).value & "") = True Then
15                            .Row = .Rows - 1
16                            .Col = i
17                            .CellPicture = objImgList.ListImages(rsTmp.Fields(i).value).ExtractIcon
18                        End If
19                    End If
20                Next
21                .Rows = .Rows + 1
22                rsTmp.MoveNext
23            Loop
24            If .Rows > .FixedRows + 1 Then .Rows = .Rows - 1
25        End With
26        vfgLoadFromRecord = True
          
27        Exit Function
vfgLoadFromRecord_Error:
28        strErr = Err.Number & " " & Err.Description
29        Call WriteErrLog("zlPublicHisCommLis", "mdlPubVSFlexGrid", "ִ��(vfgLoadFromRecord)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
30        Err.Clear
End Function
Public Function CheckImgListKey(Vfgimg As ImageList, strKey As String) As Boolean
    '����           ���ؼ����Ƿ���ͼ���б��д��ڣ�������ڷ���Ϊ��
    '����
    '               Vfgimg �����ͼ�����
    '               strKey Ҫ��鵱ǰ�����Key�Ƿ����
    '����           �з����棬û�з��ؼ�
    Dim intloop As Integer
    On Error Resume Next
    If Vfgimg Is Nothing Then Exit Function
    With Vfgimg
        For intloop = 1 To .ListImages.Count
            If .ListImages(intloop).Key = strKey Then
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
Public Function VsfColAllSelAllcls(objVSF As VSFlexGrid, intCol As Integer, Optional intSel As Integer, Optional strErr As String) As Boolean
          '����               ȫѡ��ȫ��ѡ���
          '����               intSel 0=����һ�н����ж� 1=ȫ��ѡ�� 2=ȫ����ѡ��

          Dim intRow As Integer
          
1         On Error GoTo VsfColAllSelAllcls_Error

2         With objVSF
3             If intSel = 0 Then
4                 If .Rows = 1 Then Exit Function
5                 intSel = .Cell(flexcpChecked, 1, intCol, 1, intCol)
6                 If intSel = 1 Then
7                     intSel = 2
8                 Else
9                     intSel = 1
10                End If
11            End If
12            For intRow = 1 To .Rows - 1
13                .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = intSel
14            Next
15        End With
16        VsfColAllSelAllcls = True

17        Exit Function
VsfColAllSelAllcls_Error:
18        strErr = "������(vfgFindRowSel),������Ϣ:" & Err.Number & " " & Err.Description
19        Call WriteErrLog("zlPublicHisCommLis", "mdlPubVSFlexGrid", "ִ��(VsfColAllSelAllcls)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
20        Err.Clear
End Function

'����δ����Ŀؼ���������

Public Sub SplitWE(LeftControls As Collection, _
                   SplitControl As Control, _
                   RightControls As Collection, _
                   X As Single, Optional minWidth As Single)
    '���ҷָ�MouseDown�¼�������
    'LeftControls:��߿ؼ���
    'RightControls:�ұ߿ؼ���
    'splitcontrol :�ָ����ؼ�
    'X            :ƫ��������Splitcontrol��MouseMove�¼��д���
    'minWidth     :��С���
    Dim objControl As Control, sinWidth As Single
    On Error GoTo errH
    If LeftControls.Count < 0 Or RightControls.Count < 0 Then Exit Sub
    sinWidth = 3000
    If minWidth > 0 Then sinWidth = minWidth
    If LeftControls.Item(1).Width + X < sinWidth Or RightControls.Item(1).Width - X < sinWidth Then Exit Sub
    
    For Each objControl In LeftControls
        objControl.Width = objControl.Width + X
    Next
    
    SplitControl.Left = SplitControl.Left + X
            
    For Each objControl In RightControls
        objControl.Left = SplitControl.Left + SplitControl.Width
        objControl.Width = objControl.Width - X
    Next
    Exit Sub
errH:
    Exit Sub
End Sub

Public Sub SplitNS(TopControls As Collection, _
                   SplitControl As Control, _
                   ButtonControls As Collection, _
                   Y As Single, _
                   Optional minHight As Single)
    '���·ָ�MouseDown�¼�������
    'TopControls:�ϱ߿ؼ���
    'ButtonControls:�±߿ؼ���
    'splitcontrol :�ָ����ؼ�
    'Y            :ƫ��������Splitcontrol��MouseMove�¼��д���
    Dim objControl As Control, sigHight As Single
    On Error GoTo errH
    If TopControls.Count < 0 Or ButtonControls.Count < 0 Then Exit Sub
    sigHight = 3000
    If minHight > 0 Then sigHight = minHight
    If TopControls.Item(1).Height + Y < sigHight Or ButtonControls.Item(1).Height - Y < sigHight Then Exit Sub
    
    For Each objControl In TopControls
        objControl.Height = objControl.Height + Y
    Next
    
    SplitControl.Top = SplitControl.Top + Y
            
    For Each objControl In ButtonControls
        objControl.Top = SplitControl.Top + SplitControl.Height
        objControl.Height = objControl.Height - Y
    Next
    Exit Sub
errH:
    Exit Sub
End Sub

