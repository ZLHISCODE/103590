Attribute VB_Name = "mdlVsGridProperty"
Option Explicit
Public Const GRD_GOTFOCUS_COLORSEL = &H8000000D '16772055 '    '����ؼ�ʱ,ѡ����ʾ��ɫ
Public Const GRD_LOSTFOCUS_COLORSEL = &HE0E0E0  '&H80000010  '�뿪����ʱ,ѡ�����ʾ��ɫ
Public Enum mTextType
    m�ı�ʽ = 0
    m����ʽ = 1
    m���ʽ = 2
    m�����ʽ = 3
End Enum
Public Function GetVsGridBoolColVal(ByVal vsGrid As VSFlexGrid, lngRow As Long, lngCol As Long) As Boolean
    '------------------------------------------------------------------------------
    '����:��ȡbool�е�ֵ
    '����:�Ǹõ�Ԫ��Ϊtrue,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/28
    '------------------------------------------------------------------------------
    
    GetVsGridBoolColVal = grid.BoolVal(vsGrid, lngRow, lngCol)
    
End Function

Public Sub VsFlxGridCheckKeyPress(ByVal objCtl As Object, Row As Long, Col As Long, KeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '����:ֻ���������ֺͻس����˸�
    '����:
    '   objctl:Vsgrid8.0�ؼ�
    '   Keyascii:
    '           Keyascii:8 (�˸�)
    '   Row-��ǰ��
    '   Col-��ǰ��
    '   TextType:(0-�ı�ʽ;1-����ʽ;2-���ʽ)
    '����:һ��KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Call grid.CheckKeyPress(objCtl, Row, Col, KeyAscii, TextType)
    
End Sub


Public Function zl_VsGridAfterSort(ByVal vsGrid As VSFlexGrid, ByVal intCol As Integer, ByVal intOrder As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�������ô��¼�(��Ҫ�Ǵ����еı���ɫ)
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-28 11:26:52
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        .Redraw = flexRDNone
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = .BackColor
        .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = 16772055
        .Redraw = flexRDBuffered
    End With
    zl_VsGridAfterSort = True
End Function

Public Sub zlVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng���� As Long = -1, Optional lngβ�� As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef lngRow As Long = -1)
    '-----------------------------------------------------------------------------------------------------------
    '����:�ƶ���Ԫ�����
    '���:blnEdit-��ǰ�����ڱ༭״̬,����������
    '     lng����-����,���<0,������Ϊ0��,����Ϊָ������
    '     lngβ��-β��,���<0,������Ϊ.cols-1,����Ϊָ������
    '����:lngRow-������ڲ�����,�򷵻ر�������к�,���򷵻�-1
    '����:
    '����:���˺�
    '����:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------
    Dim lngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    
    'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
    If lng���� <> -1 Then
        lngCol = lng����
    Else
        lngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If lngCol = -1 Then lngCol = 0
    lngLastCol = IIf(lngβ�� < 0, vsGrid.Cols - 1, lngβ��)
    lngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = lngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        lngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                arrSplit = Split(.ColData(i) & "||", "||")
                If .ColHidden(i) Or Val(arrSplit(1)) >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = lngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    lngRow = .Row
                                End If
                            End If
                            .Col = lngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.Width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
Errhand:
End Sub
Public Function zlVsInsertIntoRow(ByVal vsGrid As VSFlexGrid, ByVal lngRow As Long, Optional blnBefor As Boolean = False, _
    Optional blnMoveNewRow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '����:������
    '����:vsGrid-�����е�������
    '     lngRow-��ǰ��
    '     blnBefor-��lngrow֮���֮��.true:֮��,false-֮��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intCol As Integer
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        If blnBefor Then
            .AddItem "", lngRow
            For intCol = 0 To .Cols - 1
                .Cell(flexcpBackColor, .Rows - 1, intCol, .Rows - 1, intCol) = .Cell(flexcpBackColor, 1, intCol, 1, intCol)
            Next
        Else
            .AddItem "", lngRow + 1
            For intCol = 0 To .Cols - 1
                .Cell(flexcpBackColor, .Rows - 1, intCol, .Rows - 1, intCol) = .Cell(flexcpBackColor, 1, intCol, 1, intCol)
            Next
        End If
        If blnMoveNewRow = True Then
            If blnBefor Then '
                .Row = lngRow
            Else
                .Row = lngRow + 1
            End If
        End If
    End With
    zlVsInsertIntoRow = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
'*********************************************************************************************************************
'**��������ؼ�
Public Sub zl_VsGridGotFocus(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���������ؼ�ʱѡ�����ɫ
    '��Σ�CustomColor-�Զ���ɫ
    '���ƣ����˺�
    '���ڣ�2010-03-23 10:52:23
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    '����ؼ�
    With vsGrid
         If CustomColor <> -1 Then
             .FocusRect = flexFocusSolid
             .HighLight = flexHighlightNever
             If .Row >= .FixedRows Then
                If .Rows - 1 > .FixedRows Then  '���ѡ����ɫ
                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
                End If
                 .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
             End If
         Else
            .FocusRect = flexFocusSolid 'IIf(vsGrid.Editable = flexEDNone, flexFocusNone, flexFocusSolid)
            .HighLight = flexHighlightNever
            .BackColorSel = GRD_GOTFOCUS_COLORSEL
        End If
    End With
    Call zl_VsGridRowChange(vsGrid, vsGrid.Row, vsGrid.Row, 0, 0)
End Sub
Public Sub zl_VsGridLOSTFOCUS(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1, Optional ForeColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
   '���ܣ��뿪����ؼ�ʱѡ�����ɫ
    '��Σ�CustomColor-�Ƿ����Զ�����ɫ������(BackColor)�ķ�ʽ������)
    '���ƣ����˺�
    '���ڣ�2010-03-23 11:03:05
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With vsGrid
        If CustomColor <> -1 Then
            If .Row >= .FixedRows Then
                .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
            End If
        Else
            .SelectionMode = flexSelectionByRow
            .FocusRect = IIf(vsGrid.Editable = flexEDNone, flexFocusHeavy, flexFocusSolid)
            If ForeColor = -1 Then .HighLight = flexHighlightAlways
            .BackColorSel = GRD_LOSTFOCUS_COLORSEL
        End If
        If ForeColor <> -1 Then
            .Cell(flexcpForeColor, .Row, .FixedCols, .Row, .Cols - 1) = ForeColor
        End If
        .ForeColorSel = .ForeColor
    End With
End Sub
Public Sub zl_VsGridRowChange(ByVal vsGrid As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngNewRow As Long, _
    ByVal lngoldCol As Long, ByVal lngNewCol As Long, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����иı�ʱ,������ص���ɫ
    '��Σ�CustomColor-�Զ�����ɫ
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-03-23 11:22:38
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    '�иı�ʱ
    Err = 0: On Error Resume Next
    If lngOldRow = lngNewRow Then
        vsGrid.Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, vsGrid.Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
        Exit Sub
    End If
    With vsGrid
        .Cell(flexcpBackColor, lngOldRow, vsGrid.FixedCols, lngOldRow, .Cols - 1) = .BackColor
        .Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, .Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
    End With
End Sub

'������
Public Sub zl_VsGridBeforeSort(ByVal vsGrid As VSFlexGrid, ByRef Col As Long, ByRef Order As Integer, Optional strSpaceRowNotCheckCol As String = "")
    '-----------------------------------------------------------------------------------------------------------
    '����:��������(����ʱ,�������հ���)
    '���:strSpaceRowNotCheckCol-���������е���Щ��(��1,��2...)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-07-25 11:38:23
    '-----------------------------------------------------------------------------------------------------------
    Dim lngStartRow As Long, lngEndRow As Long, lngStartCol As Long, lngEndCol As Long
    Dim lngRow As Long, lngCol As Long
    Dim blnAllowSelect As Boolean, blnAllowBigSel As Boolean
    Dim lngOldBackColor As Long

    If vsGrid.ExplorerBar > &H1000& Then Exit Sub
    '���浱ǰ��ѡ������
    vsGrid.GetSelection lngStartRow, lngStartCol, lngEndRow, lngEndCol
    vsGrid.Redraw = flexRDNone
    blnAllowBigSel = vsGrid.AllowBigSelection: blnAllowSelect = vsGrid.AllowSelection
    
    '������հ���
    With vsGrid
        For lngRow = .Rows - 1 To .FixedRows Step -1
            For lngCol = 0 To .Cols - 1
               If InStr(1, "," & strSpaceRowNotCheckCol & ",", "," & lngCol & ",") > 0 Then
               Else
                    If Trim(.TextMatrix(lngRow, lngCol)) <> "" Then GoTo GoNext:
               End If
            Next
        Next
GoNext:
        If lngRow > .FixedRows Then
            
             .Select .FixedRows, Col, lngRow, Col
            .Sort = Order
        End If
        ' �ָ���ǰѡ�������
        .Select lngStartRow, lngStartCol, lngEndRow, lngEndCol
            
        .Redraw = flexRDDirect
    End With
    Order = 0
End Sub


Public Function zl_vsGrid_Para_Save(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional blnǿ�Ʊ��� As Boolean = False, Optional blnHaveParaPrivs As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '����:����vsFlex�Ŀ�ȵ�ע���
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '����:����ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    If blnSaveToDataBase = False Then
        zl_vsGrid_Para_Save = True
        If blnǿ�Ʊ��� = False Then
            If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
        End If
    End If
    zl_vsGrid_Para_Save = False
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '�����ʽ:������,�п�,������|������,�п�,������|...
    If blnSaveToDataBase Then
        zlDatabase.SetPara strKey, strCol, glngSys, lngModule, blnHaveParaPrivs
    Else
        Call SaveRegInFor(g˽��ģ��, strCaption, strKey, strCol)
    End If
    zl_vsGrid_Para_Save = True
End Function

Public Function zl_vsGrid_Para_Restore(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String, _
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
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, arrtemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    If blnSaveToDataBase = False Then
        'ֻ���ڱ���ע����вŻᴦ����Ի�����
        zl_vsGrid_Para_Restore = True
        If blnǿ�ƻָ����� = False Then
            If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g˽��ģ��, strCaption, strKey, strParaValue)
    Else
        strParaValue = zlDatabase.GetPara(strKey, glngSys, lngModule)
    End If
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:�����ʽ:������,�п�,������|������,�п�,������|...
    Err = 0: On Error GoTo Errhand:
    arrReg = Split(strParaValue, "|")
    If vsGrid.Cols <> UBound(arrReg) + 1 Then Exit Function
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrtemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrtemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(arrtemp(1))
                If Val(arrtemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
Errhand:
End Function

Public Function zl_vsGrid_GetCols_Property(ByVal vsGrid As VSFlexGrid) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ͷ���
    '����:vsGrid-��Ӧ������ؼ�
    '����:������ͷ��Ϣ,��ʽΪ:������,�п�,������|������,�п�,������|....
    '����:���˺�
    '����:2014-10-09 12:08:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    zl_vsGrid_GetCols_Property = strCol
End Function

Public Sub zl_vsGrid_RestoreCols_Property(ByVal vsGrid As VSFlexGrid, ByVal strColsInfor As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������п�
    '����:vsGrid-��Ӧ������ؼ�
    '     strColsInfor-����Ϣ:������,�п�,������|������,�п�,������|....
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-09 12:34:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCols As Integer, arrReg As Variant, arrtemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    If strColsInfor = "" Then Exit Sub
    Err = 0: On Error GoTo Errhand:
    arrReg = Split(strColsInfor, "|")
    
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrtemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrtemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(arrtemp(1))
                If Val(arrtemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    Exit Sub
Errhand:
End Sub

