Attribute VB_Name = "mdlVsFlexGrid"
Option Explicit

Public Enum TVsStyle
    vsZL���� = 0
End Enum

Public Type POINTAPI
        X As Long
        Y As Long
End Type

'--��������ģ��ϲ���Ԫ��ı���ߴ���
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const ETO_OPAQUE = 2
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Public Sub vfgDrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, _
                       ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByRef Done As Boolean, _
                       ByVal LeftCol As Long, ByVal RightCol As Long, ByVal topRow As Long, ByVal BottomRow As Long, _
                       ByRef vfg As VSFlexGrid)
                       
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    
    Dim vRect As RECT
    
    With vfg
        
        If Not Between(Col, LeftCol, RightCol) Then Exit Sub
        If Not Between(Row, topRow, BottomRow) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = topRow Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = BottomRow Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Row = .RowSel Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            If .Cell(flexcpBackColor, Row, Col) <> 0 Then
                SetBkColor hDC, SysColor2RGB(.Cell(flexcpBackColor, Row, Col))
            Else
                SetBkColor hDC, SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, 0, 0, 0
        Done = True
    End With

End Sub

Public Sub vfgDrawProgress(ByRef vfg As VSFlexGrid, ByVal Row As Long, ByVal Col As Long, ByVal hDC As Long, _
                           ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
                           ByVal Percent As Long)
    '���ƽ�����
    
    Dim clr As Long, clr1 As Long, strWord As String, lngLeft As Long, lngRight As Long, i As Integer
    
    
    If Percent >= 0 And Percent < 70 Then
        clr = SetBkColor(hDC, vbBlue) '
    ElseIf Percent >= 70 And Percent < 90 Then
        clr = SetBkColor(hDC, &H80FF&) '��ɫ
    ElseIf Percent >= 90 And Percent < 100 Then
        clr = SetBkColor(hDC, vbRed)
    Else
        clr = SetBkColor(hDC, &H6A00&) '����
    End If
    Dim rc As RECT
    
    SetRect rc, Left, Top, Right, Bottom
    
    rc.Top = Top + (Bottom - Top) / 2 - 5
    lngLeft = Left + 3
    lngRight = Left + (Right - Left) * (Percent / 100) - 3
    rc.Bottom = Top + (Bottom - Top) / 2 + 5
    
    rc.Left = lngLeft: rc.Right = lngRight
    If rc.Left >= Left And rc.Right > Left Then
        ExtTextOut hDC, rc.Left, rc.Top, ETO_OPAQUE, rc, 0, 0, 0
    End If
    With vfg
        If Row = .RowSel Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            If .Cell(flexcpBackColor, Row, Col) <> 0 Then
                SetBkColor hDC, SysColor2RGB(.Cell(flexcpBackColor, Row, Col))
            Else
                SetBkColor hDC, SysColor2RGB(.BackColor)
            End If
        End If
    End With
    'strWord = vfg.TextMatrix(Row, Col)
    For i = 2 To 50
        rc.Left = rc.Left + (lngRight - lngLeft) / 50 + 0.5
        rc.Right = rc.Left + (lngRight - lngLeft) / 50 + 0.5
        If Mid(Format(i / 2, "0.00"), InStr(Format(i / 2, "0.00"), ".") + 1) <= 0 Then
            If rc.Right <= lngRight Then
                ExtTextOut hDC, rc.Left, rc.Top, ETO_OPAQUE, rc, 0, 0, ByVal 0&
            End If
        Else
        
        End If
    Next

    SetBkColor hDC, clr
    
End Sub

Public Sub SetVsFlexGridHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid)
    '���ܣ���ʼvsFlexGrid
    '           ��һ�̶��У���ʼ����ֻ��һ�м�¼���޹̶��С�
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
    'vsGrid:    Ҫ��ʼ���Ŀؼ�

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
         
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ
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
        .RowHeight(0) = 500
        
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
        .Redraw = True
    End With
End Sub

Public Sub SetStyle(ByVal STYLE As TVsStyle, ByRef vsGrid As VSFlexGrid)
    '�������
    With vsGrid
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserFreezing = flexFreezeColumns
        .AllowUserResizing = flexResizeColumns
        .AutoResize = False
        .AutoSizeMode = flexAutoSizeRowHeight
        .BackColorBkg = vbWindowBackground
        .BackColorSel = &HFFEBD7
        .Editable = flexEDKbdMouse
        .ForeColorSel = vbWindowText
        .GridColor = vbApplicationWorkspace
        .GridColorFixed = vbApplicationWorkspace
        .GridLinesFixed = flexGridFlat
        .OwnerDraw = flexODOver
        .RowHeightMax = 2000
        .RowHeightMin = 250
        .ScrollTrack = True
        .SelectionMode = flexSelectionByRow
        .SheetBorder = &H0&
        .WordWrap = True
    End With

End Sub

'Public Function ShowPubSelect(ByVal frmParent As Object, _
'                                ByVal obj As Object, _
'                                ByVal bytStyle As Byte, _
'                                ByVal strLvw As String, _
'                                ByVal strSavePath As String, _
'                                ByVal strDescrible As String, _
'                                ByVal rsData As ADODB.Recordset, _
'                                ByRef rsResult As ADODB.Recordset, _
'                                Optional ByVal lngCX As Long = 9000, _
'                                Optional ByVal lngCY As Long = 4500, _
'                                Optional ByVal blnMuliSel As Boolean = False, _
'                                Optional ByVal strInitKey As String = "") As Byte
'    '******************************************************************************************************************
'    '���ܣ�������+�б�ṹ,Ӧ���ڱ��ؼ�
'    '������
'    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
'    '���أ�0:ȡ��ѡ��;1:ѡ��;2:�����ݷ���
'    '******************************************************************************************************************
'
'    Dim lngX As Long
'    Dim lngY As Long
'    Dim lngObjHeight As Long
'    Dim rs As New ADODB.Recordset
'    Dim objPoint As POINTAPI
'
'    On Error GoTo errHand
'
'    If rsData.BOF Then
'        ShowPubSelect = 2
'        Exit Function
'    End If
'
'    Call ClientToScreen(obj.hwnd, objPoint)
'
'    Select Case TypeName(obj)
'    Case "TextBox", "CommandButton"
'        lngX = objPoint.x * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
'        lngY = obj.Height + objPoint.y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
'        lngObjHeight = obj.Height
'
'    Case Else
'        lngX = objPoint.x * Screen.TwipsPerPixelX + obj.CellLeft
'        lngY = objPoint.y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
'        lngObjHeight = obj.CellHeight
'    End Select
'
'    ShowPubSelect = frmPubSelDialog.ShowDialog(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, , False, blnMuliSel)
'
'    If ShowPubSelect = 1 Then Set rsResult = rsData
'
'    Exit Function
'
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'End Function

Private Function Between(X, a, b) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

Private Function SysColor2RGB(ByVal lngColor As Long) As Long
'���ܣ���VB��ϵͳ��ɫת��ΪRGBɫ
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function
