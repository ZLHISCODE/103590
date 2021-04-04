Attribute VB_Name = "mdlVsFlexGrid"
Option Explicit

Public Enum TVsStyle
    vsZL常用 = 0
End Enum

Public Type POINTAPI
        X As Long
        Y As Long
End Type

'--以下用于模拟合并单元格的表格线处理
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
                       
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    
    Dim vRect As RECT
    
    With vfg
        
        If Not Between(Col, LeftCol, RightCol) Then Exit Sub
        If Not Between(Row, topRow, BottomRow) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = topRow Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = BottomRow Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
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
    '绘制进度条
    
    Dim clr As Long, clr1 As Long, strWord As String, lngLeft As Long, lngRight As Long, i As Integer
    
    
    If Percent >= 0 And Percent < 70 Then
        clr = SetBkColor(hDC, vbBlue) '
    ElseIf Percent >= 70 And Percent < 90 Then
        clr = SetBkColor(hDC, &H80FF&) '橙色
    ElseIf Percent >= 90 And Percent < 100 Then
        clr = SetBkColor(hDC, vbRed)
    Else
        clr = SetBkColor(hDC, &H6A00&) '深绿
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
    '功能：初始vsFlexGrid
    '           有一固定行，初始化后，只有一行记录，无固定列。
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'vsGrid:    要初始化的控件

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
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
        Next
        
        '固定行文字居中
        If .FixedRows > 0 And .Cols > 0 Then
            .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        .RowHeight(0) = 500
        
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
    End With
End Sub

Public Sub SetStyle(ByVal STYLE As TVsStyle, ByRef vsGrid As VSFlexGrid)
    '公共外观
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
'    '功能：打开树型+列表结构,应用于表格控件
'    '参数：
'    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
'    '返回：0:取消选择;1:选择;2:无数据返回
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
'功能：判断x是否在a和b之间
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

Private Function SysColor2RGB(ByVal lngColor As Long) As Long
'功能：将VB的系统颜色转换为RGB色
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function
