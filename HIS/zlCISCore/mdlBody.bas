Attribute VB_Name = "mdlBodyNarcosis"
Option Explicit

Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

Public Sub DrawLine(pic As PictureBox, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1)
    '在(X1,Y1),(X2,Y2)之间使用ForeColor色画一直线
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    
    lngSaveForeColor = pic.ForeColor
    bytSaveLineWidth = pic.DrawWidth
    pic.ForeColor = ForeColor
    pic.DrawStyle = DrawStyle
    pic.DrawWidth = LineWidth
    pic.Line (X2, Y2)-(X1, Y1)
    pic.ForeColor = lngSaveForeColor
    pic.DrawWidth = bytSaveLineWidth
End Sub

Public Sub DrawText(pic As PictureBox, ByVal X As Single, ByVal Y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0)
    '在(X,Y)处输出Text文本
    Dim lngSaveForeColor As Long
    
    With pic
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        .CurrentX = X
        .CurrentY = Y
        pic.Print Text
        .ForeColor = lngSaveForeColor
    End With
End Sub

'=========================================
Public Sub ClearSpecRowCol(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
'功能: 清除指定网格的指定行指定列的数据
'参数: obj=要操作的网格控件
'      intRow=要清除的行号
'      intCol=要清除的列号列表如Array(1,2,3),若所有列则可以表示为Array()
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Public Sub SetColumnText(fgd As Object, intRow As Integer, ByVal varColText As Variant)
'功能: 设置指定网格控件的列头文本
'参数: fgd=网格控件
'      intRow=行号
'      varColText=列头文本数组
    Dim i As Integer
    For i = 0 To fgd.Cols - 1
        fgd.TextMatrix(intRow, i) = varColText(i)
    Next
End Sub

Public Sub SetColAlignment(fgd As Object, varColAlignment As Variant)
'功能: 设置指定网格控件的列对齐方式
'参数: fgd=网格控件
'      varColAlignment=列对齐方式数组
    Dim i As Long
    For i = 0 To UBound(varColAlignment)
        fgd.ColAlignment(i) = varColAlignment(i)
    Next
End Sub

Public Sub SetColData(fgd As Object, varColData As Variant)
'功能: 设置指定网格控件的列数据来源方式
'参数: fgd=网格控件
'      varColData=列数据来源方式数组
    Dim i As Long
    For i = 0 To UBound(varColData)
        fgd.ColData(i) = varColData(i)
    Next
End Sub

Public Sub SetFixColAlignment(fgd As Object, varFixColAlignment As Variant)
'功能: 设置指定网格控件的固定列对齐方式
'参数: fgd=网格控件
'      varColAlignment=固定列对齐方式数组
    Dim i As Long
    For i = 0 To UBound(varFixColAlignment)
        fgd.ColAlignmentFixed(i) = varFixColAlignment(i)
    Next
End Sub

Public Sub SetColumnWidth(fgd As Object, ByVal varColWidth As Variant)
'功能: 设置指定网格控件的列宽
'参数: fgd=网格控件
'      varColWidth=列宽数组
    Dim i As Integer
    For i = 0 To fgd.Cols - 1
        fgd.ColWidth(i) = varColWidth(i)
    Next
End Sub

Public Sub SetRowForeColor(mshObject As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    Dim i As Integer
    Dim blnPre As Boolean
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshObject
        blnPre = .Redraw
        intRow = .Row
        intCol = .Col
        .Redraw = False
        .Row = lngRow
        For i = 0 To .Cols - 1
            .Col = i
            .CellForeColor = lngColor
        Next
        
        .Row = intRow
        .Col = intCol
        .Redraw = blnPre
    End With
End Sub

Public Sub CalcXY(objFrm As Object, objMSH As Object, objX As Single, objY As Single, sglX As Single, sglY As Single)
    sglX = objFrm.Left + objX + objMSH.CellLeft + Screen.TwipsPerPixelX
    sglY = objFrm.Top + objFrm.Height - objFrm.ScaleHeight + objY + objMSH.CellTop + objMSH.CellHeight
    If sglX + 5895 > Screen.Width Then
        sglX = Screen.Width - 5895
    End If
    If sglY + 3420 > Screen.Height Then
        sglY = sglY - objMSH.CellHeight - 3420
    End If
End Sub

