Attribute VB_Name = "mEPRPicture"
Option Explicit

Private Type Size
    cx As Long
    cy As Long
End Type

Public glngPen As Long          '当前画笔对象
Public glngBrush As Long        '当前刷子对象

'以下变量取值与API参数对应
Public gcurPenColor As Long     '当前使用的线条色
Public gcurPenStyle As Byte     '当前使用的线型
Public gcurPenWidth As Byte     '当前使用的线宽
Public gcurFillColor As Long    '当前使用的填充色
Public gcurFillStyle As Integer '当前使用的填充样式

'################################################################################################################
'## 功能：  根据指定值设置当前的画笔的画刷
'##
'## 参数：  lngHDc          :   IN  ，编辑控件
'##         PenColor        :   IN  ，线条色
'##         PenStyle        :   IN  ，0-实线,1-划线,2-点线,3-点划线,4-双点划线
'##         PenWidth        :   IN  ，线宽
'##         FillColor       :   IN  ，填充色
'##         FillStyle       :   IN  ，-1-不填充,-2-实心,0-水平线,1-垂直线,2-左斜线,3-右斜线,4-水平和垂直线,5-交叉线
'################################################################################################################
Public Sub SetDrawStyleFromValue(lngHDc As Long, _
    PenColor As OLE_COLOR, _
    PenStyle As Byte, _
    PenWidth As Byte, _
    FillColor As OLE_COLOR, _
    FillStyle As Integer)
    
    Dim vBrush As LOGBRUSH
    Dim lngPen As Long, lngBrush As Long
    
    If glngBrush <> 0 Then DeleteObject glngBrush
    If glngPen <> 0 Then DeleteObject glngPen
    
    '画笔
    lngPen = CreatePen(PenStyle, IIf(PenWidth < 1, 1, PenWidth), PenColor)
    glngPen = SelectObject(lngHDc, lngPen)
    
    '画刷
    vBrush.lbColor = FillColor
    If FillStyle = -1 Then
        vBrush.lbStyle = BS_NULL
    ElseIf FillStyle = -2 Then
        vBrush.lbStyle = BS_SOLID
    Else
        vBrush.lbStyle = BS_HATCHED
        vBrush.lbHatch = FillStyle
    End If
    lngBrush = CreateBrushIndirect(vBrush)
    glngBrush = SelectObject(lngHDc, lngBrush)
End Sub

'################################################################################################################
'## 功能：  将标记缩放保存到另一个PicMarks对象中
'##
'## 参数：  picMarksSource  ：源
'##         picMarksDest    ：目标
'##         zoomFactor      ：缩放因子
'################################################################################################################
Public Function ScalePicMarks(picMarksSource As cTabPicMarks, ZoomFactor As Double) As cTabPicMarks
    Dim i As Long, j As Long, T As Variant, x As Long, y As Long
    Dim strTmp As String
    Dim picMarksDest As cTabPicMarks
    
    Set picMarksDest = picMarksSource.Clone
    For i = 1 To picMarksDest.Count
        picMarksDest(i).X1 = picMarksDest(i).X1 * ZoomFactor
        picMarksDest(i).Y1 = picMarksDest(i).Y1 * ZoomFactor
        picMarksDest(i).X2 = picMarksDest(i).X2 * ZoomFactor
        picMarksDest(i).Y2 = picMarksDest(i).Y2 * ZoomFactor
        strTmp = ""
        T = Split(picMarksDest(i).点集, ";")
        For j = 0 To UBound(T)
            x = CLng(Split(T(j), ",")(0))
            y = CLng(Split(T(j), ",")(1))
            x = x * ZoomFactor
            y = y * ZoomFactor
            If j = 0 Then
                strTmp = CStr(x) & "," & CStr(y)
            Else
                strTmp = strTmp & ";" & CStr(x) & "," & CStr(y)
            End If
        Next
        picMarksDest(i).点集 = strTmp
    Next
    Set ScalePicMarks = picMarksDest
End Function

'################################################################################################################
'## 功能：  显示标记图结果内容
'##
'## 参数：  objPic          :   IN  ，绘图介质，图片框控件
'##         objMarks        :   IN  ，图形标记集合
'################################################################################################################
Public Function ShowPicMark(objPic As PictureBox, objMark As cTabPicMark) As StdPicture
    Dim arrTmp() As String, arrXY() As POINTAPI
    Dim i As Integer, j As Integer, objFnt As New StdFont
    
    objPic.DrawMode = vbCopyPen
    objFnt.Name = IIf(objMark.字体 = "", "宋体", objMark.字体) '目前此参数传入未使用
            
    With objMark '具体标记元素
        Call SetDrawStyleFromValue(objPic.hdc, .线条色, .线型, .线宽, .填充色, .填充方式)
        Select Case .类型
            Case 0 '文本
                Call TextOut(objPic, .内容, .X1, .Y1, .X2, .Y2, objFnt)
            Case 1 '线条
                MoveToEx objPic.hdc, .X1, .Y1, 0
                LineTo objPic.hdc, .X2, .Y2
            Case 2 '折线
                arrTmp = Split(.点集, ";")
                For j = 0 To UBound(arrTmp)
                    ReDim Preserve arrXY(j)
                    arrXY(j).x = CLng(Split(arrTmp(j), ",")(0))
                    arrXY(j).y = CLng(Split(arrTmp(j), ",")(1))
                Next
                Polyline objPic.hdc, arrXY(0), UBound(arrXY) + 1
            Case 3 '矩形
                Rectangle objPic.hdc, .X1, .Y1, .X2, .Y2
            Case 4 '多边形
                arrTmp = Split(.点集, ";")
                For j = 0 To UBound(arrTmp)
                    ReDim Preserve arrXY(j)
                    arrXY(j).x = CLng(Split(arrTmp(j), ",")(0))
                    arrXY(j).y = CLng(Split(arrTmp(j), ",")(1))
                Next
                Polygon objPic.hdc, arrXY(0), UBound(arrXY) + 1
            Case 5 '圆
                Ellipse objPic.hdc, .X1, .Y1, .X2, .Y2
            Case 6 '序列编号
                If .填充色 = 0 Then
                    Call SetDrawStyleFromValue(objPic.hdc, RGB(255, 255, 0), 0, 1, RGB(255, 255, 0), -2)
                Else
                    Call SetDrawStyleFromValue(objPic.hdc, RGB(255, 255, 0), 0, 1, .填充色, -2)
                End If
                Ellipse objPic.hdc, .X1 - 7, .Y1 - 7, .X1 + 7, .Y1 + 7
                If .线条色 = 0 Then
                    Call SetDrawStyleFromValue(objPic.hdc, vbBlack, 0, 1, vbBlack, -1)
                Else
                    Call SetDrawStyleFromValue(objPic.hdc, .线条色, 0, 1, .线条色, -1)
                End If
                Ellipse objPic.hdc, .X1 - 7, .Y1 - 7, .X1 + 7, .Y1 + 7
                objFnt.Bold = True
                Call TextOut(objPic, .内容, IIf(Len(.内容) > 1, .X1 - 6, .X1 - 2), .Y1 - 6, .X1 + 14, .Y1 + 14, objFnt)
        End Select
    End With
    objPic.Refresh

    
'    Set ShowPicMark = objPic.Image
    
End Function

'################################################################################################################
'## 功能：  判断矩形与椭圆相交与否
'##
'## 参数：  (X1,Y1),(X2,Y2) :矩形左上角与右下角点坐标
'##         (X3,Y3),(X4,Y4) :椭圆左上角与右下角点坐标
'################################################################################################################
Public Function 矩形与椭圆相交(X1 As Long, Y1 As Long, _
    X2 As Long, Y2 As Long, _
    X3 As Long, Y3 As Long, _
    X4 As Long, Y4 As Long) As Boolean
    
    Dim MyRgn As Long, OutRgn As Long, InRgn As Long, R As Long
    MyRgn = CreateRectRgn(0, 0, 0, 0) '矩形
    OutRgn = CreateRectRgn(X1, Y1, X2, Y2)       '椭圆
    InRgn = CreateEllipticRgn(X3, Y3, X4, Y4)
    R = CombineRgn(MyRgn, OutRgn, InRgn, RGN_AND)

    If R = NULLREGION Or R = 0 Then  '0：失败！NULLREGION：无交点
'        If (X3 > X1 And X3 < X2 And Y3 > Y1 And Y3 < Y2) Or (X4 > X1 And X4 < X2 And Y4 > Y1 And Y4 < Y2) Then
'            矩形与椭圆相交 = True
'        Else
            矩形与椭圆相交 = False
'        End If
    Else
        矩形与椭圆相交 = True
    End If
End Function

'################################################################################################################
'## 功能：  判断矩形与矩形相交与否
'##
'## 参数：  (X1,Y1),(X2,Y2) :矩形1左上角与右下角点坐标
'##         (X3,Y3),(X4,Y4) :矩形2左上角与右下角点坐标
'################################################################################################################
Public Function 矩形与矩形相交(X1 As Long, Y1 As Long, _
    X2 As Long, Y2 As Long, _
    X3 As Long, Y3 As Long, _
    X4 As Long, Y4 As Long) As Boolean
    
    Dim MyRgn As Long, OutRgn As Long, InRgn As Long, R As Long
    MyRgn = CreateRectRgn(0, 0, 0, 0) '矩形
    OutRgn = CreateRectRgn(X1, Y1, X2, Y2)       '椭圆
    InRgn = CreateRectRgn(X3, Y3, X4, Y4)
    R = CombineRgn(MyRgn, OutRgn, InRgn, RGN_AND)

    If R = NULLREGION Or R = 0 Then  '0：失败！NULLREGION：无交点
        矩形与矩形相交 = False
    Else
        矩形与矩形相交 = True
    End If
End Function

'################################################################################################################
'## 功能：  判断矩形与多边形相交与否
'##
'## 参数：  (X1,Y1),(X2,Y2) :矩形左上角与右下角点坐标
'##         Points()        :多边形顶点坐标集合
'################################################################################################################
Public Function 矩形与多边形相交(X1 As Long, Y1 As Long, _
    X2 As Long, Y2 As Long, _
    Points() As POINTAPI) As Boolean
    
    Dim MyRgn As Long, OutRgn As Long, InRgn As Long, R As Long
    MyRgn = CreateRectRgn(0, 0, 0, 0)           '矩形
    OutRgn = CreateRectRgn(X1, Y1, X2, Y2)      '椭圆
    InRgn = CreatePolygonRgn(Points(0), UBound(Points), WINDING) '根据多边形顶点数据创建多边形
    R = CombineRgn(MyRgn, OutRgn, InRgn, RGN_AND)

    If R = NULLREGION Or R = 0 Then  '0：失败！NULLREGION：无交点
        矩形与多边形相交 = False
    Else
        矩形与多边形相交 = True
    End If
End Function

'################################################################################################################
'## 功能：  在指定设备的指定范围内输出文字
'##
'## 参数：  objOut          :绘图对象--图片框控件
'##         strOut          :文本内容
'##         (X1,Y1),(X2,Y2) :矩形区域
'##         sFont           :字体对象
'################################################################################################################
Public Sub TextOut(objOut As Object, _
    ByVal strOut As String, _
    ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, _
    ByRef sFont As StdFont)
    
    Dim R As RECT
    
    If Trim(Replace(strOut, vbCrLf, "")) = "" Then Exit Sub
    
    R.Left = X1: R.Right = X2
    R.Top = Y1: R.Bottom = Y2
    
    DrawTextEx objOut.hdc, strOut, LenB(StrConv(strOut, vbFromUnicode)), R, DT_EDITCONTROL Or DT_WORDBREAK, 0&
    
    objOut.Refresh
End Sub

'################################################################################################################
'## 功能：  将指定矩形坐标强行调整成正方形
'##
'## 参数：  (X1,Y1)     :IN         原始矩形左上角坐标
'##         (X2,Y2)     :IN/OUT     新的矩形右下角坐标
'################################################################################################################
Public Sub ForceSquare(ByVal X1 As Long, ByVal Y1 As Long, X2 As Long, Y2 As Long)
    If Abs(Y2 - Y1) > Abs(X2 - X1) Then
        If X2 < X1 Then
            X2 = X1 - Abs(Y2 - Y1)
        Else
            X2 = X1 + Abs(Y2 - Y1)
        End If
        Y2 = Y2
    Else
        If Y2 < Y1 Then
            Y2 = Y1 - Abs(X2 - X1)
        Else
            Y2 = Y1 + Abs(X2 - X1)
        End If
        X2 = X2
    End If
End Sub




