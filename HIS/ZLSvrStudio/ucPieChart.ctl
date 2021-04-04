VERSION 5.00
Begin VB.UserControl ucPieChart 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ucPieChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mtpItem() As tpItem
Private mdblTotal As Double
Private mtpColor(5) As tpColorRGB
Private mstType As Show_Type '0-饼图（默认），1-折线图，2-柱状图
Private msymType As Symbol_Type '0-什么都不显示(默认)，1-显示数量，2-显示百分比
Private moleLineColor As OLE_COLOR, moleItemColor As OLE_COLOR, moleTitleColor As OLE_COLOR
Private mstdItemFont As New StdFont, mstdTitleFont As New StdFont
Private mstrTitle As String '标题
Private mblnIsShow As Boolean '判断是否已经将图片展示出来
Private mblnLegend As Boolean '判断是否显示图例
Private mlngLevColor As Long   '颜色平滑梯度

Public Enum Show_Type
    饼图 = 0
    折线图 = 1
    柱状图 = 2
End Enum

Public Enum Symbol_Type
    不显示 = 0
    显示数量 = 1
    显示百分比 = 2
End Enum

Private Type tpItem
    PartNumber As Long
    Color As OLE_COLOR
    Title As String
End Type

Private Type tpColorRGB
    coR As Byte
    coG As Byte
    coB As Byte
    coLevel As Long
End Type

Private Const mInstructStart = 0.66 '第一段指示线起点到圆心距离与r的比例
Private Const mInstructBaseY = 0.5 '第一段指示线y轴偏移量基准值
Private Const mInstructBaseX = 0.5 '第一段指示线x轴偏移量基准值

'设置颜色显示梯度
Public Property Let LevColor(ByVal lngLevColor As Long)
    If lngLevColor = 0 Then lngLevColor = 30
    mlngLevColor = lngLevColor
End Property

Public Property Get LevColor() As Long
    LevColor = mlngLevColor
End Property

'设置显示类型
Public Property Let ShowType(ByVal stType As Show_Type)
    mstType = stType
    UserControl_Resize
End Property

Public Property Get ShowType() As Show_Type
    ShowType = mstType
End Property

'设置指示线颜色
Public Property Let LineColor(ByVal oleLineColor As OLE_COLOR)
    moleLineColor = oleLineColor
    UserControl_Resize
End Property

Public Property Get LineColor() As OLE_COLOR
    LineColor = moleLineColor
End Property

'设置标题
Public Property Let Title(ByVal strTitle As String)
    mstrTitle = strTitle
    UserControl_Resize
End Property

Public Property Get Title() As String
    Title = mstrTitle
End Property

'设置标题字体
Public Property Set TitleFont(ByVal stdTitleFont As StdFont)
    Set mstdTitleFont = stdTitleFont
    UserControl_Resize
End Property

Public Property Get TitleFont() As StdFont
    Set TitleFont = mstdTitleFont
End Property

'设置标题字体颜色
Public Property Let TitleColor(ByVal oleTitleColor As OLE_COLOR)
    moleTitleColor = oleTitleColor
    UserControl_Resize
End Property

Public Property Get TitleColor() As OLE_COLOR
    TitleColor = moleTitleColor
End Property

'设置项目字体
Public Property Set ItemFont(ByVal stdItemFont As StdFont)
    Set mstdItemFont = stdItemFont
    UserControl_Resize
End Property

Public Property Get ItemFont() As StdFont
    Set ItemFont = mstdItemFont
End Property

'设置项目字体颜色
Public Property Let ItemColor(ByVal oleItemColor As OLE_COLOR)
    moleItemColor = oleItemColor
    UserControl_Resize
End Property

Public Property Get ItemColor() As OLE_COLOR
    ItemColor = moleItemColor
End Property

'设置每个项目比例显示方式：1-什么都不显示，2-显示数量，3-显示百分比
Public Property Let SymbolType(ByVal symType As Symbol_Type)
    msymType = symType
    UserControl_Resize
End Property

Public Property Get SymbolType() As Symbol_Type
    SymbolType = msymType
End Property

'设置是否显示图例
Public Property Let Legend(ByVal blnLegend As Boolean)
    mblnLegend = blnLegend
    UserControl_Resize
End Property

Public Property Get Legend() As Boolean
    Legend = mblnLegend
End Property

Public Sub addItem(Optional ByVal strItemTitle As String, Optional ByVal oleItemColor As OLE_COLOR, Optional ByVal lngItemNumber As Long)
    Dim lngColor As Long
    Dim lngLevel As Long
    
    If lngItemNumber = 0 Then Exit Sub
    
    If strItemTitle = "" Then
        strItemTitle = "项目" & UBound(mtpItem) + 1
    End If
    If oleItemColor = 0 Then
        lngColor = UBound(mtpItem) Mod (UBound(mtpColor) + 1)
        oleItemColor = RGB(mtpColor(lngColor).coR + mtpColor(lngColor).coLevel * mlngLevColor, mtpColor(lngColor).coG + mtpColor(lngColor).coLevel * mlngLevColor, mtpColor(lngColor).coB + mtpColor(lngColor).coLevel * mlngLevColor)
        mtpColor(lngColor).coLevel = mtpColor(lngColor).coLevel + 1
    End If
    
    ReDim Preserve mtpItem(UBound(mtpItem) + 1)
    mtpItem(UBound(mtpItem)).Title = strItemTitle
    mtpItem(UBound(mtpItem)).Color = oleItemColor
    mtpItem(UBound(mtpItem)).PartNumber = lngItemNumber
    mdblTotal = mdblTotal + lngItemNumber
End Sub

Public Sub Clear()
    Dim i As Long

    UserControl.Cls
    ReDim mtpItem(0)
    mdblTotal = 0
    mblnIsShow = False
    For i = 0 To UBound(mtpColor)
        mtpColor(i).coLevel = 0
    Next
End Sub

Public Sub PaintChart(Optional bolType As Boolean = True)
    'bolType:判断是内部调用还是外部调用
    '清除原有的图像
    UserControl.Cls
    
    '判断显示方式
    If mstType = 饼图 Then
        Call showCircle
    End If
    mblnIsShow = True
End Sub

'以折线图形式显示
Private Sub showPolygon()

End Sub

'以饼图形式显示
Private Sub showCircle()
    Dim i As Long, K As Long
    Dim dblPi As Double 'π
    Dim R As Double  '半径
    Dim x As Double  '圆心x坐标
    Dim y As Double  '圆心y坐标
    Dim x0 As Double, y0 As Double '指示线起点坐标相对圆心增减量
    Dim x1 As Double, y1 As Double '指示线起点坐标
    Dim dblRadianLine As Double '指示线起点弧度
    Dim dblAccumulate As Double
    Dim strTitle As String
    Dim dblLegendW As Double, dblLegendH As Double '图例宽，高
    Dim dblLegendX As Double, dblLegendY As Double '图例起点坐标
    Dim dblRadianStart As Double, dblRadianEnd As Double '饼图单个项目起始与终止弧度

    If UBound(mtpItem) = 0 Then Exit Sub

    '以窗体中心为原点，选窗体长和宽中最小的那个的1/4为半径
    R = IIf(UserControl.ScaleWidth > UserControl.ScaleHeight, UserControl.ScaleHeight / 4, UserControl.ScaleWidth / 4)
    x = UserControl.ScaleWidth / 2
    y = UserControl.ScaleHeight / 2
    dblPi = 4 * Atn(1)
    
    '以实心方式填充
    UserControl.FillStyle = vbFSSolid

    Set UserControl.Font = mstdItemFont
    UserControl.ForeColor = moleItemColor
    UserControl.DrawStyle = 5
    
    '画扇形
    dblAccumulate = 0
    For i = 1 To UBound(mtpItem)
        dblAccumulate = dblAccumulate + mtpItem(i).PartNumber
        '判断弧度的正负
        dblRadianStart = 1 / 4 - dblAccumulate / mdblTotal
        If dblRadianStart <= 0 Then
            dblRadianStart = dblRadianStart + 1
        End If
        dblRadianEnd = 1 / 4 - dblAccumulate / mdblTotal + mtpItem(i).PartNumber / mdblTotal
        If dblRadianEnd <= 0 Then
            dblRadianEnd = dblRadianEnd + 1
        End If
        UserControl.FillColor = mtpItem(i).Color
        UserControl.Circle (x, y), R, mtpItem(i).Color, -dblRadianStart * 2 * dblPi, -dblRadianEnd * 2 * dblPi
    Next
    
    '画指示线以及项目名称，分为只有一个子项和有很多子项两种情况
    UserControl.DrawStyle = 0
    If UBound(mtpItem) = 1 Then
        '指示线，以圆心为起点，圆心x+2r为终点
        '因为当时选定圆心坐标时是以1/2长和宽为原点，1/4长或宽为半径，即圆心距边界最短距离为2r
        '所以选定指示线终点为x+2r
        UserControl.Line (x, y)-(x + R * mInstructBaseX, y + R * mInstructBaseY), moleLineColor
        UserControl.Line (x + R * mInstructBaseX, y + R * mInstructBaseY)-(x + R * 2, y + R * mInstructBaseY), moleLineColor

        '判断是显示百分比还是数字，或是不显示
        If msymType = 1 Then
            strTitle = mtpItem(1).Title & ":" & mdblTotal
        ElseIf msymType = 2 Then
            strTitle = mtpItem(1).Title & ":" & "100%"
        Else
            strTitle = mtpItem(1).Title
        End If

        UserControl.CurrentX = x + R * 2 - UserControl.TextWidth(strTitle)
        UserControl.CurrentY = y + R / 2 - UserControl.TextHeight("TT")
        UserControl.Print strTitle
    Else
        For i = 1 To UBound(mtpItem)
            '根据饼图中每个扇形所占角度计算偏移量x0，y0
            '偏移量计算公式为：2 / 3 * r * cos(1 / 2 * π - 1 / 2 * θo - θ1) 和 -2 / 3 * r * sin(1 / 2 * π - 1 / 2 * θo - θ1)
            '其中θo为当前项目扇形所占弧度，θ1为该扇形之前所有扇形所占弧度之和
            'x1，y1为指示线起点坐标
            
            dblRadianLine = 1 / 4 - (mtpItem(i).PartNumber / 2 + dblAccumulate) / mdblTotal
            x0 = mInstructStart * R * Cos(dblRadianLine * 2 * dblPi)
            y0 = -mInstructStart * R * Sin(dblRadianLine * 2 * dblPi)
            
            x1 = x0 + x
            y1 = y0 + y

            '指示线分为两部分
            '第一部分以x1，y1为起点，终点x轴偏移量设置为r/2，正好超出半径r/6的距离，y轴偏移量不是一个固定值，是根据弧度计算得来的。
            'y轴偏移量计算公式：1 / 2 * r * abs(sin(1 / 2 * π - 1 / 2 * θo - θ1))
            '其中θo为当前项目扇形所占弧度，θ1为该扇形之前所有扇形所占弧度之和
            '第二部分是以第一部分终点为起点，y轴偏移量为0，x轴偏移量为3/2r，因为第一部分x轴偏移量为1/2r，而圆心距边界距离为2r，所以第二部分偏移量即为2r-1/2r
            UserControl.Line (x1, y1)-(x1 + mInstructBaseX * R * Sgn(x0), y1 + mInstructBaseY * R * Sgn(y0) * Abs(Sin(dblRadianLine * 2 * dblPi))), moleLineColor
            UserControl.Line (x1 + mInstructBaseX * R * Sgn(x0), y1 + mInstructBaseY * R * Sgn(y0) * Abs(Sin(dblRadianLine * 2 * dblPi)))-(x + R * 2 * Sgn(x0), y1 + mInstructBaseY * R * Sgn(y0) * Abs(Sin(dblRadianLine * 2 * dblPi))), moleLineColor

            If msymType = 1 Then
                strTitle = mtpItem(i).Title & ":" & mtpItem(i).PartNumber
            ElseIf msymType = 2 Then
                strTitle = mtpItem(i).Title & ":" & Round(mtpItem(i).PartNumber / mdblTotal * 100, 1) & "%"
            Else
                strTitle = mtpItem(i).Title
            End If

            '打印项目标题
            '如果指示线在左侧，那么标题开始位置即为指示线最左端，即上面画第二条指示线时的终点坐标
            '如果指示线在右侧，那么标题开始位置即为指示线最右端-标题长度
            UserControl.CurrentX = x + R * 2 * Sgn(x0) - IIf(Sgn(x0) = -1, 0, UserControl.TextWidth(strTitle))
            UserControl.CurrentY = y1 + mInstructBaseY * R * Sgn(y0) * Abs(Sin(dblRadianLine * 2 * dblPi)) - UserControl.TextHeight("TT")
            UserControl.Print strTitle
            dblAccumulate = dblAccumulate + mtpItem(i).PartNumber
        Next
    End If
    
    '显示标题
    Set UserControl.Font = mstdTitleFont
    UserControl.ForeColor = moleTitleColor
    '设置标题横向居中，因为饼图所占高度最大为窗体高度的1/2，另外还有指示线所占高度，所以我选定标题y值为窗体高度的1/8-标题高度
    UserControl.CurrentX = x - UserControl.TextWidth(mstrTitle) / 2
    UserControl.CurrentY = UserControl.Height / 8 - UserControl.TextHeight("TT")
    UserControl.Print mstrTitle
    
    '画图例
    If mblnLegend = True Then
        Set UserControl.Font = mstdItemFont
        '设置图例宽度为高度的2倍，高度为字体高度
        dblLegendW = UserControl.TextHeight("TT") * 2
        dblLegendH = UserControl.TextHeight("TT")
        '图例起点为指示线最左侧，指示线最下侧 - 一个图例高度
        dblLegendX = x - R * 2
        dblLegendY = y + R * IIf((mInstructStart + mInstructBaseY) > 1, (mInstructStart + mInstructBaseY), 1) + dblLegendH
        
        For i = 1 To UBound(mtpItem)
            UserControl.Line (dblLegendX, dblLegendY)-(dblLegendX + dblLegendW, dblLegendY + dblLegendH), mtpItem(i).Color, BF
            UserControl.CurrentX = dblLegendX + dblLegendW * 1.5
            UserControl.CurrentY = dblLegendY
            strTitle = mtpItem(i).Title & "(" & mtpItem(i).PartNumber & ")"
            UserControl.Print strTitle
            dblLegendX = dblLegendX + dblLegendW * 2 + UserControl.TextWidth(strTitle)
            If i = UBound(mtpItem) Then Exit For
            strTitle = mtpItem(i + 1).Title & "(" & mtpItem(i + 1).PartNumber & ")"
            '判断下一个图例是否超出了最右侧指示线，如果超出了，则另起一行
            If dblLegendX + dblLegendW * 1.5 + UserControl.TextWidth(strTitle) > x + 2 * R Then
                dblLegendX = x - R * 2
                dblLegendY = dblLegendY + dblLegendH * 2
            End If
        Next
    End If
End Sub

Private Sub UserControl_InitProperties()
    Set mstdItemFont = New StdFont
    Set mstdTitleFont = New StdFont
    ReDim mtpItem(0)
    mlngLevColor = 30
    mstrTitle = "标题"
    mblnLegend = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set mstdItemFont = New StdFont
    Set mstdTitleFont = New StdFont
    ReDim mtpItem(0)
    
    '初始化内置颜色
    mtpColor(0).coR = 158
    mtpColor(0).coG = 65
    mtpColor(0).coB = 62
    mtpColor(0).coLevel = 0
    mtpColor(1).coR = 127
    mtpColor(1).coG = 154
    mtpColor(1).coB = 72
    mtpColor(1).coLevel = 0
    mtpColor(2).coR = 105
    mtpColor(2).coG = 81
    mtpColor(2).coB = 133
    mtpColor(2).coLevel = 0
    mtpColor(3).coR = 60
    mtpColor(3).coG = 141
    mtpColor(3).coB = 163
    mtpColor(3).coLevel = 0
    mtpColor(4).coR = 204
    mtpColor(4).coG = 123
    mtpColor(4).coB = 56
    mtpColor(4).coLevel = 0
    mtpColor(5).coR = 79
    mtpColor(5).coG = 129
    mtpColor(5).coB = 189
    mtpColor(5).coLevel = 0
    
    mstType = PropBag.ReadProperty("ShowType")
    msymType = PropBag.ReadProperty("SymbolType")
    mstrTitle = PropBag.ReadProperty("Title")
    moleLineColor = PropBag.ReadProperty("LineColor")
    Set mstdTitleFont = PropBag.ReadProperty("TitleFont")
    Set mstdItemFont = PropBag.ReadProperty("ItemFont")
    moleItemColor = PropBag.ReadProperty("ItemColor")
    moleTitleColor = PropBag.ReadProperty("TitleColor")
    mblnLegend = PropBag.ReadProperty("Legend")
    mlngLevColor = PropBag.ReadProperty("LevColor")
End Sub

Private Sub UserControl_Resize()
    '只有将图片展示出来时，才提供resize功能
    If mblnIsShow Then
        PaintChart False
    Else
        '展示示例效果
        UserControl.Cls
        ReDim mtpItem(0)
        Call addItem(, vbRed, 1)
        Call addItem(, vbGreen, 1)
        Call addItem(, vbBlue, 1)
        mdblTotal = 3
        If mstType = 饼图 Then
            Call showCircle
        End If
        ReDim mtpItem(0)
        mdblTotal = 0
        mblnIsShow = False
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ShowType", mstType)
    Call PropBag.WriteProperty("SymbolType", msymType)
    Call PropBag.WriteProperty("Title", mstrTitle)
    Call PropBag.WriteProperty("LineColor", moleLineColor)
    Call PropBag.WriteProperty("TitleFont", mstdTitleFont)
    Call PropBag.WriteProperty("ItemFont", mstdItemFont)
    Call PropBag.WriteProperty("ItemColor", moleItemColor)
    Call PropBag.WriteProperty("TitleColor", moleTitleColor)
    Call PropBag.WriteProperty("Legend", mblnLegend)
    Call PropBag.WriteProperty("LevColor", mlngLevColor)
End Sub


