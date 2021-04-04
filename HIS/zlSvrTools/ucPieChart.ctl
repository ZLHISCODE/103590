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
Private mstType As Show_Type '0-��ͼ��Ĭ�ϣ���1-����ͼ��2-��״ͼ
Private msymType As Symbol_Type '0-ʲô������ʾ(Ĭ��)��1-��ʾ������2-��ʾ�ٷֱ�
Private moleLineColor As OLE_COLOR, moleItemColor As OLE_COLOR, moleTitleColor As OLE_COLOR
Private mstdItemFont As New StdFont, mstdTitleFont As New StdFont
Private mstrTitle As String '����
Private mblnIsShow As Boolean '�ж��Ƿ��Ѿ���ͼƬչʾ����
Private mblnLegend As Boolean '�ж��Ƿ���ʾͼ��
Private mlngLevColor As Long   '��ɫƽ���ݶ�

Public Enum Show_Type
    ��ͼ = 0
    ����ͼ = 1
    ��״ͼ = 2
End Enum

Public Enum Symbol_Type
    ����ʾ = 0
    ��ʾ���� = 1
    ��ʾ�ٷֱ� = 2
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

Private Const mInstructStart = 0.66 '��һ��ָʾ����㵽Բ�ľ�����r�ı���
Private Const mInstructBaseY = 0.5 '��һ��ָʾ��y��ƫ������׼ֵ
Private Const mInstructBaseX = 0.5 '��һ��ָʾ��x��ƫ������׼ֵ

'������ɫ��ʾ�ݶ�
Public Property Let LevColor(ByVal lngLevColor As Long)
    If lngLevColor = 0 Then lngLevColor = 30
    mlngLevColor = lngLevColor
End Property

Public Property Get LevColor() As Long
    LevColor = mlngLevColor
End Property

'������ʾ����
Public Property Let ShowType(ByVal stType As Show_Type)
    mstType = stType
    UserControl_Resize
End Property

Public Property Get ShowType() As Show_Type
    ShowType = mstType
End Property

'����ָʾ����ɫ
Public Property Let LineColor(ByVal oleLineColor As OLE_COLOR)
    moleLineColor = oleLineColor
    UserControl_Resize
End Property

Public Property Get LineColor() As OLE_COLOR
    LineColor = moleLineColor
End Property

'���ñ���
Public Property Let Title(ByVal strTitle As String)
    mstrTitle = strTitle
    UserControl_Resize
End Property

Public Property Get Title() As String
    Title = mstrTitle
End Property

'���ñ�������
Public Property Set TitleFont(ByVal stdTitleFont As StdFont)
    Set mstdTitleFont = stdTitleFont
    UserControl_Resize
End Property

Public Property Get TitleFont() As StdFont
    Set TitleFont = mstdTitleFont
End Property

'���ñ���������ɫ
Public Property Let TitleColor(ByVal oleTitleColor As OLE_COLOR)
    moleTitleColor = oleTitleColor
    UserControl_Resize
End Property

Public Property Get TitleColor() As OLE_COLOR
    TitleColor = moleTitleColor
End Property

'������Ŀ����
Public Property Set ItemFont(ByVal stdItemFont As StdFont)
    Set mstdItemFont = stdItemFont
    UserControl_Resize
End Property

Public Property Get ItemFont() As StdFont
    Set ItemFont = mstdItemFont
End Property

'������Ŀ������ɫ
Public Property Let ItemColor(ByVal oleItemColor As OLE_COLOR)
    moleItemColor = oleItemColor
    UserControl_Resize
End Property

Public Property Get ItemColor() As OLE_COLOR
    ItemColor = moleItemColor
End Property

'����ÿ����Ŀ������ʾ��ʽ��1-ʲô������ʾ��2-��ʾ������3-��ʾ�ٷֱ�
Public Property Let SymbolType(ByVal symType As Symbol_Type)
    msymType = symType
    UserControl_Resize
End Property

Public Property Get SymbolType() As Symbol_Type
    SymbolType = msymType
End Property

'�����Ƿ���ʾͼ��
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
        strItemTitle = "��Ŀ" & UBound(mtpItem) + 1
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
    'bolType:�ж����ڲ����û����ⲿ����
    '���ԭ�е�ͼ��
    UserControl.Cls
    
    '�ж���ʾ��ʽ
    If mstType = ��ͼ Then
        Call showCircle
    End If
    mblnIsShow = True
End Sub

'������ͼ��ʽ��ʾ
Private Sub showPolygon()

End Sub

'�Ա�ͼ��ʽ��ʾ
Private Sub showCircle()
    Dim i As Long, K As Long
    Dim dblPi As Double '��
    Dim R As Double  '�뾶
    Dim x As Double  'Բ��x����
    Dim y As Double  'Բ��y����
    Dim x0 As Double, y0 As Double 'ָʾ������������Բ��������
    Dim x1 As Double, y1 As Double 'ָʾ���������
    Dim dblRadianLine As Double 'ָʾ����㻡��
    Dim dblAccumulate As Double
    Dim strTitle As String
    Dim dblLegendW As Double, dblLegendH As Double 'ͼ������
    Dim dblLegendX As Double, dblLegendY As Double 'ͼ���������
    Dim dblRadianStart As Double, dblRadianEnd As Double '��ͼ������Ŀ��ʼ����ֹ����

    If UBound(mtpItem) = 0 Then Exit Sub

    '�Դ�������Ϊԭ�㣬ѡ���峤�Ϳ�����С���Ǹ���1/4Ϊ�뾶
    R = IIf(UserControl.ScaleWidth > UserControl.ScaleHeight, UserControl.ScaleHeight / 4, UserControl.ScaleWidth / 4)
    x = UserControl.ScaleWidth / 2
    y = UserControl.ScaleHeight / 2
    dblPi = 4 * Atn(1)
    
    '��ʵ�ķ�ʽ���
    UserControl.FillStyle = vbFSSolid

    Set UserControl.Font = mstdItemFont
    UserControl.ForeColor = moleItemColor
    UserControl.DrawStyle = 5
    
    '������
    dblAccumulate = 0
    For i = 1 To UBound(mtpItem)
        dblAccumulate = dblAccumulate + mtpItem(i).PartNumber
        '�жϻ��ȵ�����
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
    
    '��ָʾ���Լ���Ŀ���ƣ���Ϊֻ��һ��������кܶ������������
    UserControl.DrawStyle = 0
    If UBound(mtpItem) = 1 Then
        'ָʾ�ߣ���Բ��Ϊ��㣬Բ��x+2rΪ�յ�
        '��Ϊ��ʱѡ��Բ������ʱ����1/2���Ϳ�Ϊԭ�㣬1/4�����Ϊ�뾶����Բ�ľ�߽���̾���Ϊ2r
        '����ѡ��ָʾ���յ�Ϊx+2r
        UserControl.Line (x, y)-(x + R * mInstructBaseX, y + R * mInstructBaseY), moleLineColor
        UserControl.Line (x + R * mInstructBaseX, y + R * mInstructBaseY)-(x + R * 2, y + R * mInstructBaseY), moleLineColor

        '�ж�����ʾ�ٷֱȻ������֣����ǲ���ʾ
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
            '���ݱ�ͼ��ÿ��������ռ�Ƕȼ���ƫ����x0��y0
            'ƫ�������㹫ʽΪ��2 / 3 * r * cos(1 / 2 * �� - 1 / 2 * ��o - ��1) �� -2 / 3 * r * sin(1 / 2 * �� - 1 / 2 * ��o - ��1)
            '���Ц�oΪ��ǰ��Ŀ������ռ���ȣ���1Ϊ������֮ǰ����������ռ����֮��
            'x1��y1Ϊָʾ���������
            
            dblRadianLine = 1 / 4 - (mtpItem(i).PartNumber / 2 + dblAccumulate) / mdblTotal
            x0 = mInstructStart * R * Cos(dblRadianLine * 2 * dblPi)
            y0 = -mInstructStart * R * Sin(dblRadianLine * 2 * dblPi)
            
            x1 = x0 + x
            y1 = y0 + y

            'ָʾ�߷�Ϊ������
            '��һ������x1��y1Ϊ��㣬�յ�x��ƫ��������Ϊr/2�����ó����뾶r/6�ľ��룬y��ƫ��������һ���̶�ֵ���Ǹ��ݻ��ȼ�������ġ�
            'y��ƫ�������㹫ʽ��1 / 2 * r * abs(sin(1 / 2 * �� - 1 / 2 * ��o - ��1))
            '���Ц�oΪ��ǰ��Ŀ������ռ���ȣ���1Ϊ������֮ǰ����������ռ����֮��
            '�ڶ��������Ե�һ�����յ�Ϊ��㣬y��ƫ����Ϊ0��x��ƫ����Ϊ3/2r����Ϊ��һ����x��ƫ����Ϊ1/2r����Բ�ľ�߽����Ϊ2r�����Եڶ�����ƫ������Ϊ2r-1/2r
            UserControl.Line (x1, y1)-(x1 + mInstructBaseX * R * Sgn(x0), y1 + mInstructBaseY * R * Sgn(y0) * Abs(Sin(dblRadianLine * 2 * dblPi))), moleLineColor
            UserControl.Line (x1 + mInstructBaseX * R * Sgn(x0), y1 + mInstructBaseY * R * Sgn(y0) * Abs(Sin(dblRadianLine * 2 * dblPi)))-(x + R * 2 * Sgn(x0), y1 + mInstructBaseY * R * Sgn(y0) * Abs(Sin(dblRadianLine * 2 * dblPi))), moleLineColor

            If msymType = 1 Then
                strTitle = mtpItem(i).Title & ":" & mtpItem(i).PartNumber
            ElseIf msymType = 2 Then
                strTitle = mtpItem(i).Title & ":" & Round(mtpItem(i).PartNumber / mdblTotal * 100, 1) & "%"
            Else
                strTitle = mtpItem(i).Title
            End If

            '��ӡ��Ŀ����
            '���ָʾ������࣬��ô���⿪ʼλ�ü�Ϊָʾ������ˣ������滭�ڶ���ָʾ��ʱ���յ�����
            '���ָʾ�����Ҳ࣬��ô���⿪ʼλ�ü�Ϊָʾ�����Ҷ�-���ⳤ��
            UserControl.CurrentX = x + R * 2 * Sgn(x0) - IIf(Sgn(x0) = -1, 0, UserControl.TextWidth(strTitle))
            UserControl.CurrentY = y1 + mInstructBaseY * R * Sgn(y0) * Abs(Sin(dblRadianLine * 2 * dblPi)) - UserControl.TextHeight("TT")
            UserControl.Print strTitle
            dblAccumulate = dblAccumulate + mtpItem(i).PartNumber
        Next
    End If
    
    '��ʾ����
    Set UserControl.Font = mstdTitleFont
    UserControl.ForeColor = moleTitleColor
    '���ñ��������У���Ϊ��ͼ��ռ�߶����Ϊ����߶ȵ�1/2�����⻹��ָʾ����ռ�߶ȣ�������ѡ������yֵΪ����߶ȵ�1/8-����߶�
    UserControl.CurrentX = x - UserControl.TextWidth(mstrTitle) / 2
    UserControl.CurrentY = UserControl.Height / 8 - UserControl.TextHeight("TT")
    UserControl.Print mstrTitle
    
    '��ͼ��
    If mblnLegend = True Then
        Set UserControl.Font = mstdItemFont
        '����ͼ�����Ϊ�߶ȵ�2�����߶�Ϊ����߶�
        dblLegendW = UserControl.TextHeight("TT") * 2
        dblLegendH = UserControl.TextHeight("TT")
        'ͼ�����Ϊָʾ������ָ࣬ʾ�����²� - һ��ͼ���߶�
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
            '�ж���һ��ͼ���Ƿ񳬳������Ҳ�ָʾ�ߣ���������ˣ�������һ��
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
    mstrTitle = "����"
    mblnLegend = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set mstdItemFont = New StdFont
    Set mstdTitleFont = New StdFont
    ReDim mtpItem(0)
    
    '��ʼ��������ɫ
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
    'ֻ�н�ͼƬչʾ����ʱ�����ṩresize����
    If mblnIsShow Then
        PaintChart False
    Else
        'չʾʾ��Ч��
        UserControl.Cls
        ReDim mtpItem(0)
        Call addItem(, vbRed, 1)
        Call addItem(, vbGreen, 1)
        Call addItem(, vbBlue, 1)
        mdblTotal = 3
        If mstType = ��ͼ Then
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


