VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'######################################################################################
'##模 块 名：cFont.cls
'##创 建 人：吴庆伟
'##日    期：2005年5月1日
'##修 改 人：
'##日    期：
'##描    述：字体属性类。对应于TOM接口中的Font。
'##版    本：
'######################################################################################

Option Explicit

Public Enum UnderLineEnum
    cprnone = 0
    cprDotted = 4
    cprDash = 5
    cprDashDot = 6
    cprDashDotDot = 7
    cprwave = 8
    cprThick = 9
    cprHair = 10
End Enum

Public Enum UnderlineColorEnum
    cprUnderLineColor_Black = &H0
    cprUnderLineColor_Blue = &H10
    cprUnderLineColor_Cyan = &H20
    cprUnderLineColor_LimeGreen = &H30
    cprUnderLineColor_Magenta = &H40
    cprUnderLineColor_Red = &H50
    cprUnderLineColor_Yellow = &H60
    cprUnderLineColor_White = &H70
    cprUnderLineColor_DarkBlue = &H80
    cprUnderLineColor_DarkCyan = &H90
    cprUnderLineColor_Green = &HA0
    cprUnderLineColor_DarkMagenta = &HB0
    cprUnderLineColor_Brown = &HC0
    cprUnderLineColor_OliveGreen = &HD0
    cprUnderLineColor_DarkGray = &HE0
    cprUnderLineColor_Gray = &HF0
End Enum

Private m_Start As Long         'Range的起始位置
Private m_End As Long           'Range的终止位置
Private m_ReadOnly As Boolean   '是否只读
Private m_bSelection As Boolean '是否是Selection对象

Private TOM As cTabTextDocument     'TOM对象

Public Property Let UnderlineColor(vData As UnderlineColorEnum)
    '注意：下划线颜色是一个不持久方法，只能在编辑过程中设置有效，一旦保存为RTF，则颜色及下划线均丢失！
    Dim fmt As CHARFORMAT2
    Dim udl As UnderLineEnum
    
    udl = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Underline, TOM.TextDocument.Range(m_Start, m_End).Font.Underline)
    If udl = tomUndefined Then GoTo LL
    fmt.cbSize = Len(fmt)
    fmt.dwMask = CFM_UNDERLINETYPE
    fmt.bUnderlineType = CByte(CByte(udl) Or CByte(vData))
    
    SendMessage TOM.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, fmt
LL:
End Property

Public Property Get UnderlineColor() As UnderlineColorEnum
    '注意：下划线颜色是一个不持久方法，只能在编辑过程中设置有效，一旦保存为RTF，则颜色及下划线均丢失！
    Dim fmt As CHARFORMAT2
    
    fmt.cbSize = Len(fmt)
    SendMessage TOM.hWnd, EM_GETCHARFORMAT, SCF_SELECTION, fmt
    If (fmt.dwMask And CFM_UNDERLINETYPE) = 0 Then UnderlineColor = cprUnderLineColor_Black
    UnderlineColor = CByte(fmt.bUnderlineType And &HF0)
End Property

Public Function GetFontFmt() As cTabFontFormat
    '获取字体属性
    Dim R As New cTabFontFormat
    R.AllCaps = Me.AllCaps
    R.BackColor = Me.BackColor
    R.Bold = Me.Bold
    R.ForeColor = Me.ForeColor
    R.Hidden = Me.Hidden
    R.Italic = Me.Italic
    R.Link = Me.Link
    R.Name = Me.Name
    R.Protected = Me.Protected
    R.Size = Me.Size
    R.Strikethrough = Me.Strikethrough
    R.Subscript = Me.Subscript
    R.Superscript = Me.Superscript
    R.Underline = Me.Underline
    Set GetFontFmt = R
End Function

Public Function SetFontFmt(ByRef vData As cTabFontFormat)
    '设置字体属性
    '参数： blnExtend 扩展模式，表示是否允许特殊设置，默认为False，表示：
    '       不允许背景色为 SPECIALBACKCOLOR
    '       不允许下划线为 cprWave 波浪线
    '       不允许前景色为 REVISION_COLOR
    '说明：
    '一律不允许设置保护、隐藏、删除线
    
'    Me.AllCaps = vData.AllCaps
'    Me.Hidden = vData.Hidden
'    Me.Link = vData.Link
'    Me.Protected = vData.Protected
'    Me.Strikethrough = vData.Strikethrough
'    If blnExtend Then
'        '扩展模式
'        Me.BackColor = vData.BackColor
'        Me.ForeColor = vData.ForeColor
'        Me.Underline = vData.Underline
'    Else
'        Me.BackColor = IIf(vData.BackColor = SPECIALBACKCOLOR, &H99FFFF, vData.BackColor)   '替换为浅黄色
'        Me.ForeColor = IIf(vData.ForeColor = REVISION_COLOR, vbRed, vData.ForeColor)        '替换为红色
'        Me.Underline = IIf(vData.Underline = cprWave, cprHair, vData.Underline)             '替换为单下划线
'    End If
    
    Me.Underline = vData.Underline
    Me.Bold = vData.Bold
    Me.Italic = vData.Italic
    Me.Name = vData.Name
    Me.Size = vData.Size
    Me.Subscript = vData.Subscript
    Me.Superscript = vData.Superscript
End Function

Public Sub Init(ByVal cTOM As cTabTextDocument, lStart As Long, lEnd As Long, Optional bReadOnly As Boolean = False, Optional bSelection As Boolean = False)
    Set TOM = cTOM
    m_Start = lStart
    m_End = lEnd
    m_ReadOnly = bReadOnly
    m_bSelection = bSelection
End Sub

Public Sub Reset()
'重置字体属性为系统默认值。
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Reset (tomDefault)
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Reset (tomDefault)
    End If
    TOM.TextDocument.DefaultTabStop = 21
End Sub

Public Property Let Underline(ByVal vData As UnderLineEnum)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Underline = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Underline = vData
    End If
End Property

Public Property Get Underline() As UnderLineEnum
    Dim l As Long
    If m_bSelection Then
        l = TOM.TextDocument.Selection.Font.Underline
    Else
        l = TOM.TextDocument.Range(m_Start, m_End).Font.Underline
    End If
    Select Case l
    Case 0, 4, 5, 6, 7, 8, 9, 10
    Case 1
        l = cprHair
    Case Else
        l = 0
    End Select
    Underline = l
End Property

Public Property Let Link(ByVal vData As Boolean)
    If m_ReadOnly Then Exit Property
    Dim tCF2 As CHARFORMAT2
    Dim lR As Long
    tCF2.dwMask = CFM_LINK
    If (vData) Then
       tCF2.dwEffects = CFE_LINK
    End If
    tCF2.cbSize = Len(tCF2)
    lR = SendMessage(TOM.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, tCF2)
End Property

Public Property Get Link() As Boolean
    Dim tCF2 As CHARFORMAT2
    Dim lR As Long
    tCF2.dwMask = CFM_LINK
    tCF2.cbSize = Len(tCF2)
    lR = SendMessage(TOM.hWnd, EM_GETCHARFORMAT, SCF_SELECTION, tCF2)
    Link = ((tCF2.dwEffects And CFE_LINK) = CFE_LINK)
End Property

Public Property Let Superscript(ByVal vData As Boolean)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Superscript = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Superscript = vData
    End If
End Property

Public Property Get Superscript() As Boolean
    Superscript = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Superscript, TOM.TextDocument.Range(m_Start, m_End).Font.Superscript)
End Property

Public Property Let Subscript(ByVal vData As Boolean)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Subscript = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Subscript = vData
    End If
End Property

Public Property Get Subscript() As Boolean
    Subscript = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Subscript, TOM.TextDocument.Range(m_Start, m_End).Font.Subscript)
End Property

Public Property Let Strikethrough(ByVal vData As Boolean)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Strikethrough = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Strikethrough = vData
    End If
End Property

Public Property Get Strikethrough() As Boolean
    Strikethrough = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Strikethrough, TOM.TextDocument.Range(m_Start, m_End).Font.Strikethrough)
End Property

Public Property Let Size(ByVal vData As Single)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Size = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Size = vData
    End If
End Property

Public Property Get Size() As Single
    Size = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Size, TOM.TextDocument.Range(m_Start, m_End).Font.Size)
End Property

Public Property Let Protected(ByVal vData As Boolean)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Protected = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Protected = vData
    End If
End Property

Public Property Get Protected() As Boolean
    Protected = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Protected, TOM.TextDocument.Range(m_Start, m_End).Font.Protected)
End Property

Public Property Let Name(ByVal vData As String)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Name = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Name = vData
    End If
End Property

Public Property Get Name() As String
    Name = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Name, TOM.TextDocument.Range(m_Start, m_End).Font.Name)
End Property

Public Property Let Italic(ByVal vData As Boolean)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Italic = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Italic = vData
    End If
End Property

Public Property Get Italic() As Boolean
    Italic = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Italic, TOM.TextDocument.Range(m_Start, m_End).Font.Italic)
End Property

Public Property Let Hidden(ByVal vData As Boolean)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Hidden = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Hidden = vData
    End If
End Property

Public Property Get Hidden() As Boolean
    Hidden = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Hidden, TOM.TextDocument.Range(m_Start, m_End).Font.Hidden)
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.ForeColor = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.ForeColor = vData
    End If
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = IIf(m_bSelection, TOM.TextDocument.Selection.Font.ForeColor, TOM.TextDocument.Range(m_Start, m_End).Font.ForeColor)
End Property

Public Sub SetForeColorRGB(ByVal R As Long, ByVal G As Long, ByVal b As Long)
    If m_ReadOnly Then Exit Sub
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.ForeColor = RGB(R, G, b)
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.ForeColor = RGB(R, G, b)
    End If
End Sub

Public Sub GetForeColorRGB(ByRef R As Long, ByRef G As Long, ByRef b As Long)
    Dim mColor As OLE_COLOR
    mColor = ForeColor
    R = RGBRed(mColor)
    G = RGBGreen(mColor)
    b = RGBBlue(mColor)
End Sub

Public Property Get ForeColorR() As Long
    Dim mColor As OLE_COLOR
    mColor = ForeColor
    ForeColorR = RGBRed(mColor)
End Property

Public Property Get ForeColorG() As Long
    Dim mColor As OLE_COLOR
    mColor = ForeColor
    ForeColorG = RGBGreen(mColor)
End Property

Public Property Get ForeColorB() As Long
    Dim mColor As OLE_COLOR
    mColor = ForeColor
    ForeColorB = RGBBlue(mColor)
End Property

Public Property Let Bold(ByVal vData As Boolean)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Bold = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Bold = vData
    End If
End Property

Public Property Get Bold() As Boolean
    Bold = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Bold, TOM.TextDocument.Range(m_Start, m_End).Font.Bold)
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.BackColor = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.BackColor = vData
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = IIf(m_bSelection, TOM.TextDocument.Selection.Font.BackColor, TOM.TextDocument.Range(m_Start, m_End).Font.BackColor)
End Property

Public Property Let AllCaps(ByVal vData As Boolean)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.AllCaps = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.AllCaps = vData
    End If
End Property

Public Property Get AllCaps() As Boolean
    AllCaps = IIf(m_bSelection, TOM.TextDocument.Selection.Font.AllCaps, TOM.TextDocument.Range(m_Start, m_End).Font.AllCaps)
End Property

