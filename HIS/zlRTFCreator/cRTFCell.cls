VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFCell"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'保持属性值的局部变量
Private mvarAlignmentV As AlignVTypeEnum    '单元格垂直对齐，1～3。默认：0
Private mvarAlignment As AlignTypeEnum      '单元格水平对齐，1～3。默认：0
Private mvarBorderLeft As Boolean       '单元格左边框
Private mvarBorderRight As Boolean      '单元格右边框
Private mvarBorderTop As Boolean        '单元格上边框
Private mvarBorderBottom As Boolean     '单元格下边框
Private mvarBorderLeftDown As Boolean   '左下内斜线\
Private mvarBorderLeftUp As Boolean     '左上内斜线/
Private mvarBorderWidth As Long         '边框宽度，0～75，默认：15。
Private mvarBorderStyle As BorderStyleEnum         '边框类型。1～28种类型。
Private mvarBorderColor As ColorEnum                   '颜色类型。
Private mvarCellPatType As PatTypeEnum             '单元格背景图案类型。1～12，默认：0
Private mvarCellPatBackColor As ColorEnum              '单元格背景图案的线条颜色。
Private mvarCellPatForeColor As ColorEnum              '单元格背景图案的背景颜色。
Private mvarCellPatPercentage As Long                  '单元格底纹明暗百分比。
Private mvarWidth As Long                              '单元格首选宽度
Private mvarWidthUnit As CellLengthUnitEnum            '单元格首选宽度单位（0～3），0：空，1：自动，2：百分比，3：缇（默认）。
Private mvarRight As Long                              '单元格右边界位置
Private mvarMergeVStart As Boolean      '垂直合并的第一个单元格
Private mvarMergeV As Boolean           '垂直与前一单元格合并
Private mvarMergeStart As Boolean       '水平合并的第一个单元格
Private mvarMerge As Boolean            '水平与前一单元格合并
Private mvarRow As Long                 '行数
Private mvarCol As Long                 '列数
Private mvarFitText As Boolean          '文本适应单元格宽度
Private mvarNoWrap As Boolean           '单元格不允许文本换行
Private mvarNested As Boolean           '是否嵌套，默认为：False
Private mvarData As Variant                             '单元格内容，文本、图片、嵌套表格
Private mvarDataType As CellContentTypeEnum             '内容类型。0：文本，1：图片，2：嵌套表格。默认：0
Private mvarText As String
Private mvarBold As Boolean
Private mvarItalic As Boolean
Private mvarForeColor As ColorEnum
Private mvarFontSizeFixup As FontSizeEnum
Private mvarFontSize As Long

Public Property Let FontSize(ByVal vData As Long)
Attribute FontSize.VB_Description = "简单文本的字体大小，为整型值。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FontSize = 5
    mvarFontSize = vData
End Property


Public Property Get FontSize() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FontSize
    FontSize = mvarFontSize
End Property




Public Property Let FontSizeFixup(ByVal vData As FontSizeEnum)
Attribute FontSizeFixup.VB_Description = "简单文本的字体大小索引值。"
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.FontSizeFixup = Form1
    mvarFontSizeFixup = vData
End Property

Public Property Get FontSizeFixup() As FontSizeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FontSizeFixup
    FontSizeFixup = mvarFontSizeFixup
End Property

Public Property Let ForeColor(ByVal vData As ColorEnum)
Attribute ForeColor.VB_Description = "简单文本的前景色。"
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.ForeColor = Form1
    mvarForeColor = vData
End Property

Public Property Get ForeColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ForeColor
    ForeColor = mvarForeColor
End Property



Public Property Let Italic(ByVal vData As Boolean)
Attribute Italic.VB_Description = "简单文本的斜体。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Italic = 5
    mvarItalic = vData
End Property


Public Property Get Italic() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Italic
    Italic = mvarItalic
End Property



Public Property Let Bold(ByVal vData As Boolean)
Attribute Bold.VB_Description = "简单文本的粗体与否。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Bold = 5
    mvarBold = vData
End Property


Public Property Get Bold() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Bold
    Bold = mvarBold
End Property




Public Property Let Text(ByVal vData As String)
Attribute Text.VB_Description = "简单文本。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Text = 5
    mvarText = vData
End Property

Public Property Get Text() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Text
    Text = mvarText
End Property

Public Property Let CellPatPercentage(ByVal vData As Long)
Attribute CellPatPercentage.VB_Description = "单元格底纹明暗百分比。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CellPatPercentage = 5
    mvarCellPatPercentage = vData
End Property

Public Property Get CellPatPercentage() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CellPatPercentage
    CellPatPercentage = mvarCellPatPercentage
End Property

Public Property Let BorderWidth(ByVal vData As Long)
Attribute BorderWidth.VB_Description = "边框的画线宽度。0～75。默认：15"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderWidth = 5
    mvarBorderWidth = vData
End Property

Public Property Get BorderWidth() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderWidth
    BorderWidth = mvarBorderWidth
End Property

Public Function GetCellDefineRTF() As String
Attribute GetCellDefineRTF.VB_Description = "获取单元格定义的RTF字符串。"
'获取单元格定义RTF
    Dim strR As String
    
    '合并单元格处理
    If mvarMergeVStart Then
        strR = strR & "\clvmgf"
    End If
    If mvarMergeStart Then
        strR = strR & "\clmgf"
    End If
    If mvarMergeV Then
        strR = strR & "\clvmrg"
    End If
    If mvarMerge Then
        strR = strR & "\clmrg"
    End If
    
    '单元格内垂直对齐方式：
    If mvarAlignmentV = alv2_居中对齐 Then
        strR = strR & "\clvertalc" & vbCrLf
    ElseIf mvarAlignmentV = alv3_底端对齐 Then
        strR = strR & "\clvertalb" & vbCrLf
    Else
        strR = strR & "\clvertalt" & vbCrLf
    End If
    Dim strBorderStyle As String, strBorderWidth As String, strBorderColor As String
    Select Case mvarBorderStyle
    Case cbs01_无边框
        strBorderStyle = "\brdrnone"
    Case cbs02_单倍厚度边框
        strBorderStyle = "\brdrs"
    Case cbs03_双倍厚度边框
        strBorderStyle = "\brdrth"
    Case cbs04_阴影边框
        strBorderStyle = "\brdrsh"
    Case cbs05_双边框
        strBorderStyle = "\brdrdb"
    Case cbs06_点线边框
        strBorderStyle = "\brdrdot"
    Case cbs07_虚线边框
        strBorderStyle = "\brdrdash"
    Case cbs08_细线边框
        strBorderStyle = "\brdrhair"
    Case cbs09_插入式边框
        strBorderStyle = "\brdrinset"
    Case cbs10_虚线细边框
        strBorderStyle = "\brdrdashsm"
    Case cbs11_点划线边框
        strBorderStyle = "\brdrdashd"
    Case cbs12_点点划线边框
        strBorderStyle = "\brdrdashdd"
    Case cbs13_初始边框
        strBorderStyle = "\brdroutset"
    Case cbs14_三重边框
        strBorderStyle = "\brdrtriple"
    Case cbs15_厚薄叠加细边框
        strBorderStyle = "\brdrtnthsg"
    Case cbs16_薄厚叠加细边框
        strBorderStyle = "\brdrthtnsg"
    Case cbs17_薄厚薄叠加细边框
        strBorderStyle = "\brdrtnthtnsg"
    Case cbs18_厚薄叠加中边框
        strBorderStyle = "\brdrtnthmg"
    Case cbs19_薄厚叠加中边框
        strBorderStyle = "\brdrthtnmg"
    Case cbs20_薄厚薄叠加中边框
        strBorderStyle = "\brdrtnthtnmg"
    Case cbs21_厚薄叠加粗边框
        strBorderStyle = "\brdrtnthlg"
    Case cbs22_薄厚叠加粗边框
        strBorderStyle = "\brdrthtnlg"
    Case cbs23_薄厚薄叠加粗边框
        strBorderStyle = "\brdrtnthtnlg"
    Case cbs24_波浪线边框
        strBorderStyle = "\brdrwavy"
    Case cbs25_双波浪线边框
        strBorderStyle = "\brdrwavydb"
    Case cbs26_条纹边框
        strBorderStyle = "\brdrdashdotstr"
    Case cbs27_浮雕边框
        strBorderStyle = "\brdremboss"
    Case cbs28_雕刻边框
        strBorderStyle = "\brdrengrave"
    Case Else   '默认单边框
        strBorderStyle = "\brdrs"
    End Select
    
        
    '边框宽度
    If mvarBorderWidth <> 0 Then
        strBorderWidth = "\brdrw" & mvarBorderWidth
    Else    '默认宽度15
        strBorderWidth = "\brdrw15"
    End If
    
    If mvarBorderColor <> clr00_默认 Then
        strBorderColor = "\brdrcf" & mvarBorderColor
    End If
    
    '生成单元格边框定义语句
    If mvarBorderTop Then
        strR = strR & "\clbrdrt" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strR = strR & "\clbrdrt\brdrnone" & vbCrLf
    End If
    If mvarBorderLeft Then
        strR = strR & "\clbrdrl" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strR = strR & "\clbrdrl\brdrnone" & vbCrLf
    End If
    If mvarBorderBottom Then
        strR = strR & "\clbrdrb" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strR = strR & "\clbrdrb\brdrnone" & vbCrLf
    End If
    If mvarBorderRight Then
        strR = strR & "\clbrdrr" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strR = strR & "\clbrdrr\brdrnone" & vbCrLf
    End If
    If mvarBorderLeftDown Then
        strR = strR & "\cldglu" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    End If
    If mvarBorderLeftUp Then
        strR = strR & "\cldgll" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    End If
    
    Dim strCellPatType As String
    '背景图案类型
    Select Case mvarCellPatType
    Case ptt00_无
        strCellPatType = "\clbghoriz"
    Case ptt02_竖线
        strCellPatType = "\clbgvert"
    Case ptt03_下斜线
        strCellPatType = "\clbgfdiag"
    Case ptt04_上斜线
        strCellPatType = "\clbgbdiag"
    Case ptt05_网格线
        strCellPatType = "\clbgcross"
    Case ptt06_斜网格线
        strCellPatType = "\clbgdcross"
    Case ptt07_粗横线
        strCellPatType = "\clbgdkhor"
    Case ptt08_粗竖线
        strCellPatType = "\clbgdkvert"
    Case ptt09_粗下斜线
        strCellPatType = "\clbgdkfdiag"
    Case ptt10_粗上斜线
        strCellPatType = "\clbgdkbdiag"
    Case ptt11_粗网格线
        strCellPatType = "\clbgdkcross"
    Case ptt12_粗斜网格线
        strCellPatType = "\clbgdkdcross"
    Case Else
        strCellPatType = "\clshdrawnil"
    End Select
    
    '底纹前景色
    If mvarCellPatForeColor <> clr00_默认 Then strR = strR & "\clcfpat" & mvarCellPatForeColor
    '底纹背景色
    If mvarCellPatBackColor <> clr00_默认 Then strR = strR & "\clcbpat" & mvarCellPatBackColor
    '底纹类型
    If mvarCellPatType <> ptt00_无 Then
        strR = strR & strCellPatType
    End If
    '背景颜色百分比
    If mvarCellPatPercentage <> 0 Then
        strR = strR & "\clshdng" & mvarCellPatPercentage
    End If
    
    '文本流向（默认）
    strR = strR & "\cltxlrtb"
    
    '单元格宽度
    strR = strR & "\clftsWidth" & mvarWidthUnit
    strR = strR & "\clwWidth" & mvarWidth
    
    '底纹背景色（样式）
    If mvarCellPatBackColor <> clr00_默认 Then strR = strR & "\clcbpatraw" & mvarCellPatBackColor
    '底纹前景色（样式）
    If mvarCellPatForeColor <> clr00_默认 Then strR = strR & "\clcfpatraw" & mvarCellPatForeColor
    '底纹类型（样式），重复一遍
    If mvarCellPatType <> ptt00_无 Then
        strR = strR & strCellPatType
    End If
    '背景颜色百分比（样式）
    If mvarCellPatPercentage <> 0 Then
        strR = strR & "\clshdngraw" & mvarCellPatPercentage
    End If
    If mvarRight = 0 Then
        mvarRight = Col * 4000
    End If
    strR = strR & "\cellx" & mvarRight & vbCrLf
    
    GetCellDefineRTF = strR
End Function

Public Function GetCellContentRTF() As String
Attribute GetCellContentRTF.VB_Description = "获取单元格内容的RTF字符串。"
'获取单元格内容RTF
    Dim strR As String, strIsInTable As String
    If mvarNested Then strIsInTable = "\intbl"
    
    Dim strAlignment As String
    Select Case mvarAlignment
    Case alm2_居中对齐
        strAlignment = "\qc"
    Case alm3_居右对齐
        strAlignment = "\qr"
    Case alm4_两端对齐
        strAlignment = "\qj"
    Case alm5_分散对齐
        strAlignment = "\qd"
    Case Else
        strAlignment = "\ql"
    End Select
    
    strR = "\pard\plain " & strAlignment & "\li0\ri0\nowidctlpar" & strIsInTable & "\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs18\lang1033\langfe2052\kerning2\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 "
    Select Case mvarDataType
    Case cct1_段落
        strR = strR & mvarData.GetTextRTF()
    Case cct2_图片
        strR = strR & mvarData.GetTextRTF()
    Case cct3_表格
        strR = strR & mvarData.GetTextRTF()
    Case Else   '简单文本
        '简单文件的处理:单元格中的图片可以搭配简单文本。简单文本只提供几个简单属性
        Dim strFontSize As String
        If mvarFontSize <> 0 Then   '字体大小优先权：FontSize > FontSizeFixup > FontSizeDefault
            strFontSize = "\fs" & mvarFontSize * 2
        Else
            If mvarFontSizeFixup <> fs00_默认 Then
                strFontSize = "\fs" & mvarFontSizeFixup * 2
            Else
                strFontSize = "\fs18"   '单元格中默认大小为 9
            End If
        End If
        strR = strR & strFontSize
        If mvarBold Then
            strR = strR & "\b"
        Else
            strR = strR & "\b0"
        End If
        If mvarItalic Then
            strR = strR & "\i"
        Else
            strR = strR & "\i0"
        End If
        If mvarForeColor <> clr00_默认 Then
            strR = strR & "\cf" & mvarForeColor
        End If
        strR = strR & StrToASC(mvarText)
    End Select
    
    
    If Nested Then
        strR = strR & "\nestcell"
    Else
        strR = strR & "\cell"
    End If
    
    GetCellContentRTF = strR
End Function

Public Property Let Nested(ByVal vData As Boolean)
Attribute Nested.VB_Description = "是否嵌套。默认:False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Nested = 5
    mvarNested = vData
End Property


Public Property Get Nested() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Nested
    Nested = mvarNested
End Property


Public Property Let NoWrap(ByVal vData As Boolean)
Attribute NoWrap.VB_Description = "单元格不允许文本换行。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.NoWrap = 5
    mvarNoWrap = vData
End Property


Public Property Get NoWrap() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.NoWrap
    NoWrap = mvarNoWrap
End Property


Public Property Let FitText(ByVal vData As Boolean)
Attribute FitText.VB_Description = "文本自动适应单元格宽度。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FitText = 5
    mvarFitText = vData
End Property


Public Property Get FitText() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FitText
    FitText = mvarFitText
End Property



Public Property Let DataType(ByVal vData As CellContentTypeEnum)
Attribute DataType.VB_Description = "单元格内容类型。0-简单文本，1-高级文本，2-图片，3-嵌套表格"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.DataType = 5
    mvarDataType = vData
End Property


Public Property Get DataType() As CellContentTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.DataType
    DataType = mvarDataType
End Property



Public Property Let Data(ByVal vData As Variant)
Attribute Data.VB_Description = "单元格内容数据。与DataType相关联，DataType取值可以是：0-简单文本，1-高级文本，2-图片，3-嵌套表格"
Attribute Data.VB_UserMemId = 0
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Data = Form1
    Select Case mvarDataType
    Case cct1_段落
        Set mvarData = New cRTFText
        Set mvarData = vData
        mvarData.IsInTable = True
    Case cct2_图片
        Set mvarData = New cRTFPicture
        Set mvarData = vData
        mvarData.IsInTable = True
    Case cct3_表格
        Set mvarData = New cRTFTable
        Set mvarData = vData
        mvarData.Nested = True
    Case Else   '否则为简单文本
        mvarData = vData
        mvarText = vData
    End Select
End Property


Public Property Set Data(ByVal vData As Variant)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Data = Form1
    Select Case mvarDataType
    Case cct1_段落
        Set mvarData = New cRTFText
        Set mvarData = vData
        mvarData.IsInTable = True
    Case cct2_图片
        Set mvarData = New cRTFPicture
        Set mvarData = vData
        mvarData.IsInTable = True
    Case cct3_表格
        Set mvarData = New cRTFTable
        Set mvarData = vData
        mvarData.Nested = True
    Case Else   '否则为简单文本
        mvarData = vData
        mvarText = vData
    End Select
End Property


Public Property Get Data() As Variant
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Data
    If IsObject(mvarData) Then
        Set Data = mvarData
    Else
        Data = mvarData
    End If
End Property



Public Property Let Col(ByVal vData As Long)
Attribute Col.VB_Description = "当前的列值。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Col = 5
    mvarCol = vData
End Property


Public Property Get Col() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Col
    Col = mvarCol
End Property



Public Property Let Row(ByVal vData As Long)
Attribute Row.VB_Description = "行数。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Row = 5
    mvarRow = vData
End Property


Public Property Get Row() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Row
    Row = mvarRow
End Property



Public Property Let Merge(ByVal vData As Boolean)
Attribute Merge.VB_Description = "水平与前一单元格合并。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Merge = 5
    mvarMerge = vData
End Property


Public Property Get Merge() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Merge
    Merge = mvarMerge
End Property



Public Property Let MergeStart(ByVal vData As Boolean)
Attribute MergeStart.VB_Description = "水平合并的第一个单元格。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.MergeStart = 5
    mvarMergeStart = vData
End Property


Public Property Get MergeStart() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.MergeStart
    MergeStart = mvarMergeStart
End Property



Public Property Let MergeV(ByVal vData As Boolean)
Attribute MergeV.VB_Description = "垂直与前一单元格合并。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.MergeV = 5
    mvarMergeV = vData
End Property


Public Property Get MergeV() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.MergeV
    MergeV = mvarMergeV
End Property



Public Property Let MergeVStart(ByVal vData As Boolean)
Attribute MergeVStart.VB_Description = "垂直合并的第一个单元格。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.MergeVStart = 5
    mvarMergeVStart = vData
End Property


Public Property Get MergeVStart() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.MergeVStart
    MergeVStart = mvarMergeVStart
End Property



Public Property Let Right(ByVal vData As Long)
Attribute Right.VB_Description = "单元格右边界位置。CellX"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Right = 5
    mvarRight = vData
End Property


Public Property Get Right() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Right
    Right = mvarRight
End Property



Public Property Let WidthUnit(ByVal vData As CellLengthUnitEnum)
Attribute WidthUnit.VB_Description = "单元格首选宽度单位，0～3，默认：3"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.WidthUnit = 5
    mvarWidthUnit = vData
End Property


Public Property Get WidthUnit() As CellLengthUnitEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.WidthUnit
    WidthUnit = mvarWidthUnit
End Property



Public Property Let Width(ByVal vData As Long)
Attribute Width.VB_Description = "单元格首选宽度。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Width = 5
    mvarWidth = vData
End Property


Public Property Get Width() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Width
    Width = mvarWidth
End Property



Public Property Let CellPatForeColor(ByVal vData As ColorEnum)
Attribute CellPatForeColor.VB_Description = "单元格底纹前景色索引值。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CellPatForeColor = 5
    mvarCellPatForeColor = vData
End Property


Public Property Get CellPatForeColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CellPatForeColor
    CellPatForeColor = mvarCellPatForeColor
End Property



Public Property Let CellPatBackColor(ByVal vData As ColorEnum)
Attribute CellPatBackColor.VB_Description = "单元格底纹线条颜色索引值。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CellPatBackColor = 5
    mvarCellPatBackColor = vData
End Property


Public Property Get CellPatBackColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CellPatBackColor
    CellPatBackColor = mvarCellPatBackColor
End Property



Public Property Let CellPatType(ByVal vData As PatTypeEnum)
Attribute CellPatType.VB_Description = "单元格底纹类型。1～12。默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CellPatType = 5
    mvarCellPatType = vData
End Property


Public Property Get CellPatType() As PatTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CellPatType
    CellPatType = mvarCellPatType
End Property



Public Property Let BorderColor(ByVal vData As ColorEnum)
Attribute BorderColor.VB_Description = "边框颜色索引值。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderColor = 5
    mvarBorderColor = vData
End Property


Public Property Get BorderColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderColor
    BorderColor = mvarBorderColor
End Property



Public Property Let BorderStyle(ByVal vData As BorderStyleEnum)
Attribute BorderStyle.VB_Description = "边框类型。1～28种类型。默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderStyle = 5
    mvarBorderStyle = vData
End Property


Public Property Get BorderStyle() As BorderStyleEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderStyle
    BorderStyle = mvarBorderStyle
End Property




Public Property Let BorderLeftUp(ByVal vData As Boolean)
Attribute BorderLeftUp.VB_Description = "是否显示单元格内左上斜线/。默认为:False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderLeftUp = 5
    mvarBorderLeftUp = vData
End Property


Public Property Get BorderLeftUp() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderLeftUp
    BorderLeftUp = mvarBorderLeftUp
End Property



Public Property Let BorderLeftDown(ByVal vData As Boolean)
Attribute BorderLeftDown.VB_Description = "是否显示单元格内左下斜线\\。默认为:False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderLeftDown = 5
    mvarBorderLeftDown = vData
End Property


Public Property Get BorderLeftDown() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderLeftDown
    BorderLeftDown = mvarBorderLeftDown
End Property



Public Property Let BorderBottom(ByVal vData As Boolean)
Attribute BorderBottom.VB_Description = "是否显示单元格下边框。默认：True"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderBottom = 5
    mvarBorderBottom = vData
End Property


Public Property Get BorderBottom() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderBottom
    BorderBottom = mvarBorderBottom
End Property



Public Property Let BorderTop(ByVal vData As Boolean)
Attribute BorderTop.VB_Description = "是否显示单元格上边框。默认：True"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderTop = 5
    mvarBorderTop = vData
End Property


Public Property Get BorderTop() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderTop
    BorderTop = mvarBorderTop
End Property



Public Property Let BorderRight(ByVal vData As Boolean)
Attribute BorderRight.VB_Description = "是否显示单元格右边框。默认：True"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderRight = 5
    mvarBorderRight = vData
End Property


Public Property Get BorderRight() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderRight
    BorderRight = mvarBorderRight
End Property



Public Property Let BorderLeft(ByVal vData As Boolean)
Attribute BorderLeft.VB_Description = "是否显示单元格左边框。默认：True"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderLeft = 5
    mvarBorderLeft = vData
End Property


Public Property Get BorderLeft() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderLeft
    BorderLeft = mvarBorderLeft
End Property

Public Property Let Alignment(ByVal vData As AlignTypeEnum)
Attribute Alignment.VB_Description = "单元格水平对齐。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Alignment = 5
    If mvarDataType = cct2_图片 Then
        mvarData.Alignment = vData
    ElseIf mvarDataType = cct3_表格 Then
        '表格对齐方式：
        'mvarData.Alignment = vData
    Else    '默认为文本方式：
        If IsObject(mvarData) Then mvarData.Alignment = vData
    End If
    mvarAlignment = vData
End Property

Public Property Get Alignment() As AlignTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Alignment
    Alignment = mvarAlignment
End Property


Public Property Let AlignmentV(ByVal vData As AlignVTypeEnum)
Attribute AlignmentV.VB_Description = "单元格垂直对齐。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.AlignmentV = 5
    mvarAlignmentV = vData
End Property


Public Property Get AlignmentV() As AlignVTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.AlignmentV
    AlignmentV = mvarAlignmentV
End Property

