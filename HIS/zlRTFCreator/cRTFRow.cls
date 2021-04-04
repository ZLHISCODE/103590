VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFRow"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarRow As Long             '行数
Private mvarRowBand As Long         '行数，参考标题行。标题行为-1
Private mvarIsLastRow As Boolean    '是否是最末一行
Private mvarCellSpace As Long       '单元格间距
Private mvarWidth As Long                           '默认行宽
Private mvarWidthUnit As CellLengthUnitEnum         '默认行宽单位，1～3，默认：0
Private mvarWidthB As Long                          '行前不可见单元格宽度
Private mvarWidthBUnit As CellLengthUnitEnum        '行前不可见单元格宽度单位，1～3，默认：0
Private mvarWidthA As Long                          '行末不可见单元格宽度
Private mvarWidthAUnit As CellLengthUnitEnum        '行末不可见单元格宽度单位，1～3，默认：0
Private mvarPatType As PatTypeEnum                         '底纹类型，1～12，默认：0
Private mvarPatForeColor As ColorEnum               '底纹前景色
Private mvarPatBackColor As ColorEnum               '底纹背景色
Private mvarPatPercentage As Long   '底纹明暗百分比
Private mvarAutoFit As Boolean      '对于行是否开启自动适应。
Private mvarIsHeadRow As Boolean    '是否标题行，该行在每页顶端显示，trhdr
Private mvarKeep As Boolean         '保持该行与前一行同页。默认：False
Private mvarKeepFollow As Boolean   '保持该行与后一行同页。默认：False
Private mvarAlignment As RowAlignTypeEnum       '行对齐方式，包含其中的单元格。1～3，默认0
Private mvarHeight As Long          '行高
Private mvarPadLeft As Long         '默认行中单元格左边距
Private mvarPadRight As Long        '默认行中单元格右边距
Private mvarPadTop As Long          '默认行中单元格上边距
Private mvarPadBottom As Long       '默认行中单元格下边距
Private mvarBorderLeft As Boolean   '行左边框，默认:True
Private mvarBorderRight As Boolean  '行右边框，默认:True
Private mvarBorderTop As Boolean    '行上边框，默认:True
Private mvarBorderBottom As Boolean '行下边框，默认:True
Private mvarBorderH As Boolean      '行水平内部边框，默认:True
Private mvarBorderV As Boolean      '行垂直内部边框，默认:True
Private mvarBorderStyle As BorderStyleEnum  '行边框类型
Private mvarBorderColor As ColorEnum            '行边框颜色
Private mvarNested As Boolean       '是否嵌套，默认:False
Private mvarCells As cRTFCells      '单元格集合
Private mvarBorderWidth As Long
Private mvarColCount As Long

Public Function LastCell() As cRTFCell
Attribute LastCell.VB_Description = "获取最后添加的单元格对象。"
    Set LastCell = mvarCells(mvarCells.Count)
End Function

Public Property Let ColCount(ByVal vData As Long)
Attribute ColCount.VB_Description = "列总数。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ColCount = 5
    mvarColCount = vData
End Property

Public Property Get ColCount() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ColCount
    ColCount = mvarCells.Count
End Property

Public Function AddCell(Optional Row As Long, Optional Col As Long, _
Optional DataType As CellContentTypeEnum = cct0_文本, _
Optional Right As Long, _
Optional Width As Long = 0, _
Optional WidthUnit As CellLengthUnitEnum = clt0_默认, _
Optional sKey As String, _
Optional AlignmentV As AlignVTypeEnum = alv0_默认对齐, _
Optional Alignment As AlignTypeEnum = alm0_默认对齐, Optional BorderLeft As Boolean = True, _
Optional BorderRight As Boolean = True, Optional BorderTop As Boolean = True, _
Optional BorderBottom As Boolean = True, Optional BorderLeftDown As Boolean = False, _
Optional BorderLeftUp As Boolean = False, Optional BorderStyle As BorderStyleEnum = cbs00_默认边框样式, _
Optional BorderColor As ColorEnum = clr00_默认, Optional CellPatType As PatTypeEnum = ptt00_无, _
Optional CellPatBackColor As ColorEnum = clr00_默认, Optional CellPatForeColor As ColorEnum = clr00_默认, _
Optional MergeVStart As Boolean = False, Optional MergeV As Boolean = False, _
Optional MergeStart As Boolean = False, Optional Merge As Boolean = False, _
Optional FitText As Boolean = True, Optional NoWrap As Boolean = False, _
Optional Nested As Boolean = False, Optional BorderWidth As Long = 0, _
Optional CellPatPercentage As Long = 100) As Long
Attribute AddCell.VB_Description = "添加一个单元格。"
    
    Dim i As Long

    i = mvarCells.Add(Row, Col, DataType, Right, Width, WidthUnit, sKey, AlignmentV, Alignment, BorderLeft, BorderRight, _
        BorderTop, BorderBottom, BorderLeftDown, BorderLeftUp, BorderStyle, BorderColor, CellPatType, _
        CellPatBackColor, CellPatForeColor, MergeVStart, MergeV, MergeStart, Merge, FitText, _
        NoWrap, Nested, BorderWidth, CellPatPercentage)
    mvarCells(i).Row = mvarRow
    mvarCells(i).Col = mvarCells.Count
    mvarColCount = mvarCells.Count
    mvarRow = Row
    AddCell = i
End Function


Public Property Let BorderWidth(ByVal vData As Long)
Attribute BorderWidth.VB_Description = "默认边框宽度。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderWidth = 5
    mvarBorderWidth = vData
End Property

Public Property Get BorderWidth() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderWidth
    BorderWidth = mvarBorderWidth
End Property


Public Property Let BorderColor(ByVal vData As ColorEnum)
Attribute BorderColor.VB_Description = "边框颜色。"
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.BorderColor = Form1
    mvarBorderColor = vData
End Property


Public Property Get BorderColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderColor
    BorderColor = mvarBorderColor
End Property


Public Property Let BorderStyle(ByVal vData As BorderStyleEnum)
Attribute BorderStyle.VB_Description = "边框样式。"
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.BorderStyle = Form1
    mvarBorderStyle = vData
End Property


Public Property Get BorderStyle() As BorderStyleEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderStyle
    BorderStyle = mvarBorderStyle
End Property




Public Function GetTextRTF() As String
Attribute GetTextRTF.VB_Description = "获取行的RTF字符串。"
    Dim strCellDefine As String, strCellData As String, strR As String, strRowDefine As String
    Dim i As Long, j As Long
    
    '行定义字符串
    strRowDefine = "\pard\plain \trowd\irow" & mvarRow
    '如果含标题行
    strRowDefine = strRowDefine & "\irowband" & mvarRow
    
    '是否最后一行
    If mvarIsLastRow Then strRowDefine = strRowDefine & "\lastrow"
    
    '默认采用样式ts15
    strRowDefine = strRowDefine & "\ts15"
    '行对齐方式
    Select Case mvarAlignment
    Case ral2_居中对齐
        strRowDefine = strRowDefine & "\trqc"
    Case ral3_居右对齐
        strRowDefine = strRowDefine & "\trqr"
    Case Else
        strRowDefine = strRowDefine & "\trql"
    End Select
    '单元格间隔
    If mvarCellSpace <> 0 Then strRowDefine = strRowDefine & "\trgaph" & mvarCellSpace
    '行高
    If mvarHeight <> 0 Then strRowDefine = strRowDefine & "\trrh" & mvarHeight
    '行最左边位置，默认为-108
    strRowDefine = strRowDefine & "\trleft-108"
    Dim strBorderStyle As String, strBorderWidth As String, strBorderColor As String
    '边框样式
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
    '边框颜色
    If mvarBorderColor <> clr00_默认 Then
        strBorderColor = "\brdrcf" & mvarBorderColor
    End If
    '生成行边框定义语句
    If mvarBorderTop Then
        strRowDefine = strRowDefine & "\trbrdrt" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrt\brdrnone" & vbCrLf
    End If
    If mvarBorderLeft Then
        strRowDefine = strRowDefine & "\trbrdrl" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrl\brdrnone" & vbCrLf
    End If
    If mvarBorderBottom Then
        strRowDefine = strRowDefine & "\trbrdrb" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrb\brdrnone" & vbCrLf
    End If
    If mvarBorderRight Then
        strRowDefine = strRowDefine & "\trbrdrr" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrr\brdrnone" & vbCrLf
    End If
    If mvarBorderH Then
        strRowDefine = strRowDefine & "\trbrdrh" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrh\brdrnone" & vbCrLf
    End If
    If mvarBorderV Then
        strRowDefine = strRowDefine & "\trbrdrv" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrv\brdrnone" & vbCrLf
    End If
    
    '单元格宽度
    strRowDefine = strRowDefine & "\trftsWidth" & mvarWidthUnit
    If mvarWidth <> 0 Then strRowDefine = strRowDefine & "\trwWidth" & mvarWidth
    strRowDefine = strRowDefine & "\trftsWidthA" & mvarWidthUnit
    If mvarWidthA <> 0 Then strRowDefine = strRowDefine & "\trwWidthA" & mvarWidthA
    strRowDefine = strRowDefine & "\trftsWidthB" & mvarWidthUnit
    If mvarWidthB <> 0 Then strRowDefine = strRowDefine & "\trwWidthB" & mvarWidthB
    '行自动适应
    If mvarAutoFit Then
        strRowDefine = strRowDefine & "\trautofit1"
    Else
        strRowDefine = strRowDefine & "\trautofit0"
    End If

    If mvarKeep Then strRowDefine = strRowDefine & "\trkeep"
    If mvarKeepFollow Then strRowDefine = strRowDefine & "\trkeepfollow"
    If mvarPadBottom <> 0 Then strRowDefine = strRowDefine & "\trpaddb" & mvarPadBottom
    If mvarPadLeft <> 0 Then strRowDefine = strRowDefine & "\trpaddl" & mvarPadLeft
    If mvarPadRight <> 0 Then strRowDefine = strRowDefine & "\trpaddr" & mvarPadRight
    If mvarPadTop <> 0 Then strRowDefine = strRowDefine & "\trpaddt" & mvarPadTop
    '间隔单位默认均为：缇，不允许重叠
    strRowDefine = strRowDefine & "\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tabsnoovrlp1"
    
    '背景图案类型，通过Cell的背景来实现
    For i = 1 To mvarCells.Count
        If mvarPatType <> ptt00_无 Then mvarCells(i).CellPatType = mvarPatType
        If mvarPatForeColor <> ptt00_无 Then mvarCells(i).CellPatForeColor = mvarPatForeColor
        If mvarPatBackColor <> ptt00_无 Then mvarCells(i).CellPatBackColor = mvarPatBackColor
        If mvarPatPercentage <> ptt00_无 Then mvarCells(i).CellPatPercentage = mvarPatPercentage
    Next
    
    strRowDefine = strRowDefine & "\tbllkhdrrows\tbllklastrow\tbllkhdrcols\tbllklastcol"
    
    strCellDefine = ""
    For i = 1 To mvarCells.Count
        '取每个单元的数据
        strCellDefine = strCellDefine & mvarCells(i).GetCellDefineRTF
    Next
    strCellData = ""
    For i = 1 To mvarCells.Count
        strCellData = strCellData & mvarCells(i).GetCellContentRTF
    Next
    
    strR = strRowDefine + strCellDefine + strCellData + strRowDefine + strCellDefine '采用<define>+<data>+<define>格式
    
    If Nested Then
        strR = strR & "\nestrow"
    Else
        strR = strR & "\row"    '嵌套表行结束
    End If
    GetTextRTF = strR
End Function

Public Property Set Cells(ByVal vData As cRTFCells)
Attribute Cells.VB_Description = "当前行的所有单元格的集合。"
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Cells = Form1
    Set mvarCells = vData
End Property


Public Property Get Cells() As cRTFCells
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Cells
    Set Cells = mvarCells
End Property


Public Property Let BorderV(ByVal vData As Boolean)
Attribute BorderV.VB_Description = "是否显示行内垂直边框。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderV = 5
    mvarBorderV = vData
End Property


Public Property Get BorderV() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderV
    BorderV = mvarBorderV
End Property



Public Property Let BorderH(ByVal vData As Boolean)
Attribute BorderH.VB_Description = "是否显示行内水平边框。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderH = 5
    mvarBorderH = vData
End Property


Public Property Get BorderH() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderH
    BorderH = mvarBorderH
End Property



Public Property Let BorderBottom(ByVal vData As Boolean)
Attribute BorderBottom.VB_Description = "是否显示行下边框。"
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
Attribute BorderTop.VB_Description = "是否显示行上边框。"
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
Attribute BorderRight.VB_Description = "是否显示行右边框。"
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
Attribute BorderLeft.VB_Description = "是否显示行左边框。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BorderLeft = 5
    mvarBorderLeft = vData
End Property


Public Property Get BorderLeft() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BorderLeft
    BorderLeft = mvarBorderLeft
End Property

Public Property Let Nested(ByVal vData As Boolean)
Attribute Nested.VB_Description = "是否是嵌套行，仅用于嵌套表。默认:False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Nested = 5
    Dim i As Long
    For i = 1 To mvarCells.Count
        mvarCells(i).Nested = vData
    Next
    mvarNested = vData
End Property

Public Property Get Nested() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Nested
    Nested = mvarNested
End Property


Public Property Let PadBottom(ByVal vData As Long)
Attribute PadBottom.VB_Description = "默认行中单元格下边距。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PatBottom = 5
    mvarPadBottom = vData
End Property


Public Property Get PadBottom() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PatBottom
    PadBottom = mvarPadBottom
End Property


Public Property Let PadTop(ByVal vData As Long)
Attribute PadTop.VB_Description = "默认行中单元格上边距。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PatTop = 5
    mvarPadTop = vData
End Property


Public Property Get PadTop() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PatTop
    PadTop = mvarPadTop
End Property


Public Property Let PadRight(ByVal vData As Long)
Attribute PadRight.VB_Description = "默认行中单元格右边距。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PadRight = 5
    mvarPadRight = vData
End Property


Public Property Get PadRight() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PadRight
    PadRight = mvarPadRight
End Property



Public Property Let PadLeft(ByVal vData As Long)
Attribute PadLeft.VB_Description = "默认行中单元格左边距。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PadLeft = 5
    mvarPadLeft = vData
End Property


Public Property Get PadLeft() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PadLeft
    PadLeft = mvarPadLeft
End Property



Public Property Let Height(ByVal vData As Long)
Attribute Height.VB_Description = "行高。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Height = 5
    mvarHeight = vData
End Property


Public Property Get Height() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Height
    Height = mvarHeight
End Property



Public Property Let Alignment(ByVal vData As RowAlignTypeEnum)
Attribute Alignment.VB_Description = "行对齐方式，包含其中单元格。1～3，默认:0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Alignment = 5
    mvarAlignment = vData
End Property


Public Property Get Alignment() As RowAlignTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Alignment
    Alignment = mvarAlignment
End Property



Public Property Let KeepFollow(ByVal vData As Boolean)
Attribute KeepFollow.VB_Description = "保持该行与后一行同页。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.KeepFollow = 5
    mvarKeepFollow = vData
End Property


Public Property Get KeepFollow() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.KeepFollow
    KeepFollow = mvarKeepFollow
End Property



Public Property Let Keep(ByVal vData As Boolean)
Attribute Keep.VB_Description = "保持该行与前一行同页。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Keep = 5
    mvarKeep = vData
End Property


Public Property Get Keep() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Keep
    Keep = mvarKeep
End Property



Public Property Let IsHeadRow(ByVal vData As Boolean)
Attribute IsHeadRow.VB_Description = "是否是标题行，该行将在每页顶端显示。默认:False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.IsHeadRow = 5
    mvarIsHeadRow = vData
End Property


Public Property Get IsHeadRow() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.IsHeadRow
    IsHeadRow = mvarIsHeadRow
End Property



Public Property Let AutoFit(ByVal vData As Boolean)
Attribute AutoFit.VB_Description = "对于行是否开启自动适应。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.AutoFit = 5
    mvarAutoFit = vData
End Property


Public Property Get AutoFit() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.AutoFit
    AutoFit = mvarAutoFit
End Property



Public Property Let PatPercentage(ByVal vData As Long)
Attribute PatPercentage.VB_Description = "底纹明暗百分比。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PatPercentage = 5
    mvarPatPercentage = vData
End Property


Public Property Get PatPercentage() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PatPercentage
    PatPercentage = mvarPatPercentage
End Property



Public Property Let PatBackColor(ByVal vData As ColorEnum)
Attribute PatBackColor.VB_Description = "底纹背景色索引值。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PatBackColor = 5
    mvarPatBackColor = vData
End Property


Public Property Get PatBackColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PatBackColor
    PatBackColor = mvarPatBackColor
End Property



Public Property Let PatForeColor(ByVal vData As ColorEnum)
Attribute PatForeColor.VB_Description = "底纹前景色索引值。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PatForeColor = 5
    mvarPatForeColor = vData
End Property


Public Property Get PatForeColor() As ColorEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PatForeColor
    PatForeColor = mvarPatForeColor
End Property



Public Property Let PatType(ByVal vData As PatTypeEnum)
Attribute PatType.VB_Description = "底纹类型。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PatType = 5
    mvarPatType = vData
End Property


Public Property Get PatType() As PatTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PatType
    PatType = mvarPatType
End Property



Public Property Let WidthAUnit(ByVal vData As CellLengthUnitEnum)
Attribute WidthAUnit.VB_Description = "行末不可见单元格宽度单位，1～3。默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.WidthAUnit = 5
    mvarWidthAUnit = vData
End Property


Public Property Get WidthAUnit() As CellLengthUnitEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.WidthAUnit
    WidthAUnit = mvarWidthAUnit
End Property



Public Property Let WidthA(ByVal vData As Long)
Attribute WidthA.VB_Description = "行末不可见单元格宽度。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.WidthA = 5
    mvarWidthA = vData
End Property


Public Property Get WidthA() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.WidthA
    WidthA = mvarWidthA
End Property



Public Property Let WidthBUnit(ByVal vData As CellLengthUnitEnum)
Attribute WidthBUnit.VB_Description = "行前不可见单元格宽度单位，1～3。默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.WidthBUnit = 5
    mvarWidthBUnit = vData
End Property


Public Property Get WidthBUnit() As CellLengthUnitEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.WidthBUnit
    WidthBUnit = mvarWidthBUnit
End Property



Public Property Let WidthB(ByVal vData As Long)
Attribute WidthB.VB_Description = "行前不可见单元格宽度。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.WidthB = 5
    mvarWidthB = vData
End Property


Public Property Get WidthB() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.WidthB
    WidthB = mvarWidthB
End Property



Public Property Let WidthUnit(ByVal vData As CellLengthUnitEnum)
Attribute WidthUnit.VB_Description = "默认行宽单位，1～3，默认：0"
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
Attribute Width.VB_Description = "默认行宽。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Width = 5
    mvarWidth = vData
End Property


Public Property Get Width() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Width
    Width = mvarWidth
End Property



Public Property Let CellSpace(ByVal vData As Long)
Attribute CellSpace.VB_Description = "单元格间距。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CellSpace = 5
    mvarCellSpace = vData
End Property


Public Property Get CellSpace() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CellSpace
    CellSpace = mvarCellSpace
End Property



Public Property Let IsLastRow(ByVal vData As Boolean)
Attribute IsLastRow.VB_Description = "是否是最末一行。默认：False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.IsLastRow = 5
    mvarIsLastRow = vData
End Property


Public Property Get IsLastRow() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.IsLastRow
    IsLastRow = mvarIsLastRow
End Property



Public Property Let RowBand(ByVal vData As Long)
Attribute RowBand.VB_Description = "行数，参考标题行。标题行为-1"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.RowBand = 5
    mvarRowBand = vData
End Property


Public Property Get RowBand() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.RowBand
    RowBand = mvarRowBand
End Property



Public Property Let Row(ByVal vData As Long)
Attribute Row.VB_Description = "当前行值。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Row = 5
    mvarRow = vData
End Property


Public Property Get Row() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Row
    Row = mvarRow
End Property

Private Sub Class_Initialize()
    Set mvarCells = New cRTFCells
End Sub

Private Sub Class_Terminate()
    Set mvarCells = Nothing
End Sub
