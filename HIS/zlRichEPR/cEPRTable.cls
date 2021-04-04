VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum TableTypeEnum
    tte_默认 = 0
    tte_医嘱项目组 = 1
    tte_报告图片组 = 2
End Enum

Private mvarKey As Long                 '对象标记
Private mvarID As Double                  'ID
Private mvar文件ID As Long              '文件ID
Private mvar父ID As Double               '父ID
Private mvar对象序号 As Double            '序号
Private mvar保留对象 As Boolean         '保留对象
Private mvar开始版 As Long              '开始版，专用于“电子病历记录”
Private mvar终止版 As Long              '终止版，专用于“电子病历记录”
Private mvar是否换行 As Boolean         '是否换行
Private mvar预制提纲ID As Long          '预制提纲ID，用于特殊预制提纲中表格处理

Private mvarRows As Long                '表格总行数
Private mvarCols As Long                '表格总列数
Private mvarSingleLine As Boolean               '是否单行显示，默认为False
Private mvarAlternateRowBackColor As OLE_COLOR  '交替色（隔行显示不同颜色）
Private mvarBackColor As OLE_COLOR              '背景色
Private mvarGridLineColor As OLE_COLOR          '网格线颜色，默认为黑色
Private mvarGridLineWidth As Long               '网格线宽度，默认为1
Private mvarBorderColor As OLE_COLOR            '边框颜色，默认为黑色
Private mvarBorderWidth As Long                 '边框宽度，默认为0
Private mvarForeColor As OLE_COLOR              '前景色
Private mvarFontQuality As FontQualityEnum      '字体质量
Private mvarAutoHeight As Boolean               '自动高度，默认为True
Private mvarWordEllipsis As Boolean             '是否在文本无法显示完时显示一个省略号
Private mvarCellMargin As Long                  '单元格边距，默认为30
Private mvarCellIndent As Long                  '单元格缩进
Private mvarWidth As Long                       '表格宽度
Private mvarHeight As Long                      '表格高度
Private mvarColWidthString As String            '列宽度
Private mvarTableType As TableTypeEnum          '表格类型
Private mvarExtendTag As String                 '扩展Tag属性，用于记录附加属性
Private mvar标记 As String                      '标记：按表格”标记”搜索和处理，实现一些对电子病历的人工控制。标记为文本，存放到”内容文本”字段

Private mvarCells As cEPRCells          '单元格对象的集合
Private mvarPictures As cEPRPictures    '图片集合
Private mvarElements As cEPRElements    '诊治要素集合

Public DicSpecSQLs As New Dictionary    '附加SQL语句数组，用于特殊预制提纲的额外数据的处理！

Public ZoomPicture As Boolean           '是否进行图片缩放

Public Property Let SingleLine(ByVal vData As Boolean)
    mvarSingleLine = vData
End Property

Public Property Get SingleLine() As Boolean
    SingleLine = mvarSingleLine
End Property

Public Property Let AlternateRowBackColor(ByVal vData As OLE_COLOR)
    mvarAlternateRowBackColor = vData
End Property

Public Property Get AlternateRowBackColor() As OLE_COLOR
    AlternateRowBackColor = mvarAlternateRowBackColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    mvarBackColor = vData
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Let GridLineColor(ByVal vData As OLE_COLOR)
    mvarGridLineColor = vData
End Property

Public Property Get GridLineColor() As OLE_COLOR
    GridLineColor = mvarGridLineColor
End Property

Public Property Let GridLineWidth(ByVal vData As Long)
    mvarGridLineWidth = vData
End Property

Public Property Get GridLineWidth() As Long
    GridLineWidth = mvarGridLineWidth
End Property

Public Property Let BorderColor(ByVal vData As OLE_COLOR)
    mvarBorderColor = vData
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mvarBorderColor
End Property

Public Property Let BorderWidth(ByVal vData As Long)
    mvarBorderWidth = vData
End Property

Public Property Get BorderWidth() As Long
    BorderWidth = mvarBorderWidth
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    mvarForeColor = vData
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mvarForeColor
End Property

Public Property Let FontQuality(ByVal vData As FontQualityEnum)
    mvarFontQuality = vData
End Property

Public Property Get FontQuality() As FontQualityEnum
    FontQuality = mvarFontQuality
End Property

Public Property Let AutoHeight(ByVal vData As Boolean)
    mvarAutoHeight = vData
End Property

Public Property Get AutoHeight() As Boolean
    AutoHeight = mvarAutoHeight
End Property

Public Property Let WordEllipsis(ByVal vData As Boolean)
    mvarWordEllipsis = vData
End Property

Public Property Get WordEllipsis() As Boolean
    WordEllipsis = mvarWordEllipsis
End Property

Public Property Let CellMargin(ByVal vData As Long)
    mvarCellMargin = vData
End Property

Public Property Get CellMargin() As Long
    CellMargin = mvarCellMargin
End Property

Public Property Let CellIndent(ByVal vData As Long)
    mvarCellIndent = vData
End Property

Public Property Get CellIndent() As Long
    CellIndent = mvarCellIndent
End Property

Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let ColWidthString(ByVal vData As String)
    mvarColWidthString = vData
End Property

Public Property Get ColWidthString() As String
    ColWidthString = mvarColWidthString
End Property

Public Property Let TableType(ByVal vData As TableTypeEnum)
    mvarTableType = vData
End Property

Public Property Get TableType() As TableTypeEnum
    TableType = mvarTableType
End Property

Public Property Let ExtendTag(ByVal vData As String)
    mvarExtendTag = vData
End Property

Public Property Get ExtendTag() As String
    ExtendTag = mvarExtendTag
End Property

Public Property Let 标记(ByVal vData As String)
    mvar标记 = vData
End Property

Public Property Get 标记() As String
    标记 = mvar标记
End Property

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Double)
    mvarID = vData
End Property

Public Property Get ID() As Double
    ID = mvarID
End Property

Public Property Let 文件ID(ByVal vData As Long)
    mvar文件ID = vData
    Dim i As Long
    For i = 1 To Me.Pictures.Count
        Me.Pictures(i).文件ID = vData
    Next
End Property

Public Property Get 文件ID() As Long
    文件ID = mvar文件ID
End Property

Public Property Let 父ID(ByVal vData As Double)
    mvar父ID = vData
End Property

Public Property Get 父ID() As Double
    父ID = mvar父ID
End Property

Public Property Let 对象序号(ByVal vData As Double)
    mvar对象序号 = vData
End Property

Public Property Get 对象序号() As Double
    对象序号 = mvar对象序号
End Property

Public Property Let 保留对象(ByVal vData As Boolean)
    mvar保留对象 = vData
End Property

Public Property Get 保留对象() As Boolean
    保留对象 = mvar保留对象
End Property

Public Property Let 开始版(ByVal vData As Long)
    mvar开始版 = vData
End Property

Public Property Get 开始版() As Long
    开始版 = IIf(mvar开始版 = 0, 1, mvar开始版)
End Property

Public Property Let 终止版(ByVal vData As Long)
    mvar终止版 = vData
End Property

Public Property Get 终止版() As Long
'    终止版 = IIf(mvar终止版 = 0, 0, mvar终止版)
    终止版 = 0
End Property

Public Property Let 是否换行(ByVal vData As Boolean)
    mvar是否换行 = vData
End Property

Public Property Get 是否换行() As Boolean
    是否换行 = mvar是否换行
End Property

Public Property Let 预制提纲ID(ByVal vData As Long)
    mvar预制提纲ID = vData
End Property

Public Property Get 预制提纲ID() As Long
    预制提纲ID = mvar预制提纲ID
End Property

Public Property Let Rows(ByVal vData As Long)
    mvarRows = vData
End Property

Public Property Get Rows() As Long
    Rows = mvarRows
End Property

Public Property Let Cols(ByVal vData As Long)
    mvarCols = vData
End Property

Public Property Get Cols() As Long
    Cols = mvarCols
End Property

Public Property Set Cells(ByVal vData As cEPRCells)
    Set mvarCells = vData
End Property

Public Property Let Cells(ByVal vData As cEPRCells)
    Set mvarCells = vData
End Property

Public Property Get Cells() As cEPRCells
    Set Cells = mvarCells
End Property

Public Property Set Pictures(ByVal vData As cEPRPictures)
    Set mvarPictures = vData
End Property

Public Property Let Pictures(ByVal vData As cEPRPictures)
    Set mvarPictures = vData
End Property

Public Property Get Pictures() As cEPRPictures
    Set Pictures = mvarPictures
End Property

Public Property Set Elements(ByVal vData As cEPRElements)
    Set mvarElements = vData
End Property

Public Property Let Elements(ByVal vData As cEPRElements)
    Set mvarElements = vData
End Property

Public Property Get Elements() As cEPRElements
    Set Elements = mvarElements
End Property

'################################################################################################################
'## 功能：  获取表格属性字符串
'################################################################################################################
Public Property Get 对象属性() As String
'    Dim strR As String
'    strR = Me.Rows
'    strR = strR & ";" & Me.Cols
'    strR = strR & ";" & Me.FixedRows
'    strR = strR & ";" & Me.FixedCols
'    strR = strR & ";" & Me.ShowGrid
'    对象属性 = strR
    Dim strR As String
    
    strR = mvarRows
    strR = strR & ";" & mvarCols
    strR = strR & ";" & mvarWidth
    strR = strR & ";" & mvarHeight
    strR = strR & ";" & mvarColWidthString
    strR = strR & ";" & mvarSingleLine
    strR = strR & ";" & mvarAlternateRowBackColor
    strR = strR & ";" & mvarBackColor
    strR = strR & ";" & mvarGridLineColor
    strR = strR & ";" & mvarGridLineWidth
    strR = strR & ";" & mvarBorderColor
    strR = strR & ";" & mvarBorderWidth
    strR = strR & ";" & mvarForeColor
    strR = strR & ";" & mvarFontQuality
    strR = strR & ";" & mvarAutoHeight
    strR = strR & ";" & mvarWordEllipsis
    strR = strR & ";" & mvarCellMargin
    strR = strR & ";" & mvarCellIndent
    strR = strR & ";" & mvarTableType
    strR = strR & ";" & mvarExtendTag
    
    对象属性 = strR
End Property

'################################################################################################################
'## 功能：  通过一个字符串来设置表格属性
'################################################################################################################
Public Property Let 对象属性(vData As String)
    On Error Resume Next
'    Dim T As Variant, i As Long
'    T = Split(vData, ";")
'    If UBound(T) > 0 Then
'        Me.Rows = IIf(T(0) = "", 0, T(0))
'        Me.Cols = IIf(T(1) = "", 0, T(1))
'        Me.FixedRows = IIf(T(2) = "", 0, T(2))
'        Me.FixedCols = IIf(T(3) = "", 0, T(3))
'        Me.ShowGrid = IIf(T(4) = "", True, T(4))
'    End If
    On Error Resume Next
    Dim T As Variant, i As Long
    T = Split(vData, ";")
    If UBound(T) = 4 Then
        Me.Rows = IIf(T(0) = "", 0, T(0))
        Me.Cols = IIf(T(1) = "", 0, T(1))
        Me.GridLineWidth = 1
        Me.GridLineColor = vbBlack
        Me.AutoHeight = True
        Me.WordEllipsis = False
        Me.CellMargin = 10
        Me.SingleLine = False
    ElseIf UBound(T) = 17 Or UBound(T) = 18 Or UBound(T) = 19 Then
        Me.Rows = IIf(T(0) = "", 0, T(0))
        Me.Cols = IIf(T(1) = "", 0, T(1))
        Me.Width = IIf(T(2) = "", 0, T(2))
        Me.Height = IIf(T(3) = "", 0, T(3))
        Me.ColWidthString = IIf(T(4) = "", "", T(4))
        Me.SingleLine = IIf(T(5) = "", 0, T(5))
        Me.AlternateRowBackColor = IIf(T(6) = "", 0, T(6))
        Me.BackColor = IIf(T(7) = "", 0, T(7))
        Me.GridLineColor = IIf(T(8) = "", 0, T(8))
        Me.GridLineWidth = IIf(T(9) = "", 0, T(9))
        Me.BorderColor = IIf(T(10) = "", 0, T(10))
        Me.BorderWidth = IIf(T(11) = "", 0, T(11))
        Me.ForeColor = IIf(T(12) = "", 0, T(12))
        Me.FontQuality = IIf(T(13) = "", 0, T(13))
        Me.AutoHeight = IIf(T(14) = "", 0, T(14))
        Me.WordEllipsis = IIf(T(15) = "", 0, T(15))
        Me.CellMargin = IIf(T(16) = "", 0, T(16))
        Me.CellIndent = IIf(T(17) = "", 0, T(17))
        Me.TableType = IIf(T(18) = "", 0, T(18))
        Me.ExtendTag = IIf(T(19) = "", 0, T(19))
    End If
End Property

'################################################################################################################
'## 功能：  得到指定行、列的单元格对象
'##
'## 参数：  (R:行 ， C:列)
'################################################################################################################
Public Function Cell(ByVal r As Long, ByVal C As Long) As cEPRCell
    Dim i As Long
    For i = 1 To Me.Cells.Count
        If Me.Cells(i).Row = r And Me.Cells(i).Col = C Then
            Set Cell = Me.Cells(i)
            Exit For
        End If
    Next
End Function

'################################################################################################################
'## 功能：  复制表格对象
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cEPRTable
    Dim cNode As New cEPRTable
    
    If blnWithKeys Then cNode.Key = Me.Key
    cNode.ID = Me.ID
    cNode.文件ID = Me.文件ID
    cNode.开始版 = Me.开始版
    cNode.终止版 = Me.终止版
    cNode.父ID = Me.父ID
    cNode.对象序号 = Me.对象序号
    cNode.保留对象 = Me.保留对象
    cNode.是否换行 = Me.是否换行
    cNode.预制提纲ID = Me.预制提纲ID
    
    cNode.Rows = Me.Rows
    cNode.Cols = Me.Cols
    cNode.Rows = Me.Width
    cNode.Cols = Me.Height
    cNode.Cols = Me.ColWidthString
    cNode.SingleLine = Me.SingleLine
    cNode.AlternateRowBackColor = Me.AlternateRowBackColor
    cNode.BackColor = Me.BackColor
    cNode.GridLineColor = Me.GridLineColor
    cNode.GridLineWidth = Me.GridLineWidth
    cNode.BorderColor = Me.BorderColor
    cNode.BorderWidth = Me.BorderWidth
    cNode.ForeColor = Me.ForeColor
    cNode.FontQuality = Me.FontQuality
    cNode.AutoHeight = Me.AutoHeight
    cNode.WordEllipsis = Me.WordEllipsis
    cNode.CellMargin = Me.CellMargin
    cNode.CellIndent = Me.CellIndent
    cNode.TableType = Me.TableType
    cNode.ExtendTag = Me.ExtendTag
    cNode.标记 = Me.标记
    
    Set cNode.Pictures = Me.Pictures.Clone
    Set cNode.Cells = Me.Cells.Clone
    Set cNode.Elements = Me.Elements.Clone(True)
    Set cNode.DicSpecSQLs = Me.DicSpecSQLs
    
    Set Clone = cNode
End Function
Public Sub TerminategfrmPublic()
'目前体检专用，因为体检动态创建对象，不初化cRichEPR.Init
    On Error Resume Next
    If Not gfrmPublic Is Nothing Then Unload gfrmPublic
    Set gfrmPublic = Nothing
End Sub
'################################################################################################################
'## 功能：  对象初始化
'################################################################################################################
Private Sub Class_initialize()
    
     ZoomPicture = False
    If gfrmPublic Is Nothing Then Set gfrmPublic = New frmPublic
    Set DicSpecSQLs = New Dictionary
    Set Cells = New cEPRCells
    Set Pictures = New cEPRPictures
    Set Elements = New cEPRElements
End Sub

'################################################################################################################
'## 功能：  对象销毁
'################################################################################################################
Private Sub Class_Terminate()
    DicSpecSQLs.RemoveAll
    Set DicSpecSQLs = Nothing
    Set Cells = Nothing
    Set Pictures = Nothing
    Set Elements = Nothing
End Sub
Public Function FillTableMember(ByVal rs As ADODB.Recordset, ByVal strTableName As String) As Boolean
'功能：用数据记录填充类成员
    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    
    If NVL(rs!对象标记, 0) <> 0 Then Me.Key = NVL(rs!对象标记, 0)
    Me.ID = NVL(rs!ID, 0)
    Me.文件ID = NVL(rs!文件ID, 0)
    Me.父ID = NVL(rs!父ID, 0)
    Me.对象序号 = NVL(rs!对象序号, 0)
    Me.保留对象 = IIf(NVL(rs!保留对象, 0) = 0, False, True)
    Me.对象属性 = NVL(rs!对象属性)
    Me.是否换行 = IIf(NVL(rs!是否换行, 0) = 0, False, True)
    Me.预制提纲ID = NVL(rs!预制提纲ID, 0)
    Me.标记 = NVL(rs!内容文本, "")
    
    If InStr(strTableName, "电子病历内容") > 0 Then '病历文件结构和病历范文内容没有这两个字段,但都共用这个函数
        Me.开始版 = NVL(rs!开始版, 1)
        Me.终止版 = NVL(rs!终止版, 0)
    End If
    
    If GetTbChildMember(Me.文件ID, Me.ID, strTableName) Then Exit Function
    
    FillTableMember = True
End Function

Public Function GetTbChildMember(ByVal lFileId As Long, ByVal lID As Double, ByVal strTableName As String) As Boolean
'功能：读取表格成员的子成员
Dim rsTemp As ADODB.Recordset, lngType As Long, lngKey As Long


    gstrSQL = "Select Level, ID, 文件id," & IIf(InStr(strTableName, "电子病历内容") > 0, "开始版, 终止版,", "") & " 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id," & vbNewLine & _
                "       诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "From (Select ID, 文件id," & IIf(InStr(strTableName, "电子病历内容") > 0, "开始版, 终止版,", "") & " 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 诊治要素id," & vbNewLine & _
                "              替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "       From " & strTableName & vbNewLine & _
                "       Where 文件id = [1] " & IIf(InStr(strTableName, "电子病历内容") > 0, "And Nvl(终止版, 0) = 0)", ")") & vbNewLine & _
                "Start With 父id + 0 = [2]" & vbNewLine & _
                "Connect By Prior ID = 父id + 0" & vbNewLine & _
                "Order By 对象标记, 对象序号, 内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRTable", lFileId, lID)
    Do While Not rsTemp.EOF
        lngType = NVL(rsTemp!对象类型, 0)
        Select Case lngType
        Case 2  '文本
            '增加单元 Cell 到 Cells集合 中
            lngKey = Me.Cells.Add(NVL(rsTemp!对象标记, 0))
            Me.Cells("K" & lngKey).ID = NVL(rsTemp!ID, 0)
            Me.Cells("K" & lngKey).文件ID = NVL(rsTemp!文件ID, 0)
            Me.Cells("K" & lngKey).父ID = NVL(rsTemp!父ID, 0)
            Me.Cells("K" & lngKey).对象序号 = NVL(rsTemp!对象序号, 0)
            Me.Cells("K" & lngKey).内容文本 = NVL(rsTemp!内容文本, "")
            Me.Cells("K" & lngKey).保留对象 = IIf(NVL(rsTemp!保留对象, 0) = 0, False, True)
            Me.Cells("K" & lngKey).对象属性 = NVL(rsTemp!对象属性)
            If InStr(strTableName, "电子病历内容") > 0 Then
                Me.Cells("K" & lngKey).开始版 = NVL(rsTemp!开始版, 1)
                Me.Cells("K" & lngKey).终止版 = NVL(rsTemp!终止版, 0)
            End If
        Case 4  '诊治要素
            '先增加单元 Cell 到 Cells集合 中
            lngKey = Me.Cells.Add(NVL(rsTemp!对象标记, 0))
            Me.Cells("K" & lngKey).ID = NVL(rsTemp!ID, 0)
            Me.Cells("K" & lngKey).文件ID = NVL(rsTemp!文件ID, 0)
            Me.Cells("K" & lngKey).父ID = NVL(rsTemp!父ID, 0)
            Me.Cells("K" & lngKey).对象序号 = NVL(rsTemp!对象序号, 0)
            Me.Cells("K" & lngKey).内容文本 = NVL(rsTemp!内容文本, "")
            Me.Cells("K" & lngKey).保留对象 = IIf(NVL(rsTemp!保留对象, 0) = 0, False, True)
            Me.Cells("K" & lngKey).对象属性 = NVL(rsTemp!对象属性)
            If InStr(strTableName, "电子病历内容") > 0 Then
                Me.Cells("K" & lngKey).开始版 = NVL(rsTemp!开始版, 1)
                Me.Cells("K" & lngKey).终止版 = NVL(rsTemp!终止版, 0)
            End If
            '再增加诊治要素 Element 到 Elements集合 中
            lngKey = Me.Elements.Add(Me.Cells("K" & lngKey).ElementKey)
            If lngKey > 0 Then
                Me.Elements("K" & lngKey).ID = rsTemp!ID                     '本行的ID与Cell的ID相同，因为是同一条记录
                Me.Elements("K" & lngKey).文件ID = NVL(rsTemp!文件ID, 0)
                Me.Elements("K" & lngKey).父ID = NVL(rsTemp!父ID, 0)
                Me.Elements("K" & lngKey).对象序号 = NVL(rsTemp!对象序号, 0)
                Me.Elements("K" & lngKey).保留对象 = IIf(NVL(rsTemp!保留对象, 0) = 0, False, True)
                Me.Elements("K" & lngKey).内容文本 = NVL(rsTemp!内容文本, "")
                Me.Elements("K" & lngKey).是否换行 = IIf(NVL(rsTemp!是否换行, 0) = 0, False, True)
                Me.Elements("K" & lngKey).诊治要素ID = NVL(rsTemp!诊治要素ID, 0)
                Me.Elements("K" & lngKey).替换域 = NVL(rsTemp!替换域, 0)
                Me.Elements("K" & lngKey).要素名称 = NVL(rsTemp!要素名称)
                Me.Elements("K" & lngKey).要素类型 = NVL(rsTemp!要素类型, 0)
                Me.Elements("K" & lngKey).要素长度 = NVL(rsTemp!要素长度, 0)
                Me.Elements("K" & lngKey).要素小数 = NVL(rsTemp!要素小数, 0)
                Me.Elements("K" & lngKey).要素单位 = NVL(rsTemp!要素单位)
                Me.Elements("K" & lngKey).要素表示 = NVL(rsTemp!要素表示, 0)
                Me.Elements("K" & lngKey).输入形态 = NVL(rsTemp!输入形态, 0)
                Me.Elements("K" & lngKey).要素值域 = NVL(rsTemp!要素值域)
                If InStr(strTableName, "电子病历内容") > 0 Then
                    Me.Elements("K" & lngKey).开始版 = NVL(rsTemp!开始版, 1)
                    Me.Elements("K" & lngKey).终止版 = NVL(rsTemp!终止版, 0)
                End If
            End If
        Case 5  '图片
            '增加图片 Picture 到 Pictures集合中
            lngKey = Me.Pictures.Add(NVL(rsTemp!对象标记, 0))
            Call Me.Pictures("K" & lngKey).FillPictureMember(rsTemp, strTableName)
        End Select
        
        rsTemp.MoveNext
    Loop
    GetTbChildMember = True
End Function
'################################################################################################################
'## 功能：  从数据库直接获取表格内容
'##
'## 参数：  eEdtType    :当前编辑方式；
'##         lngFileID   :文件ID
'##         lngID       :数据库中的对象ID号
'##         blnWithKey  :是否包含Key值一起恢复
'##         strTableName:表名，如果省略，则根据当前的eEdtType来判断表名
'################################################################################################################
Public Function GetTableFromDB(ByVal eEdtType As EditTypeEnum, _
    ByVal lngFileID As Long, _
    ByVal lngID As Double, _
    Optional blnWithKey As Boolean = False, _
    Optional ByVal strTableName As String = "") As Boolean
    
    Dim rs As New ADODB.Recordset, RsSub As New ADODB.Recordset, strTmp As String, lngType As Long, lngKey As Long
    If strTableName = "" Then
        Select Case eEdtType
        Case cprET_病历文件定义
            strTableName = "病历文件结构"
        Case cprET_全文示范编辑
            strTableName = "病历范文内容"
        Case cprET_单病历编辑, cprET_单病历审核
            strTableName = "电子病历内容"
        End Select
    End If
    
    gstrSQL = "Select ID, 文件id," & IIf(InStr(strTableName, "电子病历内容") > 0, "开始版, 终止版, ", "") & "父id, 对象序号,对象标记, 保留对象, 对象属性, 内容文本, 是否换行, 预制提纲id " & vbNewLine & _
            "From " & strTableName & vbNewLine & _
            "Where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRTable", lngID)
    If rs.EOF Then GetTableFromDB = False: Exit Function
    
    If blnWithKey And NVL(rs("对象标记"), 0) <> 0 Then Me.Key = NVL(rs("对象标记"), 0)
    Me.ID = NVL(rs("ID"), 0)
    Me.文件ID = NVL(rs("文件ID"), 0)
    Me.父ID = NVL(rs("父ID"), 0)
    Me.对象序号 = NVL(rs("对象序号"), 0)
    Me.保留对象 = IIf(NVL(rs("保留对象"), 0) = 0, False, True)
    Me.对象属性 = NVL(rs("对象属性"))
    Me.是否换行 = IIf(NVL(rs("是否换行"), 0) = 0, False, True)
    Me.预制提纲ID = NVL(rs("预制提纲ID"), 0)
    Me.标记 = NVL(rs("内容文本"), "")
    
    If eEdtType = cprET_单病历审核 Or eEdtType = cprET_单病历编辑 Then
        Me.开始版 = NVL(rs("开始版"), 1)
        Me.终止版 = NVL(rs("终止版"), 0)
    End If

    '读取表格内部单元、图片、诊治要素的子成员信息
    Call GetTbChildMember(lngFileID, Me.ID, strTableName)
    
    GetTableFromDB = True
End Function

'################################################################################################################
'## 功能：  从数据库直接获取指定版本单元集合，仅用于回退时重新获取上以版本的单元
'##
'## 参数：  lngVersion  :指定的版本
'################################################################################################################
Public Function ReGetCellsFromDB(lngVersion As Long) As Boolean
    Dim i As Long
    Dim RsSub As New ADODB.Recordset, strTmp As String, lngType As Long, lngKey As Long
    
    '先清除单元集合及其关联的要素和图片
    For i = Me.Cells.Count To 1 Step -1
        Me.Cells.Remove i
    Next
    For i = Me.Elements.Count To 1 Step -1
        Me.Elements.Remove i
    Next
    
    '再重新获取单元集合
    gstrSQL = "Select Level, t.*" & vbNewLine & _
            "From (Select t.* From 电子病历内容 t Where t.文件id = [1] And Nvl(t.开始版, 0) < [3] And (Nvl(终止版, 0) = 0 Or Nvl(终止版, 0) >= [3])) t" & vbNewLine & _
            "Start With 父id + 0 = [2]" & vbNewLine & _
            "Connect By Prior Id = 父id + 0" & vbNewLine & _
            "Order By 对象标记, 对象序号, 内容行次"
    Set RsSub = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRTable", Me.文件ID, Me.ID, lngVersion)
    Do While Not RsSub.EOF
        lngType = NVL(RsSub("对象类型"), 0)
        Select Case lngType
        Case 2  '文本
            '增加单元 Cell 到 Cells集合 中
            lngKey = Me.Cells.Add(NVL(RsSub("对象标记"), 0))
            Me.Cells("K" & lngKey).ID = NVL(RsSub("ID"), 0)
            Me.Cells("K" & lngKey).文件ID = NVL(RsSub("文件ID"), 0)
            Me.Cells("K" & lngKey).父ID = NVL(RsSub("父ID"), 0)
            Me.Cells("K" & lngKey).对象序号 = NVL(RsSub("对象序号"), 0)
            Me.Cells("K" & lngKey).内容文本 = NVL(RsSub("内容文本"), "")
            Me.Cells("K" & lngKey).保留对象 = IIf(NVL(RsSub("保留对象"), 0) = 0, False, True)
            Me.Cells("K" & lngKey).对象属性 = NVL(RsSub("对象属性"))
            Me.Cells("K" & lngKey).开始版 = NVL(RsSub("开始版"), 1)
            Me.Cells("K" & lngKey).终止版 = NVL(RsSub("终止版"), 0)
        Case 4  '诊治要素
            '先增加单元 Cell 到 Cells集合 中
            lngKey = Me.Cells.Add(NVL(RsSub("对象标记"), 0))
            Me.Cells("K" & lngKey).ID = NVL(RsSub("ID"), 0)
            Me.Cells("K" & lngKey).文件ID = NVL(RsSub("文件ID"), 0)
            Me.Cells("K" & lngKey).父ID = NVL(RsSub("父ID"), 0)
            Me.Cells("K" & lngKey).对象序号 = NVL(RsSub("对象序号"), 0)
            Me.Cells("K" & lngKey).内容文本 = NVL(RsSub("内容文本"), "")
            Me.Cells("K" & lngKey).保留对象 = IIf(NVL(RsSub("保留对象"), 0) = 0, False, True)
            Me.Cells("K" & lngKey).对象属性 = NVL(RsSub("对象属性"))
            Me.Cells("K" & lngKey).开始版 = NVL(RsSub("开始版"), 1)
            Me.Cells("K" & lngKey).终止版 = NVL(RsSub("终止版"), 0)
            
            '再增加诊治要素 Element 到 Elements集合 中
            lngKey = Me.Elements.Add(Me.Cells("K" & lngKey).ElementKey)
            If lngKey > 0 Then
                Me.Elements("K" & lngKey).ID = RsSub("ID")                     '本行的ID与Cell的ID相同，因为是同一条记录
                Me.Elements("K" & lngKey).文件ID = NVL(RsSub("文件ID"), 0)
                Me.Elements("K" & lngKey).父ID = NVL(RsSub("父ID"), 0)
                Me.Elements("K" & lngKey).对象序号 = NVL(RsSub("对象序号"), 0)
                Me.Elements("K" & lngKey).保留对象 = IIf(NVL(RsSub("保留对象"), 0) = 0, False, True)
                Me.Elements("K" & lngKey).内容文本 = NVL(RsSub("内容文本"), "")
                Me.Elements("K" & lngKey).是否换行 = IIf(NVL(RsSub("是否换行"), 0) = 0, False, True)
                Me.Elements("K" & lngKey).诊治要素ID = NVL(RsSub("诊治要素ID"), 0)
                Me.Elements("K" & lngKey).替换域 = NVL(RsSub("替换域"), 0)
                Me.Elements("K" & lngKey).要素名称 = NVL(RsSub("要素名称"))
                Me.Elements("K" & lngKey).要素类型 = NVL(RsSub("要素类型"), 0)
                Me.Elements("K" & lngKey).要素长度 = NVL(RsSub("要素长度"), 0)
                Me.Elements("K" & lngKey).要素小数 = NVL(RsSub("要素小数"), 0)
                Me.Elements("K" & lngKey).要素单位 = NVL(RsSub("要素单位"))
                Me.Elements("K" & lngKey).要素表示 = NVL(RsSub("要素表示"), 0)
                Me.Elements("K" & lngKey).输入形态 = NVL(RsSub("输入形态"), 0)
                Me.Elements("K" & lngKey).要素值域 = NVL(RsSub("要素值域"))
                Me.Elements("K" & lngKey).开始版 = NVL(RsSub("开始版"), 1)
                Me.Elements("K" & lngKey).终止版 = NVL(RsSub("终止版"), 0)
            End If
        Case 5  '图片
        End Select
        RsSub.MoveNext
    Loop
    RsSub.Close
    ReGetCellsFromDB = True
End Function

'################################################################################################################
'## 功能：  获取最终的表格替代图片
'##
'## 返回：  返回结果图片
'################################################################################################################
Public Function GetFinalPic(Optional ByVal BorderLine As Boolean = True) As StdPicture
    On Error Resume Next
    Dim f As New frmTablePicCreator
    f.ZoomPicture = Me.ZoomPicture
    Set GetFinalPic = f.GetFinalPic(Me, BorderLine)
    Unload f
    Set f = Nothing
End Function

'################################################################################################################
'## 功能：  检查对象的ID是否有效
'##
'## 参数：  eEditType :当前的编辑模式
'##
'## 说明：  如果没有ID，则到数据库中提取一个唯一ID号。
'################################################################################################################
Public Sub CheckValidID(eEdtType As EditTypeEnum)
    Select Case eEdtType
    Case cprET_病历文件定义
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("病历文件结构")
    Case cprET_全文示范编辑
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("病历范文内容")
    Case cprET_单病历编辑, cprET_单病历审核
        If Me.ID = 0 Then Me.ID = GetEPRContentNextId
    End Select
End Sub

'################################################################################################################
'## 功能：  获取适当的缩放比例，使得图片保持最接近原始尺寸的大小（只用于Pacs报告图片）
'################################################################################################################
Public Function GetFitZoomFactor() As Double
    Dim Factor As Double, i As Integer, f1 As Double, f2 As Double, f As Double
    Factor = 999999#
    For i = 1 To Pictures.Count
        f1 = Pictures(i).OrigWidth / Pictures(i).Width
        f2 = Pictures(i).OrigHeight / Pictures(i).Height
        If (f1 >= 1# And f2 >= 2#) Then
            f = IIf(f1 < f2, f1, f2)
            If (f < Factor) Then Factor = f
        ElseIf f1 >= 1# Then
            f = f1
            If (f < Factor) Then Factor = f
        ElseIf f2 >= 1# Then
            f = f2
            If (f < Factor) Then Factor = f
        End If
    Next
    If (Factor > 100# Or Factor <= 1#) Then
        Factor = 1#
    End If
    GetFitZoomFactor = Factor
End Function

'################################################################################################################
'## 功能：  插入图片到Editor中
'##
'## 参数：  edtThis         :当前的编辑器控件
'##         lngStartPos     :起始位置，默认为当前位置
'##         FinalPic        :最终图片，如果为0，则需要获取最终图片；如果已经有图片，则直接绘图
'##         blnFirstIns     :是否第一次插入表格
'##         blnResetPara    :是否将段落格式设置为初始状态（正文）
'##         blnForceInsert  :是否强制插入，在修改要素时如果前后都保护，这时需要强制插入
'##         bNotChangeSelPos:是否不改变光标位置
'################################################################################################################
Public Function InsertIntoEditor(ByRef edtThis As Object, _
    Optional ByVal lngStartPos As Long = -1, _
    Optional ByRef FinalPic As StdPicture, _
    Optional ByVal blnFirstIns As Boolean = False, _
    Optional ByVal blnResetPara As Boolean = False, _
    Optional ByVal blnForceInsert As Boolean = False, _
    Optional ByVal bNotChangeSelPos As Boolean = False) As Boolean
    
        On Error Resume Next
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '保证不能插入关键字内部
    If blnForceInsert = False And edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR _
        And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then Exit Function

    Dim p As Long, lngKey As Long, blnForce As Boolean
    blnForce = edtThis.ForceEdit

    With edtThis
        Dim lS As Long, lE As Long
        .InProcessing = True
        If bNotChangeSelPos = False Then .Freeze
        .Tag = "Table:InsertIntoEditor"
        .ForceEdit = True
        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        '此时插入图片，那么当前选中位置会改变
        If FinalPic Is Nothing Then
            p = .InsertPicture(Me.GetFinalPic(), mvarWidth, mvarHeight, lngStartPos, lngStartPos)
        Else
            p = .InsertPicture(FinalPic, lStart:=lngStartPos, lEnd:=lngStartPos)
        End If
                
        If .Range(p - 2, p) = vbCrLf Then   ' And .Range(p - 2, p).Font.Protected = True
            .Range(p - 2, p) = ""
            p = p - 2
        End If
        If blnFirstIns Then
            '第一次插入图片，后面跟一个空格。
            .Range(p, p) = vbCrLf & "TS(" & Format(Me.Key, "00000000") & ",1,0)"
            .Range(p + 19, p + 19) = "TE(" & Format(Me.Key, "00000000") & ",1,0)" & " "
            .Range(p + 2, p + 18).Font.Hidden = True
            .Range(p + 18, p + 19).Font.Hidden = False
            .Range(p + 19, p + 35).Font.Hidden = True
            .Range(p, p + 35).Font.Protected = True
            .Range(p + 35, p + 36).Font.Protected = False
            .Range(p + 35, p + 36).Font.Hidden = False
            .Range(p + 36, p + 36).Selected
            '手工设置为单倍行距
            .Range(p + 18, p + 19).Para.SetLineSpacing cprLSSignle, 0
        Else
            '
            .Range(p, p) = vbCrLf & "TS(" & Format(Me.Key, "00000000") & ",1,0)"
            .Range(p + 19, p + 19) = "TE(" & Format(Me.Key, "00000000") & ",1,0)" & IIf(Me.是否换行, vbCrLf, "")
            .Range(p + 2, p + 18).Font.Hidden = True
            .Range(p + 18, p + 19).Font.Hidden = False
            .Range(p + 19, p + 35).Font.Hidden = True
            .Range(p, p + 35).Font.Protected = True
            If blnResetPara Then
                SetCommonStyle edtThis, "正文", p + 18, p + 19, True
                '手工设置为单倍行距
                .Range(p + 18, p + 19).Para.SetLineSpacing cprLSSignle, 0
            End If
            If Me.是否换行 Then
                .Range(p + 35, p + 37).Font.Protected = False
                .Range(p + 35, p + 37).Font.Hidden = False
                If bNotChangeSelPos = False Then .Range(p + 37, p + 37).Selected
            Else
                .Range(p + 35, p + 35).Font.Protected = False
                .Range(p + 35, p + 35).Font.Hidden = False
                If bNotChangeSelPos = False Then .Range(p + 35, p + 35).Selected
            End If
        End If
        .ForceEdit = blnForce
        If bNotChangeSelPos = False Then .UnFreeze
        .LockAllOLEObjectSize
        .InProcessing = False
        .Tag = ""
    End With
    InsertIntoEditor = True
End Function

'################################################################################################################
'## 功能：  刷新显示
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function Refresh(ByRef edtThis As Object, Optional FinalPic As StdPicture = Nothing, Optional bNotChangeSelPos As Boolean = False) As Boolean
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "T", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Tag = "Refresh"
            .Freeze
            .InProcessing = True
            blnForce = .ForceEdit
            .ForceEdit = True
            Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat
            Set ParaFmt = edtThis.Range(lKSE, lKSE).Para.GetParaFmt
            Set FontFmt = edtThis.Range(lKSE, lKSE).Font.GetFontFmt
            
            .Range(lKSS, lKEE) = ""
            If Me.是否换行 Then Me.是否换行 = False
            Me.InsertIntoEditor edtThis, lKSS, FinalPic, , , True, bNotChangeSelPos
            .Range(lKSE, lKES).Font.SetFontFmt FontFmt
            .Range(lKSE, lKSE).Para.SetParaFmt ParaFmt
            .ForceEdit = blnForce
            .InProcessing = False
            .UnFreeze
            .Tag = ""
        End With
        Refresh = True
    Else
        Refresh = False
    End If
End Function

'################################################################################################################
'## 功能：  从编辑器中删除
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function DeleteFromEditor(ByRef edtThis As Object) As Boolean
    '保存诊治要素编辑结果
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

        On Error Resume Next
    bFinded = FindKey(edtThis, "T", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = ""
            .Range(lKSS, lKSS).Selected
            .ForceEdit = blnForce
            .UnFreeze
        End With
        DeleteFromEditor = True
    Else
        DeleteFromEditor = False
    End If
End Function

'################################################################################################################
'## 功能：  获取表格保存到数据库的SQL语句
'##
'## 参数：  eEdtType        :IN，       当前编辑模式
'##         ArraySQL()      :IN/OUT，   SQL数组
'##         RecID           :IN，       记录ID
'################################################################################################################
Public Function GetSaveSQL(ByVal eEdtType As EditTypeEnum, ByRef ArraySQL() As String, Optional ByVal RecID As Long = 0) As Boolean
    '1、保存 Table：
    Dim lngCount As Long
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String
    
    Call CheckValidID(eEdtType)         '检查ID有效性！
    Select Case eEdtType
    Case cprET_病历文件定义
        gstrSQL = "Zl_病历文件结构_Update(" & Me.ID & "," & Me.文件ID & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            Me.对象序号 & ",3," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.对象属性 & "',NULL,'" & mvar标记 & "'," & IIf(Me.是否换行, 1, 0) & _
            "," & IIf(Me.预制提纲ID = 0, "NULL", Me.预制提纲ID) & ")"
    Case cprET_全文示范编辑
        gstrSQL = "Zl_病历范文内容_Update(" & Me.ID & "," & Me.文件ID & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            Me.对象序号 & ",3," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.对象属性 & "',NULL,'" & mvar标记 & "'," & IIf(Me.是否换行, 1, 0) & _
            ",NULL," & IIf(Me.预制提纲ID = 0, "NULL", Me.预制提纲ID) & ")"
    Case cprET_单病历编辑, cprET_单病历审核
        '加了 开始版,终止版 字段，其他相同
        gstrSQL = "ZL_电子病历内容_Update(" & Me.ID & "," & Me.文件ID & "," & Me.开始版 & "," & Me.终止版 & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            Me.对象序号 & ",3," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.对象属性 & "',NULL,'" & mvar标记 & "'," & IIf(Me.是否换行, 1, 0) & _
            ",NULL," & IIf(Me.预制提纲ID = 0, "NULL", Me.预制提纲ID) & ")"
    End Select
    ArraySQL(lngCount) = gstrSQL
    
    '2、保存Cells：
    Dim i As Long, lng行次 As Long, lngEleKey As Long
    lng行次 = 1
    For i = 1 To Me.Cells.Count
        lngCount = UBound(ArraySQL) + 1
        ReDim Preserve ArraySQL(1 To lngCount) As String
        
        Me.Cells(i).CheckValidID (eEdtType)
        If Me.Cells(i).ElementKey > 0 Then
            '诊治要素
            lngEleKey = Me.Cells(i).ElementKey
            Select Case eEdtType
            Case cprET_病历文件定义
                gstrSQL = "Zl_病历文件结构_Update(" & Me.Cells(i).ID & "," & Me.文件ID & "," & Me.ID & ",NULL," & _
                    "4," & Me.Cells(i).Key & ",NULL,'" & Me.Cells(i).对象属性 & "'," & lng行次 & ",'" & Me.Cells(i).内容文本 & "',NULL,NULL," & _
                    "NULL,NULL," & _
                    IIf(Me.Elements("K" & lngEleKey).诊治要素ID = 0, "NULL", Me.Elements("K" & lngEleKey).诊治要素ID) & _
                    "," & Me.Elements("K" & lngEleKey).替换域 & ",'" & Me.Elements("K" & lngEleKey).要素名称 & "'," & Me.Elements("K" & lngEleKey).要素类型 & "," & Me.Elements("K" & lngEleKey).要素长度 & "," & _
                    Me.Elements("K" & lngEleKey).要素小数 & ",'" & Me.Elements("K" & lngEleKey).要素单位 & "'," & Me.Elements("K" & lngEleKey).要素表示 & "," & Me.Elements("K" & lngEleKey).输入形态 & ",'" & Me.Elements("K" & lngEleKey).要素值域 & "')"
            Case cprET_全文示范编辑
                gstrSQL = "Zl_病历范文内容_Update(" & Me.Cells(i).ID & "," & Me.文件ID & "," & Me.ID & ",NULL," & _
                    "4," & Me.Cells(i).Key & ",NULL,'" & Me.Cells(i).对象属性 & "'," & lng行次 & ",'" & Me.Cells(i).内容文本 & "',NULL,NULL,NULL," & _
                    "NULL,NULL," & _
                    IIf(Me.Elements("K" & lngEleKey).诊治要素ID = 0, "NULL", Me.Elements("K" & lngEleKey).诊治要素ID) & _
                    "," & Me.Elements("K" & lngEleKey).替换域 & ",'" & Me.Elements("K" & lngEleKey).要素名称 & "'," & Me.Elements("K" & lngEleKey).要素类型 & "," & Me.Elements("K" & lngEleKey).要素长度 & "," & _
                    Me.Elements("K" & lngEleKey).要素小数 & ",'" & Me.Elements("K" & lngEleKey).要素单位 & "'," & Me.Elements("K" & lngEleKey).要素表示 & "," & Me.Elements("K" & lngEleKey).输入形态 & ",'" & Me.Elements("K" & lngEleKey).要素值域 & "')"
            Case cprET_单病历编辑, cprET_单病历审核
                gstrSQL = "ZL_电子病历内容_Update(" & Me.Cells(i).ID & "," & Me.文件ID & "," & Me.Cells(i).开始版 & ",0," & Me.ID & "," & Me.Cells(i).ID & "," & _
                    "4," & Me.Cells(i).Key & ",NULL,'" & Me.Cells(i).对象属性 & "'," & lng行次 & ",'" & Me.Cells(i).内容文本 & "',NULL,NULL,NULL," & _
                    "NULL,NULL," & _
                    IIf(Me.Elements("K" & lngEleKey).诊治要素ID = 0, "NULL", Me.Elements("K" & lngEleKey).诊治要素ID) & _
                    "," & Me.Elements("K" & lngEleKey).替换域 & ",'" & Me.Elements("K" & lngEleKey).要素名称 & "'," & Me.Elements("K" & lngEleKey).要素类型 & "," & Me.Elements("K" & lngEleKey).要素长度 & "," & _
                    Me.Elements("K" & lngEleKey).要素小数 & ",'" & Me.Elements("K" & lngEleKey).要素单位 & "'," & Me.Elements("K" & lngEleKey).要素表示 & "," & Me.Elements("K" & lngEleKey).输入形态 & ",'" & Me.Elements("K" & lngEleKey).要素值域 & "')"
            End Select
        Else
            '纯文本
            Select Case eEdtType
            Case cprET_病历文件定义
                gstrSQL = "Zl_病历文件结构_Update(" & Me.Cells(i).ID & "," & Me.文件ID & "," & Me.ID & ",NULL," & _
                    "2," & Me.Cells(i).Key & ",NULL,'" & Me.Cells(i).对象属性 & "'," & lng行次 & ",'" & Replace(Me.Cells(i).内容文本, "'", "' || chr(39) || '") & "')"
            Case cprET_全文示范编辑
                gstrSQL = "Zl_病历范文内容_Update(" & Me.Cells(i).ID & "," & Me.文件ID & "," & Me.ID & ",NULL," & _
                    "2," & Me.Cells(i).Key & ",NULL,'" & Me.Cells(i).对象属性 & "'," & lng行次 & ",'" & Replace(Me.Cells(i).内容文本, "'", "' || chr(39) || '") & "')"
            Case cprET_单病历编辑, cprET_单病历审核
                gstrSQL = "ZL_电子病历内容_Update(" & Me.Cells(i).ID & "," & Me.文件ID & "," & Me.Cells(i).开始版 & ",0," & Me.ID & "," & Me.Cells(i).ID & "," & _
                    "2," & Me.Cells(i).Key & ",NULL,'" & Me.Cells(i).对象属性 & "'," & lng行次 & ",'" & Replace(Me.Cells(i).内容文本, "'", "' || chr(39) || '") & "')"
            End Select
        End If
        ArraySQL(lngCount) = gstrSQL
        lng行次 = lng行次 + 1
    Next
    
    '3、保存Pictures：
    For i = 1 To Me.Pictures.Count
        Me.Pictures(i).父ID = Me.ID
        Me.Pictures(i).GetSaveSQL eEdtType, ArraySQL, True, lng行次
        lng行次 = lng行次 + 1
    Next
  
    '4、如果是诊断等特殊对象，则保存附加SQL
    If eEdtType = cprET_单病历编辑 Or eEdtType = cprET_单病历审核 Then
        For i = 1 To Me.DicSpecSQLs.Count
            lngCount = UBound(ArraySQL) + 1
            ReDim Preserve ArraySQL(1 To lngCount) As String
            ArraySQL(lngCount) = Replace(Me.DicSpecSQLs("K" & i), "[1]", Me.ID)                 '诊断的第一个参数为病历元素ID（表格ID）
            If RecID > 0 Then ArraySQL(lngCount) = Replace(ArraySQL(lngCount), "[2]", RecID)    '第二个参数代表文件记录的ID（用于其他附加表格）
        Next
    End If
    
    GetSaveSQL = True
End Function

Public Sub InsertRow(ByVal lrow As Long)
    '插入到指定行之后
    Dim i As Long, lKey As Long
    mvarRows = mvarRows + 1
    For i = 1 To mvarCells.Count
        If mvarCells(i).Row > lrow Then
            mvarCells(i).Row = mvarCells(i).Row + 1
        End If
    Next
    For i = 1 To mvarCols
        lKey = mvarCells.Add(, lrow + 1, i)
    Next
End Sub

Public Sub DeleteRow(ByVal lrow As Long)
    '删除指定行
    On Error GoTo LL
    Dim i As Long, lKey As Long
    For i = 1 To mvarCols
        '删除图片、要素
        lKey = Cell(lrow, i).Key
        If mvarCells("K" & lKey).ElementKey > 0 Then Me.Elements.Remove "K" & mvarCells("K" & lKey).ElementKey
        If mvarCells("K" & lKey).PictureKey > 0 Then Me.Pictures.Remove "K" & mvarCells("K" & lKey).PictureKey
        mvarCells.Remove "K" & lKey
    Next
    mvarRows = mvarRows - 1
    For i = 1 To mvarCells.Count
        If mvarCells(i).Row > lrow Then
            mvarCells(i).Row = mvarCells(i).Row - 1
        End If
    Next
    Exit Sub
LL:
    Debug.Assert False
End Sub

Public Sub InsertCol(ByVal lCol As Long)
    '插入到指定列之后
    Dim i As Long, lKey As Long
    mvarCols = mvarCols + 1
    For i = 1 To mvarCells.Count
        If mvarCells(i).Col > lCol Then
            mvarCells(i).Col = mvarCells(i).Col + 1
        End If
    Next
    For i = 1 To mvarCols
        lKey = mvarCells.Add(, i, lCol + 1)
    Next
End Sub

Public Sub DeleteCol(ByVal lCol As Long)
    '删除指定列
    On Error GoTo LL
    Dim i As Long, lKey As Long
    For i = 1 To mvarRows
        '删除图片、要素
        lKey = Cell(i, lCol).Key
        If mvarCells("K" & lKey).ElementKey > 0 Then Me.Elements.Remove "K" & mvarCells("K" & lKey).ElementKey
        If mvarCells("K" & lKey).PictureKey > 0 Then Me.Pictures.Remove "K" & mvarCells("K" & lKey).PictureKey
        mvarCells.Remove "K" & lKey
    Next
    mvarCols = mvarCols - 1
    For i = 1 To mvarCells.Count
        If mvarCells(i).Col > lCol Then
            mvarCells(i).Col = mvarCells(i).Col - 1
        End If
    Next
    Exit Sub
LL:
    Debug.Assert False
End Sub

