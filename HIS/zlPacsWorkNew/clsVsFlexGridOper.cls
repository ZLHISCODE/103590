VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVsFlexGridOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




Private Const mstrCheckColName As String = "□"
Private Const mstrAdjustColName As String = "≡"
Private Const mstrNameSplitChar As String = ">"
Private Const mstrPropertySplitChar As String = "@"


Private Const mstrColProp_Hide As String = "hide"
Private Const mstrColProp_Read As String = "read"
Private Const mstrColProp_Btn As String = "btn"
Private Const mstrColProp_Merge As String = "merge"
Private Const mstrColProp_Check As String = "check"
Private Const mstrColProp_RowCheck As String = "rowcheck"
Private Const mstrColProp_Key As String = "key"
Private Const mstrColProp_Cbx As String = "cbx"
Private Const mstrColProp_Align As String = "align"
Private Const mstrColProp_TxtLeft As String = "txtleft"
Private Const mstrColProp_TxtRight As String = "txtright"
Private Const mstrColProp_TxtCenter As String = "txtcenter"
Private Const mstrColProp_ColLeft As String = "colleft"
Private Const mstrColProp_ColRight As String = "colright"
Private Const mstrColProp_ColCenter As String = "colcenter"
Private Const mstrColProp_ChkLeft As String = "chkleft"
Private Const mstrColProp_ChkRight As String = "chkright"
Private Const mstrColProp_ChkCenter As String = "chkcenter"
Private Const mstrColProp_WidthTag As String = "w"
Private Const mstrColProp_Tdate As String = "tdate"
Private Const mstrColProp_Tnum As String = "tnum"
Private Const mstrColProp_Tstr As String = "tstr"



'显示列信息
Private Type colInf
    ColName As String
    FieldName As String
    IsHide As Boolean
    Width As Long
    IsRead As Boolean
    IsMerge As Boolean
    IsButton As Boolean
    IsCheck As Boolean
    IsKey As Boolean
    IsCombox As Boolean
    IsRowCheck As Boolean
End Type



'对象显示位置的类型
Public Enum ObjPostionType
    optLeft = 0  '靠左
    optRight = 1 '靠右
End Enum


'数据行状态
Public Enum RowState
    Normal = 0  '正常
    Add = 1 '新增
    Update = 2  '更新
    Del = 3 '删除
End Enum


Public mCurFlexGrid As VSFlexGrid

Private mCols() As colInf                     '数据列信息
Private mstrKey As String                     '关键列名称
Private mstrKeyField As String                '关键字段名
Private mlngKeepRows As Long                  '需要保持的行数量
Private mblnIsAllowExtCol As Boolean          '是否允许扩展列
Private mblnIsShowNumber As Boolean           '是否显示行编号
Private mlngDisableColor As Long              '不可编辑单元格颜色


Private mlngOldDataRowHeight As Long




'不可编辑的单元格颜色
Property Get DisableColor() As OLE_COLOR
    DisableColor = mlngDisableColor
End Property

Property Let DisableColor(value As OLE_COLOR)
    mlngDisableColor = value
End Property




'是否显示行编号
Property Get IsShowNumber() As Boolean
    IsShowNumber = mblnIsShowNumber
End Property


Property Let IsShowNumber(value As Boolean)
    mblnIsShowNumber = value
End Property





'关键字（对应列表中的显示名称）
Property Get KeyName() As String
    KeyName = mstrKey
End Property


Property Let KeyName(value As String)
    mstrKey = value
End Property



'关键字段（直接对应数据库中的字段名）
Property Get KeyField() As String
    KeyField = mstrKeyField
End Property

Property Let KeyField(value As String)
    mstrKeyField = value
End Property





'保持的行数量
Property Get KeepRows() As Long
    KeepRows = mlngKeepRows
End Property


Property Let KeepRows(value As Long)
    mlngKeepRows = value
End Property



'check列名称
Property Get CheckColName() As String
    CheckColName = mstrCheckColName
End Property





'adjust列名称
Property Get AdjustColName() As String
    AdjustColName = mstrAdjustColName
End Property





Public Sub InitVsFlexGridList(vfgList As Object, ByVal strCols As String)
'初始化列表对象
    Dim i As Integer
    Dim Cols() As String
    Dim colInf() As String
    Dim colCount As Integer
    
    Dim strTemp As String
    Dim strSubPro As String

    Set mCurFlexGrid = vfgList
    
    If Trim(strCols) = "" Then Exit Sub
    
    If mblnIsAllowExtCol Then
        Cols() = Split("|≡" & strCols, "|")
    Else
        Cols() = Split(strCols, "|")
    End If
        
    
    colCount = 0
    
    vfgList.Cols = 0
    vfgList.Rows = 1
    
    
'    '词句的作用是为了再执行自定义排序后能够触发aftersort事件
'
'    If vfgList.ExplorerBar <= &H1000& Then
'        vfgList.ExplorerBar = vfgList.ExplorerBar + &H1000&
'    End If
    
    For i = LBound(Cols()) To UBound(Cols())
        If Trim(Cols(i)) <> "" Then
            strTemp = Cols(i)
            strTemp = Replace(strTemp, " ", "")
            
            colInf() = Split(strTemp & ",,,,,,", ",")
            
            
            colCount = colCount + 1
            ReDim Preserve mCols(colCount)
            
            
            
            '读取列属性---------------------
            
            '列名称
            mCols(colCount).ColName = Split(colInf(0) & mstrNameSplitChar, mstrNameSplitChar)(0)
            '数据库字段名
            mCols(colCount).FieldName = Split(colInf(0) & mstrNameSplitChar, mstrNameSplitChar)(1)
            If Trim(mCols(colCount).FieldName) = "" Then mCols(colCount).FieldName = mCols(colCount).ColName
            
            '是否隐藏列
            mCols(colCount).IsHide = IIf(InStr(1, UCase(strTemp), UCase("," & mstrColProp_Hide)) >= 1, True, False)
            '是否只读列
            mCols(colCount).IsRead = IIf(InStr(1, UCase(strTemp), UCase("," & mstrColProp_Read)) >= 1, True, False)
            '是否按钮列
            mCols(colCount).IsButton = IIf(InStr(1, UCase(strTemp), UCase("," & mstrColProp_Btn)) >= 1, True, False)
            '是否合并列
            mCols(colCount).IsMerge = IIf(InStr(1, UCase(strTemp), UCase("," & mstrColProp_Merge)) >= 1, True, False)
            '是否check列
            mCols(colCount).IsCheck = IIf(InStr(1, UCase(strTemp), UCase("," & mstrColProp_Check)) >= 1, True, False)
            '是否关键字
            mCols(colCount).IsKey = IIf(InStr(1, UCase(strTemp), UCase("," & mstrColProp_Key)) >= 1, True, False)
            '是否可选列
            mCols(colCount).IsCombox = IIf(InStr(1, UCase(strTemp), UCase("," & mstrColProp_Cbx)) >= 1, True, False)
            '是否为行的check
            mCols(colCount).IsRowCheck = IIf(InStr(1, UCase(strTemp), UCase("," & mstrColProp_RowCheck)) >= 1, True, False)
            
            
            '列宽度设置
            If InStr(1, UCase(strTemp), UCase("," & mstrColProp_WidthTag)) > 1 Then
                strSubPro = Mid(strTemp, InStr(1, UCase(strTemp), UCase("," & mstrColProp_WidthTag)) + 1, 100)
                strSubPro = Replace(strSubPro, "|", ",") & ","
                strSubPro = Mid(strSubPro, 1, InStr(1, strSubPro, ","))

                mCols(colCount).Width = GetNumber(strSubPro)
            End If
            
            
            vfgList.Cols = vfgList.Cols + 1
            vfgList.Cell(flexcpText, 0, colCount - 1) = mCols(colCount).ColName
            
            '保存字段名和当前列的属性字符串
            vfgList.Cell(flexcpData, 0, colCount - 1) = mCols(colCount).FieldName & mstrPropertySplitChar & strTemp

            
            If mCols(colCount).Width > 0 Then vfgList.ColWidth(colCount - 1) = mCols(colCount).Width
            
            '设置列的关键字
            vfgList.ColKey(colCount - 1) = mCols(colCount).ColName
                
            '设置列的对齐方式
            If InStr(1, UCase(strTemp), "," & UCase(mstrColProp_ColRight)) >= 1 Then
                vfgList.Cell(flexcpAlignment, 0, colCount - 1) = flexAlignRightCenter
            ElseIf InStr(1, UCase(strTemp), "," & UCase(mstrColProp_ColCenter)) >= 1 Then
                vfgList.Cell(flexcpAlignment, 0, colCount - 1) = flexAlignCenterCenter
            ElseIf InStr(1, UCase(strTemp), "," & UCase(mstrColProp_ColLeft)) >= 1 Then
                vfgList.Cell(flexcpAlignment, 0, colCount - 1) = flexAlignLeftCenter
            End If
            
            
            '隐藏列
            If mCols(colCount).IsHide Then
                vfgList.ColHidden(colCount - 1) = True
            End If
            
            
            'button列
            If mCols(colCount).IsButton Then
                vfgList.ColComboList(colCount - 1) = "..."
            End If
            
            '合并列
            If mCols(colCount).IsMerge Then
                vfgList.MergeCol(colCount - 1) = True
            End If
            
            '设置关键字段
            If mCols(colCount).IsKey Then
                mstrKey = mCols(colCount).ColName
                mstrKeyField = mCols(colCount).FieldName
            End If
            
            
            
            '设置该列为combox列
            If mCols(colCount).IsCombox Then
                strSubPro = Mid(strTemp, InStr(1, UCase(strTemp), UCase("," & mstrColProp_Cbx & "<")) + Len("," & mstrColProp_Cbx & "<"), 100)
                strSubPro = Mid(strSubPro, 1, InStr(1, strSubPro, ">") - 1)
                strSubPro = Replace(strSubPro, ",", "|")
                
                vfgList.ColComboList(colCount - 1) = strSubPro
            End If
            
            
            '设置该列为CheckBox列
            If mCols(colCount).ColName = mstrCheckColName Then
                vfgList.ColDataType(colCount - 1) = flexDTBoolean
                vfgList.ColWidth(colCount - 1) = 300
                vfgList.ColAlignment(colCount - 1) = flexAlignCenterCenter
            End If
            
            
            '设置该列为扩展调整列
            If mCols(colCount).ColName = mstrAdjustColName Then
                vfgList.ColWidth(colCount - 1) = 500
                vfgList.ColAlignment(colCount - 1) = flexAlignCenterCenter
            End If
        End If
    Next i
    
    '如果未设置关键字，则选取第一个字段
    If Trim(mstrKey) = "" Then
        mstrKey = mCols(1).ColName
        mstrKeyField = mCols(1).FieldName
    End If
    
    '当允许改变行高度时，则第一列为固定列
    If mblnIsAllowExtCol Then
        vfgList.FixedCols = 1
        vfgList.AllowUserResizing = flexResizeBoth
    End If
    
    '初始化已有行的状态
    For i = 1 To mCurFlexGrid.Rows - 1
        mCurFlexGrid.RowData(i) = RowState.Normal
    Next i
    
End Sub











Public Sub UpdateRowNumber()
'更新列表编号
    Dim i As Long
    Dim lngNumber As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    If Not mblnIsAllowExtCol Then Exit Sub
    
    If mblnIsAllowExtCol And Not mblnIsShowNumber Then
        For i = 1 To mCurFlexGrid.Rows - 1
            mCurFlexGrid.TextMatrix(i, 0) = ""
        Next i
        
        Exit Sub
    End If
    
    If mblnIsAllowExtCol And mblnIsShowNumber Then
        lngNumber = 1
        
        For i = 1 To mCurFlexGrid.Rows - 1
            If Not mCurFlexGrid.RowHidden(i) Then
                mCurFlexGrid.TextMatrix(i, 0) = lngNumber
                lngNumber = lngNumber + 1
            End If
        Next i
    End If
End Sub


Public Function GetOrder(ByVal lngCurSortCol As Long, ByVal lngCurOrder As Long)
'取得排序order（参考vsflexgrid的排序demo）
    GetOrder = lngCurOrder
'
'    ' no flags? apply custom sort
'    If mCurFlexGrid.ExplorerBar > &H1000& Then Exit Function
'
'    '
'    ' the 'ignore blanks' flag isn't set, so do it with custom code
'    '
'
'    ' save selection
    
     '没有数据时退出排序
    If mCurFlexGrid.Rows = 1 Then Exit Function
    
    Dim R&, c&, RS&, cs&
    mCurFlexGrid.GetSelection R, c, RS, cs
    mCurFlexGrid.Redraw = flexRDNone

    ' apply sort to non-empty range
    Dim Row%
    For Row = mCurFlexGrid.Rows - 1 To mCurFlexGrid.FixedRows Step -1
        '整行数据为空时，不参与排序
        If Len(mCurFlexGrid.TextMatrix(Row, lngCurSortCol)) Or Not KeyEmpty(Row) Then Exit For
    Next
    
    If Row > mCurFlexGrid.FixedRows Then
        mCurFlexGrid.Select mCurFlexGrid.FixedRows, lngCurSortCol, Row, lngCurSortCol
        mCurFlexGrid.Sort = lngCurOrder
    End If
    
    ' restore selection
    mCurFlexGrid.Select R, c, RS, cs
    mCurFlexGrid.Redraw = flexRDDirect
    
    ' cancel default sort
    GetOrder = 0
End Function


Public Sub ClearDataList()
'清除数据列表
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    mCurFlexGrid.Rows = 1
    If mlngKeepRows > 0 Then mCurFlexGrid.Rows = mlngKeepRows
End Sub


Public Function FindRowIndex(ByVal strFindValue As String, ByVal strColName As String, _
    Optional ByVal blnIsPrecise As Boolean = False) As Long
'查找指定值并返回所在行索引
    Dim i As Long
    Dim lngCol As Long
    
    FindRowIndex = -1
    If mCurFlexGrid Is Nothing Then Exit Function
    If Trim(strFindValue) = "" Then Exit Function
    
    lngCol = GetColumnIndex(strColName)
    
    For i = 1 To mCurFlexGrid.Rows - 1
        If Not mCurFlexGrid.RowHidden(i) Then

            If UCase(mCurFlexGrid.TextMatrix(i, lngCol)) Like IIf(blnIsPrecise, UCase(strFindValue), "*" & UCase(strFindValue) & "*") Then
                FindRowIndex = i
                Exit Function
            End If
        End If
    Next i
End Function


Public Sub ShowAllDataRow()
'显示所有数据行
    Dim i As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    For i = 1 To mCurFlexGrid.Rows - 1
        mCurFlexGrid.RowHidden(i) = False
    Next i
End Sub


Public Sub SelectAll()
'全选
    Dim iCheckIndex As Long
    Dim i As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    iCheckIndex = GetColumnIndex(CheckColName)
    If iCheckIndex <= 0 Then Exit Sub
    
    '隐藏的行数据不能被选择
    For i = 1 To mCurFlexGrid.Rows - 1
        If Not mCurFlexGrid.RowHidden(i) Then
            mCurFlexGrid.TextMatrix(i, iCheckIndex) = -1
        End If
    Next i
End Sub


Public Sub ClearSelect()
'清除选择
    Dim iCheckIndex As Long
    Dim i As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    iCheckIndex = GetColumnIndex(CheckColName)
    If iCheckIndex < 0 Then Exit Sub
    
    For i = 1 To mCurFlexGrid.Rows - 1
        If Not mCurFlexGrid.RowHidden(i) Then
            mCurFlexGrid.TextMatrix(i, iCheckIndex) = 0
        End If
    Next i
End Sub




Public Sub ReSelect()
'反向选择
    Dim iCheckIndex As Long
    Dim i As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    iCheckIndex = GetColumnIndex(CheckColName)
    If iCheckIndex <= 0 Then Exit Sub
    
    For i = 1 To mCurFlexGrid.Rows - 1
        If Not mCurFlexGrid.RowHidden(i) Then
            If Val(mCurFlexGrid.TextMatrix(i, iCheckIndex)) = 0 Then
                mCurFlexGrid.TextMatrix(i, iCheckIndex) = -1
            Else
                mCurFlexGrid.TextMatrix(i, iCheckIndex) = 0
            End If
        End If
    Next i
End Sub


Public Function DisplayDataRowCount() As Integer
'取得当前显示行的数量
    Dim i As Long
    Dim lngCount As Long
    
    DisplayDataRowCount = 0
    If mCurFlexGrid Is Nothing Then Exit Function
    
    lngCount = 0
    For i = 1 To mCurFlexGrid.Rows - 1
        If Not mCurFlexGrid.RowHidden(i) Then lngCount = lngCount + 1
    Next i
    
    DisplayDataRowCount = lngCount
End Function


Public Function RowChecked(ByVal lngRowIndex As Long) As Boolean
'判断行的check状态
    RowChecked = False
    If mCurFlexGrid Is Nothing Then Exit Function
    
    RowChecked = IIf(Val(mCurFlexGrid.TextMatrix(lngRowIndex, GetColumnIndex(mstrCheckColName))) = 0, False, True)
End Function

Public Sub SetRowCheck(ByVal lngRowIndex As Long, ByVal blnIsChecked As Boolean)
'设置行的checked状态
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    mCurFlexGrid.TextMatrix(lngRowIndex, GetColumnIndex(mstrCheckColName)) = IIf(blnIsChecked, -1, 0)
End Sub



Public Function CurRowChecked() As Boolean
'判断当前行的选中状态
    CurRowChecked = False
    If mCurFlexGrid Is Nothing Then Exit Function
    
    CurRowChecked = RowChecked(mCurFlexGrid.RowSel)
End Function



Public Sub CopyGrid(vfgSource As Object, vfgTarget As Object)
'    If vfgSource Is Nothing Or vfgTarget Is Nothing Then Exit Sub
'
'    Dim i As Long
'    Dim j As Long
'
'    vfgTarget.cols = vfgSource.cols
'    vfgTarget.Rows = vfgSource.Rows
'    vfgTarget.FixedCols = vfgSource.FixedCols
'    vfgTarget.FixedRows = vfgSource.FixedRows
'    vfgTarget.ColWidthMax = vfgSource.ColWidthMax
'    vfgTarget.ColWidthMin = vfgSource.ColWidthMin
'    vfgTarget.RowHeightMax = vfgSource.RowHeightMax
'    vfgTarget.RowHeightMin = vfgSource.RowHeightMin
'    vfgTarget.MergeCells = vfgSource.MergeCells
'    vfgTarget.MergeCellsFixed = vfgSource.MergeCellsFixed
'    vfgTarget.BackColor = vfgSource.BackColor
'    vfgTarget.CellBackColor = vfgSource.CellBackColor
'    vfgTarget.BackColorAlternate = vfgSource.BackColorAlternate
'    vfgTarget.BackColorBkg = vfgSource.BackColorBkg
'    vfgTarget.BackColorFixed = vfgSource.BackColorFixed
'    vfgTarget.BackColorFrozen = vfgSource.BackColorFrozen
'    vfgTarget.BackColorSel = vfgSource.BackColorSel
'    vfgTarget.CellFloodColor = vfgSource.CellFloodColor
'    vfgTarget.CellForeColor = vfgSource.CellForeColor
'    vfgTarget.FloodColor = vfgSource.FloodColor
'    vfgTarget.ForeColor = vfgSource.ForeColor
'    vfgTarget.BorderStyle = vfgSource.BorderStyle
'
'    For j = 0 To vfgSource.cols - 1
'        vfgTarget.ColHidden(j) = vfgSource.ColHidden(j)
'        vfgTarget.ColData(j) = vfgSource.ColData(j)
'        vfgTarget.ColDataType(j) = vfgSource.ColDataType(j)
'        vfgTarget.ColComboList(j) = vfgSource.ColComboList(j)
'        vfgTarget.ColFormat(j) = vfgSource.ColFormat(j)
'        vfgTarget.ColEditMask(j) = vfgSource.ColEditMask(j)
'        vfgTarget.ColKey(j) = vfgSource.ColKey(j)
'        vfgTarget.ColWidth(j) = vfgSource.ColWidth(j)
'        vfgTarget.ColAlignment(j) = vfgSource.ColAlignment(j)
'        vfgTarget.MergeCol(j) = vfgSource.MergeCol(j)
'    Next j
'
'
'    For i = 0 To vfgSource.Rows - 1
'        For j = 0 To vfgSource.cols - 1
'            vfgTarget.Cell(flexcpAlignment, i, j) = vfgSource.Cell(flexcpAlignment, i, j)
'            vfgTarget.Cell(flexcpBackColor, i, j) = vfgSource.Cell(flexcpBackColor, i, j)
'            vfgTarget.Cell(flexcpChecked, i, j) = vfgSource.Cell(flexcpChecked, i, j)
'            vfgTarget.Cell(flexcpCustomFormat, i, j) = vfgSource.Cell(flexcpCustomFormat, i, j)
'            vfgTarget.Cell(flexcpData, i, j) = vfgSource.Cell(flexcpData, i, j)
'            vfgTarget.Cell(flexcpFloodColor, i, j) = vfgSource.Cell(flexcpFloodColor, i, j)
'            vfgTarget.Cell(flexcpFloodPercent, i, j) = vfgSource.Cell(flexcpFloodPercent, i, j)
'            vfgTarget.Cell(flexcpFont, i, j) = vfgSource.Cell(flexcpFont, i, j)
'            vfgTarget.Cell(flexcpFontBold, i, j) = vfgSource.Cell(flexcpFontBold, i, j)
'            vfgTarget.Cell(flexcpFontItalic, i, j) = vfgSource.Cell(flexcpFontItalic, i, j)
'            vfgTarget.Cell(flexcpFontName, i, j) = vfgSource.Cell(flexcpFontName, i, j)
'            vfgTarget.Cell(flexcpFontSize, i, j) = vfgSource.Cell(flexcpFontSize, i, j)
'            vfgTarget.Cell(flexcpFontStrikethru, i, j) = vfgSource.Cell(flexcpFontStrikethru, i, j)
'            vfgTarget.Cell(flexcpFontUnderline, i, j) = vfgSource.Cell(flexcpFontUnderline, i, j)
'            vfgTarget.Cell(flexcpFontWidth, i, j) = vfgSource.Cell(flexcpFontWidth, i, j)
'            vfgTarget.Cell(flexcpForeColor, i, j) = vfgSource.Cell(flexcpForeColor, i, j)
'            'vfgTarget.Cell(flexcpHeight, i, j) = vfgSource.Cell(flexcpHeight, i, j)
'            'vfgTarget.Cell(flexcpLeft, i, j) = vfgSource.Cell(flexcpLeft, i, j)
'            vfgTarget.Cell(flexcpPicture, i, j) = vfgSource.Cell(flexcpPicture, i, j)
'            vfgTarget.Cell(flexcpPictureAlignment, i, j) = vfgSource.Cell(flexcpPictureAlignment, i, j)
''            vfgTarget.Cell(flexcpRefresh, i, j) = vfgSource.Cell(flexcpRefresh, i, j)
''            vfgTarget.Cell(flexcpSort, i, j) = vfgSource.Cell(flexcpSort, i, j)
'            vfgTarget.Cell(flexcpText, i, j) = vfgSource.Cell(flexcpText, i, j)
''            vfgTarget.Cell(flexcpTextDisplay, i, j) = vfgSource.Cell(flexcpTextDisplay, i, j)
'            vfgTarget.Cell(flexcpTextStyle, i, j) = vfgSource.Cell(flexcpTextStyle, i, j)
''            vfgTarget.Cell(flexcpTop, i, j) = vfgSource.Cell(flexcpTop, i, j)
''            vfgTarget.Cell(flexcpValue, i, j) = vfgSource.Cell(flexcpValue, i, j)
''            vfgTarget.Cell(flexcpVariantValue, i, j) = vfgSource.Cell(flexcpVariantValue, i, j)
''            vfgTarget.Cell(flexcpWidth, i, j) = vfgSource.Cell(flexcpWidth, i, j)
'        Next j
'
'        vfgTarget.RowHidden(i) = vfgSource.RowHidden(i)
'        vfgTarget.RowData(i) = vfgSource.RowData(i)
'        vfgTarget.RowHeight(i) = vfgSource.RowHeight(i)
'        vfgTarget.MergeRow(i) = vfgSource.MergeRow(i)
'    Next i
End Sub


Public Function IsReadCol(ByVal lngCol As Long) As Boolean
'判断指定列是否为Read列
    Dim strColProperty As String
    
    IsReadCol = True
    If lngCol < 0 Then Exit Function
    
    '如果是隐藏列，则read为true
    If mCurFlexGrid.ColHidden(lngCol) Then
        IsReadCol = True
        Exit Function
    End If
    
    strColProperty = mCurFlexGrid.Cell(flexcpData, 0, lngCol)
    
    IsReadCol = IIf(InStr(1, strColProperty, "," & mstrColProp_Read) >= 1, True, False)
End Function

Public Sub RefreshReadColColor()
'刷新不能编辑列的颜色
On Error Resume Next

    Dim i As Long
    Dim j As Long
    Dim strColProperty As String
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    If mCurFlexGrid.Editable = flexEDNone Then Exit Sub
    
    For i = 0 To mCurFlexGrid.Cols - 1
        If IsReadCol(i) Then
            mCurFlexGrid.Cell(flexcpBackColor, 1, i, mCurFlexGrid.Rows - 1, i) = mlngDisableColor
        End If
    Next i
End Sub


Public Sub ShowObject(obj As Object, ByVal lngRow As Long, ByVal lngCol As Long, Optional ByVal postionType As ObjPostionType = optRight)
'在指定单元格显示对象
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    obj.Left = mCurFlexGrid.Left + mCurFlexGrid.Cell(flexcpLeft, lngRow, lngCol) + mCurFlexGrid.Cell(flexcpWidth, lngRow, lngCol) - obj.Width + 10
    obj.Top = mCurFlexGrid.Top + mCurFlexGrid.Cell(flexcpTop, lngRow, lngCol) + 25
    obj.Height = mCurFlexGrid.Cell(flexcpHeight, lngRow, lngCol) - 15
    
    obj.Visible = True
End Sub


Public Sub ShowObjectWithEditCell(obj As Object, Optional ByVal postionType As ObjPostionType = optRight)
'在当前编辑单元格显示对象
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    Call ShowObject(obj, mCurFlexGrid.Row, mCurFlexGrid.Col, postionType)
End Sub


'Public Sub AppendButton(cmdApp As Object, ByVal lngRowIndex As Long, ByVal lngColIndex As Long)
''附加按钮到Grid中
'    If lngColIndex >= mCurFlexGrid.cols Then Exit Sub
'
'    cmdApp.Left = mCurFlexGrid.Cell(flexcpLeft, lngRowIndex, lngColIndex) + 10
'    cmdApp.Width = mCurFlexGrid.Cell(flexcpWidth, lngRowIndex, lngColIndex) - 10
'    cmdApp.Top = mCurFlexGrid.Cell(flexcpTop, lngRowIndex, lngColIndex) + 10
'    cmdApp.Height = mCurFlexGrid.Cell(flexcpHeight, lngRowIndex, lngColIndex) - 10
'End Sub


Public Sub SetRowBackColor(ByVal lngRow As Long, ByVal Col As OLE_COLOR)
'设置行的背景颜色
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    If mblnIsAllowExtCol Then
        mCurFlexGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, mCurFlexGrid.Cols - 1) = Col
    Else
        mCurFlexGrid.Cell(flexcpBackColor, lngRow, 0, lngRow, mCurFlexGrid.Cols - 1) = Col
    End If
    
    Call RefreshReadColColor
End Sub


Public Sub SetCurRowBackColor(ByVal Col As OLE_COLOR)
'设置当前行的背景颜色
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    Call SetRowBackColor(mCurFlexGrid.RowSel, Col)
End Sub


Public Sub SetColBackColor(ByVal strColName As String, ByVal Col As OLE_COLOR)
'设置列的背景颜色
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    mCurFlexGrid.Cell(flexcpBackColor, 1, GetColumnIndex(strColName), mCurFlexGrid.Rows - 1, GetColumnIndex(strColName)) = Col
End Sub


Public Sub SetCurColBackColor(ByVal Col As OLE_COLOR)
'设置当前列的背景颜色
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    mCurFlexGrid.Cell(flexcpBackColor, 1, mCurFlexGrid.ColSel, mCurFlexGrid.Rows - 1, mCurFlexGrid.ColSel) = Col
End Sub

Public Function ColEnable(ByVal lngCol As Long) As Boolean
'读取列可编辑状态
    ColEnable = False
    If mCurFlexGrid Is Nothing Then Exit Function
    
    ColEnable = ColEnableWithColName(GetColumnName(lngCol))
End Function


Public Function ColEnableWithColName(ByVal strColName As String) As Boolean
'读取列可编辑状态
    Dim i As Long
    
    ColEnableWithColName = False
    If mCurFlexGrid Is Nothing Then Exit Function
    
    '如果是隐藏列，则不允许编辑
    If mCurFlexGrid.ColHidden(GetColumnIndex(strColName)) Then
        ColEnableWithColName = False
        Exit Function
    End If
    
    For i = LBound(mCols) To UBound(mCols)
        If UCase(mCols(i).ColName) = UCase(strColName) Then
            ColEnableWithColName = mCols(i).IsRead
            Exit Function
        End If
    Next i
End Function


Private Function IsCheckCol(ByVal lngCol As Long) As Boolean
'临时方法，可将该类中的方法合并到ucFlexGrid单元中

'判断指定列是否为check列
    Dim strColProperty As String
    
    IsCheckCol = False
    
    If mCurFlexGrid Is Nothing Then Exit Function
    If lngCol < 0 Then Exit Function
    
    '如果是隐藏列，则check为false
    If mCurFlexGrid.ColHidden(lngCol) Then
        IsCheckCol = False
        Exit Function
    End If
    
    strColProperty = mCurFlexGrid.Cell(flexcpData, 0, lngCol)
    
    IsCheckCol = IIf(InStr(1, UCase(strColProperty), "," & UCase(mstrColProp_Check)) >= 1 Or _
                        InStr(1, UCase(strColProperty), "," & UCase(mstrColProp_RowCheck)) >= 1 Or _
                        InStr(1, UCase(strColProperty), UCase(mstrCheckColName)) >= 1, True, False)
End Function


Public Sub EditNextCell(ByVal lngRow As Long, Optional ByVal blnAutoNextRow As Boolean = True)
'编辑下一列
    If mCurFlexGrid Is Nothing Then Exit Sub
    If mCurFlexGrid.Editable = flexEDNone Then Exit Sub
    
    Do While mCurFlexGrid.ColSel + 1 < mCurFlexGrid.Cols
        If Not IsReadCol(mCurFlexGrid.ColSel + 1) Then
            Exit Do
        Else
            Call mCurFlexGrid.Select(lngRow, mCurFlexGrid.ColSel + 1)
        End If
    Loop
    
nextCell:
    
    If mCurFlexGrid.ColSel + 1 >= mCurFlexGrid.Cols Then
        If blnAutoNextRow Then
            Dim iRow As Long
            Dim iCol As Long
            
            iRow = GetNextRowIndex(lngRow)
            
            If iRow > 0 Then
                For iCol = IIf(mblnIsAllowExtCol, 1, 0) To mCurFlexGrid.Cols - 1
                    If Not IsReadCol(iCol) Then
                        If Not mCurFlexGrid.ColHidden(iCol) Then Exit For
                    End If
                Next iCol
                
                If iRow < mCurFlexGrid.Rows Then
                    If iCol = mCurFlexGrid.Cols Then iCol = mCurFlexGrid.Cols - 1
                    
                    Call mCurFlexGrid.Select(iRow, iCol)
                    Call mCurFlexGrid.ShowCell(iRow, iCol)
                End If
            End If
            
            If Not IsCheckCol(iCol) Then Call mCurFlexGrid.EditCell
        End If
        
        Exit Sub
    End If
    
    
    
    Call mCurFlexGrid.Select(lngRow, mCurFlexGrid.ColSel + 1)
        
    If Not IsCheckCol(iCol) Then Call mCurFlexGrid.EditCell
End Sub


Public Sub EditNextCellWithCurRow(Optional ByVal blnAutoNextRow As Boolean = True)
'编辑当前行的下一列
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    Call EditNextCell(mCurFlexGrid.RowSel, blnAutoNextRow)
End Sub


Public Sub RestoreList(Optional ByVal blnKeepRowCount As Boolean = True)
'恢复列表数据
    Dim R As Long
    Dim c As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    R = mCurFlexGrid.Rows - 1
    If R = 0 Then Exit Sub
    
    While R > 0
        mCurFlexGrid.RowData(R) = RowState.Normal
        
        '恢复修改前得数据
        For c = 0 To mCurFlexGrid.Cols - 1
            If Not mCurFlexGrid.ColHidden(c) Then
                mCurFlexGrid.TextMatrix(R, c) = mCurFlexGrid.Cell(flexcpData, R, c)
            End If
        Next c
        
        '恢复删除过的数据
        If mCurFlexGrid.RowHidden(R) Then
            If KeyEmpty(R) Then
                Call mCurFlexGrid.RemoveItem(R)
            Else
                mCurFlexGrid.RowHidden(R) = False
            End If
        End If
               
        
        R = R - 1
    Wend
    
    If mblnIsAllowExtCol Then
        mCurFlexGrid.Cell(flexcpBackColor, 1, 1, mCurFlexGrid.Rows - 1, mCurFlexGrid.Cols - 1) = mCurFlexGrid.BackColor
    Else
        mCurFlexGrid.Cell(flexcpBackColor, 1, 0, mCurFlexGrid.Rows - 1, mCurFlexGrid.Cols - 1) = mCurFlexGrid.BackColor
    End If
    
    If blnKeepRowCount Then mCurFlexGrid.Rows = IIf(mlngKeepRows <= -1, mCurFlexGrid.Rows, mlngKeepRows)
    
    Call LocateRow(1)
End Sub


Public Sub ApplyNormalState()
'更新列表的应用状态为normal
    Dim i As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    For i = 1 To mCurFlexGrid.Rows - 1
        mCurFlexGrid.RowData(i) = RowState.Normal
    Next i
End Sub






Public Function GetNextRowIndex(ByVal lngRow As Long) As Long
'取得下一行的索引
    Dim i As Long
    
    GetNextRowIndex = -1
    If mCurFlexGrid Is Nothing Then Exit Function
    
    For i = lngRow + 1 To mCurFlexGrid.Rows - 1
        If Not mCurFlexGrid.RowHidden(i) Then
            GetNextRowIndex = i
            Exit Function
        End If
    Next i
    
    If GetNextRowIndex = -1 Then
        i = lngRow - 1
        Do While i > 0
            If Not mCurFlexGrid.RowHidden(i) Then
                GetNextRowIndex = i
                Exit Function
            End If
            
            i = i - 1
        Loop
    End If
End Function



Public Sub DelRow(ByVal lngRow As Long, Optional ByVal blnKeepRowCount As Boolean = True)
'删除指定行
    Dim iNextIndex As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    mCurFlexGrid.RowHidden(lngRow) = True
    iNextIndex = GetNextRowIndex(lngRow)
    
    
    '重新设置行状态
    If mCurFlexGrid.RowData(lngRow) = RowState.Add Then
       mCurFlexGrid.RowData(lngRow) = RowState.Normal
    Else
        mCurFlexGrid.RowData(lngRow) = RowState.Del
    End If
    
    If blnKeepRowCount Then mCurFlexGrid.Rows = mCurFlexGrid.Rows + 1
    
    If iNextIndex > 0 Then Call LocateRow(iNextIndex)
End Sub



Public Sub DelCurRow(Optional ByVal blnKeepRowCount As Boolean = True)
'删除当前行
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    Call DelRow(mCurFlexGrid.RowSel, blnKeepRowCount)
End Sub




Public Function KeyEmpty(ByVal lngRow As Long) As Boolean
'检查key是否为空
    KeyEmpty = True
    If mCurFlexGrid Is Nothing Then Exit Function
    
    If Trim(mCurFlexGrid.TextMatrix(lngRow, GetColumnIndex(mstrKey))) <> "" Then
        KeyEmpty = False
    End If
End Function


Public Function CurKeyEmpty() As Boolean
'检查当前Key是否为空
    CurKeyEmpty = True
    If mCurFlexGrid Is Nothing Then Exit Function
    
    CurKeyEmpty = KeyEmpty(mCurFlexGrid.RowSel)
End Function


Public Sub SetKeyValue(ByVal lngRow As Long, ByVal strValue As String)
'设置关键值
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    mCurFlexGrid.TextMatrix(lngRow, GetColumnIndex(mstrKey)) = strValue
End Sub


Public Sub SetCurKeyValue(ByVal strValue As String)
'设置当前关键值
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    Call SetKeyValue(mCurFlexGrid.RowSel, strValue)
End Sub

Public Function GetKeyValue(ByVal lngRow As Long) As String
'取得行关键值
    GetKeyValue = ""
    If mCurFlexGrid Is Nothing Then Exit Function
    
    GetKeyValue = mCurFlexGrid.TextMatrix(lngRow, GetColumnIndex(mstrKey))
End Function


Public Function GetCurKeyValue() As String
'取得当前行关键值
    GetCurKeyValue = ""
    If mCurFlexGrid Is Nothing Then Exit Function
    
    If mCurFlexGrid.Rows = 1 Then Exit Function
    If mCurFlexGrid.RowSel <= 0 Then Exit Function
    
    GetCurKeyValue = GetKeyValue(mCurFlexGrid.RowSel)
End Function


Public Function CheckIsNullRow(ByVal lngRow As Long) As Boolean
'检查指定行是否为空行
    Dim i As Long
    
    CheckIsNullRow = True
    If mCurFlexGrid Is Nothing Then Exit Function
    
    If mblnIsAllowExtCol Then
        CheckIsNullRow = IIf(Len(mCurFlexGrid.Cell(flexcpText, lngRow, 1, lngRow, mCurFlexGrid.Cols - 1)) = mCurFlexGrid.Cols - 2, True, False)
    Else
        CheckIsNullRow = IIf(Len(mCurFlexGrid.Cell(flexcpText, lngRow, 0, lngRow, mCurFlexGrid.Cols - 1)) = mCurFlexGrid.Cols - 1, True, False)
    End If
    
'    For i = 0 To mCurFlexGrid.cols - 1
'        If mCurFlexGrid.TextMatrix(lngRow, i) <> "" Then
'            CheckIsNullRow = False
'            Exit Function
'        End If
'    Next i
End Function


Public Function CheckIsNullRowWithCurRow() As Boolean
'检查当前行是否为空行
    CheckIsNullRowWithCurRow = True
    If mCurFlexGrid Is Nothing Then Exit Function
    
    CheckIsNullRowWithCurRow = CheckIsNullRow(mCurFlexGrid.RowSel)
End Function


Public Function CheckIsUpdate(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'检查数据是否更新
    Dim blnUpdate As Boolean
    
    blnUpdate = False
    If mCurFlexGrid Is Nothing Then Exit Function
    
    If mCurFlexGrid.TextMatrix(lngRow, lngCol) <> mCurFlexGrid.Cell(flexcpData, lngRow, lngCol) Then
        blnUpdate = True
    End If
    
    CheckIsUpdate = blnUpdate
End Function


Public Function CheckIsUpdateWithCurRow(ByVal lngCol As Long) As Boolean
'检查当前行数据是否有更新
    CheckIsUpdateWithCurRow = False
    If mCurFlexGrid Is Nothing Then Exit Function
    
    Call CheckIsUpdate(mCurFlexGrid.RowSel, lngCol)
End Function


Public Sub UpdateRowState(ByVal lngRow As Long, Optional ByVal rsState As RowState = Normal)
'更新行状态
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    mCurFlexGrid.RowData(lngRow) = rsState
End Sub


Public Sub UpdateCurRowState(Optional ByVal rsState As RowState = Normal)
'更新当前行状态
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    Call UpdateRowState(mCurFlexGrid.RowSel, rsState)
End Sub


Public Function GetRowState(ByVal lngRow As Long) As RowState
'取得行状态
    GetRowState = RowState.Normal
    If mCurFlexGrid Is Nothing Then Exit Function
    
    GetRowState = mCurFlexGrid.RowData(lngRow)
End Function


Public Function GetCurRowState() As RowState
'取得当前行状态
    GetCurRowState = RowState.Normal
    If mCurFlexGrid Is Nothing Then Exit Function
    
    Call GetRowState(mCurFlexGrid.RowSel)
End Function


Public Sub RestoreRowText(ByVal lngRow As Long)
'恢复文本数据
    Dim i As Integer
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    For i = 0 To mCurFlexGrid.Cols - 1
        If Not mCurFlexGrid.ColHidden(i) Then
            mCurFlexGrid.TextMatrix(lngRow, i) = mCurFlexGrid.Cell(flexcpData, lngRow, i)
        End If
    Next i
End Sub


Public Sub RestoreCurRowText()
'恢复当前文本数据
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    Call RestoreRowText(mCurFlexGrid.RowSel)
End Sub




Private Function GetCharCount(ByVal strSource As String, ByVal strChar As String) As Long
'获取相同字符数量
    GetCharCount = Len(strSource) - Len(Replace(strSource, strChar, ""))
End Function


Public Sub AutoDataRowHeight(ByVal lngRow As Long, ByVal strText As String)
    Dim lngCharCount As Long
    Dim lngCharHeight As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    
    If mlngOldDataRowHeight <= 0 Then mlngOldDataRowHeight = mCurFlexGrid.RowHeight(lngRow)
    
    lngCharCount = GetCharCount(strText, vbCr)
    
    If lngCharCount = 0 Then Exit Sub
    
    lngCharHeight = mlngOldDataRowHeight * (lngCharCount + 1)
    
    If lngCharHeight > mCurFlexGrid.RowHeight(lngRow) Then
        mCurFlexGrid.RowHeight(lngRow) = lngCharHeight
    End If
End Sub


Public Sub SetText(ByVal lngRow As Long, ByVal lngCol As Long, ByVal strText As String)
'根据行列信息设置当前值
    Dim iCol As Long
    Dim strColProperty As String
    Dim strValue As String
    
    If mCurFlexGrid Is Nothing Then Exit Sub
        
    If lngCol <= -1 Then Exit Sub
    

    strColProperty = mCurFlexGrid.Cell(flexcpData, 0, lngCol)
    strValue = FormatValue(strColProperty, strText)
    
    '设置数据值
    mCurFlexGrid.TextMatrix(lngRow, lngCol) = strValue
    '对行单元格的data赋值，可用于以后直接判断该单元格是否进行更新
    mCurFlexGrid.Cell(flexcpData, lngRow, lngCol) = strValue
    
    
    If Not mCurFlexGrid.ColHidden(lngCol) Then
        '调整数据行高度
        Call AutoDataRowHeight(lngRow, strText)
    End If
End Sub


Public Sub SetTextWithColName(ByVal lngRow As Long, ByVal strCol As String, ByVal strText As String)
'根据行列信息设置当前值
    Dim iCol As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
        
    iCol = GetColumnIndex(strCol)
    If iCol <= -1 Then Exit Sub

    Call SetText(lngRow, iCol, strText)
End Sub


Public Function FormatValue(ByVal strFormats As String, ByVal strValue As String) As String
    Dim strColProperty As String
    Dim lngAlign As Long
    Dim lngCount As Long
    Dim strFill As String
    
    
    strColProperty = strFormats
    
    '格式化时间样式
    If IsDate(strValue) Then
        If InStr(1, strColProperty, ",fulldatetime") >= 1 Then
            strValue = Format(strValue, "yyyy-mm-dd hh:mm:ss")
        ElseIf InStr(1, strColProperty, ",onlydate") >= 1 Then
            strValue = Format(strValue, "yyyy-mm-dd")
        ElseIf InStr(1, strColProperty, ",onlytime") >= 1 Then
            strValue = Format(strValue, "hh:mm:ss")
        ElseIf InStr(1, strColProperty, ",shortdatetime") >= 1 Then
            strValue = Format(strValue, "yyyy-mm-dd hh:mm")
        End If
    End If
    
    '补齐字符长度
    lngAlign = InStr(1, UCase(strColProperty), UCase(mstrColProp_Align & "<"))
    If lngAlign > 0 Then
        strColProperty = Mid(strColProperty, lngAlign + Len(mstrColProp_Align) + 1, 10)
        strColProperty = Mid(strColProperty, 1, InStr(1, UCase(strColProperty), ">") - 1)
        
        lngCount = Val(strColProperty)
        strFill = Mid(strColProperty, InStr(1, UCase(strColProperty), ",") + 1, 3)
        
        If Len(strValue) < lngCount Then
            strValue = Lpad(strValue, lngCount, strFill)
        End If
    End If
    
    FormatValue = strValue
End Function


Public Sub SetCurText(ByVal lngCol As Long, ByVal strText As String)
'根据当前选择行，设置指定列值

    If mCurFlexGrid Is Nothing Then Exit Sub
    
    Call SetText(mCurFlexGrid.RowSel, lngCol, strText)
End Sub


Public Sub SetCurTextWithColName(ByVal strCol As String, ByVal strText As String)
'根据当前选择行，设置指定列值
    Dim iCol As Long
        
    If mCurFlexGrid Is Nothing Then Exit Sub
        
    iCol = GetColumnIndex(strCol)
    If iCol <= -1 Then Exit Sub
        
    Call SetText(mCurFlexGrid.RowSel, iCol, strText)
End Sub


Public Function CheckEquateValue(ByVal lngRow As Long, ByVal lngCol As Long) As String
'检查相同的值，如果有相同的，则返回新值
    Dim strValue As String
    Dim i As Long
    Dim num As Long
    Dim maxNum As Long
    Dim blnFind As Boolean
    
    maxNum = 0
    num = 0
    blnFind = False
    
    CheckEquateValue = ""
    If mCurFlexGrid Is Nothing Then Exit Function
    
    strValue = mCurFlexGrid.TextMatrix(lngRow, lngCol)
    If strValue <> "" Then
        For i = 1 To mCurFlexGrid.Rows - 1
            If mCurFlexGrid.TextMatrix(i, lngCol) <> "" And i <> lngRow And Not mCurFlexGrid.RowHidden(i) Then
                If mCurFlexGrid.TextMatrix(i, lngCol) Like strValue & "*" Then
                    num = Val(Substr(mCurFlexGrid.TextMatrix(i, lngCol), _
                        InStr(1, mCurFlexGrid.TextMatrix(i, lngCol), strValue & "_") + LenB(StrConv(strValue & "_", vbFromUnicode)), 3))
                    
                    If num >= maxNum Then maxNum = num + 1
                End If
                
                If mCurFlexGrid.TextMatrix(i, lngCol) = strValue Then
                    blnFind = True
                End If
            End If
        Next i
    End If
    
    If maxNum > 0 Then CheckEquateValue = IIf(blnFind, strValue & "_" & maxNum, "")
End Function


Public Function GetColumnName(ByVal lngColIndex As Long) As String
'获取列名称
    GetColumnName = ""
    If mCurFlexGrid Is Nothing Then Exit Function
    
    GetColumnName = mCurFlexGrid.ColKey(lngColIndex)
End Function


Public Function GetFieldName(ByVal lngColIndex As Long) As String
'取得数据库字段名称
    Dim strColProperty As String
    
    GetFieldName = ""
    If mCurFlexGrid Is Nothing Then Exit Function
    
    strColProperty = mCurFlexGrid.Cell(flexcpData, 0, lngColIndex)
    
    GetFieldName = Mid(strColProperty, 1, InStr(1, strColProperty, mstrPropertySplitChar) - 1)
End Function


Public Function GetColumnIndex(ByVal strColName As String) As Long
'获取列索引
    GetColumnIndex = -1
    If mCurFlexGrid Is Nothing Then Exit Function
    
    GetColumnIndex = mCurFlexGrid.ColIndex(strColName)
End Function


Public Function GetColDateTimeFormat(ByVal lngCol As Long) As String
'获取日期时间格式
    Dim strColProperty As String
    
    GetColDateTimeFormat = ""
    If mCurFlexGrid Is Nothing Then Exit Function
    
    strColProperty = mCurFlexGrid.Cell(flexcpData, 0, lngCol)
    
    If InStr(1, strColProperty, ",fulldatetime") >= 1 Then
        GetColDateTimeFormat = "yyyy-mm-dd hh:mm:ss"
    ElseIf InStr(1, strColProperty, ",onlydate") >= 1 Then
        GetColDateTimeFormat = "yyyy-mm-dd"
    ElseIf InStr(1, strColProperty, ",onlytime") >= 1 Then
        GetColDateTimeFormat = "hh:mm:ss"
    ElseIf InStr(1, strColProperty, ",shortdatetime") >= 1 Then
        GetColDateTimeFormat = "yyyy-mm-dd hh:mm"
    Else
        GetColDateTimeFormat = ""
    End If
End Function




Public Function GetText(ByVal lngRow As Long, ByVal lngCol As Long) As String
'根据列索引获取显示内容
    GetText = ""
    If mCurFlexGrid Is Nothing Then Exit Function
    
    GetText = mCurFlexGrid.TextMatrix(lngRow, lngCol)
End Function


Public Function GetTextWithColName(ByVal lngRow As Long, ByVal strCol As String) As String
'根据列名称获取显示内容
    Dim iCol As Long
    
    GetTextWithColName = ""
    If mCurFlexGrid Is Nothing Then Exit Function
    
    iCol = GetColumnIndex(strCol)
    If iCol <= -1 Then Exit Function
    
    GetTextWithColName = mCurFlexGrid.TextMatrix(lngRow, iCol)
End Function


Public Function GetCurText(ByVal lngCol As Long) As String
'取得当前选定行中指定列的显示内容
    GetCurText = ""
    If mCurFlexGrid Is Nothing Then Exit Function
    
    GetCurText = mCurFlexGrid.TextMatrix(mCurFlexGrid.RowSel, lngCol)
End Function


Public Function GetCurTextWithColName(ByVal strCol As String) As String
'取得当前选定行中指定列的显示内容
    Dim iCol As Long
    
    GetCurTextWithColName = ""
    If mCurFlexGrid Is Nothing Then Exit Function
    
    iCol = GetColumnIndex(strCol)
    If iCol <= -1 Then Exit Function
    
    GetCurTextWithColName = mCurFlexGrid.TextMatrix(mCurFlexGrid.RowSel, iCol)
End Function


Public Sub ShowNullDataFace(vfgList As Object)
'显示无数据的界面
    vfgList.Cols = 10
    vfgList.Rows = 10
    vfgList.FixedCols = 1
    vfgList.FixedRows = 1
End Sub



Public Function GetNullRowIndex() As Long
'获取空数据行索引
    Dim i As Long
    
    GetNullRowIndex = -1
    If mCurFlexGrid Is Nothing Then Exit Function
    
    For i = 1 To mCurFlexGrid.Rows - 1
        If Not mCurFlexGrid.RowHidden(i) Then
            If CheckIsNullRow(i) Then
                GetNullRowIndex = i
                Exit Function
            End If
        End If
    Next i
End Function



Public Sub CopyRowData(ByVal lngSourceRow As Long, ByVal lngTargetRow As Long)
'复制行数据
    Dim i As Long
    
    For i = 0 To mCurFlexGrid.Cols - 1
        mCurFlexGrid.TextMatrix(lngTargetRow, i) = mCurFlexGrid.TextMatrix(lngSourceRow, i)
        mCurFlexGrid.Cell(flexcpData, lngTargetRow, i) = mCurFlexGrid.Cell(flexcpData, lngSourceRow, i)
    Next i
    
End Sub


Public Function AddRow() As Long
'添加行并返回行索引
    AddRow = -1
    If mCurFlexGrid Is Nothing Then Exit Function
    
    mCurFlexGrid.Rows = mCurFlexGrid.Rows + 1
    
    mCurFlexGrid.RowSel = mCurFlexGrid.Rows - 1
    Call mCurFlexGrid.ShowCell(mCurFlexGrid.RowSel, 0)
    
    '添加行状态
    mCurFlexGrid.RowData(mCurFlexGrid.RowSel) = RowState.Normal
    
    AddRow = mCurFlexGrid.RowSel
End Function



Public Sub LocateRow(Optional ByVal lngRowIndex As Long = -1)
'定位指定行，默认定位为最后一位
    Dim lngRow As Long
    Dim iCol As Long
    
    If mCurFlexGrid Is Nothing Then Exit Sub
    If mCurFlexGrid.Rows <= 1 Then Exit Sub
    
    lngRow = lngRowIndex
    If lngRow < 0 Then
        lngRow = mCurFlexGrid.Rows - 1
    End If
    
    '取得第一个未隐藏的列
    For iCol = 0 To mCurFlexGrid.Cols - 1
        If Not mCurFlexGrid.ColHidden(iCol) Then Exit For
    Next iCol
    
    Call mCurFlexGrid.Select(lngRow, iCol)
    Call mCurFlexGrid.ShowCell(lngRow, iCol)
End Sub



Public Function IsRowSelected() As Boolean
'判断是否有数据行被选择
    IsRowSelected = False
    If mCurFlexGrid Is Nothing Then Exit Function
    
    If mCurFlexGrid.Rows <= 1 Then Exit Function
    
    If mCurFlexGrid.RowHidden(mCurFlexGrid.RowSel) = True Then Exit Function
    
    IsRowSelected = IIf(mCurFlexGrid.RowSel <= 0 Or mCurFlexGrid.RowSel > mCurFlexGrid.Rows - 1, False, True)
End Function



Public Function IsRowChecked() As Boolean
'判断是否有数据行被Check
    Dim i As Long
    
    IsRowChecked = False
    If mCurFlexGrid Is Nothing Then Exit Function
    
    For i = 1 To mCurFlexGrid.Rows - 1
        If RowChecked(i) Then
            IsRowChecked = True
            
            Exit Function
        End If
    Next i
End Function




Private Sub Class_Initialize()
    Set mCurFlexGrid = Nothing
    
    mstrKey = ""
    mstrKeyField = ""
    mlngKeepRows = -1
    mblnIsAllowExtCol = True
    mlngOldDataRowHeight = -1
    mblnIsShowNumber = True
    mlngDisableColor = &HC0FFFF
End Sub



