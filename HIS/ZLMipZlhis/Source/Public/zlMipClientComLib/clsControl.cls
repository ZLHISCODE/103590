VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------
'模块名称：clsControl
'包含对象：操作控件的函数
'命名规则：控件简名+操作类型 + 操作属性
'说明    ：控件简名按编码规范中的书写；操作类型一般为动词；操作属性为控件属性，对于涉及多个或不好确定的，可以用便于理解的词汇替代。
'例      ：LvwSortColumn，LvwFlatColumnHeader，FrmCheckInput
'--------------------------------------------------------------------------
Public Enum mTextAlign
    taLeftAlign = 0
    taCenterAlign = 1
    taRightAlign = 2
End Enum

Public Enum mTextType
    m文本式 = 0
    m数字式 = 1
    m金额式 = 2
    m负金额式 = 3
End Enum

Public Sub CboSetIndex(ByVal hWnd_combo As Long, ByVal lngIndex As Long)
'功能：设置Combo控件的Index值
'为一个Combo控件选择列表项，但又不触发其Click事件
    Const CB_SETCURSEL = &H14E
    
    SendMessage hWnd_combo, CB_SETCURSEL, lngIndex, 0
End Sub

Public Sub CboSetWidth(ByVal hWnd_combo As Long, ByVal lngWidth As Long)
'功能：设置Combo控件下拉列表的宽度
'此处的宽度是批下拉列表的宽度，并且是以TWIP为单位
    Const CB_SETDROPPEDWIDTH As Long = &H160

    SendMessage hWnd_combo, CB_SETDROPPEDWIDTH, lngWidth / Screen.TwipsPerPixelX, 0
End Sub

Public Sub CboSetHeight(cboControl As Object, ByVal lngHeight As Long)
'功能：设置Combo控件的下拉列表的高度
'此处的宽度是批下拉列表的高度，并且是以TWIP为单位
    SetWindowPos cboControl.hWnd, 0, 0, 0, cboControl.Width / Screen.TwipsPerPixelX, lngHeight / Screen.TwipsPerPixelY, SWP_NOMOVE
End Sub

Public Sub CboSetText(cboControl As Object, ByVal strText As String, Optional ByVal blnAfter As Boolean = True, Optional strSplit As String = "-")
'功能：根据文本串更新Combo控件的当前值
'参数：cboControl  准备设置的ComboBox控件
'      strText     输入的文本串
'      blnAfter    表示在分隔符之前或之后取值。如果没有分隔符，则取之后
'      strSplit    分隔符，通常为-
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cboControl.ListCount - 1
        strTemp = cboControl.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            '直接返回整个字符串
            If strText = strTemp Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                '圆点之前
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '已经找到
        cboControl.ListIndex = lngCount
    Else
        If blnAfter = True Then
            '这才是实际内容，如果为前则只是编码
            If strText <> "" Then
                cboControl.AddItem strText
                cboControl.ListIndex = cboControl.NewIndex
            End If
        End If
    End If
End Sub

Public Sub CboAppendText(cboControl As Object, KeyAscii As Integer)
'功能：对ComboBox实现输入过程中自动完成的功能
'说明：在Combox.KeyPress事件中调用
    Dim strInput As String
    Dim lngStart As Long
    Dim lngIndex As Long
    Const CB_FINDSTRING = &H14C
    
    If cboControl.STYLE <> 0 Then Exit Sub
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    strInput = Chr(KeyAscii): KeyAscii = 0

    With cboControl
        '记录上次的插入点位置
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)

        '接着得到用户击键完成后文本框中出现的内容
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)

        '根据假想的内容得到可能的列表项
        lngIndex = SendMessage(cboControl.hWnd, CB_FINDSTRING, -1, ByVal strInput)
        If lngIndex >= 0 Then
            '.ListIndex = lngIndex
            .Text = .List(lngIndex)
            .SelStart = Len(strInput)
            .SelLength = 100
        Else
            .Text = strInput
            .SelStart = lngStart
        End If
    End With
End Sub

Public Function CboLocate(ByVal cboObj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
    'blnItem:True-表示根据ItemData的值定位下拉框;False-表示根据文本的内容定位下拉框
    Dim lngLocate As Long
    CboLocate = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If blnItem Then
            If cboObj.ItemData(lngLocate) = Val(strValue) Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        Else
            If Mid(cboObj.List(lngLocate), InStr(1, cboObj.List(lngLocate), "-") + 1) = strValue Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        End If
    Next
End Function

Public Sub CboAddData(objSource As Object, ByVal rsDataSource As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
    '-----------------------------------------------------------------------------------
    '功能:装载数据入指定的组合下拉框或网格中的下拉框中
    '参数:objSource   要装载记录集的下拉框控件
    '     rsDataSource     记录集数据,只能有两个数据项,第一个是下拉数据,第二个是对应的关键字,如id
    '     blnClear    装载时是否清楚原有的下拉数据,缺省为True
    '返回:
    '-----------------------------------------------------------------------------------
    If blnClear = True Then objSource.Clear
    
    If rsDataSource.BOF = False Then
        rsDataSource.MoveFirst
        While Not rsDataSource.EOF
            objSource.AddItem rsDataSource.Fields(0).value
            objSource.ItemData(objSource.NewIndex) = Val(rsDataSource.Fields(1).value)
            rsDataSource.MoveNext
        Wend
        rsDataSource.MoveFirst
    End If
End Sub

Public Function CboMatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：lngHwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：匹配的索引(含不匹配的索引)
'说明：只返回索引,但未设置索引。请将该函数在KeyPress事件中调用。
    Static lngPreTime As Single, lngPreHwnd As Long, strFind As String
    Dim sngTime As Single, lngR As Long
    Dim lngCount As Long, strText As String * 255
    Dim lngIdx As Long, strTmp As String, i As Long
    
    If lngHwnd <> 0 Then
        If KeyAscii = 13 Then '回车不处理,直接返回当前索引
            CboMatchIndex = SendMessage(lngHwnd, CB_GETCURSEL, 0, 0)
            Exit Function
        Else
            '弹出列表内容
            If SendMessage(lngHwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 Then
                Call gobjComLib.zlCommFun.PressKey(vbKeyF4)
            End If
        End If
    End If
    
    If lngPreHwnd <> lngHwnd Or lngHwnd = 0 Then
        lngPreTime = Empty: strFind = Empty
    End If
    lngPreHwnd = lngHwnd
    
    If lngHwnd = 0 Then
        CboMatchIndex = -1: Exit Function
    End If
    
    sngTime = timer
    If Abs(sngTime - lngPreTime) > sngInterval Then '输入间隔(缺省为0.5秒)
        strFind = ""
    End If
    strFind = strFind & Chr(KeyAscii)
    lngPreTime = timer
    KeyAscii = 0 '使ComboBox本身的单字匹配功能失效
        
    '用API从开头进行查找
    lngIdx = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
    
    '无匹配再从中间按简码匹配
    If lngIdx = -1 Then
        lngCount = SendMessage(lngHwnd, CB_GETCOUNT, 0, 0)
        For i = 0 To lngCount - 1
            strText = String(SendMessage(lngHwnd, CB_GETLBTEXTLEN, i, 0), Chr(32)) & Chr(0)
            Call SendMessage(lngHwnd, CB_GETLBTEXT, i, ByVal strText)
            
            If InStr(strText, Chr(0)) > 0 Then
                strTmp = Mid(strText, 1, InStr(strText, Chr(0)) - 1)
            Else
                strTmp = strText
            End If
            
            If InStr(strTmp, "]") > 0 And InStr(strTmp, "-") = 0 Then
                strTmp = LTrim(Mid(strTmp, InStr(strTmp, "]") + 1))
            ElseIf InStr(strTmp, ")") > 0 And InStr(strTmp, "-") = 0 Then
                strTmp = LTrim(Mid(strTmp, InStr(strTmp, ")") + 1))
            Else
                strTmp = LTrim(Mid(strTmp, InStr(strTmp, "-") + 1))
            End If
            
            If gobjComLib.zlCommFun.SpellCode(strTmp) Like UCase(strFind) & "*" Then
                lngIdx = i: Exit For
            End If
        Next
    End If
    
    If lngIdx = -1 Then Beep
    CboMatchIndex = lngIdx
End Function

Public Function FormCheckInput(objForm As Object) As Boolean
'功能：检查指定窗体中是否有非法字符输入及输入长度是否超长
'参数：objForm=要检查的窗体
    Dim obj As Object, strText As String
    
    On Error Resume Next
    For Each obj In objForm.Controls
        If InStr("TextBox,ComboBox", TypeName(obj)) > 0 Then
            If obj.Visible And obj.Enabled And Not obj.Locked Then
                Select Case TypeName(obj)
                Case "TextBox"
                    strText = obj.Text
                Case "ComboBox"
                    If obj.STYLE = 0 Then strText = obj.Text
                End Select
                If InStr(strText, "'") > 0 _
                    Or InStr(strText, ",") > 0 _
                    Or InStr(strText, ";") > 0 _
                    Or InStr(strText, "|") > 0 _
                    Or InStr(strText, "~") > 0 _
                    Or InStr(strText, "^") > 0 Then
                    MsgBox "输入数据中包含非法字符！", vbInformation, gstrSysName
                    obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                    obj.SetFocus: Exit Function
                End If
            End If
        End If
    Next
    FormCheckInput = True
End Function

'Public Function LvwSelectColumns(objSet As Object, ByVal strColumn As String, Optional ByVal blnInit As Boolean = False) As Boolean
''功能:对列表控件的列进行设置
''参数:
''   objSet：要设置的对象,目前只支持ListView，以后再加上FlexGrid,DataGrid。
''   strColumn；列串。格式是"列名,列宽,对齐数值,列特性;列名,列宽,对齐数值,列特性"    注意列之间是用分号
''      比如 "名称,2000,0,1;编码,800,0,0;简码,1440,0,0"
''      对ListView而言：列特性为1表示该列不可删除，列特性为0表示该列可以删除
''      对FlexGridView而言：列特性还要表示是否属于固定列，以便不能和其它列进行顺序调整
''   blnInit：True,不显示选择窗口，直接初始化
'    Dim varColumns As Variant, varColumn As Variant
'    Dim lngCol As Long
'
'    If blnInit Then
'        varColumns = Split(strColumn, ";")
'        Select Case TypeName(objSet)
'            Case "ListView"
'                With objSet.ColumnHeaders
'                    .Clear
'                    For lngCol = LBound(varColumns) To UBound(varColumns)
'                        varColumn = Split(varColumns(lngCol), ",")
'                        .Add , "_" & varColumn(0), varColumn(0), varColumn(1), varColumn(2)
'                    Next
'                End With
'            Case "MSHFlexGrid"
'            Case "DataGrid"
'        End Select
'    Else
'        LvwSelectColumns = frmSelectColumn.设置列(objSet, strColumn)
'    End If
'End Function

Public Sub LvwFlatColumnHeader(ByVal lvw As Object)
'功能：使用ListView的列标题成为平面
    Const strHeaderClass As String = "msvb_lib_header"
    Const HDS_BUTTONS   As Long = 2
    
    Dim lngChild As Long, lngLen As Long, lngStyle As Long
    Dim strName As String * 255

    
    lngChild = GetWindow(lvw.hWnd, GW_CHILD)
    Do While lngChild <> 0
        lngLen = GetClassName(lngChild, strName, 255)
        If lngLen > 0 Then
            If Mid(strName, 1, lngLen) = strHeaderClass Then
                lngStyle = GetWindowLong(lngChild, GWL_STYLE)
                lngStyle = lngStyle And (Not HDS_BUTTONS)
                SetWindowLong lngChild, GWL_STYLE, lngStyle
                Exit Sub
            End If
        End If
        lngChild = GetWindow(lngChild, GW_HWNDNEXT)
    Loop

End Sub

Public Sub LvwSortColumn(lvw As Object, lngColumn As Long)
'功能：根据ListView列头的点击进行排序
'参数：lvw            ListView控件
'      lngColumn      当前点击的列头序号，即ColumnHeader.Index值
    If lvw.Sorted = False Then lvw.Sorted = True

    If lvw.SortKey = lngColumn - 1 Then '仍是刚才那列
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvw.SortKey = lngColumn - 1  '排序关键字是从0开始的
        lvw.SortOrder = lvwAscending
    End If
    
    If Not lvw.SelectedItem Is Nothing Then
        lvw.SelectedItem.EnsureVisible
    End If
End Sub

Public Function LvwSaveItem(lvwObj As Object) As String
    '-----------------------------------------------------------------------------------
    '功能:保存ListView中的选中项
    '参数:lvwobj      要保存的ListView控件
    '返回:选中项的关键字
    '-----------------------------------------------------------------------------------
    If Not (lvwObj.SelectedItem Is Nothing) Then LvwSaveItem = lvwObj.SelectedItem.Key
End Function

Public Sub LvwRestoreItem(lvwObj As Object, svrKey As String)
    '-----------------------------------------------------------------------------------
    '功能:恢复ListView中的选中项
    '参数:lvwobj      要恢复的ListView控件
    '     svrKey      要恢复的ListView项的关键字
    '返回:
    '-----------------------------------------------------------------------------------
    On Error GoTo EndP
    If lvwObj.ListItems.count > 0 Then
        If Not (lvwObj.ListItems(svrKey) Is Nothing) Then
            lvwObj.ListItems(svrKey).Selected = True
            lvwObj.ListItems(svrKey).EnsureVisible
        End If
    End If
    Exit Sub
EndP:
    If lvwObj.ListItems.count > 0 Then lvwObj.ListItems(1).Selected = True
End Sub

Public Sub LvwSetColWidth(objLvw As Object, Optional blnHideNullCol As Boolean)
'功能：根据ListView中当前的内容自动调整列为最小匹配宽度,并保持至少可以显示列头文字的宽度
'参数：objLvw=要调整的ListView对象
'      blnHideNullCol=是否隐藏没有任何数据的列
    Dim i As Integer, lngW As Long
    For i = 1 To objLvw.ColumnHeaders.count
        SendMessage objLvw.hWnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If blnHideNullCol Then If objLvw.ColumnHeaders(i).Width < 200 Then objLvw.ColumnHeaders(i).Width = 0
        If objLvw.ColumnHeaders(i).Width < (gobjComLib.zlCommFun.ActualLen(objLvw.ColumnHeaders(i).Text) + 2) * 90 And objLvw.ColumnHeaders(i).Width <> 0 Then
            objLvw.ColumnHeaders(i).Width = (gobjComLib.zlCommFun.ActualLen(objLvw.ColumnHeaders(i).Text) + 2) * 90
        End If
    Next
End Sub

Public Sub MshSetColWidth(objMsh As Object, frmParent As Object)
'功能：自动调整MSHFlexGrid表格的各列宽度,以最小匹配为准,并保持列头文字可以全部显示
'参数：msh=要调整列宽的MSHFlexGrid控件
'      frmParent=父窗体,用于计算文字宽度
    Dim arrWidth() As Long, blnRedraw As Boolean
    Dim lngRows As Long
    Dim i As Long, j As Long
    
    ReDim arrWidth(objMsh.Cols - 1)
    
    blnRedraw = objMsh.Redraw
    lngRows = objMsh.Rows - 1
    If lngRows > 200 Then lngRows = 200
    objMsh.Redraw = False
    For i = 0 To objMsh.Cols - 1
        For j = IIf(objMsh.FixedRows = 0, 0, objMsh.FixedRows - 1) To lngRows
            If Not objMsh.MergeRow(j) Then
                If frmParent.TextWidth(objMsh.TextMatrix(j, i)) + 90 > arrWidth(i) Then
                    arrWidth(i) = frmParent.TextWidth(objMsh.TextMatrix(j, i)) + 90
                End If
            End If
        Next
    Next
    For i = 0 To objMsh.Cols - 1
        objMsh.ColWidth(i) = IIf(arrWidth(i) > 3000, 3000, arrWidth(i))
    Next
    objMsh.Redraw = blnRedraw
End Sub

Public Sub MshColumnSort(ByVal mshGrd As Object, ByVal IntSorCol As Integer, _
    ByRef intPreCol As Integer, ByRef intPreSort As Integer, _
    Optional blnNum As Boolean = False, Optional intMachingCol As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对列进行排序
    '参数:
    '       mshGrd-MSHFlexGrid网络控件
    '       intPreCol-前一次单据头的排序列
    '       intPreSort-前一次单据头的排序
    '       blnNum-是否为数量列
    '       intMachingCol-匹配列,主要用于排序后重新定位到指定的列
    '---------------------------------------------------------------------------------------------------------------------------------------------------------
    Const flexSortNumericAscending  As Long = 3
    Const flexSortNumericDescending  As Long = 4
    Const flexSortStringNoCaseAscending As Long = 5
    Const flexSortStringNoCaseDescending As Long = 6
    
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strTmp As String
    
    Err = 0
    On Error GoTo errHand:
    With mshGrd
        If IntSorCol > .Cols - 1 Or IntSorCol < 0 Then Exit Sub
        If .Rows > 1 Then
            .Redraw = False
            intCol = IntSorCol   '.MouseCol
            .Col = intCol
            .ColSel = intCol
            If intMachingCol > .Cols - 1 Or intMachingCol < 0 Then
                strTmp = .TextMatrix(.Row, 0)
            Else
                strTmp = .TextMatrix(.Row, intMachingCol)
            End If
            If blnNum Then
                    If intCol = intPreCol And intPreSort = flexSortNumericDescending Then
                       .Sort = flexSortNumericAscending
                       intPreSort = flexSortNumericAscending
                    Else
                       .Sort = flexSortNumericDescending
                       intPreSort = flexSortNumericDescending
                    End If
            Else
                    If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       intPreSort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       intPreSort = flexSortStringNoCaseDescending
                    End If
            End If
            
            intPreCol = intCol
            If intMachingCol > .Cols - 1 Or intMachingCol < 0 Then
                .Row = MshGrdFindRow(mshGrd, strTmp, 0)
            Else
                .Row = MshGrdFindRow(mshGrd, strTmp, intMachingCol)
            End If
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
    Exit Sub
errHand:
    mshGrd.Redraw = True
End Sub

Public Function MshGrdFindRow(ByVal mshGrd As Object, _
        ByVal strTmp As Variant, ByVal intCol As Integer) As Integer
    '----------------------------------------------------------------------------
    '功能:寻找与某一单元值相等的行
    '参数:
    '       mshGrd-网格控件
    '       strTmp-匹配串
    '       intCol-匹配列
    '返回:查找到的行,如未找到,返位第一行
    '----------------------------------------------------------------------------
    
    Dim i As Integer
    Err = 0
    On Error GoTo errHand:
    
    With mshGrd
        For i = 1 To .Rows - 1
            If IsDate(strTmp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(strTmp, "yyyy-mm-dd") Then
                  MshGrdFindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = strTmp Then
                  MshGrdFindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    MshGrdFindRow = 1
    Exit Function
errHand:
    MshGrdFindRow = 1
    Exit Function
End Function

Public Sub MshClearData(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
    '-----------------------------------------------------------------------------------
    '功能: 清除指定网格的指定行指定列的数据
    '参数: obj=要操作的网格控件
    '      intRow=要清除的行号
    '      intCol=要清除的列号列表如Array(1,2,3),若所有列则可以表示为Array()
    '返回：
    '-----------------------------------------------------------------------------------
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

Public Sub MshSetFormat(ByVal msfObj As Object, ByVal strColumns As String, ByVal strCaption As String, _
Optional ByVal rsData As ADODB.Recordset, Optional ByVal blnData As Boolean = False, Optional ByVal blnRestore As Boolean = False)
    Dim intCol As Integer, intCols As Integer
    '各参数含义：msfObj是表格对象；strColumns是列名格式串（列名|宽度|列对齐方式）；blnData说明是设置格式还是绑定数据；blnRestore说明是否恢复原始宽度
    '列名格式：NO|1000|4,日期|1600|4,检查单位|2000|4,检查人|800|4,简要记录|2000|4,结论|1000|4
    
    '设置表格控件格式
    intCols = UBound(Split(strColumns, ",")) + 1
    msfObj.Redraw = False
    
    '设置单据头，单据体的格式
    If Not blnData Then
        With msfObj
            .Clear
            .Rows = 2
            .Cols = intCols
            
            For intCol = 0 To .Cols - 1
                .TextMatrix(0, intCol) = Split(Split(strColumns, ",")(intCol), "|")(0)
                .ColAlignmentFixed(intCol) = 4
            Next
            If blnRestore Then
                For intCol = 0 To .Cols - 1
                    .ColWidth(intCol) = Split(Split(strColumns, ",")(intCol), "|")(1)
                Next
                Call gobjComLib.RestoreFlexState(msfObj, strCaption)
            End If
        End With
    Else
        If Not rsData.EOF Then
            With msfObj
                Set .DataSource = rsData
                For intCol = 0 To .Cols - 1
                    .ColAlignment(intCol) = Split(Split(strColumns, ",")(intCol), "|")(2)
                Next
            End With
        End If
    End If
    
    With msfObj
        .Row = 1: .Col = 0
        .RowSel = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
End Sub

Public Sub VSFSetFontSize(ByRef vsf As Object, ByVal lngFontSize As Long, Optional ByVal lngCol As Long)
'功能：设置vsfflexgrid字体的大小，并自动调整列宽和行高
'参数：lngFontSize常用设置：9号即小五，12号即小四
'      lngCol,如果要设置自动调整行高，则传入行高依据的列，要求AutoSizeMode = flexAutoSizeRowHeight,WordWrap =True
    Dim i As Long, lngRate As Double, lngTmp As Long
    
    If lngFontSize < 5 And lngFontSize > 50 Then Exit Sub
    With vsf
        lngRate = lngFontSize / .FontSize
        lngTmp = .Redraw
        
        .Redraw = flexRDNone
        .FontSize = lngFontSize
        
        For i = 0 To .Cols - 1
           If .ColWidth(i) > 0 Then
             .ColWidth(i) = .ColWidth(i) * lngRate
           End If
        Next
        
        If .AutoSizeMode = flexAutoSizeRowHeight And .WordWrap And lngCol > 0 Then
            .AutoSize lngCol
        Else
            .RowHeightMin = .RowHeightMin * lngRate
            .RowHeight(-1) = .RowHeightMin
        End If
        .Redraw = lngTmp
    End With
End Sub

Public Sub PicShowFlat(objPic As Object, Optional intStyle As Integer = -1, Optional strName As String = "", Optional intAlign As mTextAlign)
'功能：将PictureBox模拟成凹下或凸起的面板
'参数：intStyle:0=平面,-1=凹下,1=凸起
'      intAlign=如果要显示文本,则指定对齐方式
    
    Dim vRect As RECT, lngTmp As Long
    
    With objPic
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            vRect.Left = .ScaleLeft
            vRect.Top = .ScaleTop
            vRect.Right = .ScaleWidth
            vRect.Bottom = .ScaleHeight
            DrawEdge .hDC, vRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If intAlign = taCenterAlign Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2 '中间对齐
            ElseIf intAlign = taRightAlign Then
                .CurrentX = .ScaleWidth - .TextWidth(strName) - 2 '右边对齐
            Else
                .CurrentX = 2 '左边对齐
            End If
            objPic.Print strName
        End If
    End With
End Sub

Public Sub StaShowPercent(sngPercent As Single, objPanel As Object, objForm As Object)
'功能：在状态条上根据百分比显示当前处理进度(█)
'参数：sngPercent=当前进度值,为0-1之间的一个小数
'      objPanel=状态条中用于显示进度的格子
'      objForm=用于计算文字宽度的窗体
    Dim intAll As Integer
    intAll = objPanel.Width / objForm.TextWidth("█") - 4
    objPanel.Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
End Sub

Public Sub TxtGetCaretPos(ByVal lngTxtHwnd As Long, lngLineNo As Long, lngColNo As Long)
'功能：获取光标在文本框中的行列位置
'参数：lngTxtHwnd=要获取光标位置的文本框句柄
'返回：lngLineNo=光标所在行号,以1开始
'      lngColNo=光标所在列号,以0开始
    Dim i As Long, j As Long, K As Long
    Dim lParam As Long, wParam As Long

    '首先向文本框传递EM_GETSEL消息以获取从起始位置到
    '光标所在位置的字符数
    i = SendMessage(lngTxtHwnd, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16
    
    '再向文本框传递EM_LINEFROMCHAR消息根据获得的字符
    '数确定光标以获取所在行数
    lngLineNo = SendMessage(lngTxtHwnd, EM_LINEFROMCHAR, j, 0) '
    lngLineNo = lngLineNo + 1
    
    '向文本框传递EM_LINEINDEX消息以获取所在列数
    K = SendMessage(lngTxtHwnd, EM_LINEINDEX, -1, 0)
    lngColNo = j - K + 1
End Sub

Public Sub TxtGetCursorPos(objMain As Object, objTxt As Object, lngX As Long, lngY As Long)
'功能：获取文本框中当前光标相对于该框的坐标
'参数：objMain=用于计算文字尺寸的窗体
'      objTXT=要获取光标坐标的文本框控件
'返回： lngX,lngY=光标坐标(Twip)
    Dim lngRow As Long, lngCol As Long
    Dim lngFirst As Long
    Dim lngRH As Long, lngCW As Long
    
    objMain.Font.Name = objTxt.Font.Name
    objMain.Font.Size = objTxt.Font.Size
    objMain.Font.Bold = objTxt.Font.Bold
    objMain.Font.Italic = objTxt.Font.Italic
    objMain.Font.Underline = objTxt.Font.Underline
    objMain.Font.Strikethrough = objTxt.Font.Strikethrough
    lngRH = objMain.TextHeight("A")
    lngCW = objMain.TextWidth("A") '是以半角计算列号
    
    lngFirst = SendMessage(objTxt.hWnd, EM_GETFIRSTVISIBLELINE, lngRow, lngCol) + 1 '以0行开始
    Call TxtGetCaretPos(objTxt.hWnd, lngRow, lngCol) '以1行,0列开始
    lngRow = lngRow - lngFirst + 1 '行偏移
    
    lngY = lngRow * lngRH
    lngX = lngCol * lngCW
End Sub

Public Sub TxtCheckKeyPress(ByVal objCtl As Object, KeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '功能:只能输入数字和回车及退格
    '参数:
    '   objctl:限制控件
    '   Keyascii:
    '           Keyascii:8 (退格)
    '   TextType:(0-文本式;1-数字式;2-金额式)
    '返回:一个KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo errHand:
    
    If TextType = m文本式 Then
        If KeyAscii = Asc("'") Then
            KeyAscii = 0
        End If
        Exit Sub
    End If

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        Select Case KeyAscii
        Case vbKeyReturn       '回车
        Case 8                 '退格
        Case Asc(".")
            If TextType = m金额式 Or TextType = m负金额式 Then
                If InStr(objCtl, ".") <> 0 Then     '只能存在一个小数点
                    KeyAscii = 0
                End If
            Else
                KeyAscii = 0
            End If
        Case Asc("-")          '负数
            Dim iRow As Long
            Dim iCol As Long
            If Trim(objCtl.Text) = "" Then Exit Sub
            If TextType <> m负金额式 Then KeyAscii = 0: Exit Sub
            If objCtl.SelStart <> 0 Then KeyAscii = 0: Exit Sub      '光标不存第一位,不能输入负数
            If InStr(1, objCtl.Text, "-") <> 0 Then   '只能存在一个负数
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
        End Select
    End If
    Exit Sub
errHand:
    KeyAscii = 0
End Sub

Public Sub TxtSelAll(objTxt As Object)
'功能：将编辑框的的文本全部选中
'参数：objTxt=需要全选的编辑控件,该控件具有SelStart,SelLength属性
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    If TypeName(objTxt) = "TextBox" Then
        If objTxt.MultiLine Then
            SendMessage objTxt.hWnd, WM_VSCROLL, SB_TOP, 0
        End If
    End If
End Sub

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hWnd, vRect)
    lngStyle = GetWindowLong(objForm.hWnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hWnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hWnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub
