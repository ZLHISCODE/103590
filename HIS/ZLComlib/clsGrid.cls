VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#功能：VSFlexGrid、MSHFlexGrid相关的处理?判断等公共函数
'命名规则：函数过程名称为通俗易懂的英文即可，VSFlexGrid方法不用加VSF关键字,其余类型表格控件需要加类型前缀
'上级类：clsControl
'##############################################################################
'表格空间滚动条的可见性，VSFHScrollVisible,VSFVScrollVisible
Private Const SB_HORZ = &H0
Private Const SB_VERT = &H1
Private Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long

Public Sub SetColVisible(ByVal frmMain As Object, ByVal strTittle As String, vsGrid As Object, _
                    Optional ByVal WinLeft As Double = 0, Optional ByVal WinTop As Double = 0, _
                    Optional ByVal lngTxtHeight As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:VsGrid的列设置(显示和隐藏)
    '入参:frmMain-父窗口
    '     strTittle-列标题设置名
    '     VsGrid的列设置(显示和隐藏):
    '           ColData存储的:列设置属性,格式为:列设置(1-固定,-1-不能选,0-可选)||光标控制(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
    '     WinLeft-WinTop:左边距;顶边距
    '     lngTxtHeight-文本框高度
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Call frmVsColSel.ShowColSet(frmMain, strTittle, vsGrid, WinLeft, WinTop, lngTxtHeight)
    Exit Sub
Errhand:
     If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Public Sub SetFontSize(ByRef vsInput As Object, ByVal lngFontSize As Long, Optional ByVal lngCol As Long)
'功能：设置vsfflexgrid字体的大小，并自动调整列宽和行高
'参数：lngFontSize常用设置：9号即小五，12号即小四
'      lngCol,如果要设置自动调整行高，则传入行高依据的列，要求AutoSizeMode = flexAutoSizeRowHeight,WordWrap =True
    Call gobjComLib.zlControl.VSFSetFontSize(vsInput, lngFontSize, lngCol)
End Sub

Public Function IsLastVisible(ByRef vsInput As Object, ByVal lngRow As Long) As Boolean
'功能：判断指定行是否最后一可见行
    Dim i As Long
    
    With vsInput
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then Exit For
        Next
        If i >= .FixedRows Then
            IsLastVisible = lngRow = i
        End If
    End With
End Function

Public Sub Init(ByRef vsInput As Object, ByVal strHead As String, Optional ByVal strRowContent As String, Optional ByVal intFixedCols As Integer, Optional ByVal intFixedRows As Integer = 1)
'功能：初始化表格内容，用在窗体个性化设置恢复之前
'参数：vsInput=要设置格式的表格
'          strHead=表格的列格式，格式为：列标题1,列宽1,对齐1,数据类型1,格式串1,隐藏1;列标题2,列宽2,对齐2,数据类型2,格式串2,隐藏2.....
'          strRowContent=表格的预定义行内容,格式为：列1,内容1,列2,内容2:行1;列1,内容1,列2,内容2:行2;
'说明：一般ColKey等于列名，如要单独设置请在列标题后加空格并加上单独的colKey,如格式串：费用 Fee,200,1表示将费用列的ColKey设置为Fee
'
    Dim i As Integer, lngRow As Long, j As Long
    Dim arrHead As Variant, arrCol As Variant, arrRow As Variant
    Dim arrTmp As Variant
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    '设置列
    With vsInput
        If strHead <> "" Then
            arrHead = Split(strHead, ";")
            .Clear: .Cols = 0: .Rows = 0
            .Rows = intFixedRows + 1: .Cols = UBound(arrHead) + 1
            .FixedRows = intFixedRows: .FixedCols = intFixedCols
            For i = LBound(arrHead) To UBound(arrHead)
                arrCol = Split(arrHead(i), ",")
                .FixedAlignment(i) = 4
                If intFixedRows <> 0 Then
                    If Not arrCol(0) Like "* ？*" Then
                        arrCol(0) = arrCol(0) & " " & arrCol(0)
                    End If
                    arrTmp = Split(arrCol(0), " ")
                    .TextMatrix(0, i) = arrTmp(0)
                    .ColKey(i) = arrTmp(1)
                End If
                
                If UBound(arrCol) > 0 Then
                    .ColWidth(i) = Val(arrCol(1))
                Else
                    .ColHidden(i) = True
                End If
                If UBound(arrCol) > 1 Then .ColAlignment(i) = Val(arrCol(2))
                If UBound(arrCol) > 2 Then .ColDataType(i) = Val(arrCol(3))
                If UBound(arrCol) > 3 Then .ColFormat(i) = arrCol(4)
                If UBound(arrCol) > 4 Then .ColHidden(i) = Val(arrCol(5))
            Next
        End If
        '设置解析行
        If strRowContent <> "" Then
            .Rows = .FixedRows
            lngRow = .FixedRows - 1: arrRow = Split(strRowContent, ";")
            For i = LBound(arrRow) To UBound(arrRow)
                arrTmp = Split(arrRow(i), ":")
                '确定行号
                lngRow = lngRow + 1 '没有行号的从上一行的行号上递增
                If UBound(arrTmp) > 0 Then lngRow = Val(arrTmp(1))
                If .Rows - 1 < lngRow Then .Rows = lngRow + 1 '填充的行号大于当前存在的行，则自动新增行
                '设置内容
                arrCol = Split(arrTmp(0), ",")
                For j = LBound(arrCol) To UBound(arrCol) Step 2
                    .TextMatrix(lngRow, Val(arrCol(j))) = arrCol(j + 1)
                Next
            Next
        End If
    End With
    Exit Sub
errH:
    Debug.Print Err.Source & "-Init:" & Err.Description
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Function HScrollVisible(vsInput As Object) As Boolean
'判断水平滚动条的可见性
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    HScrollVisible = False
    i = GetScrollRange(vsInput.hWnd, SB_HORZ, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos Then HScrollVisible = True
End Function

Public Function VScrollVisible(vsInput As Object) As Boolean
'判断垂直滚动条的可见性
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    VScrollVisible = False
    i = GetScrollRange(vsInput.hWnd, SB_VERT, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos Then VScrollVisible = True
End Function

Public Function CopyTo(ByRef vsfSource As Object, Optional ByRef vsfInput As Object) As Object
'功能: 将vsfSource的可见行列的数据复制到vsfInput或另一个公共表格中 , 便于Excel输出
'参数:
'     vsfInput-复制后的对象
'     vsfSource-复制到的对象
'     strMsg -提示信息
'返回：一个复制对象，该函数可以将数据复制到制定表格，也可以自动返回一个表格
    Dim i As Long
    Dim j As Long
    Dim lngCol As Long
    Dim lngRow As Long
    Dim lngTmp As Long
    Dim vsfReturn As VSFlexGrid
    Dim objTmp As frmPubTmp
    
    On Error GoTo errH:
    If vsfInput Is Nothing Then
        Set objTmp = New frmPubTmp
        Set vsfReturn = objTmp.vsfTmp
    End If
    With IIf(vsfInput Is Nothing, vsfReturn, vsfInput)
        .Rows = 0: .Cols = 0
        .Rows = vsfSource.Rows
        .FixedRows = vsfSource.FixedRows
        
        '复制
        lngCol = 0
        For i = 0 To vsfSource.Cols - 1 '列
            If Not vsfSource.ColHidden(i) Then
                .Cols = .Cols + 1
                .ColWidth(lngCol) = vsfSource.ColWidth(i)
                lngRow = 0: lngTmp = 0
                
                For j = 0 To vsfSource.Rows - 1 '行
                    If Not vsfSource.RowHidden(j) Then
                        .ColAlignment(i) = vsfSource.ColAlignment(i)
                        .Cell(flexcpAlignment, lngRow, lngCol) = vsfSource.Cell(flexcpAlignment, j, i)  '对齐方式
                        .TextMatrix(lngRow, lngCol) = vsfSource.TextMatrix(j, i)
                        lngRow = lngRow + 1
                    Else
                        lngTmp = lngTmp + 1  '记录隐藏行
                    End If
                Next
                lngCol = lngCol + 1
            End If
        Next
        '
        .Rows = .Rows - lngTmp '删除隐藏行
        .FixedCols = vsfSource.FixedCols
        .RowHeight(-1) = vsfSource.RowHeightMin
    End With
    Set CopyTo = IIf(vsfInput Is Nothing, vsfReturn, vsfInput)
    Exit Function
errH:
    Set CopyTo = Nothing
    Err.Clear
    If 1 = 2 Then
        Resume
    End If
End Function

Public Sub CheckKeyPress(ByVal vsInput As Object, ByRef lngRow As Long, ByRef lngCol As Long, ByRef intKeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '功能:只能输入数字和回车及退格
    '参数:
    '   objctl:Vsgrid8.0控件
    '   intKeyascii:
    '           Keyascii:8 (退格)
    '   Row-当前行
    '   Col-当前列
    '   TextType:(0-文本式;1-数字式;2-金额式)
    '返回:一个KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    
    If TextType = m文本式 Then
        If intKeyAscii = Asc("'") Then
            intKeyAscii = 0
        End If
        Exit Sub
    End If

    If intKeyAscii < Asc("0") Or intKeyAscii > Asc("9") Then
        Select Case intKeyAscii
        Case vbKeyReturn       '回车
        
        Case 8                 '退格
        
        Case Asc(".")
            If TextType = m金额式 Or TextType = m负金额式 Then
                If InStr(vsInput.EditText, ".") <> 0 Then     '只能存在一个小数点
                    intKeyAscii = 0
                End If
            Else
                intKeyAscii = 0
            End If
        Case Asc("-")          '负数
            Dim iRow As Long
            Dim icol As Long
            If Trim(vsInput.EditText) = "" Then Exit Sub
            If TextType <> m负金额式 Then intKeyAscii = 0: Exit Sub
            If vsInput.EditSelStart <> 0 Then intKeyAscii = 0: Exit Sub      '光标不存第一位,不能输入负数
            If InStr(1, vsInput.EditText, "-") <> 0 Then   '只能存在一个负数
                intKeyAscii = 0
            End If
        Case Else
            intKeyAscii = 0
        End Select
    End If
    Exit Sub
Errhand:
    Err.Clear
    intKeyAscii = 0
End Sub

Public Function BoolVal(ByVal vsGrid As Object, lngRow As Long, lngCol As Long) As Boolean
    '------------------------------------------------------------------------------
    '功能:获取bool列的值
    '返回:是该单元格为true,返回true,否则返回False
    '------------------------------------------------------------------------------
    Dim strTemp As String
    Err = 0: On Error GoTo errH:
    With vsGrid
        strTemp = .TextMatrix(lngRow, lngCol)
    End With
    If UCase(strTemp) = UCase("True") Then
        BoolVal = True: Exit Function
    End If
    BoolVal = Val(strTemp) <> 0
    Exit Function
errH:
    Err.Clear
End Function

Public Sub BandRec(objGrid As Object, rsData As ADODB.Recordset, Optional ByVal blnBandMark As Boolean)
'功能：将数据绑定到一个表格控件上
'参数：blnBandMark=是否绑定书签，书签绑定在RowData上
    Dim blnPre As Boolean, i As Long, j As Long
    
    objGrid.Clear: objGrid.Rows = 2: objGrid.Cols = 2
    objGrid.FixedRows = 1: objGrid.FixedCols = 0
    
    If rsData Is Nothing Then Exit Sub
    If rsData.State = adStateClosed Then Exit Sub
    
    blnPre = objGrid.Redraw
    objGrid.Redraw = False
    
    objGrid.Cols = rsData.Fields.count
    objGrid.Rows = IIf(rsData.RecordCount = 0, 2, rsData.RecordCount + 1)
    objGrid.FixedRows = 1
    
    For j = 0 To rsData.Fields.count - 1
        objGrid.TextMatrix(0, j) = rsData.Fields(j).Name
        If UCase(TypeName(objGrid)) <> UCase("MSHFlexGrid") Then objGrid.ColKey(j) = rsData.Fields(j).Name
    Next
    
    If rsData.RecordCount = 0 Then objGrid.Redraw = blnPre: Exit Sub
    
    rsData.MoveFirst
    For i = 1 To rsData.RecordCount
        For j = 0 To rsData.Fields.count - 1
            objGrid.TextMatrix(i, j) = "" & rsData.Fields(j).value
        Next
        If blnBandMark Then objGrid.RowData(i) = rsData.Bookmark '记录书签
        rsData.MoveNext
    Next
End Sub

Public Sub AdjustCols(ByRef vsTmp As Object, Optional ByVal strOtherCol As String)
'功能：修正表格列宽，改变表格字体后，需要表格列宽进行调整。该函数就是处理这种情况。
' 参数：vsTmp=表格对象
'          strOtherCol=不应对字体变化的列
'          lngBaseFontSize=原始字体大小
    Dim lngFixed As Long
    Dim dbAdd As Double, dbMinus As Double
    Dim arrAdd As Variant, arrMinus As Variant
    Dim lngCount As Long, lngTmp As Long
    Dim i As Long
    Dim lngLen As Long

    On Error GoTo errH
    lngFixed = gobjComLib.zlControl.OneCharWidth(vsTmp.Font) '获取一个字符宽
    arrMinus = Array(): arrAdd = Array()
    With vsTmp
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) And InStr(strOtherCol, "," & i & ",") = 0 Then
                lngLen = gobjComLib.zlStr.ActualLen(.TextMatrix(0, i))
                If lngLen <= 2 Then lngLen = 2
                lngLen = lngLen + 1 '增加边界宽度
                If .ColWidth(i) < lngLen * lngFixed Then  '小于字体增长后的最小宽度，则需要加宽
                    ReDim Preserve arrAdd(UBound(arrAdd) + 1)
                    arrAdd(UBound(arrAdd)) = i
                    dbAdd = dbAdd + lngLen * lngFixed - .ColWidth(i) '记录需要增宽的总量
                Else '大于字体增长后的最小宽度，则需要减少宽度
                    ReDim Preserve arrMinus(UBound(arrMinus) + 1)
                    arrMinus(UBound(arrMinus)) = i
                    dbMinus = dbMinus + .ColWidth(i) - lngLen * lngFixed '记录需要减少的总量
                End If
            End If
        Next
        If UBound(arrAdd) = -1 Then Exit Sub '没有可以加宽的列，则不用设置
        If UBound(arrMinus) = -1 Then Exit Sub '没有可以减少宽度的列，则不用设置
        For i = LBound(arrAdd) To UBound(arrAdd)
            lngLen = gobjComLib.zlStr.ActualLen(.TextMatrix(0, arrAdd(i)))
            If lngLen <= 2 Then lngLen = 2 '最小两个英文字符宽度
            lngLen = lngLen + 1 '增加边界宽度
            .ColWidth(arrAdd(i)) = lngLen * lngFixed '加宽的列设置为最小宽度
        Next
        dbAdd = IIf(dbAdd > dbMinus, dbMinus, dbAdd) '计算实际加宽的总量
        For i = LBound(arrMinus) To UBound(arrMinus)
            lngLen = gobjComLib.zlStr.ActualLen(.TextMatrix(0, arrMinus(i)))
            If lngLen <= 2 Then lngLen = 2 '最小两个英文字符宽度
            lngLen = lngLen + 1 '增加边界宽度
            '增量平摊到每一个可减少宽度的列，若某一列减少后会小于最小宽度，则设置为最小宽度
            If .ColWidth(arrMinus(i)) - dbAdd / (UBound(arrMinus) - i + 1) < lngLen * lngFixed Then
                dbAdd = dbAdd - (lngLen * lngFixed - .ColWidth(arrMinus(i))) '增量总量减去减少宽度
                .ColWidth(arrMinus(i)) = lngLen * lngFixed  '设置为最小宽度
            Else
                dbAdd = dbAdd - dbAdd / (UBound(arrMinus) - i + 1) '增量总量减去平摊量
                .ColWidth(arrMinus(i)) = .ColWidth(arrMinus(i)) - dbAdd / (UBound(arrMinus) - i + 1) '宽度为当前宽度减去平摊宽度
            End If
        Next
    End With
    Exit Sub
errH:
    Debug.Print Err.Source & "-AdjustCols:" & Err.Description
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Sub MshSetFormat(ByVal mshGrd As Object, ByVal strColumns As String, ByVal strCaption As String, Optional ByVal rsData As ADODB.Recordset, Optional ByVal blnData As Boolean = False, Optional ByVal blnRestore As Boolean = False)
    '功能：设置表格控件的列头，或绑定数据到表格控件
    '各参数含义：msfObj是表格对象；strColumns是列名格式串（列名|宽度|列对齐方式）；blnData说明是设置格式还是绑定数据；blnRestore说明是否恢复原始宽度
    '列名格式：NO|1000|4,日期|1600|4,检查单位|2000|4,检查人|800|4,简要记录|2000|4,结论|1000|4
    Call gobjComLib.zlControl.MshSetFormat(mshGrd, strColumns, strCaption, rsData, blnData, blnRestore)
End Sub

Public Sub MshSetColWidth(mshGrd As Object, frmParent As Object)
'功能：自动调整MSHFlexGrid表格的各列宽度,以最小匹配为准,并保持列头文字可以全部显示
'参数：mshGrd=要调整列宽的MSHFlexGrid控件
'      frmParent=父窗体,用于计算文字宽度
    Call gobjComLib.zlControl.MshSetColWidth(mshGrd, frmParent)
End Sub

Public Function MshGrdFindRow(ByVal mshGrd As Object, ByVal strTmp As Variant, ByVal intCol As Integer) As Integer
'功能:寻找与某一单元值相等的行
'参数:
'       mshGrd-网格控件
'       strTmp-匹配串
'       intCol-匹配列
'返回:查找到的行,如未找到,返位第一行
    MshGrdFindRow = gobjComLib.zlControl.MshGrdFindRow(mshGrd, strTmp, intCol)
End Function

Public Sub MshColumnSort(ByVal mshGrd As Object, ByVal IntSorCol As Integer, ByRef intPreCol As Integer, ByRef intPreSort As Integer, Optional blnNum As Boolean = False, Optional intMachingCol As Integer = 0)
'功能:对列进行排序
'参数:
'       mshGrd-MSHFlexGrid网络控件
'       intPreCol-前一次单据头的排序列
'       intPreSort-前一次单据头的排序
'       blnNum-是否为数量列
'       intMachingCol-匹配列,主要用于排序后重新定位到指定的列
    Call gobjComLib.zlControl.MshColumnSort(mshGrd, IntSorCol, intPreCol, intPreSort, blnNum, intMachingCol)
End Sub

Public Sub MshClearData(mshGrd As Object, ByVal intRow As Integer, Optional intCol As Variant)
    '功能: 清除指定网格的指定行指定列的数据
    '参数: mshGrd=要操作的网格控件
    '      intRow=要清除的行号
    '      intCol=要清除的列号列表如Array(1,2,3),若所有列则可以表示为Array()
    '返回：
    Call gobjComLib.zlControl.MshClearData(mshGrd, intRow, intCol)
End Sub

