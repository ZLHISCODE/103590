Attribute VB_Name = "mdlControlCommon"

'String
Public Function StringsDelItem(ByVal strItem As String, ByVal strList As String, Optional ByVal strSplit As String = ",") As String
'功能:移除字符串中的一项
'参数:blnSplit-分隔符
    Dim i As Long, arrTmp As Variant
    
    If strList = "" Then Exit Function
    
    arrTmp = Split(strList, strSplit)
    For i = 0 To UBound(arrTmp)
        If arrTmp(i) <> strItem Then
            StringsDelItem = StringsDelItem & "," & arrTmp(i)
        End If
    Next
    StringsDelItem = Mid(StringsDelItem, 2)
End Function

'ReportControl
Public Function GetRptColumn(ByRef rpt As ReportControl, ByVal strColumn As String) As Long
'功能：根据列名返回列序号，没有找到时返回-1
    Dim arrTmp As Variant, i As Long
    
    GetRptColumn = -1
    For i = 0 To rpt.Columns.Count - 1
        If rpt.Columns(i).Caption = strColumn Then GetRptColumn = i: Exit For
    Next
End Function

Public Sub rptRemoveGroupsItem(ByRef rpt As ReportControl, ByVal strColumn As String)
'功能:移除分组中的一个分组项,重新分组
'参数:rco-分组列集合
    Dim i As Long, arrTmp As Variant
    Dim strGroup As String
    
    With rpt
        .Columns(GetRptColumn(rpt, strColumn)).Visible = True
        For i = 0 To .GroupsOrder.Count - 1
            If .GroupsOrder.Column(i).Caption <> strColumn Then
                strGroup = strGroup & IIf(strGroup = "", "", ",") & .GroupsOrder.Column(i).Index
            End If
        Next
        
        .GroupsOrder.DeleteAll
        arrTmp = Split(strGroup, ",")
        For i = 0 To UBound(arrTmp)
            .GroupsOrder.Add .Columns(arrTmp(i))
        Next
        .Populate
    End With
End Sub

'ComboBox
Public Sub CboAddByStrings(cbo As ComboBox, strList As String, Optional blnSelectOne As Boolean = True)
'功能:根据字符串值初始化Combobox控件
'参数:strTmp-以逗号分隔的字符串,blnSelectOne-是否缺省选中第一项
    Dim i As Long, arrTmp As Variant
    
    If strList = "" Then Exit Sub
    
    cbo.Clear
    arrTmp = Split(strList, ",")
    For i = 0 To UBound(arrTmp)
        cbo.AddItem arrTmp(i)
    Next
    If blnSelectOne Then cbo.ListIndex = 0
End Sub

Public Sub CboRemoveItem(cbo As ComboBox, strListName As String, Optional blnSelectOne As Boolean = True)
'功能:根据字符串删除Combobox控件中的项
'参数:strTmp-要删除的项的文字,blnSelectOne-如果删除当前显示的项,则删除后Listindex=-1,是否缺省选中第一项
    Dim i As Long, lngRow As Long
    
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = strListName Then
            cbo.RemoveItem (i)
            Exit For
        End If
    Next
    If cbo.ListIndex = -1 And blnSelectOne Then cbo.ListIndex = 0
End Sub



'vsfFlexGrid
Public Function VsfGetColNum(vsf As VSFlexGrid, strColName As String) As Long
'功能:根据列名查找vsfFlexGrid控件中的列序号,没有找到时返回-1(使用vsfFee.ColIndex方法无效)
'参数:strColName-列名
    Dim i As Long
    
    For i = 0 To vsf.Cols - 1
        If vsf.TextMatrix(0, i) = strColName Then VsfGetColNum = i: Exit Function
    Next
    VsfGetColNum = -1
End Function

'MSHFlexGrid
Public Function MshGetColNum(msh As MSHFlexGrid, strColName As String) As Long
'功能:根据列名查找MSHFlexGrid控件中的列序号,没有找到时返回-1
'参数:strColName-列名
    Dim i As Long
    
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strColName Then MshGetColNum = i: Exit Function
    Next
    MshGetColNum = -1
End Function


Public Function MshLackComplement(msh As MSHFlexGrid, txtLack As TextBox, ByVal strDec As String, ByVal curMustSum As Currency, _
    ByVal lngCurrencyCol As Long, ByVal lngDefaultRow As Long, ByVal lngCurrentRow As Long, ByVal blnDefaultRowIsCash As Boolean, _
    Optional ByVal lngStartRow As Long = 1) As Currency
'功能:根据当前输入已改变的单元格的值,将"应付总金额"-"已付总金额"的差额补加到缺省行上
'     如果当前行是缺省行,则差额显示到差额文本框中
'依赖:CentMoney-分币处理函数,gBytMoney-分币处理规则
'参数:strDec-差额文本框显示的金额小数格式(如:0.00),curMustSum-应付总金额
'     lngCurrencyCol-金额列,lngDefaultRow-缺省行,lngCurrentRow-当前行,blnDefaultRowIsCash-缺省行是否进行分币处理,lngStartRow-起始数据行(除标题行外)
'返回:误差金额
    Dim i As Long, curSum As Currency, curLack As Currency, curTmp As Currency, curError As Currency
        
    For i = lngStartRow To msh.Rows - 1
        curSum = curSum + Val(msh.TextMatrix(i, lngCurrencyCol))
    Next
    curLack = curMustSum - curSum
    
    If lngCurrentRow <> lngDefaultRow Then
        curTmp = Val(msh.TextMatrix(lngDefaultRow, lngCurrencyCol)) + curLack
        If curTmp = 0 Then
            msh.TextMatrix(lngDefaultRow, lngCurrencyCol) = ""
        Else
            If blnDefaultRowIsCash Then
                msh.TextMatrix(lngDefaultRow, lngCurrencyCol) = Format(CentMoney(curTmp), "0.00")
            Else
                msh.TextMatrix(lngDefaultRow, lngCurrencyCol) = Format(curTmp, "0.00")
            End If
        End If
        curLack = curTmp - Val(msh.TextMatrix(lngDefaultRow, lngCurrencyCol))
        curError = -1 * curLack
        txtLack.Text = Format(0, strDec)
    Else
        curError = 0
        txtLack.Text = Format(curLack, strDec)
        If curLack <> 0 And (Abs(curLack) < 0.1 Or gBytMoney = 5 And Abs(curLack) < 0.3) Then
            '1.可能是现金分币处理产生的误差,三七作五二舍八入时最大可能有0.29的误差,0.79作0.5,0.29作0
            If blnDefaultRowIsCash Then
                curTmp = Val(msh.TextMatrix(lngDefaultRow, lngCurrencyCol))
                If CentMoney(curTmp + curLack) = curTmp Then txtLack.Text = Format(0, strDec): curError = -curLack
            Else
            '2.可能金额小数位数取舍产生的误差
                If Abs(curLack) < 0.005 Then txtLack.Text = Format(0, strDec): curError = -curLack
            End If
        End If
    End If
    
    MshComplement = curError
End Function
