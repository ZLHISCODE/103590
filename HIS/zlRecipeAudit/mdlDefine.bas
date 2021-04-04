Attribute VB_Name = "mdlDefine"
Option Explicit

'模块号
Public Enum enuModule
    门诊处方审查_1351 = 1351
    住院药嘱审查_1352 = 1352
    处方审查项目_1353 = 1353
    处方审查条件_1354 = 1354
    处方审查统计_1355 = 1355
End Enum

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Enum enuMenus
    文件 = 1
        打印设置 = 101
        打印预览 = 102
        打印 = 103
        输出Excel = 104
        参数设置 = 181
        退出 = 191
    编辑 = 2
        开启审查 = 8401
        停止审查 = 8402
        合格 = 3950
        不合格 = 3951
    报表 = 4
    查看 = 5
        工具栏 = 701
            标准按钮 = 7011
            文本标签 = 7012
            大图标 = 7013
        状态栏 = 702
        字体大小 = 509
            小字体 = 4041
            大字体 = 4042
        刷新 = 791
        查看PASS结果 = 3944
    帮助 = 6
        帮助主题 = 901
        WEB上的中联 = 902
            中联主页 = 9021
            中联论坛 = 9023
            发送反馈 = 9022
        关于 = 991
End Enum

Public gobjPubAdvice As zlPublicAdvice.clsPublicAdvice     '临床公共方法
Public gcnOracle As ADODB.Connection
Public gcnBusiness As ADODB.Connection
Public gstrSQL As String
Public glngSys As Long
Public glngModule As Long
Public gstrUnitName As String
Public gstrSysName As String                '系统名称
Public gstrProductName As String            'OEM产品名称
Public gintHoursRecipe As Integer           '待审查，参考多少小时内的处方药品
Public gstrErrInfo As String

Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
'功能：将SQL写入集合
'参数：
'  cllData：集合对象
'  strSQL：SQL字符串

    Dim l As Long
    
    l = cllData.Count + l
    cllData.Add strSQL, "K" & l
End Sub

Public Sub ExecuteProcedureArray(ByVal varArr As Variant, ByVal strCaption As String, Optional blnNoTrans As Boolean = False)
'功能:执行多条存储过程
'参数:
'  varArr：SQL集合对象
'  strCaption：窗体标题
'  blnNoTrans：是否不存在事务

    Dim i As Long, strSQL As String
    
    If blnNoTrans = False Then gcnOracle.BeginTrans
    For i = 1 To varArr.Count
        strSQL = varArr(i)
        zlDatabase.ExecuteProcedure strSQL, strCaption
    Next
    
    If blnNoTrans = False Then gcnOracle.CommitTrans
End Sub

'Public Function FormatString(ByVal strFormat As String, ParamArray arrParams() As Variant) As String
''功能：格式化字符串
''参数：
''  strFormat：表达式；[1-x]为参数号关键字；例子："测试值为：[1]"
''  arrParams：表达式的参数，对应strFormat中的参数号关键字
''返回：格式化后的字符串
'
'    Dim i As Integer, intSN As Integer
'    Dim strKey As String, strTmp As String
'    Dim blnStart As Boolean
'
'    FormatString = strFormat
'
'    If Len(strFormat) > 60000 Then Exit Function
'    If Not strFormat Like "*[[]*[]]*" Then Exit Function
'    If UBound(arrParams) < 0 Then Exit Function
'
'    On Error GoTo errHandle
'
'    For i = 1 To Len(strFormat)
'        If Mid(strFormat, i, 1) = "[" Then
'            blnStart = True
'        End If
'        If blnStart Then
'            If Mid(strFormat, i, 1) = "]" Then
'                intSN = Val(Mid(strKey, 2))
'                If intSN > 0 Then
'                    If UBound(arrParams) >= intSN - 1 Then
'                        strTmp = strTmp & arrParams(intSN - 1)
'                    End If
'                Else
'                    strTmp = strTmp & Mid(strKey, 2)
'                End If
'                blnStart = False
'                strKey = ""
'            Else
'                strKey = strKey & Mid(strFormat, i, 1)
'            End If
'        Else
'            strTmp = strTmp & Mid(strFormat, i, 1)
'        End If
'    Next
'
'    FormatString = strTmp
'    Exit Function
'
'errHandle:
'End Function

Public Sub SetLVColumnHeaders(ByRef lvwVar As ListView, ByVal strHeader As String)
'功能：统一设置ListView的列头
'参数：
'  lvwVar：要设置的ListView控件
'  strHeader：列头标准字串
'    格式：列名,Key值,宽度,对齐方式,图标号[|列名1,...]
'    说明：Key值不填，表示用列名代；宽度不填，表示隐藏；对齐不填，默认左齐；

    Dim i As Integer, j As Integer
    Dim arrCols As Variant, arrElements As Variant
    Dim strText As String, strKey As String
    Dim intWidth As Integer, intAlignment As Integer, intIcon As Integer

    If Trim(strHeader) = "" Then Exit Sub
    If lvwVar Is Nothing Then Exit Sub
    
    arrCols = Split(strHeader, "|")
    With lvwVar
        .ColumnHeaders.Clear
        
        For i = LBound(arrCols) To UBound(arrCols)
            arrElements = Split(arrCols(i), ",")
            If UBound(arrElements) < 2 Then
                MsgBox zlStr.FormatString("设置“[1]”控件列头的参数不正确！", lvwVar.Name), vbInformation, gstrSysName
                Exit Sub
            End If
            '列名
            strText = Trim(arrElements(0))
            If strText = "" Then
                MsgBox zlStr.FormatString("设置“[1]”控件列头名称的参数不正确！", lvwVar.Name), vbInformation, gstrSysName
                Exit Sub
            End If
            'Key
            If UBound(arrElements) > 0 Then
                strKey = arrElements(1)
            End If
            If Trim(strKey) = "" Then
                strKey = strText
            End If
            '宽度
            If UBound(arrElements) > 1 Then
                intWidth = Val(arrElements(2))
            Else
                intWidth = 0
            End If
            '对齐
            If UBound(arrElements) > 2 Then
                intAlignment = Val(arrElements(3))
                If intAlignment > 2 Then intAlignment = 0
            Else
                intAlignment = 0
            End If
            '图标号
            If UBound(arrElements) > 3 Then
                intIcon = Val(arrElements(4))
            Else
                intIcon = 0
            End If
            
            .ColumnHeaders.Add i + 1, strText, strKey, intWidth, intAlignment, intIcon
        Next
    End With
    
End Sub

Public Function GetLVColumnIndex(ByVal lvwVar As ListView, ByVal strKey As String) As Integer
'功能：获取指定ListView控件列的Index
'参数：
'  lvwVar：指定ListView控件
'  strKey：要获取列的Key值
'返回：列的Index

    Dim i As Integer

    With lvwVar
        For i = 0 To .ColumnHeaders.Count - 1
            If UCase(strKey) = UCase(.ColumnHeaders.Item(i).Key) Then
                GetLVColumnIndex = i
                Exit Function
            End If
        Next
    End With

    GetLVColumnIndex = -1
End Function

Public Sub FillLVData(ByRef rsVar As ADODB.Recordset, ByRef lvwVar As ListView, _
    Optional ByVal strCheckCol As String, _
    Optional ByVal strKey As String)
'功能：给ListView控件填充数据
'参数：
'  rsVar：记录集对象
'  lvwVar：指定要填充的ListView的控件
'  strCheckCol：Checkbox的列
'  strKey：指定记录作为Key的字段

    If rsVar Is Nothing Then Exit Sub
    If rsVar.State <> adStateOpen Then Exit Sub
    If lvwVar Is Nothing Then Exit Sub
    
    Dim limTmp As ListItem
    Dim i As Integer, j As Integer
    Dim arrFields As Variant
    Dim strMasterCol As String, strTmp As String
    
    strCheckCol = "," & strCheckCol & ","
    
    '主列
    strMasterCol = lvwVar.ColumnHeaders(1).Key
    
    '填数据
    i = 1
    If rsVar.RecordCount > 0 Then rsVar.MoveFirst
    Do While rsVar.EOF = False
        '主列
        strTmp = rsVar.Fields(strMasterCol).Value
        If InStr(strCheckCol, "," & strMasterCol & ",") > 0 Then
            '不显示文本
            strTmp = ""
        End If
        If strKey = "" Then
            Set limTmp = lvwVar.ListItems.Add(i, , strTmp)
        Else
            Set limTmp = lvwVar.ListItems.Add(i, "_" & rsVar.Fields(strKey).Value, strTmp)
        End If
        If lvwVar.Checkboxes Then
            limTmp.Checked = zlCommFun.NVL(rsVar.Fields(strMasterCol).Value, 0) = 1
        End If
        '子列
        For j = 1 To lvwVar.ColumnHeaders.Count
            If j > 1 Then
                strTmp = lvwVar.ColumnHeaders(j).Key
                On Error Resume Next
                limTmp.ListSubItems.Add , , rsVar.Fields(strTmp).Value
                Err.Clear
                On Error GoTo 0
            End If
        Next
    
        rsVar.MoveNext: i = i + 1
    Loop

End Sub

Public Sub MergeVSFHead(ByRef strNew As String, ByVal strConst As String, ByVal strRegsiter As String)
'功能：合并VSF列头字串，strConst有，就为strRegsiter添加；strConst没有，strRegsiter就删除
'参数：
'  strNew：合并后的字串
'  strConst：变量字串
'  strRegister：注册表的字串

    Dim i As Integer, j As Integer
    Dim arrConst As Variant, arrReg As Variant
    Dim strI As String, strJ As String
    Dim blnFind As Boolean
    
    strNew = strRegsiter
    
    arrConst = Split(strConst, "|")
    arrReg = Split(strRegsiter, "|")
    For i = LBound(arrConst) To UBound(arrConst)
        If Split(arrConst(i), ",")(0) <> "" Then
            strI = Split(arrConst(i), ",")(0)
        Else
            strI = Split(arrConst(i), ",")(1)
        End If
        
        blnFind = False
        For j = LBound(arrReg) To UBound(arrReg)
            If Split(arrReg(j), ",")(0) <> "" Then
                strJ = Split(arrReg(j), ",")(0)
            Else
                strJ = Split(arrReg(j), ",")(1)
            End If
            If strI = strJ Then
                blnFind = True
                Exit For
            End If
        Next
        
        If blnFind = False Then
            strNew = strNew & "|" & arrConst(i)
        End If
    Next
    
    strRegsiter = strNew
    strNew = ""
    
    arrConst = Split(strConst, "|")
    arrReg = Split(strRegsiter, "|")
    For i = LBound(arrReg) To UBound(arrReg)
        If Split(arrReg(i), ",")(0) <> "" Then
            strI = Split(arrReg(i), ",")(0)
        Else
            strI = Split(arrReg(i), ",")(1)
        End If
        
        blnFind = False
        For j = LBound(arrConst) To UBound(arrConst)
            If Split(arrConst(j), ",")(0) <> "" Then
                strJ = Split(arrConst(j), ",")(0)
            Else
                strJ = Split(arrConst(j), ",")(1)
            End If
            If strI = strJ Then
                blnFind = True
                Exit For
            End If
        Next
        
        If blnFind Then
            strNew = strNew & arrReg(i) & "|"
        End If
    Next
    
    strNew = Left(strNew, Len(strNew) - 1)
End Sub

Public Sub SetVSFHead(ByVal vsfObject As VSFlexGrid, ByVal strHead As String)
'--------------------------------
'功能：初始化VSFlexGrid控件表格头
'参数：
'  vsfObject：目标控件；
'  strHead：表格头的初始化字串
'
'格式： "剂型,,3,1000,s|..."
'   元素1：Key值；
'   元素2：Caption值（默认为Key值）；
'   元素3：列属性（0：内部显示，可移动；1：内部隐藏，不可移动，不可显示；2：用户隐藏；3：用户显示(默认值)）
'   元素4：列宽度（默认0）；
'   元素5：显示格式；s(默认)：字符串； n：数字； d：日期； t：时间； dt：日期时间
'--------------------------------
    Dim arrCols As Variant, arrRows As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    arrRows = Split(strHead, "|")
    With vsfObject
        If .Rows = 0 Then .Rows = 1
        .Cols = UBound(arrRows) + 1
        For i = LBound(arrRows) To UBound(arrRows)
            If arrRows(i) <> "" Then
                arrCols = Split(arrRows(i), ",")
                '第1元素：Key值
                .ColKey(i) = arrCols(0)
                
                '第2元素：Caption值
                If arrCols(1) = "" Then
                    .TextMatrix(0, i) = arrCols(0)
                Else
                    .TextMatrix(0, i) = arrCols(1)
                End If
                
                '第3元素：列属性
                If arrCols(2) = "" Then
                    .ColData(i) = 3
                Else
                    .ColData(i) = Val(arrCols(2))
                End If
                
                '第4元素：宽度
                .ColWidth(i) = Val(arrCols(3))
                
                '第5元素：显示格式
                If UBound(arrCols) > 3 Then
                    If UCase(arrCols(4)) = "D" Then
                        .ColFormat(i) = "yyyy-mm-dd"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "T" Then
                        .ColFormat(i) = "hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "DT" Then
                        .ColFormat(i) = "yyyy-mm-dd hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "N" Then
                        .ColAlignment(i) = flexAlignRightCenter
                    Else
                        .ColAlignment(i) = flexAlignLeftCenter
                    End If
                Else
                    .ColAlignment(i) = flexAlignLeftCenter
                End If
                
                '隐藏列
                If Val(arrCols(2)) = 1 Or Val(arrCols(2)) = 2 Or Val(arrCols(2)) = 0 Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                End If
                
            End If
        Next
        
        If .Cols > 0 Then .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    Exit Sub
    
errHandle:
    MsgBox Err.Description, vbInformation, gstrSysName
End Sub

Public Function GetCurrentVSFHead(ByVal vsfObject As VSFlexGrid) As String
'-------------------------------------
'功能：获取VSF目标控件当前的表格头字串
'参数：vsfObject：目标控件
'返回：表格头字串
'-------------------------------------
    Dim i As Integer
    Dim strHead As String, strCol As String
    
    With vsfObject
        strHead = ""
        For i = 0 To .Cols - 1
            '第1元素：Key
            strCol = .ColKey(i) & ","
            '第2元素：Caption
            If strCol = .TextMatrix(0, i) & "," Then
                strCol = strCol & ","
            Else
                strCol = strCol & .TextMatrix(0, i) & ","
            End If
            '第3元素：列属性
            If Val(.ColData(i)) = 3 Then
                If .ColHidden(i) Then
                    strCol = strCol & "2,"
                Else
                    strCol = strCol & ","
                End If
            Else
                If .ColHidden(i) = False And Val(.ColData(i)) = 2 Then
                    strCol = strCol & "3,"
                Else
                    strCol = strCol & .ColData(i) & ","
                End If
            End If
            '第4元素：列宽
            If Val(.ColWidth(i)) = 0 Then
                strCol = strCol & ","
            Else
                strCol = strCol & .ColWidth(i) & ","
            End If
            '第5元素：显示格式
            If Trim(.ColFormat(i)) = "" Then
                If .ColAlignment(i) = flexAlignRightCenter Then
                    strCol = strCol & "n"
                Else
                    strCol = Left(strCol, Len(strCol) - 1)
                End If
            Else
                If .ColFormat(i) = "yyyy-mm-dd" Then
                    strCol = strCol & "d"
                ElseIf .ColFormat(i) = "hh:mm:ss" Then
                    strCol = strCol & "t"
                ElseIf .ColFormat(i) = "yyyy-mm-dd hh:mm:ss" Then
                    strCol = strCol & "dt"
                End If
            End If
            '各列组合
            strHead = strHead & strCol & IIf(i = .Cols - 1, "", "|")
        Next
    End With
    GetCurrentVSFHead = strHead
End Function

Public Sub FillVSFData(ByRef vsfVar As VSFlexGrid, ByRef rsVar As ADODB.Recordset)
'功能：将记录集对象的数据填充至vsf控件中
'参数：
'  vsfVar：要填充数据的Vsf控件
'  rsVar：记录集对象

    If rsVar Is Nothing Then Exit Sub
    If rsVar.State <> adStateOpen Then Exit Sub
    If vsfVar Is Nothing Then Exit Sub
    
    Dim i As Integer, intCol As Integer
    Dim lngRow As Long
    
    With rsVar
        vsfVar.Redraw = flexRDNone
        vsfVar.Rows = .RecordCount + 1
        vsfVar.Clear 1
        
        lngRow = 1
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            For i = 0 To .Fields.Count - 1
                intCol = vsfVar.ColIndex(.Fields(i).Name)
                If intCol >= 0 Then
                    'vsf列存在该字段
                    vsfVar.TextMatrix(lngRow, intCol) = zlCommFun.NVL(.Fields(i).Value)
                End If
            Next
            
            lngRow = lngRow + 1
            .MoveNext
        Loop
        vsfVar.Redraw = flexRDDirect
    End With

End Sub

Public Sub SetTextMaxLen(ByRef txtVal As TextBox, ByVal strTableField As String)
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = zlStr.FormatString("Select [2] as 字段 From [1] Where Rownum < 1 ", _
                        CStr(Split(strTableField, ".")(0)), _
                        CStr(Split(strTableField, ".")(1)))
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取字段信息")
    txtVal.MaxLength = rsTmp.Fields(0).DefinedSize
    rsTmp.Close

    Exit Sub
    
errHandle:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Sub

Public Sub SetRecordsetStructure(ByVal bytClass As Byte, ByRef rsVar As ADODB.Recordset)
'功能：设置不合格记录集对象的字段
'参数：
'  bytClass：1-待提交
'  rsVar：不合格记录集对象
    
    '格式：字段名,字段类型,长度
    '  3-adInteger；20-adBigInt；200-adVarchar；201-adLongVarchar
    Const STR_NG_PROP     As String = "药名ID;20|审查项目;200;100|药品名称;200;100|商品名;200;100|规格;200;100|单位;200;100"
    Const STR_SUBMIT_PROP As String = "发药药房ID;20|审查项目ID;20|编码;200;100|简称;200;100|医嘱ID;20|审查结果;3"
    
    Dim strFieldsProp As String
    Dim arrFields As Variant, arrProp As Variant
    Dim i As Integer

    If Not rsVar Is Nothing Then
        If rsVar.State <> adStateClosed Then rsVar.Close
        Set rsVar = Nothing
    End If
    
    Set rsVar = New ADODB.Recordset
    
    With rsVar
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockPessimistic
        
        If bytClass = 1 Then
            strFieldsProp = STR_SUBMIT_PROP
        End If
        
        '新建
        arrFields = Split(strFieldsProp, "|")
        For i = LBound(arrFields) To UBound(arrFields)
            arrProp = Split(arrFields(i), ";")
            Select Case Val(arrProp(1))
                Case DataTypeEnum.adVarChar
                    If UBound(arrProp) >= 2 Then
                        .Fields.Append arrProp(0), adVarChar, Val(arrProp(2))
                    Else
                        .Fields.Append arrProp(0), adVarChar, 100           '默认长度
                    End If
                Case DataTypeEnum.adLongVarChar
                    .Fields.Append arrProp(0), adLongVarChar                'LongVarchar动态长度
                Case DataTypeEnum.adBigInt
                    .Fields.Append arrProp(0), adBigInt
                Case DataTypeEnum.adInteger
                    .Fields.Append arrProp(0), adInteger
            End Select
        Next
        .Open
    End With
End Sub

Public Function SetPublicFontSize(ByRef frmMe As Object, ByVal bytSize As Byte, Optional ByVal strOther As String)
'功能：设置窗体及所有控件的字体大小
'参数：frmMe=需要设置字体的窗体对象
'      bytSize:设置为9号字体,0:设置为9号字体,1,设置为12号字体
'      strOther:不进行字体设置的控件父容器的集合,格式为：容器名字1,容器名字2,容器名字3,....
'说明：1.如果涉及到VsFlexGrid等表格控件，需要根据所在的环境重新调整列宽和行高
'      2.如果存在未列出的其他控件或自定义控件,需要用特定方法指定字体大小及相关处理的，需另外单独设置

    Dim objCtrol As Control
    Dim CtlFont As StdFont
    Dim i As Long, lngOldSize As Long
    Dim lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    Dim strContainer As String
    
    lngFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    frmMe.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In frmMe.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "SpeedButton", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox", "IDKindNew", _
                "VSFlexGrid", "StatusBar"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '对于CommandBars用户自定义控件读取objCtrol.Container会出错
            strContainer = ""
            On Error Resume Next
            strContainer = objCtrol.Container.Name
            Err.Clear: On Error GoTo 0
            If InStr(1, strOther, "," & strContainer & ",") > 0 Then
                 blnDo = False
            End If
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
                Case "TabStrip"
                        objCtrol.Font.Size = lngFontSize
                Case "Label"
                        If Not LCase(objCtrol.Name) Like "*_fixed" Then
                            lngOldSize = objCtrol.Font.Size
                            dblRate = lngFontSize / lngOldSize
                            
                            objCtrol.Font.Size = lngFontSize
                            objCtrol.Height = frmMe.TextHeight("字") + 20
                            'Label宽度需要自行调整
                        End If
               Case "ComboBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "ListView"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        For i = 1 To objCtrol.ColumnHeaders.Count
                            objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                        Next
                Case "OptionButton"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("字体" & objCtrol.Caption)
                        objCtrol.Height = objCtrol.Height * dblRate
                Case "CheckBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "DTPicker"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("2012-01-01    ")
                        objCtrol.Height = frmMe.TextHeight("字") + IIf(bytSize = 0, 100, 120)
                Case "TextBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("字")
                Case "MaskEdBox"
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = frmMe.TextWidth(objCtrol.Mask)
                        objCtrol.Height = frmMe.TextHeight("字")
'                Case "ReportControl"
'                        lngOldSize = objCtrol.PaintManager.TextFont.Size
'                        dblRate = lngFontSize / lngOldSize
'
'                        Set CtlFont = objCtrol.PaintManager.CaptionFont
'                        CtlFont.Size = lngFontSize
'                        Set objCtrol.PaintManager.CaptionFont = CtlFont
'                        Set CtlFont = objCtrol.PaintManager.TextFont
'                        CtlFont.Size = lngFontSize
'                        Set objCtrol.PaintManager.TextFont = CtlFont
'                        For Each objrptCol In objCtrol.Columns
'                            objrptCol.Width = objrptCol.Width * dblRate
'                        Next
'                        objCtrol.Redraw
                Case "SpeedButton"
                        Dim objFont As New StdFont
                        
                        Set objFont = frmMe.Font
                        If bytSize = 0 Then
                            objFont.Size = 12
                            dblRate = 0.8
                        Else
                            objFont.Size = 15.75
                            dblRate = 1 / 0.8
                        End If
                        Set objCtrol.Font = objFont
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "VSFlexGrid"
                        Set objCtrol.Font = frmMe.Font
                        objCtrol.Font.Size = IIf(bytSize = 0, 9, 12)
                Case "DockingPane"
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        
                        Set CtlFont = objCtrol.TabPaintManager.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.TabPaintManager.Font = CtlFont
        
                        Set CtlFont = objCtrol.PanelPaintManager.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PanelPaintManager.Font = CtlFont
                Case "CommandBars"
                        Set CtlFont = objCtrol.Options.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.Options.Font = CtlFont
                Case "TabControl"
                        Set CtlFont = objCtrol.PaintManager.Font
                        If CtlFont Is Nothing Then  '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.Font = CtlFont
                        objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
                Case "CommandButton"
                        If Not LCase(objCtrol.Name) Like "*_fixed" Then
                            lngOldSize = objCtrol.FontSize
                            dblRate = lngFontSize / lngOldSize
    
                            objCtrol.FontSize = lngFontSize
                            objCtrol.Width = dblRate * objCtrol.Width
                            objCtrol.Height = dblRate * objCtrol.Height
                        End If
                Case "Frame"
                        objCtrol.FontSize = lngFontSize
                Case "IDKindNew"
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "StatusBar"
                        objCtrol.Font.Size = lngFontSize
            End Select
        End If
    Next
End Function

