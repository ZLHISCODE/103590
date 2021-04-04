Attribute VB_Name = "mdlEinvoice"
Option Explicit
Public gstrProductName As String, gstrSysName As String
Public gcnOracle As ADODB.Connection
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    部门名称 As String
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public gobjEInvProviders As clsEInvProviders   '当前提供者集
Public gobjEinvProvider As clsEInvProvider  '当前使用提供者
Public glngInstanceCount As Long '实例数
Public Const GRD_GOTFOCUS_COLORSEL = &H8000000D '16772055 '    '进入控件时,选择显示颜色
Public Const GRD_LOSTFOCUS_COLORSEL = &HE0E0E0  '&H80000010  '离开焦点时,选择的显示颜色

Public Sub InitEInvProviders()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化接口提供商数据
    '编制:刘兴洪
    '日期:2020-03-04 15:25:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim objEinvProvider As clsEInvProvider
    If Not gobjEInvProviders Is Nothing Then Exit Sub
    
    Set gobjEInvProviders = New clsEInvProviders
    strSQL = "Select 编号,名称,简码,是否启用,部件,包名称 From 电子票据类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取电子票据类别")
    With rsTemp
        Do While Not .EOF
            Set objEinvProvider = gobjEInvProviders.Add(Val(Nvl(!编号)), Nvl(!名称), Val(Nvl(!是否启用)) = 1, Nvl(!部件), True, "K" & zlStr.LPAD(Val(Nvl(!编号)), 3, "0"))
            If objEinvProvider.是否启用 Then Set gobjEinvProvider = objEinvProvider
            .MoveNext
        Loop
    End With
End Sub

Public Function GetPubEInvoiceObject(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
    objPubEInvoice As Object, Optional ByVal byt场合 As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取电子票据公共接口部件
    '入参:
    '   frmMain：调用的主窗体
    '   lngModule：当前调用模块号
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '   blnDeviceSet：设备设置调用的初始化
    '出参:
    '返回:初始化成功返回true,否则返回False
    '说明:
    '   1.使用本部件前,必须先调用本接口进行初始化
    '   2.初始化接口,在HIS进入模块时调用(例如：进入收费管理界面)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExtend As String
    
    If objPubEInvoice Is Nothing Then
        On Error Resume Next
        Set objPubEInvoice = CreateObject("zlPublicExpense.clsPubEInvoice")
        If Err <> 0 Then
            MsgBox "不存在可用的电子票据接口部件(zlPublicExpense.clsPubEInvoice)，请与系统管理员联系。详细的错误信息为:" & vbCrLf & Err.Description, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If objPubEInvoice Is Nothing Then Exit Function
    
    GetPubEInvoiceObject = objPubEInvoice.zlInitialize(frmMain, byt场合, gcnOracle, lngSys, lngModule, False, strExtend)
End Function

Public Function load开票点(cboControl As Object, ByRef rs开票点 As ADODB.Recordset, ByRef rs收费员 As ADODB.Recordset)
    On Error GoTo ErrHandler
    cboControl.Clear
    If Get开票点(rs开票点) = False Then Exit Function
    If rs开票点.RecordCount > 0 Then
        Do While Not rs开票点.EOF
            cboControl.AddItem rs开票点!编码 & "-" & rs开票点!名称
            rs开票点.MoveNext
        Loop
        load开票点 = True: Exit Function
    End If
    
    If Load收费员(cboControl, rs收费员) = False Then Exit Function
    load开票点 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Load收费员(cboControl As Object, ByRef rs收费员 As ADODB.Recordset)

    On Error GoTo ErrHandler
    cboControl.Clear
    If Get收费员(rs收费员) = False Then Exit Function
    
    Do While Not rs收费员.EOF
        cboControl.AddItem rs收费员!编号 & "-" & rs收费员!姓名
        rs收费员.MoveNext
    Loop
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Select开票点(frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
    cboControl As Object, rs开票点 As ADODB.Recordset) As Boolean
    '模糊查找开票点
    Dim lngCount As Long
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset, strAdded As String
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    
    '先复制记录集
    On Error GoTo ErrHandler
    Set rsTemp = zlDatabase.zlCopyDataStructure(rs开票点)
    
    strText = cboControl.Text
    strCompents = strText & "*"
    If IsNumeric(strText) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strText) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    
    rs开票点.Filter = strFilter: lngCount = 0
    With rs开票点
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not rs开票点.EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编码输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编码01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编码完全相同,则直接就定位到该名称
                If Nvl(!编码) = strText Then strResult = Nvl(!名称): lngCount = 0: Exit Do
                
                '1.编码输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编码)) = Val(strText) Then
                    If lngCount = 0 Then strResult = Nvl(!名称)
                    lngCount = lngCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Val(rs开票点!编码) Like strText & "*" Then
                    If CheckComBoxExists(cboControl, Nvl(!名称)) And InStr(strAdded, "," & Nvl(!编码) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs开票点, rsTemp)
                        strAdded = strAdded & "," & Nvl(!编码) & ","
                    End If
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strText Then
                    If lngCount = 0 Then strResult = Nvl(!名称)   '可能存在多个相同简码
                    lngCount = lngCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If CheckComBoxExists(cboControl, Nvl(!名称)) And InStr(strAdded, "," & Nvl(!编码) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs开票点, rsTemp)
                        strAdded = strAdded & "," & Nvl(!编码) & ","
                    End If
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编码类似于N001简码可能有ZYK01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或名称 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strText Or Trim(!简码) = strText Or Trim(!名称) = strText Then
                    If lngCount = 0 Then strResult = Nvl(!名称)   '可能存在多个相同的多个
                    lngCount = lngCount + 1
                End If
                '2.简码或编码或名称 根据参数来匹配数(但编码只能左匹配)
                If Trim(!编码) Like strText & "*" Or Trim(Nvl(!简码)) Like strCompents Or Trim(Nvl(!名称)) Like strCompents Then
                    If CheckComBoxExists(cboControl, Nvl(!名称)) And InStr(strAdded, "," & Nvl(!编码) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs开票点, rsTemp)
                        strAdded = strAdded & "," & Nvl(!编码) & ","
                    End If
                End If
            End Select
            rs开票点.MoveNext
        Loop
    End With
    
    If lngCount > 1 Then strResult = ""
    If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!名称)
    '直接定位
    If strResult <> "" Then
        rsTemp.Close: Set rsTemp = Nothing
        If CheckComBoxExists(cboControl, strResult, True) Then zlCommFun.PressKey vbKeyTab
        Select开票点 = True: Exit Function
    End If
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then
        '未找到
        rsTemp.Close: Set rsTemp = Nothing
        Exit Function
    End If

    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = "简码"
    Case Else
        rsTemp.Sort = "名称"
    End Select
    
    '弹出选择器
    Dim rsReturn As ADODB.Recordset
    If zlDatabase.zlShowListSelect(frmMain, lngSys, lngModule, cboControl, rsTemp, True, "", "ID", rsReturn) Then
        Call zlControl.ControlSetFocus(cboControl)
        If Not rsReturn Is Nothing Then
            If rsReturn.RecordCount <> 0 Then
                '进行定位
                If CheckComBoxExists(cboControl, Nvl(rsReturn!名称), True) Then zlCommFun.PressKey vbKeyTab
                rsTemp.Close: Set rsTemp = Nothing
                Select开票点 = True: Exit Function
            End If
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Select收费员(frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
    cboControl As Object, rs收费员 As ADODB.Recordset) As Boolean
    '模糊查找收费员
    Dim lngCount As Long
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset, strAdded As String
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    
    '先复制记录集
    On Error GoTo ErrHandler
    Set rsTemp = zlDatabase.zlCopyDataStructure(rs收费员)
    
    strText = cboControl.Text
    strCompents = strText & "*"
    If IsNumeric(strText) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strText) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    
    rs收费员.Filter = strFilter: lngCount = 0
    With rs收费员
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not rs收费员.EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编号) = strText Then strResult = Nvl(!姓名): lngCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编号)) = Val(strText) Then
                    If lngCount = 0 Then strResult = Nvl(!姓名)
                    lngCount = lngCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Val(rs收费员!编号) Like strText & "*" Then
                    If CheckComBoxExists(cboControl, Nvl(!姓名)) And InStr(strAdded, "," & Nvl(!编号) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs收费员, rsTemp)
                        strAdded = strAdded & "," & Nvl(!编号) & ","
                    End If
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strText Then
                    If lngCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同简码
                    lngCount = lngCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If CheckComBoxExists(cboControl, Nvl(!姓名)) And InStr(strAdded, "," & Nvl(!编号) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs收费员, rsTemp)
                        strAdded = strAdded & "," & Nvl(!编号) & ","
                    End If
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编号) = strText Or Trim(!简码) = strText Or Trim(!姓名) = strText Then
                    If lngCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同的多个
                    lngCount = lngCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If Trim(!编号) Like strText & "*" Or Trim(Nvl(!简码)) Like strCompents Or Trim(Nvl(!姓名)) Like strCompents Then
                    If CheckComBoxExists(cboControl, Nvl(!姓名)) And InStr(strAdded, "," & Nvl(!编号) & ",") = 0 Then
                        Call zlDatabase.zlInsertCurrRowData(rs收费员, rsTemp)
                        strAdded = strAdded & "," & Nvl(!编号) & ","
                    End If
                End If
            End Select
            rs收费员.MoveNext
        Loop
    End With
    
    If lngCount > 1 Then strResult = ""
    If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!姓名)
    '直接定位
    If strResult <> "" Then
        rsTemp.Close: Set rsTemp = Nothing
        If CheckComBoxExists(cboControl, strResult, True) Then zlCommFun.PressKey vbKeyTab
        Select收费员 = True: Exit Function
    End If
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then
        '未找到
        rsTemp.Close: Set rsTemp = Nothing
        Exit Function
    End If

    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = "编号"
    Case 1 '输入全拼音
        rsTemp.Sort = "简码"
    Case Else
        rsTemp.Sort = "姓名"
    End Select
    
    '弹出选择器
    Dim rsReturn As ADODB.Recordset
    If zlDatabase.zlShowListSelect(frmMain, lngSys, lngModule, cboControl, rsTemp, True, "", "ID", rsReturn) Then
        Call zlControl.ControlSetFocus(cboControl)
        If Not rsReturn Is Nothing Then
            If rsReturn.RecordCount <> 0 Then
                '进行定位
                If CheckComBoxExists(cboControl, Nvl(rsReturn!姓名), True) Then zlCommFun.PressKey vbKeyTab
                rsTemp.Close: Set rsTemp = Nothing
                Select收费员 = True: Exit Function
            End If
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckComBoxExists(cboControl As Object, ByVal strText As String, _
    Optional ByVal blnLocateItem As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在开单人下拉列表中.
    '入参:
    '     blnLocateItem:是否直接定位
    '返回:存在返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    For i = 0 To cboControl.ListCount - 1
        If zlStr.NeedName(cboControl.List(i)) = strText Then
            If blnLocateItem Then cboControl.ListIndex = i
            CheckComBoxExists = True
            Exit Function
        End If
    Next
End Function

Private Function Get开票点(ByRef rs开票点 As ADODB.Recordset) As Boolean
    '加载开票点
    Dim strSQL As String

    On Error GoTo ErrHandler
    strSQL = _
        " Select a.ID, a.编码, a.简码, a.名称" & _
        " From 电子票据开票点 A" & _
        " Where Nvl(A.撤档时间, Sysdate + 1) > Sysdate And A.末级 = 1" & _
        "           And (a.院区 Is Null Or a.院区='" & gstrNodeNo & "')"
    Set rs开票点 = zlDatabase.OpenSQLRecord(strSQL, "获取开票点数据")
    Get开票点 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get收费员(ByRef rs收费员 As ADODB.Recordset) As Boolean
    '加载开票点
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select Distinct a.Id, a.编号, a.简码, a.姓名" & _
        " From 人员表 A, 人员性质说明 B" & _
        " Where a.Id = b.人员id And Nvl(a.撤档时间, Sysdate + 1) > Sysdate" & _
        "           And b.人员性质 In ('门诊挂号员', '门诊收费员', '预交收款员', '住院结帐员','入院登记员','发卡登记人','医生','护士')" & _
        "           And (a.站点 Is Null Or a.站点='" & gstrNodeNo & "')"
    Set rs收费员 = zlDatabase.OpenSQLRecord(strSQL, "获取收费员数据")
    Get收费员 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    zlCloseWindows = (Forms.Count = 0)
End Function

Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:释放资源
    '入参:objPati-病人信息集
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-04 17:50:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '实例数为0时，才放资源
    If glngInstanceCount > 0 Then Exit Function
    Call zlCloseWindows   '关闭窗体
    Err = 0: On Error Resume Next
    Set gobjEInvProviders = Nothing
    Set gobjEinvProvider = Nothing
    Set gcnOracle = Nothing
    
    zlReleaseResources = True
End Function

Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False, Optional blnNotTran As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '     blnNotTran-不处理事务
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    
    If blnNotTran = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub

Public Function GetUserInfo(ByVal strDBUser As String) As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.用户名 = strDBUser
    UserInfo.姓名 = strDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.部门名称 = "" & rsTmp!部门名
        UserInfo.简码 = "" & rsTmp!简码
        UserInfo.姓名 = "" & rsTmp!姓名
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetEInvoiceData(ByVal byt业务场合 As Byte, ByVal dt开始时间 As Date, ByVal dt结束时间 As Date, _
    ByRef rsEInvoice As ADODB.Recordset, Optional ByVal byt票据状态 As Byte, Optional ByVal byt时间类型 As Byte, _
    Optional ByVal bytQueryType As Byte, Optional ByVal varQueryValue As Variant, Optional ByVal str开票点 As String) As Boolean
    '获取电子票据数据
    '入参：
    '   byt业务场合 0-所有，1-收费，2-预交，3-结帐，4-挂号，5-就诊卡
    '   byt票据状态 0-所有，1-正常，2-冲红，3-有效
    '   byt时间类型 0-票据生成时间，1-费用时间
    '   bytQueryType 查询类型，0-所有，1-按病人ID查询，2-按费用单据号查询，3-按电子票据号查询
    '   varQueryValue 查询调整值，与 bytQueryType 配合使用
    '   str开票点 开票点编码
    Dim strSQL As String, strWhere As String
    Dim strSqlSub As String
    
    On Error GoTo ErrHandler
    If byt时间类型 = 0 Then strWhere = strWhere & " And a.生成时间 Between [1] And [2]"
    
    Select Case byt票据状态
    Case 0 '0-所有
    Case 1 '1-正常
        strWhere = strWhere & " And a.记录状态 In(1,3)"
    Case 2 '2-冲红
        strWhere = strWhere & " And a.记录状态 = 2"
    Case 3 '3-有效
        strWhere = strWhere & " And a.记录状态 = 1"
    End Select
    
    Select Case bytQueryType
    Case 0 '0-所有
    Case 1 '1-按病人ID查询
        strWhere = strWhere & " And a.病人ID = [5]"
    Case 2 '2-按费用单据号查询
        strWhere = strWhere & " And b.NO = [5]"
    Case 3 '3-按电子票据号查
        strWhere = strWhere & " And a.号码 = [5] "
    End Select
    
    If str开票点 <> "" Then strWhere = strWhere & " And a.开票点 = [6]"
    
    '1)预交款
    If byt业务场合 = 0 Or byt业务场合 = 2 Then
        strSQL = _
            " Select a.ID, b.收款时间 As 收费时间,a.票种,a.记录状态 As 票据状态, b.No, a.代码 As 票据代码, a.号码 As 票据号码,a.检验码, Decode(a.记录状态, 2, -1, 1) * a.票据金额 As 票据金额," & _
            "           a.结算ID,a.病人id,a.姓名,a.性别,a.年龄,a.门诊号,a.住院号,a.是否换开,a.纸质发票号,a.开票点, a.原票据ID," & _
            "           To_Char(To_Date(Substr(a.生成时间, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss') As 开票时间,a.退款ID,0 As 补结算" & _
            " From 电子票据使用记录 A,病人预交记录 B" & _
            " Where a.结算ID =b.ID And a.票种=2 And b.记录性质=1" & strWhere & _
                        IIf(byt时间类型 = 0, "", " And b.收款时间 Between [3] And [4]")
        '余额退款
        strSQL = strSQL & " Union All " & _
            " Select a.ID, b.收款时间 As 收费时间,a.票种,a.记录状态 As 票据状态, b.No, a.代码 As 票据代码, a.号码 As 票据号码,a.检验码, Decode(a.记录状态, 2, -1, 1) * a.票据金额 As 票据金额," & _
            "           a.结算ID,a.病人id,a.姓名,a.性别,a.年龄,a.门诊号,a.住院号,a.是否换开,a.纸质发票号,a.开票点, a.原票据id," & _
            "           To_Char(To_Date(Substr(a.生成时间, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss') As 开票时间,a.退款ID,0 As 补结算" & _
            " From 电子票据使用记录 A,病人预交记录 B" & _
            " Where a.退款ID =b.ID And a.票种=2 And b.记录性质=11" & strWhere & _
                        IIf(byt时间类型 = 0, "", " And b.收款时间 Between [3] And [4]")
    End If
    '2)就诊卡
    If byt业务场合 = 0 Or byt业务场合 = 5 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select a.ID, b.登记时间 As 收费时间,a.票种,a.记录状态 As 票据状态, b.No, a.代码 As 票据代码, a.号码 As 票据号码,a.检验码, Decode(a.记录状态, 2, -1, 1) * a.票据金额 As 票据金额," & _
            "           a.结算ID,a.病人id,a.姓名,a.性别,a.年龄,a.门诊号,a.住院号,a.是否换开,a.纸质发票号,a.开票点, a.原票据ID," & _
            "           To_Char(To_Date(Substr(a.生成时间, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss') As 开票时间,a.退款ID,0 As 补结算" & _
            " From 电子票据使用记录 A,住院费用记录 B" & _
            " Where a.结算ID =b.结帐ID And a.票种=5 And b.记录性质=5 And b.记录状态 In(1,3)" & strWhere & _
                        IIf(byt时间类型 = 0, "", " And b.登记时间 Between [3] And [4]")
    End If
    '3)结帐
    If byt业务场合 = 0 Or byt业务场合 = 3 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select a.ID, b.收费时间,a.票种,a.记录状态 As 票据状态, b.No, a.代码 As 票据代码, a.号码 As 票据号码,a.检验码, Decode(a.记录状态, 2, -1, 1) * a.票据金额 As 票据金额," & _
            "           a.结算ID,a.病人id,a.姓名,a.性别,a.年龄,a.门诊号,a.住院号,a.是否换开,a.纸质发票号,a.开票点, a.原票据ID," & _
            "           To_Char(To_Date(Substr(a.生成时间, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss') As 开票时间,a.退款ID,0 As 补结算" & _
            " From 电子票据使用记录 A,病人结帐记录 B" & _
            " Where a.结算ID =b.ID And a.票种=3 And b.记录状态 In(1,3)" & strWhere & _
                        IIf(byt时间类型 = 0, "", " And b.收费时间 Between [3] And [4]")
    End If
    '4)挂号、收费
    If byt业务场合 = 0 Or byt业务场合 = 1 Or byt业务场合 = 4 Then
        strSqlSub = _
            " Select a.ID, b.登记时间 As 收费时间,a.票种,a.记录状态 As 票据状态, Min(b.No) As No, a.代码 As 票据代码, a.号码 As 票据号码,a.检验码, Decode(a.记录状态, 2, -1, 1) * a.票据金额 As 票据金额," & _
            "           a.结算ID,a.病人id,a.姓名,a.性别,a.年龄,a.门诊号,a.住院号,a.是否换开,a.纸质发票号,a.开票点, a.原票据ID," & _
            "           Max(To_Char(To_Date(Substr(a.生成时间, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss')) As 开票时间,a.退款ID,0 As 补结算" & _
            " From 电子票据使用记录 A,门诊费用记录 B" & _
            " Where a.结算ID =b.结帐ID And a.票种=[票种] And b.记录性质=[记录性质] And b.记录状态 In(1,3)" & strWhere & _
                        IIf(byt时间类型 = 0, "", " And b.登记时间 Between [3] And [4]") & _
            " Group By b.登记时间,a.ID,a.票种,a.记录状态, a.代码, a.号码,a.检验码, a.票据金额," & _
            "           a.结算ID,a.病人id,a.姓名,a.性别,a.年龄,a.门诊号,a.住院号,a.是否换开,a.纸质发票号,a.开票点,a.原票据id,a.退款ID"
        
        '保险补充结算
        strSqlSub = strSqlSub & " Union All " & _
            " Select a.ID, b.登记时间 As 收费时间,a.票种,a.记录状态 As 票据状态, b.No, a.代码 As 票据代码, a.号码 As 票据号码,a.检验码, Decode(a.记录状态, 2, -1, 1) * a.票据金额 As 票据金额," & _
            "           a.结算ID,a.病人id,a.姓名,a.性别,a.年龄,a.门诊号,a.住院号,a.是否换开,a.纸质发票号,a.开票点, a.原票据ID," & _
            "           To_Char(To_Date(Substr(a.生成时间, 1, 14), 'yyyymmddHH24miss'), 'yyyy-mm-dd HH24:mi:ss') As 开票时间,a.退款ID,1 As 补结算" & _
            " From 电子票据使用记录 A,费用补充记录 B" & _
            " Where a.结算ID =b.结算ID And a.票种=[票种]  And b.记录性质=[记录性质] And Nvl(b.附加标志,0)=[附加标志] And b.记录状态 In(1,3)" & strWhere & _
                        IIf(byt时间类型 = 0, "", " And b.登记时间 Between [3] And [4]")
                    
        If byt业务场合 = 0 Or byt业务场合 = 1 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(strSqlSub, "[记录性质]", 1), "[附加标志]", 0), "[票种]", 1)
        End If
        
        If byt业务场合 = 0 Or byt业务场合 = 4 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(strSqlSub, "[记录性质]", 4), "[附加标志]", 1), "[票种]", 4)
        End If
    End If
    strSQL = strSQL & " Order By 收费时间"
    Set rsEInvoice = zlDatabase.OpenSQLRecord(strSQL, "获取电子票据数据", _
        Format(dt开始时间, "yyyyMMddHHmmss"), Format(dt结束时间, "yyyyMMddHHmmss"), dt开始时间, dt结束时间, varQueryValue, str开票点)
    GetEInvoiceData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEInvoiceExse(ByVal byt业务场合 As Byte, ByVal lng结算ID As Long, ByRef rsExse As ADODB.Recordset) As Boolean
    '获取电子票据费用数据
    '入参：
    '   byt业务场合 1-收费，2-预交，3-结帐，4-挂号，5-就诊卡
    '   lng结算ID byt业务场合=2，病人预交记录.ID；其它，结帐ID
    '出参：
    '   rsExse 记录集：NO,序号,开单科室,开单人,费别,类别,名称,商品名,规格,单位,执行科室,数量,单价,应收金额,实收金额,结帐金额
    Dim strSQL As String, strWhere As String
    
    On Error GoTo ErrHandler
    Select Case byt业务场合
    '1)就诊卡，结帐
    Case 3, 5
        strSQL = _
            " Select a.No, Nvl(a.序号, a.价格父号) As 序号, a.开单部门id, a.开单人, a.费别," & _
            "        a.收费类别, a.收费细目id, a.计算单位 As 单位, a.执行部门id," & _
            "        Avg(Nvl(a.付数, 1) * a.数次) As 数量, Sum(a.标准单价) As 单价," & _
            "        Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额,Sum(a.结帐金额) As 结帐金额" & _
            " From 住院费用记录 A, 住院费用记录 A1" & _
            " Where a.记录性质 = a1.记录性质 And a.No = a1.No And a.序号 = a1.序号 And a1.结帐id = [1]" & _
            " Group By a.No, a.记录性质, a.记录状态, Nvl(a.序号, a.价格父号), a.开单部门id, a.开单人, a.费别," & _
            "       a.收费类别, a.收费细目id, a.计算单位, a.执行部门id"
    '2)挂号，收费
    Case 1, 4
        strSQL = _
        " Select a.No, Nvl(a.序号, a.价格父号) As 序号, a.开单部门id, a.开单人, a.费别," & _
        "        a.收费类别, a.收费细目id, a.计算单位 As 单位, a.执行部门id," & _
        "        Avg(Nvl(a.付数, 1) * a.数次) As 数量, Sum(a.标准单价) As 单价," & _
        "        Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额" & _
        " From 门诊费用记录 A," & _
        "      (Select a.记录性质, a.No, a.序号" & _
        "        From 门诊费用记录 A" & _
        "        Where a.结帐id = [1] And Not Exists (Select 1 From 费用补充记录 Where 收费结帐id = a.结帐id)" & _
        "        Union All" & _
        "        Select a.记录性质, a.No, a.序号" & _
        "        From 门诊费用记录 A, 费用补充记录 B" & _
        "        Where a.结帐id = b.收费结帐id And b.结算id = [1]) A1" & _
        " Where Mod(a.记录性质, 10) = a1.记录性质 And a.No = A1.No And a.序号 = A1.序号" & _
        " Group By a.No, a.记录性质, a.记录状态, Nvl(a.序号, a.价格父号), a.开单部门id, a.开单人, a.费别," & _
        "       a.收费类别, a.收费细目id, a.计算单位, a.执行部门id"
    Case Else
        Exit Function
    End Select
    
    strSQL = _
        " Select a.No, a.序号, b.名称 As 开单科室, a.开单人, a.费别, c.名称 As 类别, e.名称, f.名称 As 商品名," & _
        "        e.规格, a.单位, B1.名称 As 执行科室, Sum(a.数量) As 数量, Avg(a.单价) As 单价," & _
        "        Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额" & _
        " From (" & strSQL & ") A, 部门表 B, 部门表 B1, 收费项目类别 C, 收费项目目录 E, 收费项目别名 F" & _
        " Where a.开单部门id = b.Id And a.执行部门id = B1.Id And a.收费类别 = c.编码 And a.收费细目id = e.Id" & _
        "       And e.Id = f.收费细目id(+) And f.码类(+) = 1 And f.性质(+) = 3" & _
        " Group By a.No, a.序号, b.名称, a.开单人, a.费别, c.名称, e.名称, f.名称, e.规格, a.单位, B1.名称" & _
        " Having Nvl(Sum(a.数量), 0) <> 0" & _
        " Order By a.No, a.序号"

    Set rsExse = zlDatabase.OpenSQLRecord(strSQL, "获取费用数据", lng结算ID)
    GetEInvoiceExse = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetExseData(ByVal byt业务场合 As Byte, ByVal str收费员 As String, _
    ByVal dt开始时间 As Date, ByVal dt结束时间 As Date, ByRef rsExse As ADODB.Recordset) As Boolean
    '获取电子票据费用数据
    '入参：
    '   byt业务场合 0-所有，1-收费，2-预交，3-结帐，4-挂号，5-就诊卡
    Dim strSQL As String, strWhere As String, strSqlSub As String
    
    On Error GoTo ErrHandler
    strWhere = " And a.收款时间 Between [1] And [2]"
    If Trim(str收费员) <> "" Then strWhere = strWhere & " And a.操作员姓名=[3]"
    
    '1)预交款
    If byt业务场合 = 0 Or byt业务场合 = 2 Then
        strSQL = _
            " Select 2 As 业务类型, a.Id As 结算ID, a.No, a.金额, a.操作员姓名, a.收款时间," & _
            "           a.病人id, a.主页id, Null As 姓名, Null As 性别, Null As 年龄, a.预交类别, Null As 冲销ID, Null As 结帐类型, Null As 补结算" & _
            " From 病人预交记录 A" & _
            " Where a.记录性质 = 1 And a.记录状态 = 1 And a.预交电子票据 = 1" & strWhere & _
            "       And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.Id And 票种 = 2 And 记录状态 = 1)"
        '余额退款
        strSQL = strSQL & " Union All" & _
            " Select 12 As 业务类型, b.ID As 结算ID, a.No, a.金额, a.操作员姓名, a.收款时间," & _
            "           a.病人id, a.主页id, Null As 姓名, Null As 性别, Null As 年龄, a.预交类别,a.Id As 冲销ID, Null As 结帐类型, Null As 补结算" & _
            " From 病人预交记录 A,病人预交记录 B" & _
            " Where a.记录性质 = 11 And a.记录状态 = 1 And a.预交电子票据 = 1" & strWhere & _
            "       And Exists(Select 1 From 病人预交记录 Where 记录性质 = 1 And 附加标志 = 1 And 结帐id = a.结帐id)" & _
            "       And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.Id And 票种 = 2 And 记录状态 = 1)" & _
            "       And a.No=b.No And b.记录性质=1 And b.记录状态 In(1,3)"
    End If
    '2)就诊卡
    If byt业务场合 = 0 Or byt业务场合 = 5 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select 5 As 业务类型, b.结帐id As 结算ID, b.No, Sum(a.结帐金额) As 金额, b.操作员姓名, b.收款时间," & _
            "           a.病人id, a.主页id, a.姓名, a.性别, a.年龄, Null As 预交类别, Null As 冲销ID, Null As 结帐类型, Null As 补结算" & _
            " From 住院费用记录 A, 住院费用记录 A1," & _
            "      (Select Distinct a.结帐id, a.操作员姓名, a.收款时间, b.No" & _
            "       From 病人预交记录 A, 住院费用记录 B" & _
            "       Where a.结帐id = b.结帐ID And b.记录性质 = 5 And b.记录状态 In(1,3) And a.是否电子票据 = 1" & strWhere & _
            "             And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.结帐Id And 票种 = 5 And 记录状态 = 1)) B" & _
            " Where a.No = a1.No And a.序号 = a1.序号 And a1.结帐id = b.结帐ID And a.记录性质 = 5" & _
            " Group By b.结帐id, b.No, b.操作员姓名, b.收款时间, a.病人id, a.主页id, a.姓名, a.性别, a.年龄" & _
            " Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0"
    End If
    '3)结帐
    If byt业务场合 = 0 Or byt业务场合 = 3 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select Distinct 3 As 业务类型, a.结帐id As 结算ID, b.No, b.结帐金额 As 金额, a.操作员姓名, a.收款时间," & _
            "           b.病人id, b.主页id, Null As 姓名, Null As 性别, Null As 年龄, Null As 预交类别, Null As 冲销ID, b.结帐类型, Null As 补结算" & _
            " From 病人预交记录 A, 病人结帐记录 B" & _
            " Where a.结帐id = b.ID And b.记录状态 = 1 And a.是否电子票据 = 1" & strWhere & _
            "       And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.结帐Id And 票种 = 3 And 记录状态 = 1)"
    End If
    '4)挂号、收费
    If byt业务场合 = 0 Or byt业务场合 = 1 Or byt业务场合 = 4 Then
        strSqlSub = _
            " Select a.结帐id, a.操作员姓名, a.收款时间, Min(b.No) As No, 0 As 补结算, Null As 结算ID" & _
            " From 病人预交记录 A, 门诊费用记录 B" & _
            " Where a.结帐id = b.结帐ID And b.记录性质 = [记录性质] And b.记录状态 In(1,3) And a.是否电子票据 = 1" & strWhere & _
            "             And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.结帐id And 票种 = [票种] And 记录状态 = 1)" & _
            "             And Not Exists(Select 1 From 费用补充记录 Where 收费结帐id = a.结帐id And 记录性质 = 1 And Nvl(附加标志,0) = [附加标志])" & _
            " Group By a.结帐id, a.操作员姓名, a.收款时间"
        '补充结算
        strSqlSub = strSqlSub & " Union All " & _
            " Select 结帐id, 操作员姓名, 收款时间, No, 1 As 补结算, 结算ID" & _
            " From (Select Distinct b.收费结帐ID As 结帐id, a.操作员姓名, a.收款时间,b.No As No, b.结算ID," & _
            "                    Row_Number() Over(Partition By b.记录性质, b.No Order By b.登记时间) As 组号" & _
            "            From 病人预交记录 A, 费用补充记录 B" & _
            "            Where a.结帐ID=b.结算ID And b.记录性质 = 1 And Nvl(b.附加标志,0) = [附加标志] And b.记录状态 In(1, 3) And a.是否电子票据 = 1" & strWhere & _
            "                       And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.结帐id And 票种 = [票种] And 记录状态 = 1))" & _
            " Where 组号 = 1"
            
        strSqlSub = _
            " Select [业务类型] As 业务类型, Nvl(b.结算ID, b.结帐id) As 结算ID, Min(b.No) As No, Sum(a.结帐金额) As 金额, b.操作员姓名, b.收款时间," & _
            "        a.病人id, a.主页id, a.姓名, a.性别, a.年龄, Null As 预交类别, Null As 冲销ID, Null As 结帐类型, b.补结算" & _
            " From 门诊费用记录 A, 门诊费用记录 A1,(" & strSqlSub & ") B" & _
            " Where a.No = a1.No And a.序号 = a1.序号 And a1.结帐id = b.结帐ID And Mod(a.记录性质,10)=[记录性质]" & _
            " Group By Nvl(b.结算ID, b.结帐id), b.操作员姓名, b.收款时间, a.病人id, a.主页id, b.补结算, a.姓名, a.性别, a.年龄" & _
            " Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0"
        
        If byt业务场合 = 0 Or byt业务场合 = 1 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(Replace(strSqlSub, "[业务类型]", 1), "[记录性质]", 1), "[票种]", 1), "[附加标志]", 0)
        End If
        
        If byt业务场合 = 0 Or byt业务场合 = 4 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(Replace(strSqlSub, "[业务类型]", 4), "[记录性质]", 4), "[票种]", 4), "[附加标志]", 1)
        End If
    End If
    
    strSQL = _
        " Select Nvl(n.姓名,Nvl(m.姓名,a.姓名)) As 姓名,Nvl(n.性别,Nvl(m.性别,a.性别)) As 性别,Nvl(n.年龄,Nvl(m.年龄,a.年龄)) As 年龄," & _
        "           m.门诊号 As 门诊号, Nvl(n.住院号,m.住院号) As 住院号, a.业务类型, a.结算id, a.No, a.金额, a.操作员姓名, a.收款时间, " & _
        "           a.病人id, a.主页id, a.预交类别, a.冲销id, a.结帐类型, a.补结算" & _
        " From (" & strSQL & ") A, 病人信息 M, 病案主页 N" & _
        " Where a.病人ID=m.病人ID(+) And a.病人ID=n.病人ID(+) And a.主页ID=n.主页ID(+)" & _
        " Order By 收款时间"
    Set rsExse = zlDatabase.OpenSQLRecord(strSQL, "获取电子票据数据", dt开始时间, dt结束时间, str收费员)
    GetExseData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub Grid_SelAllRecord(vsfGrid As VSFlexGrid, ByVal blnSel As Boolean, Optional ByVal strColName As String = "选择")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:全选/全清记录
    '入参:
    '   blnSel-选择
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsfGrid
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex(strColName)) = blnSel
        Next
    End With
End Sub

Public Function GetSwapCollectFromBalanceID(ByVal byt场合 As Byte, ByVal lng原结算ID As Long, _
    ByRef cllSwapData_Out As Collection, Optional ByVal bln补结算 As Boolean, _
    Optional ByVal lng冲销ID As Long, Optional ByVal blnShowMsg As Boolean = True, _
    Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算ID获取交易结算信息
    '入参:
    '    byt场合-1-收费, 2-预交, 3-结帐, 4-挂号;5-就诊卡
    '   lng原结算ID byt场合=2，病人预交记录.ID；其它，结帐ID
    '出参:
    '   cllSwapData_Out-返回结算信息
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类),key(_节点名称)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection, cllBalanceInfo As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, strInsureSql As String
    
    On Error GoTo ErrHandler
    Select Case byt场合
    Case 1, 4
        If bln补结算 Then
            strWhere = " And b.结帐id In(Select 收费结帐ID From 费用补充记录 Where 结算ID=[1])"
        Else
            strWhere = " And b.结帐id = [1]"
        End If
    
        strSQL = _
            " Select Max(a.病人id) As 病人ID, Max(a.主页id) As 主页ID, Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄," & _
            "        f_List2Str(Cast(Collect(a.No) As t_StrList)) As NO, Sum(a.结帐金额) As 结帐金额, Max(a.登记时间) As 收费时间" & _
            " From (Select a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.No, a.序号, Sum(a.结帐金额) As 结帐金额, Max(b.登记时间) As 登记时间" & _
            "        From 门诊费用记录 A, 门诊费用记录 B" & _
            "        Where Mod(a.记录性质, 10) = Mod(b.记录性质, 10) And a.No = b.No And a.序号 = b.序号" & strWhere & _
            "        Group By a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.No, a.序号" & _
            "        Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0) A"
        
        strInsureSql = "Select Max(险类) As 险类 From 保险结算记录 Where 性质 = 1 And 记录id = [1]"
        
        strSQL = _
            " Select a.病人id, a.主页id, a.姓名, a.性别, a.年龄, m.门诊号, Nvl(n.住院号, m.住院号) As 住院号," & _
            "           a.No, a.结帐金额, a.收费时间, b.险类" & _
            " From (" & strSQL & ") A, (" & strInsureSql & ") B, 病人信息 M, 病案主页 N" & _
            " Where a.病人id = m.病人id(+) And a.病人id = n.病人id(+) And a.主页id = n.主页id(+) And a.No Is Not Null"
    Case 2
        strSQL = _
            "   Select a.Id, a.No, a.病人id, a.主页id, Sum(A.金额) As 结帐金额, Max(A.预交电子票据) As 是否电子票据, " & _
            "          Max(Nvl(d.姓名, c.姓名)) As 姓名, " & _
            "          Max(Nvl(d.性别, c.性别)) As 性别, Max(Nvl(d.年龄, c.年龄)) As 年龄, Max(Nvl(d.住院号, c.住院号)) As 住院号, Max(c.门诊号) As 门诊号, " & _
            "          max(M.险类) as 险类,to_char(max(A.收款时间),'yyyy-mm-dd hh24:mi:ss') as 收费时间,max(a.预交类别) as 预交类别" & _
            "   From  病人预交记录 A, 病人信息 C, 病案主页 D,(Select 记录ID, 险类 From 保险结算记录 where 性质=3  and 记录ID=[1] ) M" & _
            "   Where a.病人id = c.病人id(+) And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And a.Id=[1]  And A.ID=M.记录ID(+)" & _
            "   Group By a.Id, a.No, a.病人id, a.主页id"
    Case 3
        strSQL = _
            "   Select a.Id, a.No, a.病人id, a.主页id, Sum(b.冲预交) As 结帐金额, Max(b.是否电子票据) As 是否电子票据, " & _
            "          Max(decode(nvl(A.病人ID,0),0,A.原因,Nvl(d.姓名, c.姓名))) As 姓名, " & _
            "          Max(Nvl(d.性别, c.性别)) As 性别, Max(Nvl(d.年龄, c.年龄)) As 年龄, Max(Nvl(d.住院号, c.住院号)) As 住院号, Max(c.门诊号) As 门诊号, " & _
            "          max(M.险类) as 险类,to_char(max(A.收费时间),'yyyy-mm-dd hh24:mi:ss') as 收费时间,max(A.结帐类型) as 结帐类型" & _
            "   From 病人结帐记录 A, 病人预交记录 B, 病人信息 C, 病案主页 D,(Select 记录ID, 险类 From 保险结算记录 where 性质=2  and 记录ID=[1] ) M" & _
            "   Where a.id=b.结帐ID and  a.病人id = c.病人id(+) And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And a.Id=[1]  And A.ID=M.记录ID(+)" & _
            "   Group By a.Id, a.No, a.病人id, a.主页id"
    Case 5
        strSQL = _
            "   Select a.结帐id As ID, b.No, a.病人id, a.主页id, Sum(a.冲预交) As 结帐金额, Max(a.是否电子票据) As 是否电子票据, Max(c.姓名) As 姓名, Max(c.性别) As 性别, " & _
            "          Max(c.年龄) As 年龄, Max(c.住院号) As 住院号, Max(c.门诊号) As 门诊号, 0 As 险类, " & _
            "          To_Char(Max(a.收款时间), 'yyyy-mm-dd hh24:mi:ss') As 收费时间 " & _
            "   From 病人预交记录 A, (Select  结帐id,No From 住院费用记录 Where 结帐id = [1]) B, 病人信息 C  " & _
            "   Where a.结帐id = b.结帐id And a.病人id = c.病人id(+)  And a.结帐id = [1] " & _
            "   Group By a.结帐id, b.No, a.病人id, a.主页id"
    Case Else
        strErrMsg_Out = "传入场合【" & byt场合 & "】参数无效。": Exit Function
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据结帐ID构建电子票据信息", IIf(byt场合 = 2 And lng冲销ID <> 0, lng冲销ID, lng原结算ID))
    If rsTemp.EOF Then strErrMsg_Out = "无剩余未退费用数据。": Exit Function
    
    '1.创建病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类)
    Set cllPati = New Collection
    cllPati.Add Val(Nvl(rsTemp!病人ID)), "_病人ID"
    cllPati.Add Val(Nvl(rsTemp!主页id)), "_主页ID"
    cllPati.Add Nvl(rsTemp!姓名), "_姓名"
    cllPati.Add Nvl(rsTemp!性别), "_性别"
    cllPati.Add Nvl(rsTemp!年龄), "_年龄"
    cllPati.Add Nvl(rsTemp!门诊号), "_门诊号"
    cllPati.Add Nvl(rsTemp!住院号), "_住院号"
    cllPati.Add Val(Nvl(rsTemp!险类)), "_险类"

    '2.创建结算信息:(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add "", "_发票号"
    cllBalanceInfo.Add lng原结算ID, "_结算ID"
    cllBalanceInfo.Add lng冲销ID, "_冲销ID"
    cllBalanceInfo.Add Nvl(rsTemp!No), "_单据号"
    cllBalanceInfo.Add Format(Nvl(rsTemp!收费时间), "yyyy-mm-dd HH:MM:SS"), "_登记时间"
    If byt场合 = 1 Or byt场合 = 4 Then
        cllBalanceInfo.Add IIf(bln补结算, 1, 0), "_是否补结算"
    Else
        cllBalanceInfo.Add 0, "_是否补结算"
    End If
    cllBalanceInfo.Add 0, "_是否部分退款"
    cllBalanceInfo.Add UserInfo.编号, "_操作员编号"
    cllBalanceInfo.Add UserInfo.姓名, "_操作员姓名"
    cllBalanceInfo.Add Val(Nvl(rsTemp!结帐金额)), "_结算金额"
    cllBalanceInfo.Add 0, "_领用ID"
    Select Case byt场合
    Case 2
        cllBalanceInfo.Add Decode(Val(Nvl(rsTemp!预交类别)) = 0, 3, Val(Nvl(rsTemp!预交类别))), "_结算类型" '预交类别:1-门诊;2-住院 ;3-门诊和住院;
        cllBalanceInfo.Add 0, "_合约单位结帐"
    Case 3
        cllBalanceInfo.Add Decode(Val(Nvl(rsTemp!结帐类型)) = 0, 3, Val(Nvl(rsTemp!结帐类型))), "_结算类型"  '结帐类型:1-门诊;2-住院 ;3-门诊和住院;
        cllBalanceInfo.Add IIf(Val(Nvl(rsTemp!病人ID)) = 0, 1, 0), "_合约单位结帐"
    Case Else
        cllBalanceInfo.Add 1, "_结算类型"
        cllBalanceInfo.Add 0, "_合约单位结帐"
    End Select
    
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    
    GetSwapCollectFromBalanceID = True
    Exit Function
ErrHandler:
    If Not blnShowMsg Then strErrMsg_Out = Err.Description: Exit Function
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get医疗卡付款方式名称(ByVal str医疗卡付款方式编码 As String) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据医疗卡付款方式编码获取名称
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    strSQL = "Select Max(名称) as 名称 From 医疗付款方式 Where 编码 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get医疗卡付款方式名称", str医疗卡付款方式编码)
    If rsTmp.RecordCount > 0 Then
        Get医疗卡付款方式名称 = Nvl(rsTmp!名称)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Get预交单据总额(ByVal strNO As String) As Double
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人预交余额
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    strSQL = "Select Sum(金额) As 票据总金额" & vbNewLine & _
            "  From (Select Sum(金额) As 金额" & vbNewLine & _
            "         From 病人预交记录" & vbNewLine & _
            "         Where NO = r_Deposit_Rec.No And 记录性质 = 1" & vbNewLine & _
            "         Union All" & vbNewLine & _
            "         Select Sum(冲预交) As 金额" & vbNewLine & _
            "         From 病人预交记录" & vbNewLine & _
            "         Where 结帐id In (Select Distinct 结帐id From 病人预交记录 Where NO = [1] And Mod(记录性质, 10) = 1) And" & vbNewLine & _
            "               Nvl(金额, 0) < 0 And Mod(记录性质, 10) = 1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get预交余额", strNO)
    If rsTmp.RecordCount > 0 Then
        Get预交单据总额 = Val(Nvl(rsTmp!票据总金额))
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Get预交余额(ByVal lng病人ID As Long, ByVal int预交类型 As Integer) As Double
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人预交余额
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    strSQL = "Select Max(预交余额) As 预交余额 From 病人余额 " & _
            " Where 病人id = [1] And 性质 = 1 And 类型 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get预交余额", lng病人ID, int预交类型)
    If rsTmp.RecordCount > 0 Then
        Get预交余额 = Val(Nvl(rsTmp!预交余额))
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetEInvoiceInfo(ByVal lngEInvoiceID As Long, strErrMsg_Out As String) As ADODB.Recordset
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    strSQL = _
    " Select ID, 票种, 记录状态, 结算id, 病人id, 姓名, 性别, 年龄, 门诊号, 住院号, 代码 As 票据代码, 号码 As 票据号码, 检验码 As 票据校验码, 票据金额," & _
    "           生成时间, Url内网, 原票据id, 是否换开, 纸质发票号, 打印id, 备注, 操作员编号, 操作员姓名, 登记时间, 开票点, 系统来源, Url外网 " & _
    " From 电子票据使用记录 Where Id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetEInvoiceInfo", lngEInvoiceID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "未找到电子票据使用记录，请检查。": Exit Function
    End If
    Set GetEInvoiceInfo = rsTmp
    Exit Function
errHand:
    strErrMsg_Out = Err.Description
End Function

Public Function GetEInvoiceWithPatiInfo(ByVal lngEInvoiceID As Long, strErrMsg_Out As String) As ADODB.Recordset
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    strSQL = "Select a.票种, a.代码 As 票据代码, a.号码 As 票据号码, a.检验码 As 票据校验码, b.手机号, b.email," & vbNewLine & _
            "        a.是否换开 " & vbNewLine & _
            "From 电子票据使用记录 a, 病人信息 b" & vbNewLine & _
            "Where a.Id =[] And a.病人id = b.病人id(+)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetEInvoiceInfo", lngEInvoiceID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "未找到电子票据使用记录，请检查。": Exit Function
    End If
    Set GetEInvoiceWithPatiInfo = rsTmp
    Exit Function
errHand:
    strErrMsg_Out = Err.Description
End Function

Public Sub zl_VsGridGotFocus(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：进入网格控件时选择的颜色
    '入参：CustomColor-自定颜色
    '编制：刘兴洪
    '日期：2010-03-23 10:52:23
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    '进入控件
    With vsGrid
         If CustomColor <> -1 Then
             .FocusRect = flexFocusSolid
             .HighLight = flexHighlightNever
             .BackColorSel = vbBlue
             If .Row >= .FixedRows Then
                If .Rows - 1 > .FixedRows Then  '清除选择颜色
                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
                End If
                 .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
             End If
              
         Else
            .FocusRect = flexFocusSolid 'IIf(vsGrid.Editable = flexEDNone, flexFocusNone, flexFocusSolid)
            .HighLight = flexHighlightNever
            .BackColorSel = GRD_GOTFOCUS_COLORSEL
        End If
    End With
    Call zl_VsGridRowChange(vsGrid, vsGrid.Row, vsGrid.Row, 0, 0)
End Sub

Public Sub zl_VsGridRowChange(ByVal vsGrid As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngNewRow As Long, _
    ByVal lngOldCol As Long, ByVal lngNewCol As Long, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：行列改变时,设置相关的颜色
    '入参：CustomColor-自定义颜色
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-03-23 11:22:38
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    '行改变时
    Err = 0: On Error Resume Next
    If lngOldRow = lngNewRow Then
        vsGrid.Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, vsGrid.Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
        Exit Sub
    End If
    With vsGrid
        .Cell(flexcpBackColor, lngOldRow, vsGrid.FixedCols, lngOldRow, .Cols - 1) = .BackColor
        .Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, .Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
    End With
End Sub

Public Sub zl_VsGridLOSTFOCUS(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1, Optional ForeColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
   '功能：离开网格控件时选择的颜色
    '入参：CustomColor-是否用自定义颜色来设置(BackColor)的方式来进行)
    '编制：刘兴洪
    '日期：2010-03-23 11:03:05
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With vsGrid
        If CustomColor <> -1 Then
            If .Row >= .FixedRows Then
                .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
            End If
        Else
            .SelectionMode = flexSelectionByRow
            .FocusRect = IIf(vsGrid.Editable = flexEDNone, flexFocusHeavy, flexFocusSolid)
            If ForeColor = -1 Then .HighLight = flexHighlightAlways
            .BackColorSel = GRD_LOSTFOCUS_COLORSEL
        End If
        If ForeColor <> -1 Then
            .Cell(flexcpForeColor, .Row, .FixedCols, .Row, .Cols - 1) = ForeColor
        End If
        .ForeColorSel = .ForeColor
    End With
End Sub
