Attribute VB_Name = "mdlIDKind"
Option Explicit
Public gobjSquare As Object '卡结部件
Public gobjCardDatabase As Object  '卡对象中的clsDataBase类
Public gobjCards As Cards    '所有的卡
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public glngModul As Long, glngSys As Long
Public gstrAviPath As String, gstrVersion As String
Public gstrProductName As String
Public gstrDBUser As String   '当前数据库用户
Public gstrUnitName As String '用户单位名称
Public gobjParent As Object

'公共部件(zl9ComLib)
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjControl As Object

'刷卡控制全局变量
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_F4 = vbKeyF4
Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20
Public glngInstanceCount As Long

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Function zlGetFromBedNumberToPatiID(ByVal cnOracle As ADODB.Connection, _
    ByVal lng病区ID As Long, _
    ByVal str床号 As String, Optional ByRef lng主页ID As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据床号获取病人ID
    '出参:lng主页ID-返回当前床号的主页ID
    '返回:成功返回病人ID,否则返回False
    '编制:刘兴洪
    '日期:2012-09-19 15:50:18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objDatabase As Object
    
    On Error GoTo errHandle
    Set objDatabase = GetCardSquareDataBaseObject(cnOracle)
    
    lng主页ID = 0
    strSQL = _
    "   Select  A.病人ID,A.主页ID" & _
    "   From 病人信息 A,床位状况记录 C" & _
    "   Where  A.病人ID=C.病人ID And A.停用时间 is NULL " & _
    "           And C.病区ID=[1] And C.床号=[2] "
    
    If objDatabase Is Nothing Then
        Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng病区ID, str床号)
    Else
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng病区ID, str床号)
    End If
    If rsTemp.EOF Then zlGetFromBedNumberToPatiID = 0: Exit Function
    lng主页ID = Val(Nvl(rsTemp!主页ID))
    zlGetFromBedNumberToPatiID = Val(Nvl(rsTemp!病人ID))
    Exit Function
errHandle:
    If Not objDatabase Is Nothing Then
        If objDatabase.ErrCenter() = 1 Then Resume
    Else
        If gobjComLib.ErrCenter() = 1 Then Resume
    End If
End Function

Private Function zlInitComponents(Optional lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化接口部件
    '编制:刘兴洪
    '日期:2012-08-16 11:09:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpand As String
    strExpand = lngCardTypeID
    If gobjSquare Is Nothing Then Exit Function
    '初始化卡结算部件
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlInitComponents (初始化接口部件)
    '入参: frmMain-调用的主窗体
    '        lngModule-HIS调用模块号
    '       lngSys-传入的系统号
    '       strDBUser-数据库用户名
    '       cnOracle -HIS/三方机构
    '       blnDeviceSet-设备设置调用初始化
    '       strExpand-扩展信息(可选:转入卡类别ID)
    zlInitComponents = gobjSquare.zlInitComponents(gobjParent, _
     glngModul, glngSys, gstrDBUser, _
      gcnOracle, False, strExpand)
End Function

Public Function zlInitCards(ByVal cnOracle As ADODB.Connection, ByVal RegType As gRegType) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡对象
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-15 16:43:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objCard As Card, bln模糊查找 As Boolean
    Dim strValue As String
    On Error GoTo errHandle
    If zlCreateSquare(cnOracle) = False Then Exit Function
    'zlGetCards(ByVal BytType As Byte) As Cards
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的卡对象
    '入参:bytType-0-所有医疗卡;
    '             1-启用的医疗卡,
    '             2-所有存在三方账户的三方卡
    '             3-启用的三方账户的医疗卡
    'Set rsTemp = gobjSquare.zlGetYLCards
    Set gobjCards = gobjSquare.zlGetCards(0)
    bln模糊查找 = False
    For Each objCard In gobjCards
        Call GetRegInFor(RegType, "医疗卡类别\" & objCard.名称, "回车符", strValue)
        Select Case strValue
            Case "启用"
                objCard.卡号长度 = objCard.卡号长度 + IIf(objCard.设备是否启用回车, 0, 1)
            Case "禁用"
                objCard.卡号长度 = objCard.卡号长度 - IIf(objCard.设备是否启用回车, 1, 0)
        End Select
        If objCard.是否模糊查找 And objCard.启用 And Not bln模糊查找 Then bln模糊查找 = True
    Next
    gobjCards.按缺省卡查找 = Not bln模糊查找
    zlInitCards = True
    Exit Function
errHandle:
    
    MsgBox Err.Description
End Function
Public Function zlGetKindCards(Optional strIDKindStr As String = "", Optional blnOnlyAccouct As Boolean = False, _
                                Optional NotAutoAppendKind As Boolean = False, Optional OnlyThreeCard As Boolean = False) As Cards
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的卡对象
    '返回: 成功,卡对象
    '编制:刘兴洪
    '日期:2012-08-15 16:58:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCards As Cards, objCard As Card
    Dim varData As Variant, i As Long, varTemp As Variant
    Dim blnFind As Boolean, j As Long
    Dim strKinds As String
    
    On Error GoTo errHandle
    If strIDKindStr = "" Then
        '缺省类别
        strIDKindStr = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0;手|手机号|0"
    End If
    Set objCards = New Cards
    varData = Split(strIDKindStr, ";")
    j = 1
    strKinds = ""
    If Not OnlyThreeCard Then
        For i = 0 To UBound(varData)
            '先找
            varTemp = Split(varData(i) & "||||||||||||", "|")
            If Trim(varTemp(1)) <> "" Then
                blnFind = False
                If Not gobjCards Is Nothing Then
                    For Each objCard In gobjCards
                        '76243,冉俊明,2014-8-5,其它开发组人员传入全角字母IC时,默认将其处理为系统中默认的IC卡类别
                        If objCard.名称 = Trim(varTemp(1)) _
                            Or (objCard.名称 Like "*IC卡*" And (varTemp(1) = "IC卡" Or varTemp(1) = "IC卡号" Or varTemp(1) Like "*ＩＣ卡*") And objCard.系统) _
                            Or (objCard.名称 Like "*身份证*" And (varTemp(1) = "二代身份证" Or varTemp(1) = "身份证" Or varTemp(1) = "身份证号") And objCard.系统) Then
                            blnFind = True
                            If InStr(strKinds & ",", "," & objCard.接口序号 & ",") = 0 Then
                                strKinds = strKinds & "," & objCard.接口序号
                                If objCard.启用 And Not objCard.消费卡 Then
                                   objCards.Add objCard, "K" & objCard.接口序号
                                End If
                            End If
                            Exit For
                        End If
                    Next
                End If
               If blnFind = False Then
                    '补充
                    Set objCard = New Card
                    '短名1|全名1|是否刷卡1|卡类别ID1|卡号长度1|缺省标志1(1-当前缺省;0-非缺省)|是否存在帐户1(1-存在帐户;0-不存在帐户)|
                    '卡号密文1(第几位至第几位加密,空为不加密)|是否扫描|是否接触式读卡|是否非接触式读卡
                    With objCard
                        .接口编码 = "-"
                        .名称 = varTemp(1)
                        .短名 = varTemp(0)
                        .是否刷卡 = Val(varTemp(2)) <> 1
                        .接口序号 = 0 ' IIf(Val(varTemp(3)) = 0, -j, Val(varTemp(3)))
                        .缺省标志 = Val(varTemp(4)) = 1
                        .是否存在帐户 = Val(varTemp(5)) = 1
                        .卡号密文规则 = Trim(varTemp(6))
                        '85565,李南春,2015/7/10:读卡性质，缺省为Fasle
                        .是否扫描 = Val(varTemp(7)) = 1
                        .是否接触式读卡 = Val(varTemp(8)) = 1
                        .是否非接触式读卡 = Val(varTemp(9)) = 1
                    End With
                    Err = 0: On Error Resume Next
                    objCards.Add objCard, "M" & objCard.名称
                    If Err <> 0 Then Err = 0: On Error GoTo 0
                    j = j + 1
               End If
            End If
        Next
    End If
    '未加入的，放入最后
    If NotAutoAppendKind = False Or OnlyThreeCard Then
        If Not gobjCards Is Nothing Then
            For Each objCard In gobjCards
                If InStr(1, strKinds & ",", "," & objCard.接口序号 & ",") = 0 And objCard.启用 And Not objCard.消费卡 Then
                    strKinds = strKinds & "," & objCard.接口序号
                    objCards.Add objCard, "K" & objCard.接口序号
                End If
            Next
        End If
    End If
    
    If Not gobjCards Is Nothing Then
        objCards.按缺省卡查找 = gobjCards.按缺省卡查找
        objCards.加密显示 = gobjCards.加密显示
    End If
    Set zlGetKindCards = objCards
    
    Err = 0: On Error Resume Next
    Erase varData '清空数组
    
    Exit Function
errHandle:
    
    MsgBox Err.Description
End Function

Public Function zlGetPati(ByVal cnOracle As ADODB.Connection, _
    ByVal lng病人ID As Long, ByRef objPati As PatiInfor, _
    ByRef strErrMsg As String, Optional strOtherName As String = "", _
    Optional strOtherValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID,重新获取数据
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-04-06 18:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere As String
    Dim objDatabase As Object
    
    On Error GoTo errHandle
    
    Set objDatabase = GetCardSquareDataBaseObject(cnOracle)
    Set objPati = New PatiInfor
    
    '检查该身份证下有此病人没有
    If strOtherName = "" Then
        strWhere = " And 病人ID=[1]"
    ElseIf strOtherName = "门诊号" Then
        strWhere = " And 门诊号=[2]"
    ElseIf strOtherName = "住院号" Then
        strWhere = " And 病人ID=(Select Max(病人ID) From 病案主页 Where 住院号 = [2])"
    Else
        strWhere = " And " & strOtherName & "=[3]"
    End If
    strSQL = "" & _
    "   Select a.病人id, a. 门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式,p.编码 as 医疗付款方式编码, a. 姓名, a.性别, a. 年龄, a.出生日期, a.出生地点, a.身份证号, a.其他证件, a.身份, " & _
    "        a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.监护人, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, " & _
    "        a.合同单位id, a.工作单位, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.担保人, a.担保额, a.担保性质, a.就诊时间, a.就诊状态, a.就诊诊室, a.在院, a.Ic卡号, " & _
    "        a.健康号, a.医保号, a.登记时间, a.停用时间, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, '' as 卡号, 0As 卡状态,'' as 密码, '' as 挂失方式, " & _
    "        a.病人类型 as 病人类型,sysdate as 挂失时间, 0  as 挂失有效天数,sysdate as 当前时间,a.手机号,a.险类,B.名称 险类名称" & _
    "   From 病人信息 A,保险类别 B,医疗付款方式 P" & _
    "   Where A.险类 = B.序号(+) And a.医疗付款方式=P.名称(+)  " & strWhere
    
    
    If Not objDatabase Is Nothing Then
        Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng病人ID, Val(strOtherValue), strOtherValue)
    Else
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng病人ID, Val(strOtherValue), strOtherValue)
    End If
    If rsTemp.EOF Then Exit Function
    objPati.病人ID = rsTemp!病人ID
    objPati.门诊号 = IIf(Val(Nvl(rsTemp!门诊号)) = 0, "", Nvl(rsTemp!门诊号))
    objPati.姓名 = Nvl(rsTemp!姓名)
    objPati.性别 = Nvl(rsTemp!性别)
    objPati.年龄 = Nvl(rsTemp!年龄)
    objPati.出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
    objPati.出生地址 = Nvl(rsTemp!出生地点)
    objPati.身份证号 = Nvl(rsTemp!身份证号)
    objPati.其他证件 = Nvl(rsTemp!其他证件)
    objPati.职业 = Nvl(rsTemp!职业)
    objPati.费别 = Nvl(rsTemp!费别)
    objPati.民族 = Nvl(rsTemp!民族)
    objPati.国籍 = Nvl(rsTemp!国籍)
    objPati.学历 = Nvl(rsTemp!学历)
    objPati.婚姻状况 = Nvl(rsTemp!婚姻状况)
    objPati.区域 = Nvl(rsTemp!婚姻状况)
    objPati.家庭地址 = Nvl(rsTemp!家庭地址)
    objPati.家庭电话 = Nvl(rsTemp!家庭电话)
    objPati.家庭邮编 = Nvl(rsTemp!家庭地址邮编)
    objPati.监护人 = Nvl(rsTemp!监护人)
    objPati.联系人 = Nvl(rsTemp!联系人姓名)
    objPati.联系人关系 = Nvl(rsTemp!联系人关系)
    objPati.联系人地址 = Nvl(rsTemp!联系人地址)
    objPati.联系人电话 = Nvl(rsTemp!联系人电话)
    objPati.工作单位 = Nvl(rsTemp!工作单位)
    objPati.工作单位电话 = Nvl(rsTemp!单位电话)
    objPati.工作单位邮编 = Nvl(rsTemp!单位邮编)
    objPati.工作单位开户行 = Nvl(rsTemp!单位开户行)
    objPati.工作单位开户行帐户 = Nvl(rsTemp!单位帐号)
    objPati.户口地址 = Nvl(rsTemp!户口地址)
    objPati.户口地址邮编 = Nvl(rsTemp!户口地址邮编)
    objPati.籍贯 = Nvl(rsTemp!籍贯)
    objPati.密码 = Nvl(rsTemp!密码)
    objPati.医疗付款方式编码 = Nvl(rsTemp!医疗付款方式编码)
    objPati.医疗付款方式 = Nvl(rsTemp!医疗付款方式)
    objPati.病人类型 = Nvl(rsTemp!病人类型)
    objPati.就诊卡号 = Nvl(rsTemp!就诊卡号)
    objPati.手机号 = Nvl(rsTemp!手机号)
    objPati.险类 = Val(Nvl(rsTemp!险类))
    objPati.险类名称 = Trim(Nvl(rsTemp!险类名称))
    zlGetPati = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function

Public Function FromObjectToCard(ByVal objCard As Object) As Card
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将Object对象换成Card对象
    '返回:成功Card对象
    '编制:刘兴洪
    '日期:2013-10-23 18:03:52
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As New Card
    With objCard
        objTemp.接口序号 = .接口序号
        objTemp.接口编码 = .接口编码
        objTemp.短名 = .短名
        objTemp.名称 = .名称
        objTemp.前缀文本 = .前缀文本
        objTemp.卡号长度 = .卡号长度
        objTemp.缺省标志 = .缺省标志
        objTemp.系统 = .系统
        objTemp.是否严格控制 = .是否严格控制
        objTemp.是否自动读取 = .是否自动读取
        objTemp.自动读取间隔 = .自动读取间隔
        objTemp.自制卡 = .自制卡
        objTemp.是否存在帐户 = .是否存在帐户
        objTemp.是否全退 = .是否全退
        objTemp.卡号重复使用 = .卡号重复使用
        objTemp.结算方式 = .结算方式
        objTemp.接口程序名 = .接口程序名
        objTemp.特定项目 = .特定项目
        objTemp.启用 = .启用
        objTemp.备注 = .备注
        objTemp.卡号密文规则 = .卡号密文规则
        objTemp.是否退现 = .是否退现
        objTemp.密码长度 = .密码长度
        objTemp.密码长度限制 = .密码长度限制
        objTemp.密码规则 = .密码规则
        objTemp.密码输入限制 = .密码输入限制
        objTemp.是否缺省密码 = .是否缺省密码
        objTemp.是否制卡 = .是否制卡
        objTemp.是否发卡 = .是否发卡
        objTemp.是否写卡 = .是否写卡
        objTemp.结算性质 = .结算性质
        '77872,李南春,2014/9/15:是否支持转帐及代扣
        objTemp.是否转帐及代扣 = .是否转帐及代扣
        objTemp.是否刷卡 = .是否刷卡    '85565,李南春,2015/7/10:读卡性质
        objTemp.是否扫描 = .是否扫描
        objTemp.是否接触式读卡 = .是否接触式读卡
        objTemp.是否非接触式读卡 = .是否非接触式读卡
        objTemp.是否持卡消费 = .是否持卡消费
        objTemp.是否退款验卡 = .是否退款验卡
        objTemp.是否缺省退现 = .是否缺省退现
    End With
    Set FromObjectToCard = objTemp
End Function

Public Function FromXMLPati(ByVal strPatiXml As String, ByRef strErrMsg As String) As PatiInfor
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从XML中获取病人信息
    '返回:病人信息对象
    '编制:刘兴洪
    '日期:2012-08-22 11:43:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strOutCardNO As String, strOutPatiInforXML As String
    Dim objNode As MSXML2.IXMLDOMElement, strExpand As String
    Dim objTempNode As MSXML2.IXMLDOMElement
    Dim strTmp As String, strValue As String
    Dim objPati As New PatiInfor
    Dim objXML As New objXML
    On Error GoTo errHandle
    
   If strPatiXml = "" Then Exit Function
   '加载病人信息
    If objXML.zlXML_Init = False Then Exit Function
    If objXML.zlXML_LoadXMLToDOMDocument(strPatiXml, False, strErrMsg) = False Then Exit Function
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    Call objXML.zlXML_GetNodeValue("卡号", , strValue)
    objPati.卡号 = strValue
    '    姓名    Varchar2    64
    Call objXML.zlXML_GetNodeValue("姓名", , strValue)
    objPati.姓名 = strValue
    '    性别    Varchar2    4
    Call objXML.zlXML_GetNodeValue("性别", , strValue)
    objPati.性别 = strValue
    '    年龄    Varchar2    10
    Call objXML.zlXML_GetNodeValue("年龄", , strValue)
    objPati.年龄 = strValue
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call objXML.zlXML_GetNodeValue("出生日期", , strValue)
    objPati.出生日期 = strValue
    '    出生地点    Varchar2    50
    Call objXML.zlXML_GetNodeValue("出生地点", , strValue)
    objPati.出生地址 = strValue
    '    身份证号    VARCHAR2    18
    Call objXML.zlXML_GetNodeValue("身份证号", , strValue)
    objPati.身份证号 = strValue
    If objPati.出生日期 = "" And strValue <> "" Then
        strTmp = gobjCommFun.GetIDCardDate(strValue)
        If IsDate(strTmp) Then objPati.出生日期 = strTmp
    End If
    '    其他证件    Varchar2    20
    Call objXML.zlXML_GetNodeValue("其他证件", , strValue)
    objPati.其他证件 = strValue
    '    职业    Varchar2    80
    Call objXML.zlXML_GetNodeValue("职业", , strValue)
    objPati.职业 = strValue
    '    民族    Varchar2    20
    Call objXML.zlXML_GetNodeValue("民族", , strValue)
    objPati.民族 = strValue
    '    国籍    Varchar2    30
    Call objXML.zlXML_GetNodeValue("国籍", , strValue)
    objPati.国籍 = strValue
    '    学历    Varchar2    10
    Call objXML.zlXML_GetNodeValue("学历", , strValue)
    objPati.学历 = strValue
    '    婚姻状况    Varchar2    4
    Call objXML.zlXML_GetNodeValue("婚姻状况", , strValue)
    objPati.婚姻状况 = strValue
    
    '    区域    Varchar2    30
    Call objXML.zlXML_GetNodeValue("区域", , strValue)
    objPati.区域 = strValue
    '    家庭地址    Varchar2    50
    Call objXML.zlXML_GetNodeValue("家庭地址", , strValue)
    objPati.家庭地址 = strValue
     '    户口地址    Varchar2    50
    Call objXML.zlXML_GetNodeValue("户口地址", , strValue)
    objPati.户口地址 = strValue
    '    家庭电话    Varchar2    20
    Call objXML.zlXML_GetNodeValue("家庭电话", , strValue)
    objPati.家庭电话 = strValue
    '    家庭地址邮编    Varchar2    6
    Call objXML.zlXML_GetNodeValue("家庭地址邮编", , strValue)
    objPati.家庭邮编 = strValue
    '    监护人  Varchar2    64
    Call objXML.zlXML_GetNodeValue("监护人", , strValue)
    objPati.监护人 = strValue
    
    '    联系人姓名  Varchar2    64
    Call objXML.zlXML_GetNodeValue("联系人姓名", , strValue)
    objPati.联系人 = strValue
    '    联系人关系  Varchar2    30
    Call objXML.zlXML_GetNodeValue("联系人关系", , strValue)
    objPati.联系人关系 = strValue
    '    联系人地址  Varchar2    50
    Call objXML.zlXML_GetNodeValue("联系人地址", , strValue)
    objPati.联系人地址 = strValue
    '    联系人电话  Varchar2    20
    Call objXML.zlXML_GetNodeValue("联系人电话", , strValue)
    objPati.联系人电话 = strValue
    '    工作单位    Varchar2    100
    Call objXML.zlXML_GetNodeValue("工作单位", , strValue)
    objPati.工作单位 = strValue
    '    单位电话    Varchar2    20
    Call objXML.zlXML_GetNodeValue("单位电话", , strValue)
    objPati.工作单位电话 = strValue
    '    单位邮编    Varchar2    6
    Call objXML.zlXML_GetNodeValue("单位邮编", , strValue)
    objPati.工作单位邮编 = strValue
    '    单位开户行  Varchar2    50
    Call objXML.zlXML_GetNodeValue("单位开户行", , strValue)
    objPati.工作单位开户行 = strValue
    '    单位帐号    Varchar2    20
    Call objXML.zlXML_GetNodeValue("单位帐号", , strValue)
   'txt单位帐号.Text = strValue
    objPati.工作单位开户行帐户 = strValue
    '    手机号    Varchar2    20
    Call objXML.zlXML_GetNodeValue("手机号", , strValue)
    objPati.手机号 = strValue
    '    照片文件    Varchar2    20
    Call objXML.zlXML_GetNodeValue("照片文件", , strValue)
    objPati.照片文件 = strValue
    '    照片
    If Trim(strValue) <> "" Then
        Err = 0: On Error Resume Next
        objPati.照片 = LoadPicture(strValue)
        If objPati.照片 = 0 Then objPati.照片 = Nothing
        Err = 0: On Error GoTo errHandle
    End If
    Set FromXMLPati = objPati
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlTranErrInfor(strErrMsg) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对错语信息进行格式化
    '返回: 返回被格式化的错语信息
    '编制:刘兴洪
    '日期:2012-08-22 14:47:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlTranErrInfor = strErrMsg
End Function

Public Function zlCreateSquare(ByVal cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建卡对象
    '编制:刘兴洪
    '日期:2012-08-15 16:40:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If Not gobjSquare Is Nothing Then zlCreateSquare = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjSquare = CreateObject("zl9CardSquare.clsCardsquare")
    If Err <> 0 Then Err = 0: Exit Function
    Call gobjSquare.zlInitComponents(gobjParent, glngModul, glngSys, gstrDBUser, cnOracle, False, strExpend)
    '初始部件不成功,则作为不存在处理
    zlCreateSquare = True
End Function

Public Function zlCreateSquareDataBaseObject(ByVal cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建clsDataBase对象(zlCardSquare部件)
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-06-03 11:02:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If Not gobjCardDatabase Is Nothing Then zlCreateSquareDataBaseObject = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjCardDatabase = CreateObject("zl9CardSquare.clsDataBase")
    If Err <> 0 Then Err = 0: Exit Function
    Call gobjCardDatabase.InitCommon(gcnOracle)
    zlCreateSquareDataBaseObject = True
    Err = 0: On Error GoTo 0
End Function
Public Function GetCardSquareDataBaseObject(cnOracle) As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡部件中的CardSquare对象
    '编制:刘兴洪
    '日期:2015-06-03 11:22:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjCardDatabase Is Nothing Then
        If zlCreateSquareDataBaseObject(cnOracle) = False Then Exit Function
    End If
    Call gobjCardDatabase.InitCommon(cnOracle)
    Set GetCardSquareDataBaseObject = gobjCardDatabase
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
    zlCloseWindows = Forms.Count = 0
End Function

Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:释放资源
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-13 10:30:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '实例数为0时，才放资源
    If glngInstanceCount > 0 Then Exit Function
    Call zlCloseWindows '释放窗体资源
    If Not gobjSquare Is Nothing Then Set gobjSquare = Nothing
    If Not gobjCardDatabase Is Nothing Then Set gobjCardDatabase = Nothing
    If Not gobjCards Is Nothing Then Set gobjCards = Nothing
    If Not gobjParent Is Nothing Then Set gobjParent = Nothing
    If Not gobjComLib Is Nothing Then Set gobjComLib = Nothing
    If Not gobjCommFun Is Nothing Then Set gobjCommFun = Nothing
    If Not gobjDatabase Is Nothing Then Set gobjDatabase = Nothing
    If Not gobjControl Is Nothing Then Set gobjControl = Nothing
    zlReleaseResources = True
End Function
