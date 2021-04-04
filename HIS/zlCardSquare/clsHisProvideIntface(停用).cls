VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHisProvideIntface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String
Private mblnStartTrans As Boolean   '是否启用了事务
Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    '初始化公共部件
    If gobjComLib Is Nothing Then
        Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
        Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
        Set gobjControl = GetObject("", "zl9Comlib.clsControl")
        Set gobjDataBase = GetObject("", "zl9Comlib.clsDatabase")
    End If
    Call zlinitSystemPara
    glngInstanceCount = glngInstanceCount + 1
End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
Public Function zlInitCompoent(ByVal cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlInitComponents (初始化接口部件)
    '入参:cnOracle-HIS/三方机构 传入连接对象
    '返回:函数返回True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 11:49:39
    '说明:如果是第三方接口主动读取数据,则在渠道更改的部件中传入连接.
    '    1.如果第三接口存在签到的情况，可以在此接口中进行签道.
    '    2.如果第三方接口读取数据,需要先初始化.
    '调用者:
    '    1.  第三方接口需要主动读取数据时,必须先初始化.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set gcnOracle = cnOracle
    InitCommon gcnOracle
    If zl_GetUserInfo = False Then
        MsgBox "注意:" & vbCrLf & "   操作员未对应相应的人员!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    zlInitCompoent = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetMedicalTreatment(ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗付款方式
    '入参:
    '出参:OutPutString_Xml-返回医院的费别信息,格式如下
    '返回:函数返回 True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 17:55:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objNode As MSXML2.IXMLDOMElement
    On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    strSQL = "Select 编码,名称, 简码,缺省标志 From 医疗付款方式"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医疗付款方式")
    With rsTemp
        Do While Not .EOF
            zlXML_InsertNode gobjXml.documentElement, "LISTDATA", "", objNode
            '       编码    Varchar2    2       Y
            zlXML_InsertNode objNode, "编码", NVL(rsTemp!编码)
            '       名称    Varchar2    10      Y
            zlXML_InsertNode objNode, "名称", NVL(rsTemp!名称)
            zlXML_InsertNode objNode, "简码", NVL(rsTemp!简码)
            '       缺省标志    Number  18          1-表示缺省使用的类别
            zlXML_InsertNode objNode, "缺省标志", Val(NVL(rsTemp!缺省标志))
            .MoveNext
        Loop
    End With
    OutPutString_Xml = zlXML_GetXMLString
    zlGetMedicalTreatment = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetFeeType(ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取费别
    '入参:
    '出参:OutPutString_Xml-返回医院的费别信息,格式如下
    '返回:函数返回 True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 17:55:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objNode As MSXML2.IXMLDOMElement
    On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    strSQL = "Select 编码,名称, 简码,缺省标志,有效开始,有效结束,说明 From 费别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取费别")
    With rsTemp
        Do While Not .EOF
            zlXML_InsertNode gobjXml.documentElement, "LISTDATA", "", objNode
            '       编码    Varchar2    2       Y
            zlXML_InsertNode objNode, "编码", NVL(rsTemp!编码)
            '       名称    Varchar2    10      Y
            zlXML_InsertNode objNode, "名称", NVL(rsTemp!名称)
            zlXML_InsertNode objNode, "简码", NVL(rsTemp!简码)
            '       缺省标志    Number  18          1-表示缺省使用的类别
            zlXML_InsertNode objNode, "缺省标志", Val(NVL(rsTemp!缺省标志))
            zlXML_InsertNode objNode, "有效开始时间", Format(rsTemp!有效开始, "yyyy-mm-dd HH:MM:SS")
            zlXML_InsertNode objNode, "有效结束时间", Format(rsTemp!有效结束, "yyyy-mm-dd HH:MM:SS")
            zlXML_InsertNode objNode, "说明", NVL(rsTemp!说明)
            .MoveNext
        Loop
    End With
    OutPutString_Xml = zlXML_GetXMLString
    zlGetFeeType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetclsPatientInfo(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-27 14:47:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, str门诊号 As String, str住院号 As String
    Dim str就诊卡号 As String, str身份证号 As String, lng卡类别ID As Long
    Dim str卡号 As String, strValue As String, str查询密码 As String, rsTemp As ADODB.Recordset
    Dim strSQL As String, strPassWord As String
    On Error GoTo errHandle
    If InputString_XML = "" Then Exit Function
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    'inputString (以下条件致少一个以上)
    '标识    数据类型    长度    精度    说明
    '病人ID  Number  18      可以传入病人ID;也可以不传,以便对建档病人的信息进行更正.
    Call zlXML_GetNodeValue("病人ID", , strValue): lng病人ID = Val(strValue)
    '门诊号  Number  18      可以为0,表示不根据门诊号查询
    Call zlXML_GetNodeValue("门诊号", , strValue): str门诊号 = zlFormatID(strValue)
    '住院号  Number  18      可以为0,表示不根据住院号查询
    Call zlXML_GetNodeValue("住院号", , strValue): str住院号 = zlFormatID(strValue)
    '就诊卡号    Varchar2    20      可以为空,表示不根据就诊卡号查询
    Call zlXML_GetNodeValue("就诊卡号", , strValue): str就诊卡号 = Trim(strValue)
    '身份证号    Varchar2    18      可以为空,表示不根据身份证号查询
    Call zlXML_GetNodeValue("身份证号", , strValue): str身份证号 = Trim(strValue)
    '卡类别ID    Number  18      可以为0:渠道可以根据三方机构固定传入
    Call zlXML_GetNodeValue("卡类别ID", , strValue): lng卡类别ID = Trim(strValue)
    '卡号    Varchar2    20      可以为空
    Call zlXML_GetNodeValue("卡号", , strValue): str卡号 = Trim(strValue)
    '查询密码    Varchar2    10
    Call zlXML_GetNodeValue("查询密码", , strValue): str查询密码 = Trim(strValue)
    If zlXML_Init = False Then Exit Function
    
    strSQL = "" & _
    "   Select A.病人id,A.门诊号,A.住院号,A.就诊卡号,A.卡验证码,A.费别,A.医疗付款方式,A.姓名,A.性别,A.年龄,A.出生日期,A.出生地点,A.身份证号, " & _
    "          A.其他证件,A.身份,A.职业,A.民族,A.国籍,A.区域,A.学历,A.婚姻状况,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.监护人,A.联系人姓名, " & _
    "          A.联系人关系,A.联系人地址,A.联系人电话,A.合同单位id,A.工作单位,A.单位电话,A.单位邮编,A.单位开户行,A.单位帐号,A.担保人, " & _
    "          A.担保额,A.担保性质,A.就诊时间,A.就诊状态,A.就诊诊室,A.住院次数,A.当前科室id,A.当前病区id,A.当前床号,A.入院时间,A.出院时间, " & _
    "          A.在院,A.Ic卡号,A.健康号,A.医保号,A.险类,A.登记时间,A.停用时间,A.锁定,A.查询密码 " & _
                IIf(lng卡类别ID <> 0, ",b.密码 ", "") & _
    "   From 病人信息 A " & IIf(lng卡类别ID <> 0, ",病人医疗卡信息 B ", "") & _
    "   Where 1=1 " & IIf(lng卡类别ID <> 0, " and A.病人ID=B.病人ID And B.卡类别ID=[1]   And  B.卡号=[2]", "") & _
                IIf(lng病人ID <> 0, " And A.病人ID=[3]", "") & _
                IIf(str身份证号 <> "", " And A.身份证号=[4]", "") & _
                IIf(str就诊卡号 <> "", " And A.就诊卡号=[5]", "") & _
                IIf(Val(str门诊号) <> 0, " And A.门诊号=[6]", "") & _
                IIf(Val(str住院号) <> 0, " And A.住院号=[7] ", "")
     Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng卡类别ID, str卡号, lng病人ID, str身份证号, str就诊卡号, str门诊号, str住院号)
    If rsTemp.EOF Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "未找到指定的病人信息"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    '检查密码是否正确
    If lng卡类别ID <> 0 Then
        strPassWord = NVL(rsTemp!密码)
    ElseIf str就诊卡号 <> "" Then
        strPassWord = NVL(rsTemp!卡验证码)
    Else
        strPassWord = NVL(rsTemp!查询密码)
    End If
    If strPassWord <> str查询密码 Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "密码输入错误!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    With rsTemp
        '病人ID  Number  18
        zlXML_InsertNode gobjXml.documentElement, "病人ID", Trim(NVL(rsTemp!病人ID))
        '门诊号  Number  18
        zlXML_InsertNode gobjXml.documentElement, "门诊号", Trim(NVL(rsTemp!门诊号))
        '住院号  Number  18
        zlXML_InsertNode gobjXml.documentElement, "住院号", Trim(NVL(rsTemp!住院号))
        '姓名    Varchar2    100
        zlXML_InsertNode gobjXml.documentElement, "姓名", Trim(NVL(rsTemp!姓名))
        '性别    Varchar2    4
        zlXML_InsertNode gobjXml.documentElement, "性别", Trim(NVL(rsTemp!性别))
        '年龄    Varchar2    10
        zlXML_InsertNode gobjXml.documentElement, "年龄", Trim(NVL(rsTemp!年龄))
        '出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
        zlXML_InsertNode gobjXml.documentElement, "出生日期", Format(rsTemp!出生日期, "yyyy-mm-dd HH:MM:SS")
        '出生地点    Varchar2    50
        zlXML_InsertNode gobjXml.documentElement, "出生地点", Trim(NVL(rsTemp!出生地点))
        '身份证号    VARCHAR2    18
        zlXML_InsertNode gobjXml.documentElement, "身份证号", Trim(NVL(rsTemp!身份证号))
        '其他证件    Varchar2    20
        zlXML_InsertNode gobjXml.documentElement, "其他证件", Trim(NVL(rsTemp!其他证件))
        '职业    Varchar2    80
        zlXML_InsertNode gobjXml.documentElement, "职业", Trim(NVL(rsTemp!职业))
        '民族    Varchar2    20
        zlXML_InsertNode gobjXml.documentElement, "民族", Trim(NVL(rsTemp!民族))
        '国籍    Varchar2    30
        zlXML_InsertNode gobjXml.documentElement, "国籍", Trim(NVL(rsTemp!国籍))
        '学历    Varchar2    10
        zlXML_InsertNode gobjXml.documentElement, "学历", Trim(NVL(rsTemp!学历))
        '婚姻状况    Varchar2    4
        zlXML_InsertNode gobjXml.documentElement, "婚姻状况", Trim(NVL(rsTemp!婚姻状况))
        '区域    Varchar2    30
        zlXML_InsertNode gobjXml.documentElement, "区域", Trim(NVL(rsTemp!区域))
        '家庭地址    Varchar2    50
        zlXML_InsertNode gobjXml.documentElement, "家庭地址", Trim(NVL(rsTemp!家庭地址))
        '家庭电话    Varchar2    20
        zlXML_InsertNode gobjXml.documentElement, "家庭电话", Trim(NVL(rsTemp!家庭电话))
        '家庭地址邮编    Varchar2    6
        zlXML_InsertNode gobjXml.documentElement, "家庭地址邮编", Trim(NVL(rsTemp!家庭地址邮编))
        '监护人  Varchar2    64
        zlXML_InsertNode gobjXml.documentElement, "监护人", Trim(NVL(rsTemp!监护人))
        '联系人姓名  Varchar2    64
        zlXML_InsertNode gobjXml.documentElement, "联系人姓名", Trim(NVL(rsTemp!联系人姓名))
        '联系人关系  Varchar2    30
        zlXML_InsertNode gobjXml.documentElement, "联系人关系", Trim(NVL(rsTemp!联系人关系))
        '联系人地址  Varchar2    50
        zlXML_InsertNode gobjXml.documentElement, "联系人地址", Trim(NVL(rsTemp!联系人地址))
        '联系人电话  Varchar2    20
        zlXML_InsertNode gobjXml.documentElement, "联系人电话", Trim(NVL(rsTemp!联系人电话))
        '工作单位    Varchar2    100
        zlXML_InsertNode gobjXml.documentElement, "工作单位", Trim(NVL(rsTemp!工作单位))
        '单位电话    Varchar2    20
        zlXML_InsertNode gobjXml.documentElement, "单位电话", Trim(NVL(rsTemp!单位电话))
        '单位邮编    Varchar2    6
        zlXML_InsertNode gobjXml.documentElement, "单位邮编", Trim(NVL(rsTemp!单位邮编))
        '单位开户行  Varchar2    50
        zlXML_InsertNode gobjXml.documentElement, "单位开户行", Trim(NVL(rsTemp!单位开户行))
        '单位帐号    Varchar2    20
        zlXML_InsertNode gobjXml.documentElement, "单位帐号", Trim(NVL(rsTemp!单位帐号))
        '费别    Varchar2    10
        zlXML_InsertNode gobjXml.documentElement, "费别", Trim(NVL(rsTemp!费别))
    End With
    OutPutString_Xml = zlXML_GetXMLString
    zlGetclsPatientInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
 
Public Function zlGetPatiPrepayAndFee(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人预交余额和消费余额
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-27 14:47:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, lng病人ID As Long
    Dim dbl预交 As Double, dbl费用 As Double, strValue As String
    On Error GoTo errHandle
    If InputString_XML = "" Then Exit Function
    If zlXML_Init = False Then Exit Function
    'InputString说明:
    'inputString
    '病人ID  Number  18
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    Call zlXML_GetNodeValue("病人ID", , strValue): lng病人ID = Val(strValue)
    If lng病人ID = 0 Then Exit Function
    
    If zlXML_Init = False Then Exit Function
    strSQL = "Select 预交余额, 费用余额 From 病人余额 Where  类型=1 And 性质 = 1 And 病人id = [1]  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人余额", lng病人ID)
    If Not rsTemp.EOF Then
        dbl预交 = Val(NVL(rsTemp!预交余额))
        dbl费用 = Val(NVL(rsTemp!费用余额))
    End If
    'OutPutString参数说明:
    'OutPutString
    '标识    数据类型    长度    精度    非空    说明
    '预交余额    number  16  5       门诊预交
    zlXML_InsertNode gobjXml.documentElement, "预交余额", Round(dbl预交, 5)
    '费用余额    number  16  5       门诊费用
    zlXML_InsertNode gobjXml.documentElement, "费用余额", Round(dbl费用, 5)
     OutPutString_Xml = zlXML_GetXMLString
    zlGetPatiPrepayAndFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPatiArchives(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建病人档案或修改病人档案
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-27 14:47:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, lng病人ID As Long
    Dim strData(0 To 30) As String, i As Long
    Dim strErrMsg As String, strErr As String, strValue As String
    Dim str门诊号 As String, str就诊卡号 As String, strIC卡号 As String
    On Error GoTo errHandle
    If InputString_XML = "" Then Exit Function
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    'InputString说明:
    'inputString
    '标识    数据类型    长度    精度    非空    说明
    '是否新增    Number  1       Y   1-增加;0-修改;
    i = 0: Call zlXML_GetNodeValue("是否新增", , strValue, strErr): strData(i) = Val(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '病人ID  Number  18          修改时传入病人ID,增加时为空 zlGetclsPatientInfo函数获取病人ID
    Call zlXML_GetNodeValue("病人ID", , strValue, strErr): strData(i) = Val(strValue): i = i + 1
    '姓名    Varchar2    100      Y
    Call zlXML_GetNodeValue("姓名", , strValue, strErr): strData(i) = strValue: i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '性别    Varchar2    4       Y
    Call zlXML_GetNodeValue("性别", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '年龄    Varchar2    10      Y
    Call zlXML_GetNodeValue("年龄", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '出生日期    Varchar2    20          yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("出生日期", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '出生地点    Varchar2    50
    Call zlXML_GetNodeValue("出生地点", , strValue): strData(i) = Trim(strValue): i = i + 1
    '身份证号    VARCHAR2    18          如果有相同号的,且姓名不一致时,将不能提交!
    Call zlXML_GetNodeValue("身份证号", , strValue): strData(i) = Trim(strValue): i = i + 1
    '其他证件    Varchar2    20
    Call zlXML_GetNodeValue("其他证件", , strValue): strData(i) = Trim(strValue): i = i + 1
    '职业    Varchar2    80
    Call zlXML_GetNodeValue("职业", , strValue): strData(i) = Trim(strValue): i = i + 1
    '民族    Varchar2    20
    Call zlXML_GetNodeValue("民族", , strValue): strData(i) = Trim(strValue): i = i + 1
    '国籍    Varchar2    30
    Call zlXML_GetNodeValue("国籍", , strValue): strData(i) = Trim(strValue): i = i + 1
    '学历    Varchar2    10
    Call zlXML_GetNodeValue("学历", , strValue): strData(i) = Trim(strValue): i = i + 1
    '婚姻状况    Varchar2    4
    Call zlXML_GetNodeValue("婚姻状况", , strValue): strData(i) = Trim(strValue): i = i + 1
    '区域    Varchar2    30
    Call zlXML_GetNodeValue("区域", , strValue): strData(i) = Trim(strValue): i = i + 1
    '家庭地址    Varchar2    50
    Call zlXML_GetNodeValue("家庭地址", , strValue): strData(i) = Trim(strValue): i = i + 1
    '家庭电话    Varchar2    20
    Call zlXML_GetNodeValue("家庭电话", , strValue): strData(i) = Trim(strValue): i = i + 1
    '家庭地址邮编    Varchar2    6
    Call zlXML_GetNodeValue("家庭地址邮编", , strValue): strData(i) = Trim(strValue): i = i + 1
    '监护人  Varchar2    64
    Call zlXML_GetNodeValue("监护人", , strValue): strData(i) = Trim(strValue): i = i + 1
    '联系人姓名  Varchar2    64
    Call zlXML_GetNodeValue("联系人姓名", , strValue): strData(i) = Trim(strValue): i = i + 1
    '联系人关系  Varchar2    30
    Call zlXML_GetNodeValue("联系人关系", , strValue): strData(i) = Trim(strValue): i = i + 1
    '联系人地址  Varchar2    50
    Call zlXML_GetNodeValue("联系人地址", , strValue): strData(i) = Trim(strValue): i = i + 1
    '联系人电话  Varchar2    20
    Call zlXML_GetNodeValue("联系人电话", , strValue): strData(i) = Trim(strValue): i = i + 1
    '工作单位    Varchar2    100
    Call zlXML_GetNodeValue("工作单位", , strValue): strData(i) = Trim(strValue): i = i + 1
    '单位电话    Varchar2    20
    Call zlXML_GetNodeValue("单位电话", , strValue): strData(i) = Trim(strValue): i = i + 1
    '单位邮编    Varchar2    6
    Call zlXML_GetNodeValue("单位邮编", , strValue): strData(i) = Trim(strValue): i = i + 1
    '单位开户行  Varchar2    50
    Call zlXML_GetNodeValue("单位开户行", , strValue): strData(i) = Trim(strValue): i = i + 1
    '单位帐号    Varchar2    20
    Call zlXML_GetNodeValue("单位帐号", , strValue): strData(i) = Trim(strValue): i = i + 1
    '门诊号  Number  18
    Call zlXML_GetNodeValue("门诊号", , strValue): strData(i) = Trim(strValue): i = i + 1
    '医疗付款方式    Varchar2    20      N   zlGetMedicalTreatment返回的名称 或渠道固定接口 , 修改时, 不能更改医疗付款方式
    Call zlXML_GetNodeValue("医疗付款方式", , strValue): strData(i) = Trim(strValue): i = i + 1
    '费别    Varchar2    10 zlGetFeeType返回的费别名称
    Call zlXML_GetNodeValue("费别", , strValue): strData(i) = Trim(strValue): i = i + 1
    
    If zlXML_Init = False Then Exit Function
    If strErrMsg <> "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    If Val(strData(0)) <> 1 Then
        '修改时,需检查传入的病人ID是否存在
        strSQL = "Select 费别,门诊号,医疗付款方式 From 病人信息 Where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查病人信息", Val(strData(1)))
        If rsTemp.EOF Then
            zlXML_InsertNode gobjXml.documentElement, "ERROR", "病人未找到,不能修改!"
            OutPutString_Xml = zlXML_GetXMLString
            Exit Function
        End If
        If NVL(rsTemp!费别) <> Trim(strData(30)) And Trim(NVL(rsTemp!费别)) <> "" Then
            zlXML_InsertNode gobjXml.documentElement, "ERROR", "不能调整病人的费别!"
            OutPutString_Xml = zlXML_GetXMLString
            Exit Function
        End If
        If NVL(rsTemp!医疗付款方式) <> Trim(strData(29)) And Trim(NVL(rsTemp!医疗付款方式)) <> "" Then
            zlXML_InsertNode gobjXml.documentElement, "ERROR", "不能调整病人的医疗付款方式!"
            OutPutString_Xml = zlXML_GetXMLString
            Exit Function
        End If
        
        lng病人ID = Val(strData(1))
        str门诊号 = strData(28)
        If Val(str门诊号) = 0 Then
            str门诊号 = zlFormatID(NVL(rsTemp!门诊号))
        End If
    Else
        lng病人ID = zlDatabase.GetNextNo(1)
    End If
    '建立病人档案
    If Val(str门诊号) = 0 Then str门诊号 = zlDatabase.GetNextNo(3)
    
    '      Zl_挂号病人病案_Insert
    strSQL = "Zl_挂号病人病案_Insert("
    '  处理类型_In     Number,
    '--处理类型：
    '--             1=新建病人信息及门诊病案(用于新挂号病人)
    '--             2=修改病人信息，新建门诊病案(用于无病案的病人)
    '--             3=修改病人信息，不处理门诊病案(用于有病案的病人,但可能修改了病案的门诊号)
    strSQL = strSQL & "" & IIf(Val(strData(0)) = 1, 1, 2) & ","
    '  病人id_In       病人信息.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  门诊号_In       病人信息.门诊号%Type,
    strSQL = strSQL & "" & zlFormatID(str门诊号) & ","
    '    就诊卡号_In     病人信息.就诊卡号%Type,
    strSQL = strSQL & "null,"
    '  卡验证码_In     病人信息.卡验证码%Type,
    strSQL = strSQL & "null,"
    '    姓名_In         病人信息.姓名%Type,
    strSQL = strSQL & "'" & Trim(strData(2)) & "',"
    '    性别_In         病人信息.性别%Type,
    strSQL = strSQL & "'" & Trim(strData(3)) & "',"
    '    年龄_In         病人信息.年龄%Type,
    strSQL = strSQL & "'" & Trim(strData(4)) & "',"
    '    费别_In         病人信息.费别%Type,
    strSQL = strSQL & "'" & Trim(strData(30)) & "',"
    '    医疗付款方式_In 病人信息.医疗付款方式%Type,
    strSQL = strSQL & "'" & Trim(strData(29)) & "',"
    '    国籍_In         病人信息.国籍%Type,
    strSQL = strSQL & "'" & Trim(strData(11)) & "',"
    '    民族_In         病人信息.民族%Type,
    strSQL = strSQL & "'" & Trim(strData(10)) & "',"
    '    婚姻_In         病人信息.婚姻状况%Type,
    strSQL = strSQL & "'" & Trim(strData(13)) & "',"
    '    职业_In         病人信息.职业%Type,
    strSQL = strSQL & "'" & Trim(strData(9)) & "',"
    '    身份证号_In     病人信息.身份证号%Type,
    strSQL = strSQL & "'" & Trim(strData(7)) & "',"
    '    工作单位_In     病人信息.工作单位%Type,
    strSQL = strSQL & "'" & Trim(strData(23)) & "',"
    '  合同单位id_In   病人信息.合同单位id%Type,
    strSQL = strSQL & "NULL,"
    '    单位电话_In     病人信息.单位电话%Type,
    strSQL = strSQL & "'" & Trim(strData(24)) & "',"
    '    单位邮编_In     病人信息.单位邮编%Type,
    strSQL = strSQL & "'" & Trim(strData(25)) & "',"
    '    家庭地址_In     病人信息.家庭地址%Type,
    strSQL = strSQL & "'" & Trim(strData(15)) & "',"
    '    家庭电话_In     病人信息.家庭电话%Type,
    strSQL = strSQL & "'" & Trim(strData(16)) & "',"
    '    家庭地址邮编_In     病人信息.家庭地址邮编%Type,
    strSQL = strSQL & "'" & Trim(strData(17)) & "',"
    '    登记时间_In     病人信息.登记时间%Type,
    strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),null,null,"
     '    出生日期_In     病人信息.出生日期%Type := Null,
    strSQL = strSQL & "to_date('" & Trim(strData(5)) & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure strSQL, "病人信息档案创建"
    
    strSQL = "Select 病人ID,门诊号 From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查病人信息", lng病人ID)
    If rsTemp.EOF Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "病人未找到,不能修改!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    '    OutPutString
    '    标识    数据类型    长度    精度    非空    说明
    '    病人ID  Number  18      Y
        zlXML_InsertNode gobjXml.documentElement, "病人ID", Val(NVL(rsTemp!病人ID))
    '    门诊号  Number  18
        zlXML_InsertNode gobjXml.documentElement, "门诊号", NVL(rsTemp!门诊号)
    OutPutString_Xml = zlXML_GetXMLString
    zlPatiArchives = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 Public Function zlGetRegisterPlanTime(ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号安排的时间段
    '入参:
    '出参:OutPutString_Xml-返回医院的费别信息,格式如下
    '返回:函数返回 True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 17:55:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objNode As MSXML2.IXMLDOMElement
    On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    strSQL = "Select 时间段,to_char(开始时间,'hh24:mi:ss') as 开始时间,to_char(终止时间,'hh24:mi:ss') as 终止时间 From 时间段"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取时间段")
    With rsTemp
        Do While Not .EOF
            zlXML_InsertNode gobjXml.documentElement, "LISTDATA", "", objNode
            'OutPutString参数说明:
            ''OutPutString
            ''标识    数据类型    长度    精度    非空    说明
            ''时间段  Varchar2    4       Y   白天;上午;…;
            zlXML_InsertNode objNode, "时间段", NVL(rsTemp!时间段)
            ''开始时间    Varchar2    8       Y   格式为:HH:MM:SS
            ''如00:00:00或08:30:00
            zlXML_InsertNode objNode, "开始时间", NVL(rsTemp!开始时间)
            ''结束时间    Varchar2    8       Y   格式为:HH:MM:SS 如果结束时间小于张终止时间,这个结束时间代表第二天的时间:比如后夜:18:00:00- 07:59:59
            zlXML_InsertNode objNode, "结束时间", NVL(rsTemp!终止时间)
            .MoveNext
        Loop
    End With
    OutPutString_Xml = zlXML_GetXMLString
    zlGetRegisterPlanTime = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlGetRegisterNOType(ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号安排的号类
    '入参:
    '出参:OutPutString_Xml-返回医院的号类信息,格式如下
    '返回:函数返回 True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 17:55:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objNode As MSXML2.IXMLDOMElement
    On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    
    strSQL = "Select 编码,名称, 简码,缺省标志,说明 From 号类"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "号类")
    With rsTemp
        Do While Not .EOF
            zlXML_InsertNode gobjXml.documentElement, "LISTDATA", "", objNode
            '       编码    Varchar2    2       Y
            zlXML_InsertNode objNode, "编码", NVL(rsTemp!编码)
            '       名称    Varchar2    10      Y
            zlXML_InsertNode objNode, "名称", NVL(rsTemp!名称)
            zlXML_InsertNode objNode, "简码", NVL(rsTemp!简码)
            '       缺省标志    Number  18          1-表示缺省使用的类别
            zlXML_InsertNode objNode, "缺省标志", Val(NVL(rsTemp!缺省标志))
            '       说明    varchar2  50
            zlXML_InsertNode objNode, "说明", Val(NVL(rsTemp!说明))
            .MoveNext
        Loop
    End With
    OutPutString_Xml = zlXML_GetXMLString
    zlGetRegisterNOType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegisterPlan(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号安排数据
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-31 10:24:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, lng病人ID As Long
    Dim strData(0 To 10) As String, i As Long, strIF As String
    Dim strErrMsg As String, strErr As String, strValue As String, strTime As String
    Dim rs序号 As ADODB.Recordset, rs挂号 As ADODB.Recordset, rs价格 As ADODB.Recordset
    Dim rs计划停用 As ADODB.Recordset, rs安排计划 As ADODB.Recordset
    Dim dRegisterDate As Date  '挂号日期
    Dim dDate As Date
    On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    i = 0
    'InputString说明:
    'inputString
    '    标识    数据类型    长度    精度    非空    说明
    '    日期    Varchar2    20          需要预约挂号的日期,格式为yyyy-mm-dd hh24:mi:ss:指定了日期表示需要确定当前的有效挂号安排.为空为:不限制日期,返回所有安排.
    Call zlXML_GetNodeValue("日期", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '    号类    Varchar2    10          普通、专家、专科、特诊…:空为不显制号类;
    Call zlXML_GetNodeValue("号类", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '    号码    Varchar2    5           指定号码,为空时,表示不限定号码
    Call zlXML_GetNodeValue("号码", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '    医生    Varchar2    64          指定医生的号别
    Call zlXML_GetNodeValue("医生", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '   不允许挂免费号 integer 1           1-不允许挂免费号;0-允许挂免费号
    Call zlXML_GetNodeValue("不允许挂免费号", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '合作单位    Varchar2    50      Y   可以渠道强制转入;也可以三方机构固定传入;内容为挂号合作单位.名称
    Call zlXML_GetNodeValue("合作单位", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If zlXML_Init = False Then Exit Function
    dRegisterDate = CDate("1992-01-01")
    If Trim(strData(0)) <> "" Then
        If IsDate(strData(0)) = False Then
            zlXML_InsertNode gobjXml.documentElement, "ERROR", "日期格式不对或不是日期型!"
            OutPutString_Xml = zlXML_GetXMLString
            Exit Function
        End If
        dRegisterDate = CDate(strData(0))
    End If
    
    If Trim(strData(1)) <> "" Then
        strIF = strIF & " And P.号类=[1]"
    End If
    If Trim(strData(2)) <> "" Then
           strIF = strIF & " And P.号码=[2]"
    End If
    If Trim(strData(3)) <> "" Then
        strIF = strIF & "   And (zlSpellCode(P.医生姓名) Like [3] or B.简码 Like [3])"
    End If
    If dRegisterDate > CDate("1992-01-01") Then
         strTime = "With 安排时间段 As   ( " & _
         "   Select 时间段 " & _
         "   From ( " & _
         "               Select 时间段, " & _
         "                      To_date(Decode(Sign(开始时间 - 终止时间), 1, '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),'3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')),'yyyy-mm-dd hh24:mi:ss') As 开始时间, " & _
         "                      To_date(Decode(Sign(开始时间 - 终止时间), 1, '3000-01-11 ' || To_Char(终止时间, 'HH24:MI:SS'),'3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')),'yyyy-mm-dd hh24:mi:ss') As 终止时间, " & _
         "                      to_date('3000-01-10 ' || To_Char([5], 'HH24:MI:SS'),'yyyy-mm-dd hh24:mi:ss') As 当前时间, " & _
         "                      to_date('3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS'),'yyyy-mm-dd hh24:mi:ss')  As 开始时间1, " & _
         "                      to_date('3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS'),'yyyy-mm-dd hh24:mi:ss')  As 终止时间1 " & _
         "               From 时间段  " & _
         "               )  " & _
         "   Where     当前时间 Between  开始时间 And  终止时间1 Or 当前时间 Between 开始时间1 And   终止时间  )"
           strSQL = _
            " Select P.ID,P.号码 ,P.号类,P.科室ID,P.项目ID,p.停用日期,p.终止时间,p.开始时间,p.序号," & _
            "       P.医生ID,P.医生姓名,P.限号数,P.限约数,Nvl(P.病案必须,0) as 病案必须," & _
            "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,nvl(P.分诊方式,0) as 分诊方式,nvl(P.序号控制,0) as 序号控制," & _
            "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL) as 排班 " & _
            " From 挂号安排 P" & _
            " Where P.停用日期 is NULL and  [5] Between Nvl(P.开始时间,To_date('1900-01-01','YYYY-MM-DD')) And Nvl(P.终止时间,To_Date('3000-01-01','YYYY-MM-DD'))" & _
            "          And not Exists(Select 1 From 挂号安排计划 where 安排ID=P.id And ([5] BETWEEN 生效时间  and 失效时间)  And 审核时间 is not NULL  ) " & _
            "          And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=P.ID and [5] between 开始停止时间 and 结束停止时间 )" & _
            " Union ALL " & _
            " Select   C.ID,C.号码,C.号类,C.科室ID,C.项目ID,p.失效时间 as 停用日期,p.失效时间 as 终止时间,p.生效时间 as 开始时间,c.序号," & _
            "       C.医生ID,C.医生姓名,P.限号数,P.限约数,Nvl(C.病案必须,0) as 病案必须," & _
            "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,nvl(P.分诊方式,0) as 分诊方式,nvl(P.序号控制,0) as 序号控制," & _
            "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL) as 排班 " & _
            " From 挂号安排计划 P, 挂号安排 C" & _
            " Where P.安排ID=C.ID And ([5] BETWEEN P.生效时间  and P.失效时间)  And P.审核时间 is not NULL " & _
            "           And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=C.ID and [5] between 开始停止时间 and 结束停止时间 )"
         
         strSQL = strTime & vbCrLf & _
            "Select Distinct " & _
            "       P.ID,P.号类,P.号码 as 号别,P.科室ID,B.编码 as 科室编码,B.名称 As 科室名称, " & _
            "       P.项目ID,C.编码 as 项目编码,C.名称 As 项目名称," & _
            "       P.医生ID,D.编号 as 医生编号,P.医生姓名," & _
            "       P.序号,P.限号数,P.限约数," & _
            "       P.周日 as 日,P.周一 as 一,P.周二 as 二,P.周三 as 三,P.周四 as 四,P.周五 as 五,P.周六 as 六," & _
            "       P.开始时间,P.终止时间, P.停用日期,p.病案必须,P.序号控制,p.分诊方式,decode(M.开始序号,NULL,NULL,M.结束序号,NULL,NULL,M.开始序号||'-'||M.结束序号)  as 序号范围" & _
            " From (" & strSQL & ") P,部门表 B,收费项目目录 C,人员表 D,合作单位安排控制 M" & _
            " Where   P.ID=M.安排ID And M.合作单位=[6] " & _
            "           And P.科室ID=B.ID And P.医生ID=d.ID(+)" & vbCrLf & _
            "           And P.科室ID=B.ID And P.项目ID=C.ID" & strIF & _
            "           And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            "           And (Nvl(P.医生ID,0)=0 Or Exists(Select 1 From 人员表 Q Where P.医生ID=Q.ID And (Q.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or Q.撤档时间 Is Null)))" & _
            "           And exists(Select 1 From 安排时间段 where 时间段=P.排班 ) " & _
            IIf(Val(strData(4)) = 1, "And Not Exists ( Select 1 From 收费价目 B,挂号安排 A  Where A.项目ID=B.收费细目ID And A.项目ID=C.ID Group by A.项目ID Having Nvl(Sum(B.现价),0)=0)", "") & _
            "  "
            'Val(strData(4)):不允许挂免费号
            
            strSQL = "Select * From (" & strSQL & ") Order by 号别"
    Else
        '所有挂号安排
          strSQL = _
            "       Select P.ID,P.号类,P.号码 as 号别,P.科室ID,B.编码 as 科室编码,B.名称 As 科室名称, " & _
            "       P.项目ID,C.编码 as 项目编码,C.名称 As 项目名称," & _
            "       P.医生ID,D.编号 as 医生编号,P.医生姓名," & _
            "       P.序号,P.限号数,P.限约数," & _
            "       P.周日 as 日,P.周一 as 一,P.周二 as 二,P.周三 as 三,P.周四 as 四,P.周五 as 五,P.周六 as 六," & _
            "       P.开始时间,P.终止时间, P.停用日期,nvl(p.病案必须,0) as 病案必须,nvl(P.序号控制,0) as 序号控制,nvl(p.分诊方式,0) as 分诊方式,decode(M.开始序号,NULL,NULL,M.结束序号,NULL,NULL,M.开始序号||'-'||M.结束序号)  as 序号范围" & _
            " From 挂号安排 P,部门表 B,收费项目目录 C,人员表 D,合作单位安排控制 M" & _
            " Where P.ID=M.安排ID And M.合作单位=[6] " & _
            "           And P.科室ID=B.ID And P.医生ID=D.ID(+)" & vbCrLf & _
            "           And P.科室ID=B.ID And P.项目ID=C.ID" & strIF & _
            "           And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            "           And (Nvl(P.医生ID,0)=0 Or Exists(Select 1 From 人员表 Q Where P.医生ID=Q.ID And (Q.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or Q.撤档时间 Is Null)))" & _
            IIf(Val(strData(4)) = 1, "And Not Exists ( Select 1 From 收费价目 B,挂号安排 A  Where A.项目ID=B.收费细目ID And A.项目ID=C.ID Group by A.项目ID Having Nvl(Sum(B.现价),0)=0)", "") & _
            " Order by  号码"
    End If
    '安排
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取挂号安排", _
       CStr(Trim(strData(1))), CStr(Trim(strData(2))), CStr(Trim(strData(3))), 0, dRegisterDate, CStr(Trim(strData(5))))
    If rsTemp.EOF Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "未找到数据!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    Dim cllIds As Collection
    Call zlGetRecodersFieldIns(rsTemp, "项目ID,ID", cllIds)
    Set rs价格 = zlItemsMoney(CStr(cllIds("项目ID")))
    
    Dim strSubTable As String, varData() As Variant
    Call zlGetSubTable(0, CStr(cllIds("ID")), strSubTable, varData)
    strSubTable = "With 安排 as (" & strSubTable & ") "
    strSQL = "" & strSubTable & _
        "   Select  /*+ rule */  A.安排ID,A.序号,A.开始停止时间,A.结束停止时间,A.制订人,A.制订日期,A.备注 " & _
        "   From 挂号安排停用状态 A,安排 B " & _
        "   Where A.安排ID=B.ID   "
    strSQL = " Select /*+ RULE */ * From (" & strSQL & ") " & _
        "   Order by 安排ID,序号"
    Set rs计划停用 = zlDatabase.OpenSQLRecordByArray(strSQL, "获取安排", varData)
    
    strSQL = "" & strSubTable & _
            "   Select   /*+ rule */  P.ID,P.安排ID, P.号码,P.生效时间," & _
            "       C.序号,P.限号数,P.限约数," & _
            "       P.周日 as 日,P.周一 as 一,P.周二 as 二,P.周三 as 三,P.周四 as 四,P.周五 as 五,P.周六 as 六," & _
            "       P.分诊方式,P.序号控制,P.安排人,P.安排时间,P.审核人,P.审核时间,P.实际生效,P.失效时间" & _
            " From 挂号安排计划 P,安排 B,挂号安排 C" & _
            " Where P.安排ID=B.ID and  p.安排ID=c.id   "
      strSQL = " Select /*+ RULE */ * From (" & strSQL & ") " & _
        "   Order by 安排时间"
     Set rs安排计划 = zlDatabase.OpenSQLRecordByArray(strSQL, "获取安排", varData)
     
    '序号使用情况
    If dRegisterDate > CDate("1992-01-01") Then
        strSQL = "" & _
            "   Select  A.号码,A.日期,A.序号,A.状态,A.操作员姓名  " & _
            "   From 挂号序号状态 A,挂号安排 P " & _
            "   Where A.日期=[1] and A.号码=P.号码" & strIF
    Else
        strSQL = "" & _
            "   Select  A.号码,A.日期,A.序号,A.状态,A.操作员姓名  " & _
            "   From 挂号序号状态 A,挂号安排 P " & _
            "   Where 日期>=trunc(sysdate) and A.号码=P.号码" & strIF
    End If
    dDate = Format(dRegisterDate, "yyyy-mm-dd")
    Set rs序号 = zlDatabase.OpenSQLRecord(strSQL, "获取序号状态", dDate)
    
    '挂号情况
    strSQL = "" & _
    "   Select P.ID as 安排ID,B.已挂数,B.已约数,B.其中已接收,B.日期" & _
    "   From 挂号安排 P,病人挂号汇总 B" & _
    "   Where P.科室ID = B.科室ID And P.项目ID = B.项目ID" & _
    "               And Nvl(P.医生ID,0)=Nvl(B.医生ID,0) " & _
    "               And Nvl(P.医生姓名,'医生')=Nvl(B.医生姓名,'医生') " & _
    "               And (P.号码=B.号码 or B.号码 is Null )     " & strIF
    
    If dRegisterDate > CDate("1992-01-01") Then
         strSQL = strSQL & " And B.日期=[1]"
    Else
         strSQL = strSQL & " And B.日期>=trunc(sysdate)"
    End If
    
    Set rs挂号 = zlDatabase.OpenSQLRecord(strSQL, "获取挂号汇总状态", dDate)
    Dim objNode As MSXML2.IXMLDOMElement
    Dim objTempNode As MSXML2.IXMLDOMElement
    
    With rsTemp
        Do While Not .EOF
            zlXML_InsertNode gobjXml.documentElement, "LISTDATA", "", objNode
            '    标识    数据类型    长度    精度    非空    说明
            '    ID  Number  18      Y   PK:安排ID
            zlXML_InsertNode objNode, "ID", NVL(rsTemp!id)
            '    号类    Varchar2    10      Y   普通、专家、专科、特诊…
            zlXML_InsertNode objNode, "号类", NVL(rsTemp!号类)
            '    号码    Varchar2    5       Y   UQ:号码
            zlXML_InsertNode objNode, "号码", NVL(rsTemp!号别)
            '    科室ID  Number  18      Y   部门唯一ID号,三方机构可以用此来进行对码.
            zlXML_InsertNode objNode, "科室ID", NVL(rsTemp!科室ID)
            '    科室编码    Varchar2    10      Y
            zlXML_InsertNode objNode, "科室编码", NVL(rsTemp!科室编码)
            '    科室名称    Varchar2    20      Y
            zlXML_InsertNode objNode, "科室名称", NVL(rsTemp!科室名称)
            '    项目ID  Number  18      Y   挂号项目唯一ID号,三方机构可以用此来进行对码.
            zlXML_InsertNode objNode, "项目ID", NVL(rsTemp!项目ID)
            '    项目编码    Varchar2    20      Y
            zlXML_InsertNode objNode, "项目编码", NVL(rsTemp!项目编码)
            '    项目名称    Varchar2    80      Y
            zlXML_InsertNode objNode, "项目名称", NVL(rsTemp!项目名称)
            '    医生ID  Number  18          医生唯一标识,可能存在外来医生坐诊情况,所以可能为空
            zlXML_InsertNode objNode, "医生ID", NVL(rsTemp!医生ID)
            '    医生编号    Varchar2    6       Y
            zlXML_InsertNode objNode, "医生编号", NVL(rsTemp!医生编号)
            '    医生姓名    Varchar2    20      Y
            zlXML_InsertNode objNode, "医生姓名", NVL(rsTemp!医生姓名)
            '    序号    Number  18          UQ: 科室ID, 项目ID, 医生姓名, 医生ID, 序号
            '    表示同科同项目,同一个医生在不同时间有两个号别;比如:周一到周五医生只能看32个病人,周六只看10个病人,此序号就是控制同一医生不同号别的一个序号
            zlXML_InsertNode objNode, "序号", NVL(rsTemp!序号)
            '    限号数  number  5
            zlXML_InsertNode objNode, "限号数", NVL(rsTemp!限号数)
            '    限约数  Number  5
            zlXML_InsertNode objNode, "限约数", NVL(rsTemp!限约数)
            '    周日    Varchar2    4           白天;上午;…;
            '    zlGetRegisterPlanTime函数获取的时间段对应
            zlXML_InsertNode objNode, "周日", NVL(rsTemp!日)
            '    周一    Varchar2    4           白天;上午;…
            zlXML_InsertNode objNode, "周一", NVL(rsTemp!一)
            '    周二    Varchar2    4           白天;上午;…
            zlXML_InsertNode objNode, "周二", NVL(rsTemp!二)
            '    周三    Varchar2    4           白天;上午;…
            zlXML_InsertNode objNode, "周三", NVL(rsTemp!三)
            '    周四    Varchar2    4           白天;上午;…
            zlXML_InsertNode objNode, "周四", NVL(rsTemp!四)
            '    周五    Varchar2    4           白天;上午;…
            zlXML_InsertNode objNode, "周五", NVL(rsTemp!五)
            '    周六    Varchar2    4           白天;上午;…
            zlXML_InsertNode objNode, "周六", NVL(rsTemp!六)
            '    开始时间    Varchar2    20          安排的开始时间,yyyy-mm-dd hh24:mi:ss
            zlXML_InsertNode objNode, "开始时间", Format(rsTemp!开始时间, "yyyy-mm-dd HH:MM:SS")
            '    终止时间    Varchar2    20          安排的终止时间yyyy-mm-dd hh24:mi:ss
            zlXML_InsertNode objNode, "终止时间", Format(rsTemp!终止时间, "yyyy-mm-dd HH:MM:SS")
            '    停用日期    Varchar2    20          停止使用安排的时间.
            zlXML_InsertNode objNode, "停用日期", Format(rsTemp!停用日期, "yyyy-mm-dd HH:MM:SS")
            '    序号控制    number  1           1-严格按序号控制;0-不严格按序号控制
            zlXML_InsertNode objNode, "序号控制", Val(NVL(rsTemp!序号控制))
            '    序号范围    Varchar2    20          比如:1-30,为空,表示不限制范围
            zlXML_InsertNode objNode, "序号范围", Val(NVL(rsTemp!序号范围))
            zlXML_InsertNode objNode, "病案必须", rsTemp!病案必须
            zlXML_InsertNode objNode, "分诊方式", rsTemp!分诊方式
            
            '    序号使用情况                    表示已经使用的序号.
            zlXML_InsertNode objNode, "序号使用情况", "", objTempNode
            rs序号.Filter = "号码='" & NVL(rsTemp!号别) & "'"
            If rs序号.RecordCount <> 0 Then rs序号.MoveFirst
            With rs序号
                Do While Not rs序号.EOF
                    '        日期    Varchar2    10          格式:yyyy-mm-dd            '    日期范围为:>=trunc(sysdate)
                    zlXML_InsertNode objTempNode, "日期", Format(!日期, "yyyy-mm-dd")
                    '        序号    number  5           已经使用的序号
                    zlXML_InsertNode objTempNode, "序号", Val(NVL(!序号))
                    '        状态    number  2           1-已挂,2-预约,3-预留
                    zlXML_InsertNode objTempNode, "状态", Val(NVL(!状态))
                    rs序号.MoveNext
                Loop
            End With
            '    挂号情况
            zlXML_InsertNode objNode, "挂号情况", "", objTempNode
            rs挂号.Filter = "安排ID=" & Val(NVL(rsTemp!id))
            If rs挂号.RecordCount <> 0 Then rs挂号.MoveFirst
            With rs挂号
                Do While Not .EOF
                '        日期    Varchar2    10          格式:yyyy-mm-dd日期范围为:>=trunc(sysdate)
                zlXML_InsertNode objTempNode, "日期", Format(!日期, "yyyy-mm-dd")
                '        已挂数  Number  5
                zlXML_InsertNode objTempNode, "已挂数", Val(NVL(!已挂数))
                '        已约数  Number  5
                zlXML_InsertNode objTempNode, "已挂数", Val(NVL(!已约数))
                '        其中已接收  Number  5           表示预约已经接收的病人数
                zlXML_InsertNode objTempNode, "其中已接收", Val(NVL(rs挂号!其中已接收))
                    .MoveNext
                Loop
            End With
            '    价格情况
            rs价格.Filter = "主项ID=" & Val(NVL(rsTemp!项目ID))
             With rs价格
                If rs价格.RecordCount <> 0 Then rs价格.MoveFirst
                Do While Not .EOF
                zlXML_InsertNode objNode, "价格情况", "", objTempNode
                '        项目ID  Number  18      Y
                zlXML_InsertNode objTempNode, "项目ID", IIf(Val(NVL(!从项目ID)) = 0, Val(NVL(!主项ID)), Val(NVL(!从项目ID)))
                '        项目编码    Varchar2    20      Y
                zlXML_InsertNode objTempNode, "项目编码", NVL(!项目编码)
                '        项目名称    Varchar2    80      Y
                zlXML_InsertNode objTempNode, "项目名称", NVL(!项目名称)
                '        单位    Varchar2    20      Y
                zlXML_InsertNode objTempNode, "单位", NVL(!计算单位)
                '        数次    Number  5       Y
                zlXML_InsertNode objTempNode, "数次", NVL(!数次)
                '        收入项目ID  Number  18      Y
                zlXML_InsertNode objTempNode, "收入项目ID", NVL(!收入项目ID)
                '        收入项目编码    Varchar2    20      Y
                zlXML_InsertNode objTempNode, "收入编码", NVL(!收入编码)
                '        收入项目名称    Varchar2    80      Y
                zlXML_InsertNode objTempNode, "收入项目名称", NVL(!收入项目)
                '        单价    number  16  5   Y   当前项目价格,挂号时,需要汇总.数次*单价
                zlXML_InsertNode objTempNode, "单价", NVL(!单价)
                    .MoveNext
                Loop
            End With
            '    计划停用情况
            zlXML_InsertNode objNode, "计划停用情况", "", objTempNode
            rs计划停用.Filter = "安排ID=" & Val(NVL(rsTemp!id))
            With rs计划停用
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    '        序号    number  18      Y   只作排列
                    zlXML_InsertNode objTempNode, "序号", NVL(!序号)
                    '        开始时间    Varchar2    20      Y   开始时间,yyyy-mm-dd hh24:mi:ss
                    zlXML_InsertNode objTempNode, "开始时间", Format(!开始停止时间, "yyyy-mm-dd HH:MM:SS")
                    '        结束时间    Varchar2    20      Y   结束时间,yyyy-mm-dd hh24:mi:ss
                    zlXML_InsertNode objTempNode, "结束时间", Format(!结束停止时间, "yyyy-mm-dd HH:MM:SS")
                    .MoveNext
                Loop
            End With
            '    计划安排情况
            zlXML_InsertNode objNode, "计划安排情况", "", objTempNode
            rs安排计划.Filter = "安排ID=" & Val(NVL(rsTemp!id))
            With rs安排计划
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                '        生效时间    Varchar2    20          开始时间,yyyy-mm-dd hh24:mi:ss生效范围为:>= sysdate
                    zlXML_InsertNode objTempNode, "生效时间", Format(!生效时间, "yyyy-mm-dd HH:MM:SS")
                '        限号数  number  5
                    zlXML_InsertNode objTempNode, "限号数", NVL(!限号数)
                '        限约数  Number  5
                    zlXML_InsertNode objTempNode, "限约数", NVL(!限约数)
                '        周日    Varchar2    4           白天;上午;…;
                    zlXML_InsertNode objTempNode, "周日", NVL(!日)
                '        周一    v   4           白天;上午;…
                    zlXML_InsertNode objTempNode, "周一", NVL(!一)
                '        周二    Varchar2    4           白天;上午;…
                    zlXML_InsertNode objTempNode, "周二", NVL(!二)
                '        周三    Varchar2    4           白天;上午;…
                    zlXML_InsertNode objTempNode, "周三", NVL(!三)
                '        周四    Varchar2    4           白天;上午;…
                    zlXML_InsertNode objTempNode, "周四", NVL(!四)
                '        周五    Varchar2    4           白天;上午;…
                    zlXML_InsertNode objTempNode, "周五", NVL(!五)
                '        周六    Varchar2    4           白天;上午;…
                    zlXML_InsertNode objTempNode, "周六", NVL(!六)
                '        序号控制    Boolean 1
                    zlXML_InsertNode objTempNode, "序号控制", Val(NVL(!序号控制))
                    .MoveNext
                Loop
            End With
            .MoveNext
        Loop
    End With
    OutPutString_Xml = zlXML_GetXMLString
    zlGetRegisterPlan = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlRegisterActualMoney(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定按排的实收金额
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-31 10:24:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strData(0 To 10) As String, i As Long
    Dim strErrMsg As String, strErr As String, strValue As String
    Dim str费别 As String, lng项目id As Long, lng病人ID As Long
    
    On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    i = 0
    'InputString说明:
    'inputString
    '标识    数据类型    长度    精度    非空    说明
    '号码    Varchar2    5       Y   当前挂号的号码
    Call zlXML_GetNodeValue("号码", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '费别    Varchar2    10      Y   zlGetFeeType返回的费别名称也可以为zlGetclsPatientInfo函数返回的病人信息中的费别
    Call zlXML_GetNodeValue("费别", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '病人ID  Long    18          费别或病人ID必须选一个条件转入.
    Call zlXML_GetNodeValue("病人ID", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If Trim(strData(0)) = "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "未传入挂号的号码!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    If zlXML_Init = False Then Exit Function
    
    str费别 = Trim(strData(1)): lng病人ID = Val(strData(2))
    If str费别 = "" And lng病人ID = 0 Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "费别或病人ID必须转其一!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    strSQL = "Select 项目ID From 挂号安排 where 号码=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取挂号项目ID", strData(0))
    If rsTemp.EOF Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "号码" & strData(0) & "未进行挂号安排,不能继续!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
     lng项目id = Val(NVL(rsTemp!项目ID))
    '获取病人ID的费别
    If lng病人ID <> 0 Then
        strSQL = "Select nvl(费别,'') as 费别 From 病人信息 Where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng病人ID)
        If rsTemp.EOF Then
             zlXML_InsertNode gobjXml.documentElement, "ERROR", "未找到指定的病人信息!"
            OutPutString_Xml = zlXML_GetXMLString
            Exit Function
        End If
        str费别 = NVL(rsTemp!费别)
    End If
    Set rsTemp = zlItemsMoney(lng项目id)
    If rsTemp.EOF Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "号码" & strData(0) & "所对应的收费项目已经被停用,不能继续!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    Dim dbl金额 As Double, dbl应收 As Double
    With rsTemp
        Do While Not .EOF
            lng项目id = Val(NVL(rsTemp!从项目ID))
            lng项目id = IIf(lng项目id = 0, Val(NVL(rsTemp!主项ID)), lng项目id)
            dbl应收 = Round(Val(NVL(!数次)) * Val(NVL(!单价)), 5)
            If NVL(!屏蔽费别, 0) = 1 Then
                dbl金额 = dbl金额 + dbl应收
            Else
                dbl金额 = Format(zlGetActualMoney(str费别, rsTemp!收入项目ID, dbl应收, lng项目id), "0.00")
            End If
            .MoveNext
        Loop
    End With
    'OutPutString参数说明:
    'OutPutString
    '标识    数据类型    长度    精度    说明
    '实收金额    Number  16  5   实收金额
    zlXML_InsertNode gobjXml.documentElement, "实收金额", dbl金额
    OutPutString_Xml = zlXML_GetXMLString
    zlRegisterActualMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlRegister(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人挂号(写入挂号数据)
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-31 10:24:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strData(0 To 14) As String, i As Long
    Dim strErrMsg As String, strErr As String, strValue As String
    Dim str费别 As String, lng项目id As Long, lng病人ID As Long
    Dim strNo As String, str登记时间 As String
    Dim cllPro As Collection, lng预交ID As Long
    Set cllPro = New Collection
    On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    i = 0
    'InputString说明:
    'inputString
    '标识    数据类型    长度    精度    非空    说明
    '0 号码    Varchar2    5       Y   当前挂号的号码
    Call zlXML_GetNodeValue("号码", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '1 日期    Varchar2    20      Y   预约的时间
    '格式yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("日期", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '2 挂号标志    Nubmer  1       Y   1-  表示挂号2-表示预约挂号,不扣款,3-表示预约挂号,扣款
    Call zlXML_GetNodeValue("挂号标志", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '3 预约方式    Varchar2    10      Y   电话;现场;网上等…
    Call zlXML_GetNodeValue("预约方式", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '4 收款总额    Double  16  5   Y   表示本次收款总额(如果涉及打折的,应该传打折的总金额),如果该金额与接口检查的本次收款总额不一致,将不能进行挂号.
    Call zlXML_GetNodeValue("收款总额", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '5 结算方式    Varchar2    20          可以某个接口传入固定的结算方式.与HIS的结算方式要同步.以便以后区分.
    '此结算方式为渠道在自己编制的部件中确定(如果在医疗卡类别中有设置,直接读取医疗卡类别中的结算方式).三方机构可以不传;
    Call zlXML_GetNodeValue("结算方式", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    ' 6 摘要    Varchar2    1000            备注
    Call zlXML_GetNodeValue("摘要", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '7 病人ID  Number  18      Y
    Call zlXML_GetNodeValue("病人ID", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '8. 合作单位    Varchar2    50      Y   挂号合作单位.名称
    Call zlXML_GetNodeValue("合作单位", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '9. 号序    Number  5           严格控制时,需要传入号序值
    Call zlXML_GetNodeValue("号序", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '10. 交易流水号  Varchar2    50          扣款的交易信息
    Call zlXML_GetNodeValue("交易流水号", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '11. 交易说明    Varchar2    500
    Call zlXML_GetNodeValue("交易说明", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    
    '12. 保险结算    Varchar2     4000 结算方式|结算金额||..... '55283
    Call zlXML_GetNodeValue("保险结算", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '13. 冲预交_In    number        '冲预存款
    Call zlXML_GetNodeValue("冲预交", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '14. 支付卡号_In    varchar2(50)
    Call zlXML_GetNodeValue("支付卡号", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    
    '14交易详细内容
    Dim lngRow As Long, lngCount As Long
    If Val(strData(1)) <> 2 Then
        '不扣款部分,不传入
        lng预交ID = zlDatabase.GetNextId("病人预交记录")
        If zlXML_GetRows("交易项目", lngCount) Then
            For lngRow = 0 To lngCount - 1
                Call zlXML_GetNodeValue("交易项目", lngRow, strValue, strErr)
                If strErr <> "" Then
                    strErrMsg = strErrMsg & strErr & vbCrLf
                    Exit For
                End If
                If Trim(strValue) = "" Then
                    strErrMsg = strErrMsg & "交易项目不能为空,请检查!"
                    Exit For
                End If
                '    Zl_三方结算交易_Insert
                strSQL = "Zl_三方结算交易_Insert("
                '      交易id_In   In 三方结算交易.交易id%Type,
                strSQL = strSQL & "" & lng预交ID & ","
                '      交易项目_In In 三方结算交易.交易项目%Type,
                '        交易项目 Varchar2    50          如: 操作类型,交易类型,卡类型,返回码,返回信息,收银机编号,柜员号,金额,结算批次,商户号,商户名称,终端号,卡号,卡有效期,发卡行编码,交易日期,交易时间,授权号,系统参照号,收银流水号,原收银流水号,系统流水号,原系统流水号等
                strSQL = strSQL & "'" & strValue & "',"
                '      交易内容_In In 三方结算交易.交易内容%Type
                '        交易内容    Varchar2    100
                Call zlXML_GetNodeValue("交易内容", lngRow, strValue, strErr)
                strSQL = strSQL & "'" & strValue & "')"
                If strErr <> "" Then
                    strErrMsg = strErrMsg & strErr & vbCrLf
                    Exit For
                End If
                zlAddArray cllPro, strSQL
            Next
        End If
    End If
    If zlXML_Init = False Then Exit Function
    If strErrMsg <> "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    If Trim(strData(8)) = "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "合作单位未传入接口"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    If Not IsDate(Trim(strData(1))) Or InStr(strData(1), "-") = 0 Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "发生日期未传入或格式(yyyy-mm-dd HH:MM:SS)不正确!"
        OutPutString_Xml = zlXML_GetXMLString
    End If
    If Val(strData(2)) = 0 Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "挂号标志传入错误!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    lng病人ID = Val(strData(7))
    If lng病人ID = 0 Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "未找到挂号的病人信息,请确认是否建立了病人信息档案"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    strSQL = "Select 项目ID From 挂号安排 where 号码=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取挂号项目ID", strData(0))
    If rsTemp.EOF Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "号码" & strData(0) & "未进行挂号安排,不能继续!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    lng项目id = Val(NVL(rsTemp!项目ID))
    '获取病人ID的费别
    strSQL = "Select nvl(费别,'') as 费别 From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng病人ID)
    If rsTemp.EOF Then
         zlXML_InsertNode gobjXml.documentElement, "ERROR", "未找到挂号的病人信息,请确认是否建立了病人信息档案!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    strNo = zlDatabase.GetNextNo(12)
    str登记时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    'Zl_三方机构挂号_Insert
    strSQL = "Zl_三方机构挂号_Insert("
    '  操作方式_In     Integer,
    strSQL = strSQL & "" & Val(strData(2)) & ","
    '  病人id_In       门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  号码_In         挂号安排.号码%Type,
    strSQL = strSQL & "'" & CStr(strData(0)) & "',"
    '  号序_In         挂号序号状态.序号%Type,
    strSQL = strSQL & "" & IIf(Val(strData(9)) = 0, "NULL", Val(strData(9))) & ","
    '  单据号_In       门诊费用记录.NO%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '  票据号_In       门诊费用记录.实际票号%Type,
    strSQL = strSQL & "NULL,"
    '  结算方式_In     病人预交记录.结算方式%Type, --现金的结算名称
    strSQL = strSQL & "'" & Trim(strData(5)) & "',"
    '  摘要_In         门诊费用记录.摘要%Type, --预约挂号摘要信息
    strSQL = strSQL & "'" & Trim(strData(6)) & "',"
    '  发生时间_In     门诊费用记录.发生时间%Type,
    strSQL = strSQL & "to_date('" & Trim(strData(1)) & "','yyyy-mm-dd hh24:mi:ss'),"
    '  登记时间_In     门诊费用记录.登记时间%Type,
    strSQL = strSQL & "to_date('" & str登记时间 & "','yyyy-mm-dd hh24:mi:ss'),"
    '  合作单位_In     挂号合作单位.名称%Type,
    strSQL = strSQL & "'" & Trim(strData(8)) & "',"
    '  挂号金额合计_In 门诊费用记录.实收金额%Type,
    strSQL = strSQL & "" & Val(strData(4)) & ","
    '  领用id_In       票据使用明细.领用id%Type,
    strSQL = strSQL & "NULL,"
    '  收费票据_In     Number := 0, --挂号是否使用收费票据
     strSQL = strSQL & "0,"
    '  交易流水号_In   病人预交记录.交易流水号%Type,
    strSQL = strSQL & "'" & Trim(strData(10)) & "',"
    '  交易说明_In     病人预交记录.交易说明%Type,
    strSQL = strSQL & "'" & Trim(strData(11)) & "',"
    '  预约方式_In     预约方式.名称%Type := Null
    strSQL = strSQL & "'" & Trim(strData(3)) & "',"
    '  预交ID_IN     病人预交记录%type:=NULL
    strSQL = strSQL & lng预交ID & ","
    '  保险结算_In     Varchar2 := Null,      格式="结算方式|结算金额||....."
    strSQL = strSQL & IIf(Trim(strData(12)) = "", "NULL", "'" & Trim(strData(12)) & "'") & ","
    '  冲预交_In Number:=Null
    strSQL = strSQL & IIf(Trim(strData(13)) = "", "NULL", Val(strData(13))) & ","
    '  支付卡号_In 病人预交记录.卡号%type:=Null
    strSQL = strSQL & IIf(Trim(strData(14)) = "", "NULL", "'" & Trim(strData(14)) & "'") & ")"
    Err = 0: On Error GoTo ErrReturn:
    If Not mblnStartTrans Then gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, "执行挂号数据"
    zlExecuteProcedureArrAy cllPro, "执行挂号数据", mblnStartTrans, True
    
    Err = 0: On Error GoTo errHandle:
    'OutPutString参数说明:
    'OutPutString
    '标识    数据类型    长度    精度    非空    说明
    '单据号  Varchar2    10      Y   挂号单号,退号时有用,三方机构需要保存.
     zlXML_InsertNode gobjXml.documentElement, "单据号", strNo
    '登记时间    Varchar2    20      Y   Yyyy-mm-dd hh24:mi:ss
     zlXML_InsertNode gobjXml.documentElement, "登记时间", str登记时间
    OutPutString_Xml = zlXML_GetXMLString
    zlRegister = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrReturn:
   If Err.Description Like "*[ZLSOFT]+*+[ZLSOFT]*" Then
        strErrMsg = Err.Description
        strErrMsg = Split(strErrMsg, "[ZLSOFT]+")(1)
        strErrMsg = Split(strErrMsg, "+[ZLSOFT]")(0)
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    Else
        strErrMsg = Err.Description
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
   End If
End Function
Public Function zlRegisterRollback(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退号操作
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-31 10:24:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strData(0 To 10) As String, i As Long
    Dim strErrMsg As String, strErr As String, strValue As String
    Dim str费别 As String, lng项目id As Long, lng病人ID As Long
    Dim strNo As String, str退号时间 As String
    Dim cllPro As New Collection, lng预交ID As Long, bln预约 As Boolean
    On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    
    i = 0
    'inputString
    '标识    数据类型    长度    精度    非空    说明
    '0.单据号  Varchar2    10      Y   挂号单号
    Call zlXML_GetNodeValue("单据号", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    strNo = Trim(strValue)
    '1姓名    Varchar2    100      Y   只有满足姓名和流水号相同的情况下,才能退号操作
    Call zlXML_GetNodeValue("姓名", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '2扣款交易流水号  Varchar2    50      Y   挂号时扣款的交易流水号,不正确,不能退号
    Call zlXML_GetNodeValue("扣款交易流水号", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '3扣款交易说明    Varchar2    500     Y   挂号时扣款的交易说明,不正确,不能退号
    Call zlXML_GetNodeValue("扣款交易说明", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '4金额    Number  16  5   Y   回退的金额;如果金额不等,也不能退号
    Call zlXML_GetNodeValue("金额", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '5退号交易流水号  Varchar2    50          退号时的交易流水号
    Call zlXML_GetNodeValue("退号交易流水号", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '6退号交易说明    Varchar2    500         退号时的交易说明
    Call zlXML_GetNodeValue("退号交易说明", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
    '7.合作单位    Varchar2    50      Y   三方合作单位名
    Call zlXML_GetNodeValue("合作单位", , strValue, strErr): strData(i) = Trim(strValue): i = i + 1
     If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf

    '8. 退号交易内容
    Dim lngRow As Long, lngCount As Long
    '不扣款部分,不传入
    lng预交ID = zlDatabase.GetNextId("病人预交记录")
    If zlXML_GetRows("交易项目", lngCount) Then
        For lngRow = 0 To lngCount - 1
            Call zlXML_GetNodeValue("交易项目", lngRow, strValue, strErr)
            If strErr <> "" Then
                strErrMsg = strErrMsg & strErr & vbCrLf
                Exit For
            End If
            If Trim(strValue) = "" Then
                strErrMsg = strErrMsg & "交易项目不能为空,请检查!"
                Exit For
            End If
            '    Zl_三方结算交易_Insert
            strSQL = "Zl_三方结算交易_Insert("
            '      交易id_In   In 三方结算交易.交易id%Type,
            strSQL = strSQL & "" & lng预交ID & ","
            '      交易项目_In In 三方结算交易.交易项目%Type,
            '        交易项目 Varchar2    50          如: 操作类型,交易类型,卡类型,返回码,返回信息,收银机编号,柜员号,金额,结算批次,商户号,商户名称,终端号,卡号,卡有效期,发卡行编码,交易日期,交易时间,授权号,系统参照号,收银流水号,原收银流水号,系统流水号,原系统流水号等
            strSQL = strSQL & "'" & strValue & "',"
            '      交易内容_In In 三方结算交易.交易内容%Type
            '        交易内容    Varchar2    100
            Call zlXML_GetNodeValue("交易内容", lngRow, strValue, strErr)
            strSQL = strSQL & "'" & strValue & "')"
            If strErr <> "" Then
                strErrMsg = strErrMsg & strErr & vbCrLf
                Exit For
            End If
            zlAddArray cllPro, strSQL
        Next
    End If
    str退号时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If zlXML_Init = False Then Exit Function
    If strErrMsg <> "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    If strData(0) = "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "单据号未转入"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    If strData(1) = "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "病人姓名未传入"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    If strData(2) = "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "挂号时的扣款流水号未传入"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
     strSQL = "" & _
    "   Select Max(decode(a.记录状态,0,1,0)) as 预约,Max(A.姓名) as 姓名,sum(A.实收金额) as 实收金额, " & _
    "              Max(nvl(J.交易流水号,B.交易流水号)) as 交易流水号, " & _
    "              Max(nvl(J.交易说明,B.交易说明)) as 交易说明, " & _
    "              Max(nvl(J.合作单位,B.合作单位)) as 合作单位" & _
    "   From 门诊费用记录 A,病人预交记录 B,病人挂号记录 J" & _
    "   Where A.记录性质=4 And A.NO=[1] and A.记录状态 in (0,1) And A.结帐ID=B.结帐ID(+)" & _
    "               And A.NO=J.NO And Decode(A.记录状态,0,1,A.记录状态)=J.记录状态 and J.合作单位=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取挂号信息", CStr(strData(0)), CStr(strData(7)))
    If rsTemp.EOF Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "挂号单未找到或已经被他人退号,不能继续!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    bln预约 = Val(NVL(rsTemp!预约)) = 1
    If NVL(rsTemp!姓名) <> strData(1) Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "你退号的病人姓名不符,不能继续!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
     
    If Round(Val(NVL(rsTemp!实收金额)), 4) <> Round(Val(strData(4)), 4) Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "退号的金额不正确,请检查!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    
    If Trim(NVL(rsTemp!交易流水号)) <> Trim(strData(2)) Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "扣款交易流水号传入不正确,需传入扣款时的交易流水号,请检查!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    
    If Trim(NVL(rsTemp!交易说明)) <> Trim(strData(3)) Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "扣款交易流水号传入不正确,需传入扣款时的交易流水号,请检查!"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    ' Zl_三方机构挂号_Delete
    strSQL = "Zl_三方机构挂号_Delete("
    '  单据号_In     门诊费用记录.NO%Type,
    strSQL = strSQL & "'" & Trim(strData(0)) & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type,
    strSQL = strSQL & "'" & Trim(strData(5)) & "',"
    '  交易说明_In   病人预交记录.交易说明%Type,
    strSQL = strSQL & "'" & Trim(strData(6)) & "',"
    '  退号时间_In   门诊费用记录.登记时间%Type := Null,
    strSQL = strSQL & "TO_Date('" & str退号时间 & "','yyyy-mm-dd hh24:mi:ss'),"
    '  预交id_In     病人预交记录.ID%Type := Null
    strSQL = strSQL & "" & lng预交ID & ")"
    Err = 0: On Error GoTo ErrReturn:
    If Not mblnStartTrans Then gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, "执行退号数据"
    If Not bln预约 Then
        '只有未扣预约,才不会存在预交记录
        zlExecuteProcedureArrAy cllPro, "执行退号数据", mblnStartTrans, True
    ElseIf Not mblnStartTrans Then
        gcnOracle.CommitTrans
    End If
    
    Err = 0: On Error GoTo errHandle
    'OutPutString参数说明:
    'OutPutString
    '标识    数据类型    长度    精度    非空    说明
    '单据号  Varchar2    10      Y   目前传的是退号的单据号(与挂号的单据号一致)
    zlXML_InsertNode gobjXml.documentElement, "单据号", strNo
    '登记时间    Varchar2    20      Y   yyyy-mm-dd hh24:mi:ss
    zlXML_InsertNode gobjXml.documentElement, "登记时间", str退号时间
    OutPutString_Xml = zlXML_GetXMLString
    zlRegisterRollback = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
    zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
    OutPutString_Xml = zlXML_GetXMLString
    Exit Function
ErrReturn:
   If Err.Description Like "*[ZLSOFT]+*+[ZLSOFT]*" Then
        strErrMsg = Err.Description
        strErrMsg = Split(strErrMsg, "[ZLSOFT]+")(1)
        strErrMsg = Split(strErrMsg, "+[ZLSOFT]")(0)
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    Else
        strErrMsg = Err.Description
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
   End If
End Function
 
Public Function zlRegisterQuery(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:挂号查询操作
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-31 10:24:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strErrMsg As String, strErr As String, strValue As String
    Dim strWhere As String, strWhereValue As String, strWhere1 As String
    Dim varPara(0 To 10) As Variant, i As Long, strSubTable As String
    Dim bln限制 As Boolean '时间范围;单据号;病人ID必须满足一个条件
    On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    
    i = 0
    'InputString说明:
    'inputString(下例条件必须选择一个进行传入,其中如果时间必须配对传入)
    '标识    数据类型    长度    精度    非空    说明
    '开始时间    Varchar2    20          yyyy-mm-dd hh24:mi:ss
    bln限制 = False
    varPara(i) = CDate("1901-01-01"): i = i + 1
     If zlXML_GetNodeValue("开始时间", , strValue, strErr) Then
        If Not IsDate(Trim(strValue)) Or Len(Trim(strValue)) <> 19 Or (Not Trim(strValue) Like "????-??-?? ??:??:??") Then             '非日期格式
            strErrMsg = strErrMsg & "开始时间不是有效的日期格式,请检查! " & vbCrLf
        End If
        varPara(i - 1) = CDate(strValue)
        strWhere = strWhere & " And A.登记时间>=[1]"
        bln限制 = True
    End If
    
    '结束时间    Varchar2    20          yyyy-mm-dd hh24:mi:ss
    varPara(i) = CDate("3000-01-01"): i = i + 1
     If zlXML_GetNodeValue("结束时间", , strValue, strErr) Then
        If Not IsDate(Trim(strValue)) Or Len(Trim(strValue)) <> 19 Or (Not Trim(strValue) Like "????-??-?? ??:??:??") Then             '非日期格式
            strErrMsg = strErrMsg & "结束时间不是有效的日期格式,请检查! " & vbCrLf
        End If
        varPara(i - 1) = CDate(strValue)
        '为了性能,只能查90天的数据
        If bln限制 = False Then varPara(0) = CDate(strValue) - 90
        strWhere = strWhere & " And A.登记时间<=[2]" & IIf(bln限制 = False, " And A.登记时间>=[1]", "")
        If bln限制 Then
            '检查时间范围不能大于三个月
            If varPara(1) - varPara(0) > 90 Then
                strErrMsg = strErrMsg & "开始时间攻与结束时间不能大于90天,请检查! " & vbCrLf
            End If
        End If
        bln限制 = True
    ElseIf bln限制 Then
        '只能查3个月范围内的数据
        varPara(i - 1) = varPara(0) + 90
        strWhere = strWhere & " And A.登记时间<=[2]"
    End If
    
    '单据号  Varchar2    10          挂号单号,允许为空
     varPara(i) = "": i = i + 1
    If zlXML_GetNodeValue("单据号", , strValue, strErr) Then
        If Trim(strValue) = "" Then
            strErrMsg = strErrMsg & "单据号不能为空,请检查! " & vbCrLf
        End If
        varPara(i - 1) = CStr(Trim(strValue))
        strWhere = strWhere & " And A.NO=[3]"
        bln限制 = True
    End If
    
    '交易流水号  Varchar2    50          交易时的流水号(交易时三方产生的)
     varPara(i) = "": i = i + 1
    If zlXML_GetNodeValue("交易流水号", , strValue, strErr) Then
        If Trim(strValue) = "" Then
            strErrMsg = strErrMsg & "交易流水号不能为空,请检查! " & vbCrLf
        End If
        varPara(i - 1) = CStr(Trim(strValue))
        strWhere = strWhere & " And A.交易流水号=[4]"
    End If
    
    '交易说明    Varchar2    500         交易说明(交易时三方产生的)
     varPara(i) = "": i = i + 1
    If zlXML_GetNodeValue("交易说明", , strValue, strErr) Then
        If Trim(strValue) = "" Then
            strErrMsg = strErrMsg & "交易流水号不能为空,请检查! " & vbCrLf
        End If
        varPara(i - 1) = CStr(Trim(strValue))
        strWhere = strWhere & " And A.交易说明=[5] "
    End If
    '病人ID  Number  18          允许为0:表示不按病人查找
     varPara(i) = 0: i = i + 1
    If zlXML_GetNodeValue("病人ID", , strValue, strErr) Then
        If Val(strValue) = 0 Then
            strErrMsg = strErrMsg & "病人ID必须大于零,请检查! " & vbCrLf
        End If
        varPara(i - 1) = Val(strValue)
        strWhere = strWhere & " And A.病人ID=[6] "
        bln限制 = True
    End If
    
    '号类    Varchar2    10
     varPara(i) = "": i = i + 1
    If zlXML_GetNodeValue("号类", , strValue, strErr) Then
        If Trim(strValue) = "" Then
            strErrMsg = strErrMsg & "号类必须输入,请检查! " & vbCrLf
        End If
        varPara(i - 1) = CStr(Trim(strValue))
        strWhere = strWhere & " And J.号类=[7] "
    End If
    
    
    '号码    Varchar2    5
   varPara(i) = "": i = i + 1
    If zlXML_GetNodeValue("号码", , strValue, strErr) Then
        If Trim(strValue) = "" Then
            strErrMsg = strErrMsg & "号码必须输入,请检查! " & vbCrLf
        End If
        varPara(i - 1) = CStr(Trim(strValue))
        strWhere = strWhere & " And A.号别=[8] "
    End If
    
    '合作单位    Varchar2    50      Y
    varPara(i) = "": i = i + 1
    If zlXML_GetNodeValue("合作单位", , strValue, strErr) Then
        If Trim(strValue) = "" Then
            strErrMsg = strErrMsg & "合作单位必须输入,请检查! " & vbCrLf
        End If
        varPara(i - 1) = CStr(Trim(strValue))
        strWhere = strWhere & " And A.合作单位=[9] "
    Else
        strErrMsg = strErrMsg & "合作单位必须输入,请检查! " & vbCrLf
    End If
    If bln限制 = False Then
        strErrMsg = strErrMsg & "时间范围,单据号,病人ID必须输入一种条件,请检查! " & vbCrLf
    End If
    If strErrMsg = "" Then
        Dim lngRow As Long, lngCount As Long
        If zlXML_GetRows("交易项目", lngCount) Then
            strWhereValue = ""
            For lngRow = 0 To lngCount - 1
                '交易详细内容(预约挂号不扣款部分,不包含这部件统计)
                '交易项目    Varchar2    50          如: 操作类型,交易类型,卡类型,返回码,返回信息,收银机编号,柜员号,金额,结算批次,商户号,商户名称,终端号,卡号,卡有效期,发卡行编码,交易日期,交易时间,授权号,系统参照号,收银流水号,原收银流水号,系统流水号,原系统流水号等
                Call zlXML_GetNodeValue("交易项目", lngRow, strValue, strErr)
                If Trim(strValue) <> "" Then
                    If InStr(1, strValue, ",") > 0 Or InStr(1, strValue, ":") > 0 Then
                        strErrMsg = strErrMsg & "交易项目为" & strValue & "中不能包含逗号和冒号,请检查 " & vbCrLf
                        Exit Function
                    End If
                    strWhereValue = strWhereValue & "," & strValue
                    '交易内容    Varchar2    100
                    Call zlXML_GetNodeValue("交易内容", lngRow, strValue, strErr)
                    If InStr(1, strValue, ",") > 0 Or InStr(1, strValue, ":") > 0 Then
                        strErrMsg = strErrMsg & "交易内容为" & strValue & "中不能包含逗号和冒号,请检查 " & vbCrLf
                        Exit Function
                    End If
                    strWhereValue = strWhereValue & ":" & strValue
                End If
            Next
            If strWhereValue <> "" Then
                strWhereValue = Mid(strWhere, 2)
                Call zlGetSubTable(3, strWhereValue, strSubTable, varPara, 8)
                If strSubTable <> "" Then strSubTable = " With 三方结算 as (" & strSubTable & ") "
            End If
        End If
    End If
    On Error GoTo errHandle
    Call zlXML_Init
    If strErrMsg <> "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    
    strSQL = "" & strSubTable & _
     "  Select Distinct  A.Id,A.No,A.病人ID,A.门诊号,A.姓名,A.性别,A.年龄,A.号别,A.号序 , " & _
     "             nvl(A.急诊,0) As 急诊,A.诊室,D.名称 As 执行部门,A.执行人, A.执行状态 ,A.执行时间,A.登记时间,A.发生时间,A.操作员编号,A.操作员姓名, " & _
     "             nvl(A.预约,1) As 预约, A.预约方式  ,A.完成时间,decode(A.记录状态,1,0,1) As 退号, " & _
     "             B.收费细目ID,E.编码 As 项目编码,E.名称 As 项目名称,B.收入项目ID,F.编码 As 收入项目编码,F.名称 As 收入项目名称, " & _
     "             B.数次,E.计算单位 as 单位 ,B.标准单价 As 单价,B.应收金额,B.实收金额,A.合作单位, A.交易流水号 ,  A.交易说明" & _
     "   From  病人挂号记录 A,门诊费用记录 B,部门表 D,收费项目目录 E,收入项目 F,挂号安排 J " & _
     "    Where A.No=B.No(+) And decode(A.记录性质,2,0,A.记录状态)=B.记录状态(+)  and B.记录性质(+)=4 " & _
                    IIf(strSubTable <> "", _
    " Exists(    Select 1 From 门诊费用记录 M,病人预交记录 H,三方结算交易 M,三方交易 J " & _
    "               Where  A.NO=M.NO And M.记录性质=4 and M.结帐ID=H.结帐ID AND h.ID=m.交易ID  " & _
    "                       And M.交易项目=J.C1 and M.交易名称=J.C2", "") & _
     "               And A.执行部门ID=D.Id(+) " & _
     "               And B.收费细目ID=E.Id(+) " & _
     "               And B.收入项目ID=F.Id(+) " & _
     "               And B.计算单位=J.号码"
     strSQL = strSQL & strWhere
    strSQL = "Select /*+ RULE */ * From (" & strSQL & ")  Order by  NO,登记时间"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "挂号查询", varPara)
    If rsTemp.EOF Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "未找到满足条件的数据,请检查"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
        Dim objNode As MSXML2.IXMLDOMElement
        Dim objTempNode As MSXML2.IXMLDOMElement
    With rsTemp
        Do While Not .EOF
           zlXML_InsertNode gobjXml.documentElement, "LISTDATA", "", objNode
            'OutPutString参数说明:
            'OutPutString
'            zlXML_InsertNode objNode, "ID", Nvl(rsTemp!ID)
            '标识    数据类型    长度    精度    说明
            '流水号  Number  18      挂号ID
            zlXML_InsertNode objNode, "流水号", NVL(rsTemp!id)
            '单据号  Varcahr2    8
            zlXML_InsertNode objNode, "单据号", NVL(rsTemp!NO)
            '病人ID      Number  18
            zlXML_InsertNode objNode, "病人ID", NVL(rsTemp!病人ID)
            '门诊号  Number  18
            zlXML_InsertNode objNode, "门诊号", NVL(rsTemp!门诊号)
            '姓名    Varchar2    100
            zlXML_InsertNode objNode, "姓名", NVL(rsTemp!姓名)
            '性别    Varchar2    4
            zlXML_InsertNode objNode, "性别", NVL(rsTemp!性别)
            '年龄    Varchar2    10
            zlXML_InsertNode objNode, "年龄", NVL(rsTemp!年龄)
            '号别    Varchar2    5
            zlXML_InsertNode objNode, "号别", NVL(rsTemp!号别)
            '号序    Number  5       挂号的序号
            zlXML_InsertNode objNode, "号序", NVL(rsTemp!号序)
            '是否急诊    Number  1       1-急诊;0-平诊
            zlXML_InsertNode objNode, "急诊", Val(NVL(rsTemp!急诊))
            '诊室    Varchar2    20
            zlXML_InsertNode objNode, "诊室", NVL(rsTemp!诊室)
            '执行部门    Varchar2    20
            zlXML_InsertNode objNode, "执行部门", NVL(rsTemp!执行部门)
            '执行人  Varchar2    20
            zlXML_InsertNode objNode, "执行人", NVL(rsTemp!执行人)
            '执行状态    Number  2       0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
            zlXML_InsertNode objNode, "执行状态", NVL(rsTemp!执行状态)
            '执行时间    Varchar2            yyyy-mm-dd hh24:mi:ss
            zlXML_InsertNode objNode, "执行时间", Format(rsTemp!执行时间, "yyyy-mm-dd HH:MM:SS")
            '登记时间    Varchar2            yyyy-mm-dd hh24:mi:ss
            zlXML_InsertNode objNode, "登记时间", Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
            '发生时间    Varchar2            yyyy-mm-dd hh24:mi:ss
            zlXML_InsertNode objNode, "发生时间", Format(rsTemp!发生时间, "yyyy-mm-dd HH:MM:SS")
            '操作员编号  Varchar2    6
            zlXML_InsertNode objNode, "操作员编号", NVL(rsTemp!操作员编号)
            '操作员姓名  Varchar2    20
            zlXML_InsertNode objNode, "操作员姓名", NVL(rsTemp!操作员姓名)
            '是否预约单  Number  1       1-是预约单;0-挂号单
            zlXML_InsertNode objNode, "是否预约单", Val(NVL(rsTemp!预约))
            '预约方式    Varchar2    10
            zlXML_InsertNode objNode, "预约方式", Trim(NVL(rsTemp!预约方式))
            '完成时间    Varchar2            yyyy-mm-dd hh24:mi:ss
            zlXML_InsertNode objNode, "完成时间", Format(rsTemp!完成时间, "yyyy-mm-dd HH:MM:SS")
            '是否退号    Number  1       是否为退号记录
            zlXML_InsertNode objNode, "是否退号", Val(NVL(rsTemp!退号))
            '项目ID  Number  18
            zlXML_InsertNode objNode, "项目ID", Val(NVL(rsTemp!收费细目ID))
            '项目编码    Varchar2    20
            zlXML_InsertNode objNode, "项目编码", Trim(NVL(rsTemp!项目编码))
            '项目名称    Varchar2    80
            zlXML_InsertNode objNode, "项目名称", Trim(NVL(rsTemp!项目名称))
            '收入项目ID  Number  18
            zlXML_InsertNode objNode, "收入项目ID", Val(NVL(rsTemp!收入项目ID))
            '收入项目编码    Varchar2    10
            zlXML_InsertNode objNode, "收入项目编码", Trim(NVL(rsTemp!收入项目编码))
            '收入项目名称    Varchar2    50
            zlXML_InsertNode objNode, "收入项目名称", Trim(NVL(rsTemp!收入项目名称))
            '单位    Varchar2    20
            zlXML_InsertNode objNode, "单位", Trim(NVL(rsTemp!单位))
            '数次    number  5
            zlXML_InsertNode objNode, "数次", Val(NVL(rsTemp!数次))
            '单价    Number  16  5
            zlXML_InsertNode objNode, "单价", Val(NVL(rsTemp!单价))
            '应收金额    Number  16  5
            zlXML_InsertNode objNode, "应收金额", Val(NVL(rsTemp!应收金额))
            '实收金额    Number  16  5
            zlXML_InsertNode objNode, "实收金额", Val(NVL(rsTemp!实收金额))
            .MoveNext
        Loop
    End With
    OutPutString_Xml = zlXML_GetXMLString
    zlRegisterQuery = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
    zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
    OutPutString_Xml = zlXML_GetXMLString
End Function
Public Function zlPrepay(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人缴预交款交易
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-31 10:24:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strErrMsg As String, strErr As String, strValue As String
    Dim strData(0 To 15) As String, cllPro As New Collection
    Dim i As Long, lng预交ID As Long, str登记时间 As String, strNo As String
   On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    i = 0
    On Error GoTo errHandle
    'InputString说明:
    'inputString
    '标识    数据类型    长度    精度    说明
    '0.日期    Varchar2    20      缴款时间格式yyyy-mm-dd hh24:mi:ss
     If zlXML_GetNodeValue("日期", , strValue, strErr) Then
        If Not IsDate(Trim(strValue)) Or Len(Trim(strValue)) <> 19 Or (Not Trim(strValue) Like "????-??-?? ??:??:??") Then             '非日期格式
            strErrMsg = strErrMsg & "日期不是有效的日期格式,请检查! " & vbCrLf
        End If
        strData(i) = CDate(strValue)
    End If
    
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '1.缴款单位    Varchar2    50
    i = i + 1
     If zlXML_GetNodeValue("缴款单位", , strValue, strErr) Then
        strData(i) = strValue
    End If
    '2单位开户行  Varchar2    50
    i = i + 1
     If zlXML_GetNodeValue("单位开户行", , strValue, strErr) Then
        strData(i) = strValue
    End If
    '3单位帐号    Varchar2    20
    i = i + 1
    If zlXML_GetNodeValue("单位帐号", , strValue, strErr) Then
        strData(i) = strValue
    End If
    '4缴款金额    Number  16  5   表示本次预交金额
    i = i + 1
    If zlXML_GetNodeValue("缴款金额", , strValue, strErr) Then
        If Val(strValue) <= 0 Then
            strErrMsg = strErrMsg & "缴款金额必须大于零 " & vbCrLf
        End If
        strData(i) = strValue
    End If
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '5结算方式    Varchar2    20      可以某个接口传入固定的结算方式.与HIS的结算方式要同步.以便以后区分.
    i = i + 1
    If zlXML_GetNodeValue("结算方式", , strValue, strErr) Then
        strData(i) = Trim(strValue)
    End If
    '6结算号码    Varchar2    20
    i = i + 1
    If zlXML_GetNodeValue("结算号码", , strValue, strErr) Then
        strData(i) = Trim(strValue)
    End If
    '7摘要    Varchar2    1000        备注
    i = i + 1
    If zlXML_GetNodeValue("摘要", , strValue, strErr) Then
        strData(i) = Trim(strValue)
    End If
    '8病人ID  Number  18
    i = i + 1
    If zlXML_GetNodeValue("病人ID", , strValue, strErr) Then
        If Val(strValue) <= 0 Then
            strErrMsg = strErrMsg & "病人ID必须传入 " & vbCrLf
        End If
        strData(i) = strValue
    End If
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '9卡类别ID    Number  18
    i = i + 1
    If zlXML_GetNodeValue("卡类别ID", , strValue, strErr) Then
        If Val(strValue) <= 0 Then
            strErrMsg = strErrMsg & "卡类别ID必须传入 " & vbCrLf
        End If
        strData(i) = strValue
    End If
    '10交易流水号  Varchar2    50
    i = i + 1
    If zlXML_GetNodeValue("交易流水号", , strValue, strErr) Then
        If Trim(strValue) = "" Then
            strErrMsg = strErrMsg & "交易流水号不能为空" & vbCrLf
        End If
        strData(i) = strValue
    End If
    '11交易说明    Varchar2    500     交易说明
    i = i + 1
    If zlXML_GetNodeValue("交易说明", , strValue, strErr) Then
        If Trim(strValue) = "" Then
            strErrMsg = strErrMsg & "交易说明不能为空" & vbCrLf
        End If
        strData(i) = strValue
    End If
    i = i + 1
    '12合作单位    Varchar2    50      Y
    If zlXML_GetNodeValue("合作单位", , strValue, strErr) Then
        If Trim(strValue) = "" Then
            strErrMsg = strErrMsg & "合作单位不能为空 " & vbCrLf
        End If
        strData(i) = strValue
    End If
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '13 病人姓名    Varchar2    100      Y   主要检查姓名是否与传入的病人ID所对应的姓名是否一致,不一致不能继续
    i = i + 1
    If zlXML_GetNodeValue("病人姓名", , strValue, strErr) Then
        If Trim(strValue) = "" Then
            strErrMsg = strErrMsg & "病人姓名不能为空 " & vbCrLf
        End If
        strData(i) = strValue
    End If
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    '14预交类别    Number  1       Y   1-门诊预交;2-住院预交
    i = i + 1
    If zlXML_GetNodeValue("预交类别", , strValue, strErr) Then
        If Val(strValue) <= 0 Or Val(strValue) >= 3 Then
            strErrMsg = strErrMsg & "预交类别必须设置为1-门诊预交或2-住院预交 " & vbCrLf
        End If
        strData(i) = strValue
    End If
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    i = i + 1
    '15卡号    Varchar2    50          如果是刷了卡的,传入刷卡的卡号
    If zlXML_GetNodeValue("卡号", , strValue, strErr) Then
        strData(i) = strValue
    End If
    '16交易详细内容
    i = i + 1
    Dim lngRow As Long, lngCount As Long
    '不扣款部分,不传入
     lng预交ID = zlDatabase.GetNextId("病人预交记录")
    If zlXML_GetRows("交易项目", lngCount) Then
        For lngRow = 0 To lngCount - 1
            Call zlXML_GetNodeValue("交易项目", lngRow, strValue, strErr)
            If strErr <> "" Then
                strErrMsg = strErrMsg & strErr & vbCrLf
                Exit For
            End If
            If Trim(strValue) = "" Then
                strErrMsg = strErrMsg & "交易项目不能为空,请检查!"
                Exit For
            End If
            '    Zl_三方结算交易_Insert
            strSQL = "Zl_三方结算交易_Insert("
            '      交易id_In   In 三方结算交易.交易id%Type,
            strSQL = strSQL & "" & lng预交ID & ","
            '      交易项目_In In 三方结算交易.交易项目%Type,
            '        交易项目 Varchar2    50          如: 操作类型,交易类型,卡类型,返回码,返回信息,收银机编号,柜员号,金额,结算批次,商户号,商户名称,终端号,卡号,卡有效期,发卡行编码,交易日期,交易时间,授权号,系统参照号,收银流水号,原收银流水号,系统流水号,原系统流水号等
            strSQL = strSQL & "'" & strValue & "',"
            '      交易内容_In In 三方结算交易.交易内容%Type
            '        交易内容    Varchar2    100
            Call zlXML_GetNodeValue("交易内容", lngRow, strValue, strErr)
            strSQL = strSQL & "'" & strValue & "')"
            If strErr <> "" Then
                strErrMsg = strErrMsg & strErr & vbCrLf
                Exit For
            End If
            zlAddArray cllPro, strSQL
        Next
    End If
    
    Call zlXML_Init
    If strErrMsg <> "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    strErrMsg = ""
    '先检查数据的合法性
    If Val(strData(9)) = 0 And strData(5) = "" Then
        '卡类别ID没有及结算方式也没有,不能进行操作
         strErrMsg = strErrMsg & "卡类别ID和结算方式必须输入其中之一 " & vbCrLf
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
     If Val(strData(9)) <> 0 Then
        strSQL = "Select 结算方式 From 医疗卡类别 Where id=[1] and nvl(是否自制,0)=0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取医疗卡的结算方式", Val(strData(9)))
        If Not rsTemp.EOF Then
            If NVL(rsTemp!结算方式) <> "" Then
                strData(5) = NVL(rsTemp!结算方式)
            End If
            If strData(5) = "" Then
                strErrMsg = strErrMsg & "结算方式不能为空 " & vbCrLf
                zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
                OutPutString_Xml = zlXML_GetXMLString
                Exit Function
            End If
        Else
            strErrMsg = strErrMsg & "卡类别ID传入错误,请检查该类别是否自制卡 " & vbCrLf
            zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
            OutPutString_Xml = zlXML_GetXMLString
            Exit Function
        End If
     End If
     strSQL = "Select 姓名  From 病人信息 where 病人ID=[1]"
     Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查病人信息是否正确!", Val(strData(8)))
     If rsTemp.EOF Then
        strErrMsg = strErrMsg & "传入的病人ID找不到病人信息" & vbCrLf
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
     End If
     If NVL(rsTemp!姓名) <> strData(13) Then
        strErrMsg = strErrMsg & "传入的病人姓名不对" & vbCrLf
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
     End If
     str登记时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
     strNo = zlDatabase.GetNextNo(11)
    'Zl_病人预交记录_Insert
    strSQL = "Zl_病人预交记录_Insert("
    '  Id_In         病人预交记录.ID%Type,
    strSQL = strSQL & "" & lng预交ID & ","
    '  单据号_In     病人预交记录.NO%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '  票据号_In     票据使用明细.号码%Type,
    strSQL = strSQL & "NULL,"
    '  病人id_In     病人预交记录.病人id%Type,
    strSQL = strSQL & Val(strData(8)) & ","
    '  主页id_In     病人预交记录.主页id%Type,
    strSQL = strSQL & "" & 0 & ","
    '  科室id_In     病人预交记录.科室id%Type,
    strSQL = strSQL & "" & UserInfo.部门ID & ","
    '  金额_In       病人预交记录.金额%Type,
    strSQL = strSQL & "" & Val(strData(4)) & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSQL = strSQL & "'" & Trim(strData(5)) & "',"
    '  结算号码_In   病人预交记录.结算号码%Type,
    strSQL = strSQL & "'" & Trim(strData(6)) & "',"
    '  缴款单位_In   病人预交记录.缴款单位%Type,
    strSQL = strSQL & "'" & Trim(strData(1)) & "',"
    '  单位开户行_In 病人预交记录.单位开户行%Type,
    strSQL = strSQL & "'" & Trim(strData(2)) & "',"
    '  单位帐号_In   病人预交记录.单位帐号%Type,
    strSQL = strSQL & "'" & Trim(strData(3)) & "',"
    '  摘要_In       病人预交记录.摘要%Type,
    strSQL = strSQL & "'" & Trim(strData(7)) & "',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  预交类别_In   病人预交记录.预交类别%Type := Null,
    strSQL = strSQL & "" & Val(strData(14)) & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(Val(strData(9)) = 0, "NULL", Val(strData(9))) & ","
    ' 结算卡序号_in 病人预交记录.结算卡序号%type:=NULL,
    strSQL = strSQL & "" & "NULL" & ","
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & strData(15) & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & strData(10) & "',"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & strData(11) & "',"
    '  合作单位_In   病人预交记录.合作单位%Type := Null
    strSQL = strSQL & "'" & strData(12) & "',"
    '收款时间_In   病人预交记录.收款时间%Type := Null
    strSQL = strSQL & "to_date('" & str登记时间 & "','yyyy-mm-dd hh24:mi:ss'))"
    Err = 0: On Error GoTo ErrReturn:
    If Not mblnStartTrans Then gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, "执行预交缴款"
    zlExecuteProcedureArrAy cllPro, "执行预交缴款", mblnStartTrans, True
    Err = 0: On Error GoTo errHandle:
    'OutPutString参数说明:
    'OutPutString
    '标识    数据类型    长度    精度    说明
    '单据号  Varchar2    10      预交单号
    zlXML_InsertNode gobjXml.documentElement, "单据号", strNo
    '病人ID  Number  10      医院建档的病人ID
    zlXML_InsertNode gobjXml.documentElement, "病人ID", Val(strData(8))
    '登记时间    Varchar2    20      Yyyy-mm-dd hh24:mi:ss
    zlXML_InsertNode gobjXml.documentElement, "登记时间", str登记时间
    '流水号  Number  18      目前传的是预交ID
    zlXML_InsertNode gobjXml.documentElement, "流水号", lng预交ID
    OutPutString_Xml = zlXML_GetXMLString
    zlPrepay = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
    zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
    OutPutString_Xml = zlXML_GetXMLString
    Exit Function
ErrReturn:
   If Err.Description Like "*[ZLSOFT]*[ZLSOFT]*" Then
        strErrMsg = Err.Description
        strErrMsg = Split(strErrMsg, "[ZLSOFT]")(1)
        strErrMsg = Split(strErrMsg, "[ZLSOFT]")(0)
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    Else
        strErrMsg = Err.Description
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
   End If
End Function
 
 Public Function zlPrepayRollback(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 预交退款
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-31 10:24:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strErrMsg As String, strErr As String, strValue As String
    Dim strData(0 To 20) As String, cllPro As New Collection
    Dim i As Long, lng预交ID As Long, str登记时间 As String, strNo As String
   On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    Dim lngRow As Long, lngCount As Long
    lng预交ID = zlDatabase.GetNextId("病人预交记录")
    i = 0
    On Error GoTo errHandle
    'InputString说明:
    'inputString
    '标识    数据类型    长度    精度    非空    说明
    '0预交类别    Number  1       Y   1-门诊预交;2-住院预交
     Call zlXML_GetNodeValue("预交类别", , strValue, strErr)
     If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
     If (Val(strValue) <= 0 Or Val(strValue) >= 3) And strErr = "" Then strErrMsg = strErrMsg & "预交类别只能为门诊(1)住院(2)" & vbCrLf
     strData(i) = strValue: i = i + 1
    '1结算方式    Varchar2    20          可以某个接口传入固定的结算方式.与HIS的结算方式要同步.以便以后区分.
    '   为空时 , 必须满足卡类别ID不为空, 以卡类别ID对应的结算方式为准.否则不能为空
     Call zlXML_GetNodeValue("结算方式", , strValue, strErr)
     strData(i) = strValue: i = i + 1
    '2退款金额    Number  16  5   Y   表示本次退预交金额(不能超过已经消费的剩于预交)>=0
     Call zlXML_GetNodeValue("退款金额", , strValue, strErr)
     If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
     If Val(strValue) <= 0 And strErr = "" Then strErrMsg = strErrMsg & "退款金额必须大于零" & vbCrLf
     strData(i) = strValue: i = i + 1
     '3摘要    Varchar2    50          备注
     Call zlXML_GetNodeValue("摘要", , strValue, strErr)
     strData(i) = strValue: i = i + 1
    '4卡类别ID    Number  18
     Call zlXML_GetNodeValue("卡类别ID", , strValue, strErr)
     If Val(strValue) <= 0 And strErr = "" Then strErrMsg = strErrMsg & "卡类别的值不正确" & vbCrLf
     strData(i) = strValue: i = i + 1
    '5卡号    Varchar2    50
     Call zlXML_GetNodeValue("卡号", , strValue, strErr)
     strData(i) = strValue: i = i + 1
    '6病人姓名    Varchar2    100      Y   退款人姓名;病人不一致时,不能退款.
     Call zlXML_GetNodeValue("病人姓名", , strValue, strErr)
     If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
     If Trim(strValue) = "" And strErr = "" Then strErrMsg = strErrMsg & "病人姓名不能为空" & vbCrLf
     strData(i) = strValue: i = i + 1
    '7病人ID  Number  18      Y   病人ID必须转入,以较对病人姓名是否一致!
     Call zlXML_GetNodeValue("病人ID", , strValue, strErr)
     If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
     If Val(strValue) <= 0 And strErr = "" Then strErrMsg = strErrMsg & "病人ID不能为空" & vbCrLf
     strData(i) = strValue: i = i + 1
    '8合作单位    Varchar2    50      Y
     Call zlXML_GetNodeValue("合作单位", , strValue, strErr)
     If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
     If Trim(strValue) = 0 And strErr = "" Then strErrMsg = strErrMsg & "合作单位不能为空" & vbCrLf
     strData(i) = strValue: i = i + 1
    '9退款交易流水号  Varchar2    50          传入退款时的交易流水号
     Call zlXML_GetNodeValue("退款交易流水号", , strValue, strErr)
     strData(i) = strValue: i = i + 1
    '10退款交易说明    Varchar2    500         传入退款时交易说明
     Call zlXML_GetNodeValue("退款交易说明", , strValue, strErr)
     strData(i) = strValue: i = i + 1
     '11结算号码    Varchar2    20
     Call zlXML_GetNodeValue("结算号码", , strValue, strErr)
     strData(i) = strValue: i = i + 1
    '11退款交易内容
    If zlXML_GetRows("交易项目", lngCount) Then
        For lngRow = 0 To lngCount - 1
            '交易项目    Varchar2    50          如: 操作类型,交易类型,卡类型,返回码,返回信息,收银机编号,柜员号,金额,结算批次,商户号,商户名称,终端号,卡号,卡有效期,发卡行编码,交易日期,交易时间,授权号,系统参照号,收银流水号,原收银流水号,系统流水号,原系统流水号等
            '交易内容    Varchar2    100
            Call zlXML_GetNodeValue("交易项目", lngRow, strValue, strErr)
            If strErr <> "" Then
                strErrMsg = strErrMsg & strErr & vbCrLf
                Exit For
            End If
            If Trim(strValue) = "" Then
                strErrMsg = strErrMsg & "交易项目不能为空,请检查!"
                Exit For
            End If
            '    Zl_三方结算交易_Insert
            strSQL = "Zl_三方结算交易_Insert("
            '      交易id_In   In 三方结算交易.交易id%Type,
            strSQL = strSQL & "" & lng预交ID & ","
            '      交易项目_In In 三方结算交易.交易项目%Type,
            '        交易项目 Varchar2    50          如: 操作类型,交易类型,卡类型,返回码,返回信息,收银机编号,柜员号,金额,结算批次,商户号,商户名称,终端号,卡号,卡有效期,发卡行编码,交易日期,交易时间,授权号,系统参照号,收银流水号,原收银流水号,系统流水号,原系统流水号等
            strSQL = strSQL & "'" & strValue & "',"
            '      交易内容_In In 三方结算交易.交易内容%Type
            '        交易内容    Varchar2    100
            Call zlXML_GetNodeValue("交易内容", lngRow, strValue, strErr)
            strSQL = strSQL & "'" & strValue & "')"
            If strErr <> "" Then
                strErrMsg = strErrMsg & strErr & vbCrLf
                Exit For
            End If
            zlAddArray cllPro, strSQL
        Next
    End If
    
    Call zlXML_Init
    If strErrMsg <> "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    strErrMsg = ""
    '先检查数据的合法性
    If Val(strData(4)) = 0 And strData(1) = "" Then
        '卡类别ID没有及结算方式也没有,不能进行操作
         strErrMsg = strErrMsg & "卡类别ID和结算方式必须输入其中之一 " & vbCrLf
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    
     If Val(strData(4)) <> 0 Then
        strSQL = "Select 结算方式 From 医疗卡类别 Where id=[1] and nvl(是否自制,0)=0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取医疗卡的结算方式", Val(strData(4)))
        If rsTemp.EOF Then
            zlXML_InsertNode gobjXml.documentElement, "ERROR", "卡类别ID未找到,请检查!"
            OutPutString_Xml = zlXML_GetXMLString
            Exit Function
        End If
        If NVL(rsTemp!结算方式) <> "" Then
            strData(1) = NVL(rsTemp!结算方式)
        End If
        If strData(1) = "" Then
            strErrMsg = "结算方式不能为空 "
            zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
            OutPutString_Xml = zlXML_GetXMLString
            Exit Function
        End If
    End If
     strSQL = "Select 姓名 ,sum(nvl(B.预交余额,0)) as 预交余额 From 病人信息 A,病人余额 B Where A.病人ID=b.病人ID(+) And A.病人ID=[1] And B.类型(+)=[2] Group by 姓名"
     Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查病人信息是否正确!", Val(strData(7)), Val(strData(0)))
     If rsTemp.EOF Then
        strErrMsg = strErrMsg & "传入的病人ID找不到病人信息" & vbCrLf
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
     End If
     
     If NVL(rsTemp!姓名) <> strData(6) Then
        strErrMsg = strErrMsg & "传入的病人姓名不对" & vbCrLf
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
     End If
     If Val(NVL(rsTemp!预交余额)) <> Val(strData(2)) Then
        strErrMsg = strErrMsg & "当前病人的预交余额为" & Format(Val(NVL(rsTemp!预交余额)), "###0.000;-###0.0000;0.00;0.00") & ",小于了当前退款额!" & vbCrLf
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
     End If
     '取剩余款
     strSQL = "" & _
     "  Select  nvl(Sum(金额),0)-nvl(sum(冲预交),0) As 金额,Max(缴款单位) as 缴款单位,max(单位开户行) 单位开户行,max(单位帐号) as 单位帐号 " & _
     "  From ( " & _
     "              Select nvl(金额,0) as 金额,nvl(冲预交,0) as 冲预交,缴款单位,单位开户行,单位帐号 From 病人预交记录  " & _
     "              Where 病人ID=[1] and 合作单位=[2] and 结算方式=[3] And nvl(预交类别,0)=[4]  Union ALL " & _
     "              Select nvl(金额,0) as 金额,nvl(冲预交,0) as 冲预交,缴款单位,单位开户行,单位帐号 From H病人预交记录  " & _
     "              Where 病人ID=[1] and 合作单位=[2] and 结算方式=[3] And nvl(预交类别,0)=[4] " & _
     "      )"
     
     Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取剩余款", Val(strData(7)), strData(8), strData(1), Val(strData(0)))
     If Val(NVL(rsTemp!金额)) < Val(strData(2)) Then
        strErrMsg = strErrMsg & "当前病人的预交余额为" & Format(Val(NVL(rsTemp!金额)), "###0.000;-###0.0000;0.00;0.00") & ",小于了当前退款额!" & vbCrLf
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
     End If
     str登记时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
     strNo = zlDatabase.GetNextNo(11)
    'Zl_病人预交记录_Insert
    strSQL = "Zl_病人预交记录_Insert("
    '  Id_In         病人预交记录.ID%Type,
    strSQL = strSQL & "" & lng预交ID & ","
    '  单据号_In     病人预交记录.NO%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '  票据号_In     票据使用明细.号码%Type,
    strSQL = strSQL & "NULL,"
    '  病人id_In     病人预交记录.病人id%Type,
    strSQL = strSQL & Val(strData(7)) & ","
    '  主页id_In     病人预交记录.主页id%Type,
    strSQL = strSQL & "" & 0 & ","
    '  科室id_In     病人预交记录.科室id%Type,
    strSQL = strSQL & "" & UserInfo.部门ID & ","
    '  金额_In       病人预交记录.金额%Type,
    strSQL = strSQL & "" & -1 * Val(strData(2)) & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSQL = strSQL & "'" & Trim(strData(1)) & "',"
    '  结算号码_In   病人预交记录.结算号码%Type,
    strSQL = strSQL & "'" & Trim(strData(11)) & "',"
    '  缴款单位_In   病人预交记录.缴款单位%Type,
    strSQL = strSQL & "'" & NVL(rsTemp!缴款单位) & "',"
    '  单位开户行_In 病人预交记录.单位开户行%Type,
    strSQL = strSQL & "'" & NVL(rsTemp!单位开户行) & "',"
    '  单位帐号_In   病人预交记录.单位帐号%Type,
    strSQL = strSQL & "'" & NVL(rsTemp!单位帐号) & "',"
    '  摘要_In       病人预交记录.摘要%Type,
    strSQL = strSQL & "'" & Trim(strData(3)) & "',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  预交类别_In   病人预交记录.预交类别%Type := Null,
    strSQL = strSQL & "" & Val(strData(0)) & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(Val(strData(4)) = 0, "NULL", Val(strData(4))) & ","
    '  结算卡序号_in 病人预交记录.结算卡序号%type:=NULL,
    strSQL = strSQL & " NULL,"
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & strData(5) & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & strData(9) & "',"
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & strData(10) & "',"
    '  合作单位_In   病人预交记录.合作单位%Type := Null
    strSQL = strSQL & "'" & strData(8) & "',"
    '收款时间_In   病人预交记录.收款时间%Type := Null
    strSQL = strSQL & "to_date('" & str登记时间 & "','yyyy-mm-dd hh24:mi:ss'))"
    Err = 0: On Error GoTo ErrReturn:
    If Not mblnStartTrans Then gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, "执行退预交"
    zlExecuteProcedureArrAy cllPro, "执行退预交", mblnStartTrans, True
    Err = 0: On Error GoTo errHandle:
    'OutPutString参数说明:
    'OutPutString
    '标识    数据类型    长度    精度    说明
    '单据号  Varchar2    10      预交退款单号
    zlXML_InsertNode gobjXml.documentElement, "单据号", strNo
    '病人ID  Number  10      医院建档的病人ID
    zlXML_InsertNode gobjXml.documentElement, "病人ID", Val(strData(7))
    '退款时间    Varchar2    20      Yyyy-mm-dd hh24:mi:ss
    zlXML_InsertNode gobjXml.documentElement, "退款时间", str登记时间
    '流水号  Number  18      目前传的是预交ID
    zlXML_InsertNode gobjXml.documentElement, "流水号", lng预交ID
    OutPutString_Xml = zlXML_GetXMLString
    zlPrepayRollback = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
    zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
    OutPutString_Xml = zlXML_GetXMLString
    zlPrepayRollback = True
ErrReturn:
   If Err.Description Like "*[ZLSOFT]*[ZLSOFT]*" Then
        strErrMsg = Err.Description
        strErrMsg = Split(strErrMsg, "[ZLSOFT]")(1)
        strErrMsg = Split(strErrMsg, "[ZLSOFT]")(0)
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    Else
        strErrMsg = Err.Description
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
   End If
End Function
 Public Function zlPrepayQuery(ByVal InputString_XML As String, ByRef OutPutString_Xml As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 预交款查询交易
    '入参:InputString_XML-输入条件
    '出参:OutPutString_Xml-输出内容
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-31 10:24:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strErrMsg As String, strErr As String, strValue As String
    Dim strData(0 To 20) As String, cllPro As New Collection
    Dim varPara() As Variant, bln限制 As Boolean, strWhere As String, strWhereValue As String
    Dim lngRow As Long, lngCount As Long, strSubTable As String
    Dim i As Long, lng预交ID As Long, str登记时间 As String, strNo As String
   On Error GoTo errHandle
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(InputString_XML) = False Then Exit Function
    i = 0
    On Error GoTo errHandle
    'inputString: 时间范围 , 单据号, 流水号, 病人ID至少要传入一个条件
    '标识    数据类型    长度    精度    说明
    '开始时间    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    '以收款时间查询
    '１. 结束时间没有:查询范围: 收款时间 between 开始时间 and 开始时间+90
    '２. 结束时间存在,则间隔时间范围不能大于90天
    bln限制 = False
    varPara(i) = CDate("1901-01-01"): i = i + 1
     If zlXML_GetNodeValue("开始时间", , strValue, strErr) Then
        If Not IsDate(Trim(strValue)) Or Len(Trim(strValue)) <> 19 Or (Not Trim(strValue) Like "????-??-?? ??:??:??") Then             '非日期格式
            strErrMsg = strErrMsg & "开始时间不是有效的日期格式,请检查! " & vbCrLf
        End If
        varPara(i - 1) = CDate(strValue)
        strWhere = strWhere & " And A.收款时间>=[1]"
        bln限制 = True
    End If
    
    '结束时间    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    '1.无开始时间,则查询范围为: 收款时间 between 结束时间-90 and 结束时间
    varPara(i) = CDate("3000-01-01"): i = i + 1
     If zlXML_GetNodeValue("结束时间", , strValue, strErr) Then
        If Not IsDate(Trim(strValue)) Or Len(Trim(strValue)) <> 19 Or (Not Trim(strValue) Like "????-??-?? ??:??:??") Then             '非日期格式
            strErrMsg = strErrMsg & "结束时间不是有效的日期格式,请检查! " & vbCrLf
        End If
        varPara(i - 1) = CDate(strValue)
        '为了性能,只能查90天的数据
        If bln限制 = False Then varPara(0) = CDate(strValue) - 90
        strWhere = strWhere & " And A.收款时间<=[2]" & IIf(bln限制 = False, " And A.收款时间>=[1]", "")
        If bln限制 Then
            '检查时间范围不能大于三个月
            If varPara(1) - varPara(0) > 90 Then
                strErrMsg = strErrMsg & "开始时间攻与结束时间不能大于90天,请检查! " & vbCrLf
            End If
        End If
        bln限制 = True
    ElseIf bln限制 Then
        '只能查3个月范围内的数据
        varPara(i - 1) = varPara(0) + 90
        strWhere = strWhere & " And A.收款时间<=[2]"
    End If
    '2单据号  Varchar2    10      预交单号,允许为空
     Call zlXML_GetNodeValue("单据号", , strValue, strErr)
     If strValue <> "" Then bln限制 = True: strWhere = strWhere & " And A.NO=[3]"
    varPara(i - 1) = strValue: i = i + 1
    '3流水号  Number  18      允许为零:表示不按流水号查找
     Call zlXML_GetNodeValue("流水号", , strValue, strErr)
    If Val(strValue) <> 0 Then bln限制 = True: strWhere = strWhere & " And A.ID=[4]"
    varPara(i - 1) = Val(strValue): i = i + 1
    '4病人ID  Number  18      允许为0:表示不按病人查找
    Call zlXML_GetNodeValue("病人ID", , strValue, strErr)
    If Val(strValue) <> 0 Then bln限制 = True: strWhere = strWhere & " And A.病人ID=[5]"
    varPara(i - 1) = Val(strValue): i = i + 1
    '5交易流水号  Varchar2    50
    Call zlXML_GetNodeValue("交易流水号", , strValue, strErr)
    If Trim(strValue) <> "" Then strWhere = strWhere & " And A.交易流水号=[6]"
    varPara(i - 1) = Trim(strValue): i = i + 1
    '6交易说明    Varchar2    500
    Call zlXML_GetNodeValue("交易流水号", , strValue, strErr)
    If Trim(strValue) <> "" Then strWhere = strWhere & " And A.交易流水号=[7]"
    varPara(i - 1) = Trim(strValue): i = i + 1
    '7.合作单位    Varchar2    50  Y
    Call zlXML_GetNodeValue("合作单位", , strValue, strErr)
    If strErr <> "" Then strErrMsg = strErrMsg & strErr & vbCrLf
    If strValue = "" And strErr = "" Then strErrMsg = strErrMsg & "合作单位不能为空,请检查!" & vbCrLf
    If Trim(strValue) <> "" Then strWhere = strWhere & " And A.合作单位=[8]"
    varPara(i - 1) = Trim(strValue): i = i + 1
    
    If strErrMsg = "" Then
        If zlXML_GetRows("交易项目", lngCount) Then
            strWhereValue = ""
            For lngRow = 0 To lngCount - 1
                '交易详细内容(预约挂号不扣款部分,不包含这部件统计)
                '交易项目    Varchar2    50          如: 操作类型,交易类型,卡类型,返回码,返回信息,收银机编号,柜员号,金额,结算批次,商户号,商户名称,终端号,卡号,卡有效期,发卡行编码,交易日期,交易时间,授权号,系统参照号,收银流水号,原收银流水号,系统流水号,原系统流水号等
                Call zlXML_GetNodeValue("交易项目", lngRow, strValue, strErr)
                If Trim(strValue) <> "" Then
                    If InStr(1, strValue, ",") > 0 Or InStr(1, strValue, ":") > 0 Then
                        strErrMsg = strErrMsg & "交易项目为" & strValue & "中不能包含逗号和冒号,请检查 " & vbCrLf
                        Exit Function
                    End If
                    strWhereValue = strWhereValue & "," & strValue
                    '交易内容    Varchar2    100
                    Call zlXML_GetNodeValue("交易内容", lngRow, strValue, strErr)
                    If InStr(1, strValue, ",") > 0 Or InStr(1, strValue, ":") > 0 Then
                        strErrMsg = strErrMsg & "交易内容为" & strValue & "中不能包含逗号和冒号,请检查 " & vbCrLf
                        Exit Function
                    End If
                    strWhereValue = strWhereValue & ":" & strValue
                End If
            Next
            
            If strWhereValue <> "" Then
                strWhereValue = Mid(strWhere, 2)
                Call zlGetSubTable(3, strWhereValue, strSubTable, varPara, 8)
                If strSubTable <> "" Then strSubTable = " With 三方结算 as (" & strSubTable & ") "
            End If
        End If
    End If
    Call zlXML_Init
    If strErrMsg <> "" Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
 
    strSQL = "" & strSubTable & _
     "  Select Distinct  A.ID,NO,病人ID,B.门诊号,A.姓名,A.性别,A.年龄,A.缴款单位,A.单位开户行," & _
     "           A.单位帐号,A.摘要,A.金额,A.结算方式,A.结算号码, " & _
     "          to_char(A.收款时间,'yyyy-mm-dd hh24:mi:ss') as 收款时间," & _
     "           A.操作员编号,A.操作员姓名" & _
     "   From  病人预交记录 A,病人信息  B" & _
     "    Where A.记录性质=1 And A.病人ID=B.病人ID " & _
                    IIf(strSubTable <> "", _
    " Exists(    Select 1 From   三方结算交易 M,三方交易 J " & _
    "               Where  A.ID=M.交易ID  " & _
    "                       And M.交易项目=J.C1 and M.交易名称=J.C2", "")
         
    strSQL = "Select /*+ RULE */ * From (" & strSQL & ")  Order by  收款时间"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "预交查询", varPara)
    If rsTemp.EOF Then
        zlXML_InsertNode gobjXml.documentElement, "ERROR", "未找到满足条件的数据,请检查"
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    End If
    'OutPutString参数说明:
    'OutPutString
    With rsTemp
        Do While Not .EOF
            '标识    数据类型    长度    精度    说明
            '流水号  Number  18      预交ID
            zlXML_InsertNode gobjXml.documentElement, "流水号", NVL(rsTemp!id)
            '单据号  Varcahr2    8
            zlXML_InsertNode gobjXml.documentElement, "单据号", NVL(rsTemp!NO)
            '病人ID      Number  18
            zlXML_InsertNode gobjXml.documentElement, "病人ID", NVL(rsTemp!病人ID)
            '门诊号  Number  18
            zlXML_InsertNode gobjXml.documentElement, "门诊号", NVL(rsTemp!门诊号)
            '姓名    Varchar2    100
            zlXML_InsertNode gobjXml.documentElement, "姓名", NVL(rsTemp!姓名)
            '性别    Varchar2    4
            zlXML_InsertNode gobjXml.documentElement, "性别", NVL(rsTemp!性别)
            '年龄    Varchar2    10
            zlXML_InsertNode gobjXml.documentElement, "年龄", NVL(rsTemp!年龄)
            '缴款单位    Varchar2    50
            zlXML_InsertNode gobjXml.documentElement, "缴款单位", NVL(rsTemp!缴款单位)
            '单位开户行  Varchar2    50
            zlXML_InsertNode gobjXml.documentElement, "单位开户行", NVL(rsTemp!单位开户行)
            '单位帐号    Varchar2    20
            zlXML_InsertNode gobjXml.documentElement, "单位帐号", NVL(rsTemp!单位帐号)
            '摘要    Varchar2    50
            zlXML_InsertNode gobjXml.documentElement, "摘要", NVL(rsTemp!摘要)
            '缴款金额    number  16  5   退款以负数体现
            zlXML_InsertNode gobjXml.documentElement, "缴款金额", NVL(rsTemp!缴款金额)
            '结算方式    Varchar2    20
            zlXML_InsertNode gobjXml.documentElement, "结算方式", NVL(rsTemp!结算方式)
            '结算号码    Varchar2    20
            zlXML_InsertNode gobjXml.documentElement, "结算号码", NVL(rsTemp!结算号码)
            '收款时间    Varchar2    20      Yyyy-mm-dd hh24:mi:ss
            zlXML_InsertNode gobjXml.documentElement, "收款时间", NVL(rsTemp!收款时间)
            '操作员编号  Varchar2    6
            zlXML_InsertNode gobjXml.documentElement, "操作员编号", NVL(rsTemp!操作员编号)
            '操作员姓名  Varchar2    20
            zlXML_InsertNode gobjXml.documentElement, "操作员姓名", NVL(rsTemp!操作员姓名)
            .MoveNext
        Loop
    End With
    OutPutString_Xml = zlXML_GetXMLString
    zlPrepayQuery = True
    Exit Function
errHandle:
   If Err.Description Like "*[ZLSOFT]*[ZLSOFT]*" Then
        strErrMsg = Err.Description
        strErrMsg = Split(strErrMsg, "[ZLSOFT]")(1)
        strErrMsg = Split(strErrMsg, "[ZLSOFT]")(0)
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
    Else
        strErrMsg = Err.Description
        zlXML_InsertNode gobjXml.documentElement, "ERROR", strErrMsg
        OutPutString_Xml = zlXML_GetXMLString
        Exit Function
   End If
End Function
Public Function zlBeginTrans(Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:事务开始接口
    '出参:strErrMsg-错误信息,为空表示成功
    '返回: Boolean 函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-06-10 16:10:20
    '说明: 表示事务开始 , 在使用事务时有效
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    mblnStartTrans = True
    zlBeginTrans = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function
Public Function zlRollbackTrans(Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:事务回退交易
    '出参:strErrMsg-错误信息,为空表示成功
    '返回: Boolean 函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-06-10 16:10:20
    '说明: 表示事务回退 , 在使用事务时有效
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    gcnOracle.RollbackTrans
    mblnStartTrans = False
    zlRollbackTrans = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function
Public Function zlCommitTrans(Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:事务提交交易
    '出参:strErrMsg-错误信息,为空表示成功
    '返回: Boolean 函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-06-10 16:10:20
    '说明: 表示事务提交 , 在使用事务时有效
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    gcnOracle.CommitTrans
    mblnStartTrans = False
    zlCommitTrans = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function
'-----------------------------------------------------------------------------------------------------
Public Function zlSendCardBeforeCheck(ByVal bytType As Byte, _
    ByVal lngCardTypeID As Long, ByVal strPatiXML As String, ByRef strErrMsg As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发卡或签约操作
    '入参:bytType-0-发卡;1-签约
    '       lngCardTypeID-卡类别ID
    '       strPatiXML-XML串(卡号;身份证号)
    '出参:strExpend-暂未用
    '编制:刘兴洪
    '日期:2012-06-21 14:40:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strCardNo As String
    On Error GoTo errHandle
     
    '检测当前卡是否已存在
    gstrSQL = "select 病人ID from 病人医疗卡信息 where 卡类别ID=[1] And 卡号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Check Card is haved", lngCardTypeID, strCardNo)
    If rsTemp.EOF = False Then
        strCardNo = "当前卡号(" & strCardNo & ")已经" & IIf(bytType = 0, "发卡", "签约")
        Exit Function
    End If
    zlSendCardBeforeCheck = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function
Public Function zlSendCard(ByVal bytType As Byte, lngCardTypeID As Long, _
    ByVal strPatiXML As String, ByRef strErrMsg As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发卡或签约操作
    '入参:bytType-0-发卡;1-签约
    '       lngCardTypeID-卡类别ID
    '       strPatiXML-发卡病人信(无根接点)
    '           病人ID,卡号,姓名,性别,出生日期,年龄,身份证号,联系电话,支付密码,发卡原因,发卡时间,充值金额
    '出参:strExpend-暂未用
    '编制:刘兴洪
    '日期:2012-06-21 14:40:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    On Error GoTo errHandle
  
    
  


    zlSendCard = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function
Private Function GetSpecialCardItemFee(ByVal strCardType As String) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取特定项目的卡费信息
    '入参:strCardType-卡类别
    '返回:卡费信息集
    '编制:刘兴洪
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = _
    " Select  A.类别 as 收费类别,A.ID as 收费细目ID,A.计算单位,C.ID as 收入项目ID,nvl(A.屏蔽费别,0) as 屏蔽费别, " & _
    "       C.名称 as 收入项目,C.收据费目,B.原价,B.现价,Nvl(B.缺省价格,0) 缺省价格,Nvl(A.是否变价,0) as 是否变价,Nvl(A.执行科室,0) as 科室标志 " & _
    " From 收费项目目录 A,收费价目 B,收入项目 C,收费特定项目 D " & _
    " Where B.收费细目ID = A.ID And B.收入项目ID = C.ID And D.特定项目=[1]" & _
    "       And D.收费细目ID=A.ID And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
    "       And ((To_Char(Sysdate,'YYYY-MM-DD HH24:MI:SS') Between To_Char(B.执行日期,'YYYY-MM-DD HH24:MI:SS') and To_Char(B.终止日期,'YYYY-MM-DD HH24:MI:SS')) or (To_Char(Sysdate,'YYYY-MM-DD HH24:MI:SS')>=To_Char(B.执行日期,'YYYY-MM-DD HH24:MI:SS') And (B.终止日期 is NULL Or B.终止日期=To_Date('3000-01-01','YYYY-MM-DD'))))"
    On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取特定项目的费用集", strCardType)
    If Not rsTemp.EOF Then Set GetSpecialCardItemFee = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
