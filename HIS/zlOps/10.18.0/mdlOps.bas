Attribute VB_Name = "mdlOps"

Option Explicit

'常量定义
'######################################################################################################################

'枚举
'----------------------------------------------------------------------------------------------------------------------
Public Enum COLOR_NativeXpPlain
    BackgroundDark = 14054755
    BackgroundLight = 15180411
    HighlightBorderBottomRight = 8388608
    HighlightBorderTopLeft = 8388608
    HighlightHot = 12775167
    HighlightPressed = 4096254
    HighlightSelected = 7323903
    NormalGroupCaptionDark = 14215660
    NormalGroupCaptionLight = 14215660
    NormalGroupCaptionTextHot = 0
    NormalGroupCaptionTextNormal = 0
    NormalGroupClient = 16244694
    NormalGroupClientBorder = 16777215
    NormalGroupClientLink = 12999969
    NormalGroupClientLinkHot = 16748098
    NormalGroupClientText = 0
    SpecialGroupCaptionDark = 14215660
    SpecialGroupCaptionLight = 14215660
    SpecialGroupCaptionTextHot = 0
    SpecialGroupCaptionTextSpecial = 0
    SpecialGroupClient = 16244694
    SpecialGroupClientBorder = 16777215
    SpecialGroupClientLink = 12999969
    SpecialGroupClientLinkHot = 16748098
    SpecialGroupClientText = 0
End Enum
'----------------------------------------------------------------------------------------------------------------------
Public Enum COLOR
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E0E0
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
    锁色 = &HF5F5F5
    启用色 = 0
    停用色 = 255
    拖动色 = &HFFE0D9

End Enum
'----------------------------------------------------------------------------------------------------------------------
Public Enum REGISTER
    注册信息
    私有模块
    私有全局
    公共模块
    公共全局
End Enum

'内部应用模块号定义
Public Enum Enum_Inside_Program
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    p辅诊记录管理 = 1256
    p医嘱附费管理 = 1257
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
End Enum

'自定义类型定义
'----------------------------------------------------------------------------------------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    部门编码 As String
    部门名称 As String
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    数据库用户 As String
    模块权限 As String
    单位名称 As String
End Type
'----------------------------------------------------------------------------------------------------------------------
Public Type TYPE_ICONS_INFO

    入出单据 As String
    血液品种 As String
    血液规格 As String
    配血规则 As String
    血液价格 As String
    
End Type

'----------------------------------------------------------------------------------------------------------------------
Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤消出院 = 17            '允许撤消病人出院
    support必须录入入出诊断 = 18    '病人入院与出院时，必须录入诊断名
    support记帐完成后上传 = 19      '要求上传在记帐数据提交后再进行
    support出院结算必须出院 = 20    '病人结帐时如果选择出院结帐，就检查必须出院才可以进行
    
    support挂号使用个人帐户 = 21    '使用医保挂号时是否使用个人帐户进行支付

    support门诊连续收费 = 22        '门诊在身份验证后，可进行多次收费操作
    support门诊收费完成后验证 = 23  '在门诊收费完成，是否再次调用身份验证
    
    support医嘱上传 = 24            '医嘱产生费用时是否实时传输
    support分币处理 = 25            '医保病人是否处理分币
    support中途结算仅处理已上传部分 = 26 '提供对已上传部分数据的结算功能
    support允许冲销已结帐的记帐单据 = 27 '是否允许冲销记帐单据，如果该单据已经结帐
    
    support允许部份冲销单据 = 28
    support出院无实际交易 = 29 '出院接口中是否要与接口商进行交易
End Enum

'系统参数信息
'----------------------------------------------------------------------------------------------------------------------
Public Type SYSPARAM_INFO
    费用金额小数位数 As String
    收费诊疗项目匹配 As String
    结帐票据号长度 As Integer
    收费票据号长度 As Integer
    就诊卡号码长度 As Integer
    就诊卡字母前缀 As String
    就诊卡密文显示 As Boolean
    项目输入匹配方式 As Integer '0-双向;1-从左
    系统号 As Long
    系统名称 As String
    产品名称 As String
    模块号 As Long
    所有者 As String
    收费票种 As Integer
    结帐票种 As Integer
    结帐票号严格控制 As Boolean
    收费票号严格控制 As Boolean
    连接HIS报告 As Byte
End Type

'公共变量定义
'----------------------------------------------------------------------------------------------------------------------
Public ParamInfo As SYSPARAM_INFO
Public gobjKernel As New clsCISKernel       '临床核心部件
Public gobjRichEPR As New cRichEPR          '病历核心部件
Public IconInfo As TYPE_ICONS_INFO
Public UserInfo As TYPE_USER_INFO
Public gcolPrivs As Collection              '记录内部模块的权限

'医保变量
'----------------------------------------------------------------------------------------------------------------------
Public gclsInsure As New clsInsure
Public gblnInsure As Boolean '是否连接医保
Public gintInsure As Integer
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrSysName As String                '系统名称
Public glngModul As Long
Public gstrDBUser As String                 '当前数据库用户
Public gstrUnitName As String '用户单位名称
Public gfrmMain As Object
Public glngTXTProc As Long '保存默认的消息函数的地址
Public gstrSQL As String
Public gblnOK As Boolean
Public gblnShowInTaskBar As Boolean
Public glngOld As Long
Public glngFormW As Long
Public glngFormH As Long
Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Private mclsUnzip As New cUnzip

'自定义过程和函数
'######################################################################################################################

Public Sub CloseRecord(rs As ADODB.Recordset)
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Public Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
    '******************************************************************************************************************
    '功能: 装载数据入指定的组合下拉框或网格中的下拉框中
    '******************************************************************************************************************
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Public Function GetUserInfo() As Boolean
    '******************************************************************************************************************
    '功能：获取登陆用户信息
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        " Select A.ID,C.部门ID,A.编号,A.简码,A.姓名,B.用户名,D.编码,D.名称 " & _
        " From 人员表 A,上机人员表 B,部门人员 C,部门表 D " & _
        " Where A.ID = B.人员ID And A.ID = C.人员ID And C.缺省 = 1 AND C.部门id=D.ID And Upper(B.用户名) = USER And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOps")
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        UserInfo.部门编码 = IIf(IsNull(rsTmp!编码), "", rsTmp!编码)
        UserInfo.部门名称 = IIf(IsNull(rsTmp!名称), "", rsTmp!名称)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitSysPara() As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strSQL = "Select DECODE(参数值,NULL,缺省值,参数值) As 参数值 From 系统参数表 Where 参数号=[1]"
    
    '费用金额小数位数
    '------------------------------------------------------------------------------------------------------------------
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", 9)
    If rs.BOF = False Then
        strTmp = Val(zlCommFun.NVL(rs.Fields(0).Value, 2))
        If Val(strTmp) > 0 Then
            strTmp = "0." & String(Val(strTmp), "0")
        Else
            strTmp = "0"
        End If
        
        ParamInfo.费用金额小数位数 = strTmp
    End If
    
    '票据号长度
    '------------------------------------------------------------------------------------------------------------------
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", 20)
    If rs.BOF = False Then
        
        strTmp = zlCommFun.NVL(rs.Fields(0).Value, "")
        If UBound(Split(strTmp, "|")) >= 2 Then ParamInfo.结帐票据号长度 = Val(Split(strTmp, "|")(2))
        If UBound(Split(strTmp, "|")) >= 0 Then ParamInfo.收费票据号长度 = Val(Split(strTmp, "|")(0))
        If UBound(Split(strTmp, "|")) >= 4 Then ParamInfo.就诊卡号码长度 = Val(Split(strTmp, "|")(4))
    End If
    
    '票号严格控制
    '------------------------------------------------------------------------------------------------------------------
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", 24)
    If rs.BOF = False Then
        strTmp = zlCommFun.NVL(rs.Fields(0).Value, "")
        If UBound(Split(strTmp, "|")) >= 2 Then ParamInfo.结帐票号严格控制 = (Val(Split(strTmp, "|")(2)) = 1)
        If UBound(Split(strTmp, "|")) >= 0 Then ParamInfo.收费票号严格控制 = (Val(Split(strTmp, "|")(0)) = 1)
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '就诊卡字母前缀
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", 27)
    If rs.BOF = False Then
        ParamInfo.就诊卡字母前缀 = zlCommFun.NVL(rs.Fields(0).Value, "")
    End If
    
    InitSysPara = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '******************************************************************************************************************
    '检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    '******************************************************************************************************************
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function GetNextCode(ByVal strTable As String, Optional ByVal strField As String = "编码", Optional ByVal strFilter As String = "") As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strFormat As String

    GetNextCode = "1"
    strFormat = "00000000000000000000"
    gstrSQL = "select nvl(max(" & strField & "),0) as 编码 from " & strTable & IIf(strFilter = "", "", " where " & strFilter)

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps")

    If rs.BOF = False Then
        strFormat = IIf(rs!编码 = 0, "0000", Mid(strFormat, 1, Len(rs!编码)))
        GetNextCode = Format(rs!编码 + 1, strFormat)
    End If
    CloseRecord rs
End Function

Public Function CalcStorage(ByVal lng药品id As Long, ByVal lng库房ID As Long, ByVal vChangePrice As Boolean, ByVal vBatch As Boolean) As Single
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset

    If lng药品id = 0 Then Exit Function

    If vChangePrice And vBatch = False Then
        '只是实价药品

        gstrSQL = "SELECT NVL(A.可用数量,0) AS 可用数量 FROM 药品库存 A WHERE A.药品id=[1] AND A.库房ID=[2]"

    ElseIf vChangePrice = False And vBatch Then
        '只是药房分批核算药品

        gstrSQL = "Select Sum(Nvl(可用数量,0)) as 可用数量 From 药品库存" & _
                    " Where 性质=1 " & _
                    " And (效期 Is NULL Or 效期>Trunc(Sysdate)) " & _
                    " And 库房ID=[2]" & _
                    " And 药品ID=[1]"

    ElseIf vChangePrice And vBatch Then
        '既是实价药品又是药房分批核算药品

        gstrSQL = "Select Sum(Nvl(可用数量,0)) as 可用数量 From 药品库存" & _
                    " Where 性质=1 " & _
                    " And (效期 Is NULL Or 效期>Trunc(Sysdate)) " & _
                    " And 库房ID=[2]" & _
                    " And 药品ID=[1]"

    Else
        '既不是实价药品又不是药房分批核算药品,和只是实价药品一样的

        gstrSQL = "SELECT NVL(A.可用数量,0) AS 可用数量 FROM 药品库存 A WHERE A.药品id=[1] AND A.库房ID=[2]"

    End If

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng药品id, lng库房ID)

    If rs.BOF = False Then CalcStorage = zlCommFun.NVL(rs("可用数量").Value, 0)

    CloseRecord rs
End Function

Public Function CheckAllNumber(ByVal strKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long

    For lngLoop = 1 To Len(strKey)
        If Mid(strKey, lngLoop, 1) < "0" Or Mid(strKey, lngLoop, 1) > "9" Then
            Exit Function
        End If
    Next

    CheckAllNumber = True
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '******************************************************************************************************************
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    If bytIsWB Then
        strSQL = "select zlWBcode('" & strInput & "') from dual"
    Else
        strSQL = "select zlSpellcode('" & strInput & "') from dual"
    End If
    On Error GoTo errHand
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function CheckHaveOrder(ByVal lngKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset

    gstrSQL = "SELECT 医嘱状态 FROM 病人医嘱记录 WHERE ID=[1]"

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)

    CheckHaveOrder = (rs.BOF = False)
    If rs.BOF = False Then
        CheckHaveOrder = (rs("医嘱状态").Value <> 4)
    End If

    CloseRecord rs
End Function

Public Function CheckAllowAudit(ByVal lngKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "SELECT 1 FROM 病人医嘱发送 WHERE 执行状态>0 AND 医嘱ID=[1]"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
    
    CheckAllowAudit = (rs.BOF = True)
    If CheckAllowAudit = False Then
        MsgBox "手术医嘱已经发送并且正在执行或已经执行完成！", vbInformation, gstrSysName
    End If
    
    CloseRecord rs
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'功能：判断是否存在指定处方单据类型的入出类别
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH

    strSQL = "Select 类别ID From 药品单据性质 Where 单据=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", bytBill)

    If Not rsTmp.EOF Then ExistIOClass = zlCommFun.NVL(rsTmp!类别ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ActualMoney(费别 As String, 收入项目ID As Long, 金额 As Currency) As Currency
'功能：根据费别,收入项目ID,金额,求打折后的金额
'说明：金额折扣范围取绝对值范围
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    ActualMoney = 金额
    If 费别 = "" Or 金额 = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = _
        "Select " & 金额 & "*实收比率/100 as 金额 From 费别明细" & _
        " Where 收入项目ID=[1] And 费别=[2]" & _
        " And [3] Between 应收段首值 and 应收段尾值"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", 收入项目ID, 费别, Abs(金额))
    If Not rsTmp.EOF Then ActualMoney = rsTmp!金额
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckChargeState(ByVal lng医嘱id As Long, ByVal lng发送号 As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '收费状态
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    CheckChargeState = False
    
    strSQL = _
        "select NVL(COUNT(1), 0) AS 计数 " & _
              "from 病人费用记录 A, " & _
              "( " & _
                   "select no from 病人医嘱发送 where 医嘱id+0=" & lng医嘱id & " and 发送号=[1] " & _
                   "Union " & _
                   "select no from 病人医嘱附费 where 医嘱id=" & lng医嘱id & " and 发送号=[1] " & _
              ") B " & _
            "Where A.NO = B.NO AND NVL(A.记录状态,0)=0"
    
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", lng发送号)
    
    If rs.BOF Then Exit Function
    If rs("计数").Value > 0 Then Exit Function
    
    CheckChargeState = True
    
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
    '******************************************************************************************************************
    '功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
    '参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    '******************************************************************************************************************
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
        End If
    End If
    FormatEx = strNumber
End Function


Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '******************************************************************************************************************
    '功能：
    '******************************************************************************************************************
    MsgBox strInfo, vbInformation, ParamInfo.系统名称
    
End Sub

Public Function ExecutePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    
    Select Case Control.ID
    Case conMenu_File_PrintSet '打印设置
    
        Call zlPrintSet
        
    Case conMenu_View_ToolBar_Button '工具栏
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text '按钮文字
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size '大图标
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar '状态栏
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_Help_Web_Home 'Web上的中联
        
        Call zlHomePage(frmMain.hWnd)
        
    Case conMenu_Help_Web_Mail '发送反馈
        
        Call zlMailTo(frmMain.hWnd)
            
    Case conMenu_Help_About '关于
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    
    Case conMenu_File_Exit '退出
        Unload frmMain
            
    End Select
    
    ExecutePublic = True
End Function

Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function

Public Sub LocationObj(ByRef objTxt As Object)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error Resume Next
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
End Sub

Public Sub LocationGrid(ByRef vsf As Object, Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error Resume Next
    
    If lngRow <> -1 Then vsf.Row = lngRow
    If lngCol <> -1 Then vsf.Col = lngCol
    
    vsf.SetFocus
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "0") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strFormat As String
    Dim lngNotPrintCols As Long
    Dim lngPrintCol As Long
    
    If strNotPrintCol <> "" Then
        lngNotPrintCols = UBound(Split(strNotPrintCol, ",")) + 1
        strNotPrintCol = "," & strNotPrintCol & ","
    End If
    
    objPrintVsf.Rows = objVsf.Rows
    objPrintVsf.Cols = objVsf.Cols - lngNotPrintCols
    objPrintVsf.FixedRows = objVsf.FixedRows
    
    lngPrintCol = -1
    For lngCol = 0 To objVsf.Cols - 1
        
        If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
            lngPrintCol = lngPrintCol + 1
            objPrintVsf.ColWidth(lngPrintCol) = objVsf.ColWidth(lngCol)
            objPrintVsf.ColAlignmentFixed(lngPrintCol) = objVsf.ColAlignment(lngCol)
            If objVsf.ColDataType(lngCol) = flexDTBoolean Then
                objPrintVsf.ColAlignment(lngPrintCol) = 4
            Else
                objPrintVsf.ColAlignment(lngPrintCol) = objVsf.ColAlignment(lngCol)
            End If
        End If
    Next
    
    
    For lngRow = 0 To objVsf.Rows - 1

        objPrintVsf.RowHeight(lngRow) = IIf(objVsf.RowHeight(lngRow) < objVsf.RowHeightMin, objVsf.RowHeightMin, objVsf.RowHeight(lngRow))
        lngPrintCol = -1
        For lngCol = 0 To objVsf.Cols - 1
            
            If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                lngPrintCol = lngPrintCol + 1
                
                If objVsf.ColDataType(lngCol) = flexDTBoolean And lngRow >= objVsf.FixedRows Then
                    objPrintVsf.TextMatrix(lngRow, lngPrintCol) = IIf(Abs(Val(objVsf.TextMatrix(lngRow, lngCol))) = 1, "√", "")
                Else
                    strFormat = objVsf.ColFormat(lngCol)
                    If strFormat = "" Then
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Trim(objVsf.TextMatrix(lngRow, lngCol))
                    Else
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Format(objVsf.TextMatrix(lngRow, lngCol), strFormat)
                    End If
                End If
            End If
        Next
        Call SetMsfForeColor(objPrintVsf, lngRow, Val(objVsf.Cell(flexcpForeColor, lngRow, 1)))
    Next
End Function

Public Sub SetMsfForeColor(ByRef msf As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    '******************************************************************************************************************
    '
    '******************************************************************************************************************
    Dim intCol As Integer
    
    With msf
        
        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = lngColor
        Next

    End With
End Sub

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '功能:获取特殊时间
    '参数:
    '******************************************************************************************************************
    Dim intDay As Integer
    
    Select Case strMode
    Case "当  时"      '当时
        GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(zlDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(zlDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前二年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365 * 2, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim strChar As String
    
    strChar = "ZXCVBNMASDFGHJKLQWERTYUIOPzxcvbnmasdfghjklqwertyuiop"
    
    Select Case bytMode
    Case 1          '全数字
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
    Case 2          '全字母
    
        For lngLoop = 1 To Len(Text)
            If InStr(strChar, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
        
    Case 99
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
    End Select
End Function

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String) As Long
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    
    gstrSQL = "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical")
    GetMaxLength = rs.Fields(0).DefinedSize

End Function

Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '功能： 将指定的信息保存在注册表中
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strKeyValue-键值
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        Call SaveSetting("ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue)
        
    Case 私有模块

        Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case 私有全局

        Call SaveSetting("ZLSOFT", "私有全局\" & UserInfo.用户名 & "\" & strSection, strKey, strKeyValue)
        
    Case 公共模块

        Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case 公共全局
        
        Call SaveSetting("ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue)
        
    End Select
    
    SetRegister = True
    
errHand:
    
End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '功能： 将指定的注册信息读取出来
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strDefKeyValue-缺省键值
    '返回： strKeyValue-键值
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        strValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, strDefKeyValue)
        
    Case 私有模块

        strValue = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 私有全局

        strValue = GetSetting("ZLSOFT", "私有全局\" & UserInfo.用户名 & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共模块

        strValue = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共全局
        
        strValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function ShowPubSelect(ByVal frmParent As Object, _
                                ByVal obj As Object, _
                                ByVal bytStyle As Byte, _
                                ByVal strLvw As String, _
                                ByVal strSavePath As String, _
                                ByVal strDescrible As String, _
                                ByVal rsData As ADODB.Recordset, _
                                ByRef rsResult As ADODB.Recordset, _
                                Optional ByVal lngCX As Long = 9000, _
                                Optional ByVal lngCY As Long = 4500, _
                                Optional ByVal blnMuliSel As Boolean = False, _
                                Optional ByVal strInitKey As String = "", _
                                Optional ByVal strFilterControl As String = "", _
                                Optional ByVal blnLeftSelect As Boolean = False) As Byte
    '******************************************************************************************************************
    '功能：打开树型+列表结构,应用于表格控件
    '参数：
    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
    '返回：0:取消选择;1:选择;2:无数据返回
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim lngObjHeight As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI

    On Error GoTo errHand
    
    If rsData.BOF Then
        ShowPubSelect = 2
        Exit Function
    End If
    
    If obj Is Nothing Then
        lngX = (Screen.Width - lngCX) / 2
        lngY = (Screen.Width - lngCY) / 2
        lngObjHeight = 0
    Else
        Call ClientToScreen(obj.hWnd, objPoint)
        
        Select Case TypeName(obj)
        Case "TextBox", "CommandButton"
        
            lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
            lngY = obj.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
            lngObjHeight = obj.Height
            
        Case Else
            lngX = objPoint.X * Screen.TwipsPerPixelX + obj.CellLeft
            lngY = objPoint.Y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
            lngObjHeight = obj.CellHeight
        End Select
    End If
    
    ShowPubSelect = frmPubSelDialog.ShowDialog(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, blnLeftSelect, False, blnMuliSel, strFilterControl)
                                
    If ShowPubSelect = 1 Then
        Set rsResult = rsData
        
        If rsResult.BOF Then
            ShowPubSelect = 0
        End If
        
    End If

    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetApplyMode(ByVal strText As String) As Byte
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    If CheckStrType(strText, 1) And Left(ParamInfo.收费诊疗项目匹配, 1) = 1 Then
        '是全数字，按编码查找
            
        GetApplyMode = 1
        
    ElseIf CheckStrType(strText, 2) And Left(ParamInfo.收费诊疗项目匹配, 2) = 1 Then
        '是全字母，按简码查找
        
        GetApplyMode = 2
    Else
        GetApplyMode = 3
    End If
End Function


Public Function AppendCode(ByVal strName As String, ByVal strCode As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If strName <> "" And strCode <> "" Then
        AppendCode = "【" & strCode & "】" & strName
    Else
        AppendCode = strName
    End If
End Function

Public Function PromptStorageWarn(ByVal dbInput As Double, _
                                    ByVal dbStorage As Double, _
                                    ByVal strDrugName As String, _
                                    ByVal strExecuteDept As String, _
                                    ByVal strUnit As String, _
                                    Optional ByVal bytWarn As Byte = 1, _
                                    Optional ByVal bytApply As Byte = 1) As Integer
    '******************************************************************************************************************
    '功能：
    '参数：bytWarn：0-不检查;1-检查,不足提醒;2-检查，不足禁
    '返回：
    '******************************************************************************************************************

    If dbInput > 0 And dbInput > dbStorage Then
        
        If bytApply = 1 Then
            Call ShowSimpleMsg("药品“" & strDrugName & "”在库房“" & strExecuteDept & "”只有" & dbStorage & strUnit & "！")
            bytWarn = 0
        Else
            Select Case bytWarn
            Case 0
                
            Case 1
                If MsgBox("药品“" & strDrugName & "”在库房“" & strExecuteDept & "”只有" & dbStorage & strUnit & "，是否继续？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                    bytWarn = 0
                Else
                    bytWarn = 1
                End If
            Case 2
                MsgBox "药品“" & strDrugName & "”在库房“" & strExecuteDept & "”只有" & dbStorage & strUnit & "，不足禁止！", vbOKOnly + vbCritical, ParamInfo.系统名称
                bytWarn = 1
            End Select
        End If
        
    End If
    
    PromptStorageWarn = bytWarn
    
End Function

Public Function BillExistBalance(ByVal strNO As String) As Boolean
    '******************************************************************************************************************
    '功能：判断指定的收费划价单是否存在已经收费的内容
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH

    strSQL = "Select ID From 病人费用记录 Where 记录性质=1 And 记录状态 IN(1,3) And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO)

    BillExistBalance = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Between(X, a, b) As Boolean
    '******************************************************************************************************************
    '功能：判断x是否在a和b之间
    '******************************************************************************************************************
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

Public Function IntEx(vNumber As Variant) As Variant
    '******************************************************************************************************************
    '功能：取大于指定数值的最小整数
    '******************************************************************************************************************
    
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function


Public Function GetDrugWarnOption(ByVal lngKey As Long, ByVal str类别 As String) As Integer
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    If str类别 = "4" Then
        gstrSQL = "SELECT 检查方式 FROM 材料出库检查 WHERE 库房ID=[1]"
    Else
        gstrSQL = "SELECT 检查方式 FROM 药品出库检查 WHERE 库房ID=[1]"
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)

    If rs.BOF = False Then
        GetDrugWarnOption = Val(IIf(IsNull(rs("检查方式").Value), 0, rs("检查方式").Value))
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function CalcTimePrice(ByVal lng药品id As Long, lng药房ID As Long, ByVal sng数量 As Single) As Currency
    '******************************************************************************************************************
    '功能：计算实价药品的实际出库价
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim sng住院包装 As Single, sng出库数量 As Single
    Dim cur指导零售价 As Currency, cur出库金额 As Currency
    
    sng出库数量 = sng数量

    gstrSQL = "Select Nvl(批次,0) as 批次,Nvl(可用数量,0) as 库存," & _
        " Nvl(Decode(Nvl(实际数量,0),0,0,实际金额/实际数量),0) as 时价" & _
        " From 药品库存" & _
        " Where 性质=1 And 库房ID=[2] And 药品ID=[1]" & _
        " And (Nvl(批次,0)=0 Or 效期 is NULL Or 效期>Trunc(Sysdate))" & _
        " Order by Nvl(批次,0)"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng药品id, lng药房ID)
    If Not rsTmp.EOF Then
        While sng数量 > 0 And Not rsTmp.EOF
            If rsTmp!库存 > sng数量 Then
                cur出库金额 = cur出库金额 + Format(sng数量 * Format(rsTmp!时价, "0.0000"), "0.00")
                sng数量 = 0
            Else
                cur出库金额 = cur出库金额 + Format(rsTmp!库存 * Format(rsTmp!时价, "0.0000"), "0.00")
                sng数量 = sng数量 - rsTmp!库存
            End If
            rsTmp.MoveNext
        Wend
        If sng数量 <= 0 Then
            If sng出库数量 <> 0 Then
                CalcTimePrice = Format(cur出库金额 / sng出库数量, "0.0000")
            Else
                CalcTimePrice = 0 '库存为0
            End If
        Else
            CalcTimePrice = 0 '库存不够
        End If
    End If

    CloseRecord rsTmp
End Function

Public Function GetWarnGrade(ByVal WarnGraded As Long, ByVal FeeClass As String, ByVal str报警方案 As String, ByVal lng病区id As Long) As Long
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    GetWarnGrade = 0
    gstrSQL = "select MAX(级别) as 级别 FROM ("
    gstrSQL = gstrSQL & "select 1 AS 级别 from 记帐报警线 where (报警标志1 like [3] or 报警标志1='-') And 适用病人=[2] AND 病区id=[1]"
    gstrSQL = gstrSQL & " union select 2 AS 级别 from 记帐报警线 where (报警标志2 like [3] or 报警标志2='-') And 适用病人=[2] AND 病区id=[1]"
    gstrSQL = gstrSQL & " union select 3 AS 级别 from 记帐报警线 where (报警标志3 like [3] or 报警标志3='-') And 适用病人=[2] AND 病区id=[1]" & _
        ") A"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病区id, str报警方案, "%" & FeeClass & "%")
    
    If rs.BOF = False Then GetWarnGrade = IIf(WarnGraded > zlCommFun.NVL(rs!级别, 0), WarnGraded, zlCommFun.NVL(rs!级别, 0))
    
End Function

Public Function 欠费情况(str姓名 As String, lng病人id As Long, lng主页id As Long, Optional ByVal curMoney As Single = 0, Optional ByVal str报警方案 As String, Optional ByVal int报警方式 As Long, Optional ByVal bln强制记帐 As Boolean, Optional str强制报警姓名 As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset, strError As String
    Dim int报警标志 As Integer, int报警方法 As Integer, sng报警值 As Single
    Dim sng剩余总额 As Single, sng发生费用 As Single, sng担保额 As Single
    
    欠费情况 = "未知"
        
    gstrSQL = "Select 报警方法,报警值 From 记帐报警线 A,病案主页 B Where A.适用病人=[3] And A.病区ID = B.当前病区ID And B.病人id =[1] And B.主页id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病人id, lng主页id, str报警方案)
    If rsTmp.BOF Then Exit Function
    sng报警值 = IIf(IsNull(rsTmp!报警值), 0, rsTmp!报警值)
    int报警方法 = IIf(IsNull(rsTmp!报警方法), 0, rsTmp!报警方法)
    int报警标志 = int报警方式
    
    gstrSQL = "Select 担保额 From 病人信息 A Where A.病人ID =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病人id)
    If Not rsTmp.BOF Then sng担保额 = zlCommFun.NVL(rsTmp!担保额, 0)
    
    Select Case int报警方法
    Case 1 '累计费用
        Set rsTmp = Get费用信息(lng病人id, lng主页id)
        If Not rsTmp.EOF Then sng剩余总额 = zlCommFun.NVL(rsTmp!剩余总额, 0)
        sng剩余总额 = sng担保额 + sng剩余总额 - curMoney
                
        Select Case int报警标志
        Case 1
            If sng剩余总额 < sng报警值 Then
                If bln强制记帐 Then
                    欠费情况 = "强制"
                    strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")已到报警值(" & FormatEx(sng报警值, 2) & ")！"
                Else
                    欠费情况 = "记帐"
                    strError = str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")已到报警值(" & FormatEx(sng报警值, 2) & ")，还需要记帐吗？"
                End If
                GoTo EndPoint
            End If
        Case 2
            If sng剩余总额 <= 0 Then
                If bln强制记帐 Then
                    欠费情况 = "强制"
                    strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "预交余额已经用完！"
                Else
                    欠费情况 = "是"
                    strError = str姓名 & "预交余额已经用完，禁止记帐！"
                End If
                GoTo EndPoint
            End If
            
            If sng剩余总额 < sng报警值 Then
                If bln强制记帐 Then
                    欠费情况 = "强制"
                    strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")小于了报警值(" & FormatEx(sng报警值, 2) & ")！"
                Else
                    欠费情况 = "记帐"
                    strError = str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")小于了报警值(" & FormatEx(sng报警值, 2) & ")，还需要记帐吗？"
                End If
                GoTo EndPoint
            End If
        Case 3
            If sng剩余总额 < sng报警值 Then
                If bln强制记帐 Then
                    欠费情况 = "强制"
                    strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")小于了报警值(" & FormatEx(sng报警值, 2) & ")！"
                Else
                    欠费情况 = "是"
                    strError = str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")小于了报警值(" & FormatEx(sng报警值, 2) & ")，禁止记帐！"
                End If
                GoTo EndPoint
            End If
        End Select
    Case 2              '每日费用
        gstrSQL = "select sum(实收金额) as 发生费用 from 病人费用记录 where 病人id=[1] and 主页id=[2] and trunc(发生时间)=[3] "
        
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病人id, lng主页id, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If rsTmp.BOF = False Then
            sng发生费用 = zlCommFun.NVL(rsTmp!发生费用, 0) + curMoney
            Select Case int报警标志
            Case 1
                If sng发生费用 > sng报警值 Then
                    If bln强制记帐 Then
                        欠费情况 = "强制"
                        strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "今天的发生费用(" & FormatEx(sng发生费用, 2) & ")已经超过了报警值(" & FormatEx(sng报警值, 2) & ")！"
                    Else
                        欠费情况 = "记帐"
                        strError = str姓名 & "今天的发生费用(" & FormatEx(sng发生费用, 2) & ")已经超过了报警值(" & FormatEx(sng报警值, 2) & "),还需要记帐吗？"
                    End If
                    GoTo EndPoint
                End If
            Case 3
                If sng发生费用 > sng报警值 Then
                    If bln强制记帐 Then
                        欠费情况 = "强制"
                        strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "今天的发生费用(" & FormatEx(sng发生费用, 2) & ")已经超过了报警值(" & FormatEx(sng报警值, 2) & ")！"
                    Else
                        欠费情况 = "是"
                        strError = str姓名 & "今天的发生费用(" & FormatEx(sng发生费用, 2) & ")已经超过了报警值(" & FormatEx(sng报警值, 2) & "),禁止记帐！"
                    End If
                    GoTo EndPoint
                End If
            End Select
        End If
    End Select
    Exit Function
EndPoint:
    If 欠费情况 = "是" Then
        MsgBox strError, vbInformation, gstrSysName
    ElseIf 欠费情况 = "强制" Then
        欠费情况 = "记帐"
        If InStr(str强制报警姓名 & ";", ";" & str姓名 & ";") = 0 Then
            str强制报警姓名 = str强制报警姓名 & ";" & str姓名
            MsgBox strError, vbInformation, gstrSysName
        End If
    Else
        If MsgBox(strError, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then 欠费情况 = "是"
    End If
End Function

'Public Function GetWarnGrade(ByVal WarnGraded As Long, ByVal FeeClass As String, ByVal bln医保 As Boolean, ByVal lng病区id As Long) As Long
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim rs As New ADODB.Recordset
'
'    GetWarnGrade = 0
'    gstrSQL = "select MAX(级别) as 级别 FROM ("
'    gstrSQL = gstrSQL & "select 1 AS 级别 from 记帐报警线 where (报警标志1 like [3] or 报警标志1='-') And 适用病人=[2] AND 病区id=[1]"
'    gstrSQL = gstrSQL & " union select 2 AS 级别 from 记帐报警线 where (报警标志2 like [3] or 报警标志2='-') And 适用病人=[2] AND 病区id=[1]"
'    gstrSQL = gstrSQL & " union select 3 AS 级别 from 记帐报警线 where (报警标志3 like [3] or 报警标志3='-') And 适用病人=[2] AND 病区id=[1]" & _
'        ") A"
'
'    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病区id, IIf(bln医保, 2, 1), "%" & FeeClass & "%")
'
'    If rs.BOF = False Then GetWarnGrade = IIf(WarnGraded > zlCommFun.NVL(rs!级别, 0), WarnGraded, zlCommFun.NVL(rs!级别, 0))
'
'End Function
'
'Public Function 欠费情况(str姓名 As String, lng病人id As Long, lng主页id As Long, _
'    Optional ByVal curMoney As Single = 0, Optional bln医保 As Boolean, Optional ByVal int报警方式 As Long, _
'    Optional ByVal bln强制记帐 As Boolean, Optional str强制报警姓名 As String) As String
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim rsTmp As New ADODB.Recordset, strError As String
'    Dim int报警标志 As Integer, int报警方法 As Integer, sng报警值 As Single
'    Dim sng剩余总额 As Single, sng发生费用 As Single, sng担保额 As Single
'
'    欠费情况 = "未知"
'
'    gstrSQL = "Select 报警方法,报警值 From 记帐报警线 A,病案主页 B Where A.适用病人=[3] And A.病区ID = B.当前病区ID And B.病人id =[1] And B.主页id = [2]"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病人id, lng主页id, IIf(bln医保, 2, 1))
'    If rsTmp.BOF Then Exit Function
'    sng报警值 = IIf(IsNull(rsTmp!报警值), 0, rsTmp!报警值)
'    int报警方法 = IIf(IsNull(rsTmp!报警方法), 0, rsTmp!报警方法)
'    int报警标志 = int报警方式
'
'    gstrSQL = "Select 担保额 From 病人信息 A Where A.病人ID =[1]"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病人id)
'    If Not rsTmp.BOF Then sng担保额 = zlCommFun.NVL(rsTmp!担保额, 0)
'
'    Select Case int报警方法
'    Case 1 '累计费用
'        Set rsTmp = Get费用信息(lng病人id, lng主页id)
'        If Not rsTmp.EOF Then sng剩余总额 = zlCommFun.NVL(rsTmp!剩余总额, 0)
'        sng剩余总额 = sng担保额 + sng剩余总额 - curMoney
'
'        Select Case int报警标志
'        Case 1
'            If sng剩余总额 < sng报警值 Then
'                If bln强制记帐 Then
'                    欠费情况 = "强制"
'                    strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")已到报警值(" & FormatEx(sng报警值, 2) & ")！"
'                Else
'                    欠费情况 = "记帐"
'                    strError = str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")已到报警值(" & FormatEx(sng报警值, 2) & ")，还需要记帐吗？"
'                End If
'                GoTo EndPoint
'            End If
'        Case 2
'            If sng剩余总额 <= 0 Then
'                If bln强制记帐 Then
'                    欠费情况 = "强制"
'                    strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "预交余额已经用完！"
'                Else
'                    欠费情况 = "是"
'                    strError = str姓名 & "预交余额已经用完，禁止记帐！"
'                End If
'                GoTo EndPoint
'            End If
'
'            If sng剩余总额 < sng报警值 Then
'                If bln强制记帐 Then
'                    欠费情况 = "强制"
'                    strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")小于了报警值(" & FormatEx(sng报警值, 2) & ")！"
'                Else
'                    欠费情况 = "记帐"
'                    strError = str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")小于了报警值(" & FormatEx(sng报警值, 2) & ")，还需要记帐吗？"
'                End If
'                GoTo EndPoint
'            End If
'        Case 3
'            If sng剩余总额 < sng报警值 Then
'                If bln强制记帐 Then
'                    欠费情况 = "强制"
'                    strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")小于了报警值(" & FormatEx(sng报警值, 2) & ")！"
'                Else
'                    欠费情况 = "是"
'                    strError = str姓名 & "剩余款总额(" & FormatEx(sng剩余总额, 2) & ")小于了报警值(" & FormatEx(sng报警值, 2) & ")，禁止记帐！"
'                End If
'                GoTo EndPoint
'            End If
'        End Select
'    Case 2              '每日费用
'        gstrSQL = "select sum(实收金额) as 发生费用 from 病人费用记录 where 病人id=[1] and 主页id=[2] and trunc(发生时间)=[3] "
'
'        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病人id, lng主页id, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
'
'        If rsTmp.BOF = False Then
'            sng发生费用 = zlCommFun.NVL(rsTmp!发生费用, 0) + curMoney
'            Select Case int报警标志
'            Case 1
'                If sng发生费用 > sng报警值 Then
'                    If bln强制记帐 Then
'                        欠费情况 = "强制"
'                        strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "今天的发生费用(" & FormatEx(sng发生费用, 2) & ")已经超过了报警值(" & FormatEx(sng报警值, 2) & ")！"
'                    Else
'                        欠费情况 = "记帐"
'                        strError = str姓名 & "今天的发生费用(" & FormatEx(sng发生费用, 2) & ")已经超过了报警值(" & FormatEx(sng报警值, 2) & "),还需要记帐吗？"
'                    End If
'                    GoTo EndPoint
'                End If
'            Case 3
'                If sng发生费用 > sng报警值 Then
'                    If bln强制记帐 Then
'                        欠费情况 = "强制"
'                        strError = "强制记帐提醒：" & vbCrLf & vbTab & str姓名 & "今天的发生费用(" & FormatEx(sng发生费用, 2) & ")已经超过了报警值(" & FormatEx(sng报警值, 2) & ")！"
'                    Else
'                        欠费情况 = "是"
'                        strError = str姓名 & "今天的发生费用(" & FormatEx(sng发生费用, 2) & ")已经超过了报警值(" & FormatEx(sng报警值, 2) & "),禁止记帐！"
'                    End If
'                    GoTo EndPoint
'                End If
'            End Select
'        End If
'    End Select
'    Exit Function
'EndPoint:
'    If 欠费情况 = "是" Then
'        MsgBox strError, vbInformation, gstrSysName
'    ElseIf 欠费情况 = "强制" Then
'        欠费情况 = "记帐"
'        If InStr(str强制报警姓名 & ";", ";" & str姓名 & ";") = 0 Then
'            str强制报警姓名 = str强制报警姓名 & ";" & str姓名
'            MsgBox strError, vbInformation, gstrSysName
'        End If
'    Else
'        If MsgBox(strError, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then 欠费情况 = "是"
'    End If
'End Function

Private Function Get费用信息(lngID As Long, Optional ByVal lngPageID As Long = 0) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：获取指定病人的剩余额
    '******************************************************************************************************************
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If lngPageID = 0 Then
        strSQL = "Select Nvl(A.费用余额,0) as 费用余额,Nvl(A.预交余额,0) as 预交余额,Nvl(A.预交余额,0)-Nvl(A.费用余额,0) AS 剩余总额 " & _
                "From 病人余额 A Where A.性质=1 And A.病人ID=[1]"
    Else
        strSQL = "Select Nvl(A.费用余额,0) as 费用余额,Nvl(A.预交余额,0) as 预交余额,Nvl(A.预交余额,0)-Nvl(A.费用余额,0) + Nvl(B.金额,0) AS 剩余总额 " & _
                "From 病人余额 A,(SELECT nvl(SUM(金额),0) as 金额 from 保险模拟结算 where 病人id=[1] AND 主页id=[2]) B Where A.性质=1 And A.病人ID=[1]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", lngID, lngPageID)
    
    Set Get费用信息 = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MakeChargeBill(ByVal lngKey As Long, ByVal int记录性质 As Integer, ByVal strMenuItem As String, Optional ByVal blnZeroBill As Boolean = False, Optional ByVal strPrivs As String) As String
    '******************************************************************************************************************
    '功能：从用药和材料中生成附加费用
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsPati As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strNO As String
    Dim int来源 As Integer
            
    Dim lng医嘱id As Long
    Dim int父号 As Integer
    Dim lng项目ID As Long
    Dim lng执行部门ID As Long
    Dim lng病人病区ID As Long
    Dim lng病人科室ID As Long
    Dim lng类别ID As Long
    Dim strDate As String
    Dim lngLoop As Long
    Dim int保险项目否 As Integer
    Dim lng保险大类ID As Long
    Dim str保险编码 As String
    Dim cur统筹金额 As Currency
    Dim cur应收 As Currency
    Dim cur实收 As Currency
    Dim strMsg As String
    Dim dbl数量 As Double
    Dim blnTran As Boolean
    Dim cur单价 As Currency
    Dim lng报警级别 As Long
    Dim str报警方案 As String
    Dim lng已报警级别 As Long
    Dim lng级别 As Long
    Dim str已强制报警姓名 As String
    Dim bln医保 As Boolean
    Dim curMoneyTotal As Currency
    Dim str费用小数位 As String
    Dim strSQL As String
    Dim rsSQL As ADODB.Recordset
    Dim bln强制记帐 As Boolean
    Dim lng病人id As Long
    Dim lng主页id As Long
    Dim lng发送号 As Long
    
    On Error GoTo errHand
    
    Screen.MousePointer = 11
    
    Call SQLRecord(rsSQL)
    
    '初始设置
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select a.病人id,a.主页id,a.病人来源,b.发送号 From 病人医嘱记录 a,病人医嘱发送 b Where a.ID=[1] And a.ID=b.医嘱id"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
    If rs.BOF Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    lng病人id = rs("病人id").Value
    lng主页id = zlCommFun.NVL(rs("主页id").Value, 0)
    int来源 = rs("病人来源").Value
    lng发送号 = zlCommFun.NVL(rs("发送号").Value, 0)
    
    '取费用金额保存小数
    '------------------------------------------------------------------------------------------------------------------
    str费用小数位 = ParamInfo.费用金额小数位数
    bln强制记帐 = (InStr(strPrivs, "欠费强制记帐") > 0)
    
    '获取病人的信息
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select A.姓名,A.性别,A.年龄,Nvl(B.费别,A.费别) as 费别," & _
        " A.门诊号,A.住院号,Nvl(A.当前床号,B.出院病床) as 床号," & _
        " Nvl(A.当前病区ID,B.当前病区ID) as 病人病区ID," & _
        " Nvl(A.当前科室ID,B.出院科室ID) as 病人科室ID," & _
        " Nvl(B.险类,A.险类) as 险类,C.编码 as 付款码" & _
        " From 病人信息 A,病案主页 B,医疗付款方式 C" & _
        " Where A.病人ID=[1] And A.病人ID=B.病人ID(+)" & _
        " And B.主页ID(+)=[2] And A.医疗付款方式=C.名称(+)"
    
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病人id, lng主页id)

    If rsPati.BOF Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    bln医保 = (Val(zlCommFun.NVL(rsPati!付款码, "0")) = 1)
    
    '可能对照费用为药品费用
    '------------------------------------------------------------------------------------------------------------------
    lng类别ID = ExistIOClass(IIf(int记录性质 = 1, 8, 9)) '8:门诊划价单;9:门诊/住院记帐单
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    gstrSQL = "SELECT B.ID AS 收费细目ID," & _
                  "A.数量,A.可否分零,A.剂量系数,A.包装," & _
                  "B.计算单位," & _
                  "B.类别," & _
                  "C.现价 AS 单价," & _
                  "D.收据费目," & _
                  "C.收入项目ID," & _
                  "A.执行科室id," & _
                  "DECODE(A.主页id,NULL,F.门诊号,0,F.门诊号,F.住院号) AS 标识号," & _
                  "F.费别," & _
                  "A.病人科室id AS 当前科室ID," & _
                  "DECODE(F.当前病区ID,NULL,A.病人科室id,F.当前病区ID) AS 当前病区ID," & _
                  "F.当前床号," & _
                  "A.病人ID," & _
                  "A.主页id," & _
                  "F.姓名," & _
                  "F.性别," & _
                  "F.年龄," & _
                  "B.名称 " & _
            "FROM   收费项目目录 B," & _
               "收费价目 C," & _
               "收入项目 D," & _
               "病人信息 F," & _
               "("
    
    Select Case strMenuItem
    '------------------------------------------------------------------------------------------------------------------
    Case "治疗"
    
        gstrSQL = gstrSQL & _
            "SELECT HH.可否分零,Decode(HH.剂量系数,0,1,Null,1,HH.剂量系数) As 剂量系数,Decode(GG.病人来源,2,HH.住院包装,HH.门诊包装) As 包装,GG.病人科室id,3 AS 序号,AA.收费细目id,AA.数量,AA.执行科室id,GG.病人id,GG.主页id ,0 AS 单价 " & _
            "FROM 病人手术计价 AA,病人手术记录 BB,药品规格 HH,病人医嘱记录 GG " & _
            "Where AA.收费细目ID = HH.药品id(+) And AA.记录id = BB.ID And BB.医嘱id = GG.ID And BB.医嘱id=[1]"
    '------------------------------------------------------------------------------------------------------------------
    Case "用药"
    
        gstrSQL = gstrSQL & _
            "SELECT HH.可否分零,Decode(HH.剂量系数,0,1,Null,1,HH.剂量系数) As 剂量系数,Decode(GG.病人来源,2,HH.住院包装,HH.门诊包装) As 包装,GG.病人科室id,1 AS 序号,AA.药品id AS 收费细目id,AA.使用总量 AS 数量,AA.执行科室id,BB.病人id,BB.主页id ,0 AS 单价 " & _
            "FROM 病人手术用药 AA,病人手术记录 BB,药品规格 HH,病人医嘱记录 GG " & _
            "Where AA.药品id = HH.药品id And AA.记录id = BB.ID And BB.医嘱id = GG.ID And BB.医嘱id=[1] "
    '------------------------------------------------------------------------------------------------------------------
    Case "材料"
    
        gstrSQL = gstrSQL & _
             "SELECT 0 As 可否分零,1 As 剂量系数,1 As 包装,II.病人科室id,2 AS 序号,CC.材料id AS 收费细目id,CC.实用数量 AS 数量,CC.执行科室id,DD.病人id,DD.主页id ,0 AS 单价 " & _
             "FROM 病人手术材料 CC,病人手术记录 DD,病人医嘱记录 II " & _
             "Where CC.记录id = DD.ID And II.ID = DD.医嘱id And DD.医嘱id =[1] "
             
    End Select
    
    gstrSQL = gstrSQL & _
               ") A " & _
            "Where C.收费细目id = B.ID " & _
               "AND C.收入项目ID = D.ID " & _
               "AND C.执行日期 <= SYSDATE " & _
               "AND A.数量 > 0 " & _
               "AND (C.终止日期 >= SYSDATE OR C.终止日期 IS NULL) " & _
               "AND A.收费细目id = B.ID " & _
               "AND F.病人id=A.病人id " & _
            "ORDER BY B.ID"
    
    Set rsCharge = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
    If rsCharge.BOF Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    '先删除原来的
    '------------------------------------------------------------------------------------------------------------------
    Select Case strMenuItem
    
    Case "治疗"
    
        If int来源 = 1 And int记录性质 = 1 Then
            
            gstrSQL = "Select Distinct c.No As No From 病人手术记录 a,病人费用记录 b,病人手术计价 c " & _
                        "Where Nvl(b.记录状态,0) In (0,1) And b.No=c.No And b.记录性质=1 And a.医嘱id=[1] And c.记录id=a.ID And c.No Is Not Null "
            
        Else
            
            gstrSQL = "Select Distinct c.No As No From 病人手术记录 a,病人费用记录 b,病人手术计价 c " & _
                        "Where Nvl(b.记录状态,0)=1 And b.No=c.No And b.记录性质=2 And a.医嘱id=[1] And c.记录id=a.ID And c.No Is Not Null "
            
        End If
        
    Case "用药"
        If int来源 = 1 And int记录性质 = 1 Then
            
            gstrSQL = "Select a.用药No As No From 病人手术记录 a,病人费用记录 b " & _
                        "Where Nvl(b.记录状态,0) In (0,1) And b.No=a.用药No And b.记录性质=1 And a.用药No Is Not Null And a.医嘱id=[1]"
            
        Else
            
            gstrSQL = "Select a.用药No As No From 病人手术记录 a,病人费用记录 b " & _
                        "Where Nvl(b.记录状态,0)=1 And b.No=a.用药No And b.记录性质=2 And a.用药No Is Not Null And a.医嘱id=[1]"
            
        End If
            
    Case "材料"
        If int来源 = 1 And int记录性质 = 1 Then
            
            gstrSQL = "Select a.材料No As No From 病人手术记录 a,病人费用记录 b " & _
                        "Where Nvl(b.记录状态,0) In (0,1) And b.No=a.材料No And b.记录性质=1 And a.材料No Is Not Null And a.医嘱id=[1]"
            
        Else
            
            gstrSQL = "Select a.材料No As No From 病人手术记录 a,病人费用记录 b " & _
                        "Where Nvl(b.记录状态,0)=1 And b.No=a.材料No And b.记录性质=2 And a.材料No Is Not Null And a.医嘱id=[1]"
            
        End If
    End Select
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
    If rs.BOF = False Then

        If int来源 = 1 Then
            If int记录性质 = 1 Then
                '划价
                strSQL = "zl_门诊划价记录_Delete('" & rs("No").Value & "','')"
                Call SQLRecordAdd(rsSQL, strSQL)
            Else
                strSQL = "zl_门诊记帐记录_Delete('" & rs("No").Value & "','','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Else
            strSQL = "zl_住院记帐记录_Delete('" & rs("No").Value & "','','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            Call SQLRecordAdd(rsSQL, strSQL)
        End If
    End If
    
    
    '
    '------------------------------------------------------------------------------------------------------------------
    With rsCharge
        
        '获取对应的医嘱信息
        gstrSQL = "Select 医嘱期效,病人科室ID,婴儿,执行频次,计价特性 From 病人医嘱记录 Where ID=[1]"
        Set rsAdvice = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
        If rsAdvice.BOF Then
            Screen.MousePointer = 0
            Exit Function
        End If
        
        strNO = zlDatabase.GetNextNo(int记录性质 + 12)
         
        '插入医嘱附加费用
        '--------------------------------------------------------------------------------------------------------------
        strSQL = "ZL_病人医嘱附费_Insert(" & lngKey & "," & lng发送号 & "," & int记录性质 & ",'" & strNO & "')"
        Call SQLRecordAdd(rsSQL, strSQL)
        
        For lngLoop = 1 To .RecordCount
            
            dbl数量 = zlCommFun.NVL(rsCharge("数量").Value, 0)
            
            
            '病人病区科室、执行科室
            '----------------------------------------------------------------------------------------------------------
            lng病人病区ID = zlCommFun.NVL(rsPati!病人病区ID, 0)
            lng病人科室ID = zlCommFun.NVL(rsPati!病人科室ID, 0)
            If lng病人科室ID = 0 Then
                lng病人病区ID = zlCommFun.NVL(rsAdvice!病人科室ID, 0)
                lng病人科室ID = zlCommFun.NVL(rsAdvice!病人科室ID, 0)
            End If
            If lng病人科室ID = 0 Then
                lng病人病区ID = UserInfo.部门ID
                lng病人科室ID = UserInfo.部门ID
            End If
            
            lng执行部门ID = !执行科室id
            
            cur单价 = rsCharge("单价").Value
            
            '检查普通收费项目的库存，计算实价药品/材料的单价
            '----------------------------------------------------------------------------------------------------------
            Select Case rsCharge("类别").Value
            Case "4", "5", "6", "7"
                Select Case rsCharge("类别").Value
                Case "4"
                    gstrSQL = "SELECT NVL(B.是否变价,0) AS 实价,NVL(在用分批,0) AS 分批 FROM 材料特性 A,收费项目目录 B WHERE A.材料id=B.ID AND A.材料id=[1] "
                Case "5", "6", "7"
                    '进行分零计算
                    dbl数量 = dbl数量
                    
                    If zlCommFun.NVL(rsCharge("可否分零").Value, 0) = 0 Then
                        dbl数量 = dbl数量 / zlCommFun.NVL(rsCharge("剂量系数").Value, 1)
                    Else
                        dbl数量 = IntEx(dbl数量 / zlCommFun.NVL(rsCharge("剂量系数").Value, 1) / zlCommFun.NVL(rsCharge("包装").Value, 1)) * zlCommFun.NVL(rsCharge("包装").Value, 1)
                    End If
                                            
                    gstrSQL = "SELECT NVL(I.是否变价,0) AS 实价,NVL(S.药房分批,0) AS 分批 FROM 收费项目目录 I,药品规格 S WHERE I.ID=S.药品id AND S.药品id=[1]"
                End Select
                
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", Val(!收费细目id))
                If rs.BOF = False Then
                    If rs("分批").Value <> 1 And rs("实价").Value <> 1 Then
                        '是普通项目,要检查库存
                        If dbl数量 > CalcStorage(!收费细目id, lng执行部门ID, False, False) Then
                            '超过库存数量
                            Select Case GetDrugWarnOption(lng执行部门ID, IIf(strMenuItem = "用药", "567", "4"))
                            Case 1          '库存不足提醒
                                If MsgBox(!名称 & "库存不足，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Screen.MousePointer = 0
                                    Exit Function
                                End If
                            Case 2          '库存不足禁止
                                MsgBox !名称 & "库存不足！", vbInformation, gstrSysName
                                Screen.MousePointer = 0
                                Exit Function
                            End Select
                        End If
                    ElseIf rs("实价") = 1 Then
                        cur单价 = CalcTimePrice(!收费细目id, lng执行部门ID, dbl数量)
                    End If
                End If
            End Select
                           
            '计算应收和实收金额
            '----------------------------------------------------------------------------------------------------------
            cur应收 = Format(dbl数量 * cur单价, str费用小数位)
            cur实收 = IIf(blnZeroBill, 0, cur应收)
            If rsPati("费别").Value <> "" And blnZeroBill = False Then cur实收 = Format(ActualMoney(rsPati("费别").Value, !收入项目ID, cur应收), str费用小数位)
            
            '每个收费项目的处理
            '----------------------------------------------------------------------------------------------------------
            If lng项目ID <> !收费细目id Then
            
                int父号 = lngLoop '获取价格父号
                
                '获取保险项目信息
                '------------------------------------------------------------------------------------------------------
                If int来源 = 2 And Not IsNull(rsPati!险类) And gblnInsure Then
                    strMsg = gclsInsure.GetItemInsure(lng病人id, !收费细目id, cur实收, False, rsPati!险类)
                    If strMsg <> "" Then
                        int保险项目否 = Val(Split(strMsg, ";")(0))
                        lng保险大类ID = Val(Split(strMsg, ";")(1))
                        cur统筹金额 = Format(Val(Split(strMsg, ";")(2)), "0.00")
                        str保险编码 = CStr(Split(strMsg, ";")(3))
                    End If
                End If
            End If
            lng项目ID = !收费细目id
            
            
            '如果是记帐单据，进行费用警告
            '----------------------------------------------------------------------------------------------------------
            
            If int记录性质 = 2 And blnZeroBill = False Then
                
                '搜索当前医嘱的最高报警级别,并与已报警级别比较
                
'                lng级别 = GetWarnGrade(lng已报警级别, !类别, bln医保, lng病人病区ID)
                
                str报警方案 = ""
                strSQL = "Select zl_PatiWarnScheme([1],[2]) As 报警方案 From Dual"
                Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", lng病人id, lng主页id)
                If rs.BOF = False Then
                    str报警方案 = zlCommFun.NVL(rs("报警方案").Value)
                End If
                lng级别 = GetWarnGrade(lng已报警级别, !类别, str报警方案, lng病人病区ID)
                
                lng报警级别 = IIf(lng报警级别 > lng级别, lng报警级别, lng级别)
                lng报警级别 = IIf(lng报警级别 > lng已报警级别, lng报警级别, lng已报警级别)
                            
                '判断是否费用是否够用
                curMoneyTotal = curMoneyTotal + cur实收
                
                If lng报警级别 > lng已报警级别 Then
                    If curMoneyTotal <> 0 Then
                        'If 欠费情况(zlCommFun.NVL(rsPati!姓名), lng病人id, lng主页id, curMoneyTotal, bln医保, lng报警级别, bln强制记帐, str已强制报警姓名) = "是" Then
                        If 欠费情况(zlCommFun.NVL(rsPati!姓名), lng病人id, lng主页id, curMoneyTotal, str报警方案, lng报警级别, bln强制记帐, str已强制报警姓名) = "是" Then
                            Screen.MousePointer = 0
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            '填写记录
            '----------------------------------------------------------------------------------------------------------
            If int来源 = 1 Then
                If int记录性质 = 1 Then
                    '生成门诊划价单据
                    '--------------------------------------------------------------------------------------------------
                    strSQL = _
                        "zl_门诊划价记录_Insert('" & strNO & "'," & lngLoop & "," & lng病人id & ",NULL," & _
                        ZVal(zlCommFun.NVL(rsPati!门诊号, 0)) & ",'" & zlCommFun.NVL(rsPati!付款码) & "','" & zlCommFun.NVL(rsPati!姓名) & "'," & _
                        "'" & zlCommFun.NVL(rsPati!性别) & "','" & zlCommFun.NVL(rsPati!年龄) & "','" & zlCommFun.NVL(rsPati!费别) & "',NULL," & _
                        lng病人病区ID & "," & lng病人科室ID & "," & UserInfo.部门ID & ",'" & UserInfo.姓名 & "'," & _
                        "NULL," & lng项目ID & ",'" & !类别 & "','" & !计算单位 & "',NULL,1," & dbl数量 & "," & _
                        "0," & ZVal(lng执行部门ID) & "," & IIf(int父号 = lngLoop, "NULL", int父号) & "," & _
                        !收入项目ID & ",'" & zlCommFun.NVL(!收据费目) & "'," & cur单价 & "," & cur应收 & "," & cur实收 & "," & _
                        strDate & "," & strDate & ",NULL,'" & UserInfo.姓名 & "'," & ZVal(lng类别ID) & ",NULL," & _
                        lngKey & ",'" & zlCommFun.NVL(rsAdvice!执行频次) & "',NULL,NULL," & zlCommFun.NVL(rsAdvice!医嘱期效, 0) & "," & _
                        zlCommFun.NVL(rsAdvice!计价特性, 0) & ",1)"
                    Call SQLRecordAdd(rsSQL, strSQL)
                Else
                    '生成门诊记帐单据
                    '--------------------------------------------------------------------------------------------------
                    strSQL = _
                        "zl_门诊记帐记录_Insert('" & strNO & "'," & lngLoop & "," & lng病人id & "," & _
                        ZVal(zlCommFun.NVL(rsPati!门诊号, 0)) & ",'" & zlCommFun.NVL(rsPati!姓名) & "','" & zlCommFun.NVL(rsPati!性别) & "'," & _
                        "'" & zlCommFun.NVL(rsPati!年龄) & "','" & zlCommFun.NVL(rsPati!费别) & "',NULL," & ZVal(rsAdvice!婴儿) & "," & _
                        lng病人病区ID & "," & lng病人科室ID & "," & UserInfo.部门ID & "," & _
                        "'" & UserInfo.姓名 & "',NULL," & lng项目ID & ",'" & !类别 & "'," & _
                        "'" & !计算单位 & "',1," & dbl数量 & ",0," & ZVal(lng执行部门ID) & "," & _
                        IIf(int父号 = lngLoop, "NULL", int父号) & "," & !收入项目ID & ",'" & zlCommFun.NVL(!收据费目) & "'," & cur单价 & "," & _
                        cur应收 & "," & cur实收 & "," & strDate & "," & strDate & ",NULL,NULL,'" & UserInfo.编号 & "'," & _
                        "'" & UserInfo.姓名 & "'," & ZVal(lng类别ID) & ",NULL,NULL," & lngKey & "," & _
                        "'" & zlCommFun.NVL(rsAdvice!执行频次) & "',NULL,NULL," & zlCommFun.NVL(rsAdvice!医嘱期效, 0) & "," & _
                        zlCommFun.NVL(rsAdvice!计价特性, 0) & ")"
                    Call SQLRecordAdd(rsSQL, strSQL)
                End If
            Else
                '生成住院记帐单据
                '------------------------------------------------------------------------------------------------------
                strSQL = _
                    "zl_住院记帐记录_Insert('" & strNO & "'," & lngLoop & "," & lng病人id & "," & ZVal(lng主页id) & "," & _
                    ZVal(zlCommFun.NVL(rsPati!住院号, 0)) & ",'" & zlCommFun.NVL(rsPati!姓名) & "','" & zlCommFun.NVL(rsPati!性别) & "'," & _
                    "'" & zlCommFun.NVL(rsPati!年龄) & "','" & Trim(zlCommFun.NVL(rsPati!床号)) & "','" & zlCommFun.NVL(rsPati!费别) & "'," & _
                    lng病人病区ID & "," & lng病人科室ID & ",NULL," & ZVal(rsAdvice!婴儿) & "," & _
                    UserInfo.部门ID & ",'" & UserInfo.姓名 & "',NULL," & lng项目ID & ",'" & !类别 & "'," & _
                    "'" & !计算单位 & "'," & int保险项目否 & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                    "1," & dbl数量 & ",0," & ZVal(lng执行部门ID) & "," & _
                    IIf(int父号 = lngLoop, "NULL", int父号) & "," & !收入项目ID & ",'" & zlCommFun.NVL(!收据费目) & "'," & cur单价 & "," & _
                    cur应收 & "," & cur实收 & "," & cur统筹金额 & "," & strDate & "," & strDate & ",NULL,NULL," & _
                    "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',NULL," & ZVal(lng类别ID) & ",NULL,NULL,NULL," & _
                    lngKey & ",'" & zlCommFun.NVL(rsAdvice!执行频次) & "',NULL,NULL," & zlCommFun.NVL(rsAdvice!医嘱期效, 0) & "," & _
                    zlCommFun.NVL(rsAdvice!计价特性, 0) & ",NULL)"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
            
            .MoveNext
            
        Next
        
        '
        '--------------------------------------------------------------------------------------------------------------
        Select Case strMenuItem
        Case "治疗"
            If .RecordCount > 0 Then
                strSQL = "zl_病人手术计价_No(" & lngKey & ",'" & strNO & "'," & int记录性质 & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
            
        Case "用药", "材料"
            If .RecordCount > 0 Then
                strSQL = "zl_病人手术记录_No(" & lngKey & ",'" & strNO & "'," & int记录性质 & ",'" & strMenuItem & "')"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        End Select
        
    End With
    
    '
    '------------------------------------------------------------------------------------------------------------------
        
    blnTran = True
    gcnOracle.BeginTrans
    
    If SQLRecordExecute(rsSQL, "mdlOps", False) = False Then GoTo errHand
        
    '在提交前进行医保传输
    '------------------------------------------------------------------------------------------------------------------
    If int来源 = 2 And Not IsNull(rsPati!险类) And gblnInsure Then
        If gclsInsure.GetCapability(support记帐上传, lng病人id, rsPati!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, lng病人id, rsPati!险类) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strMsg, rsPati!险类) Then
                gcnOracle.RollbackTrans
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    blnTran = False
    
    '在提交后进行医保传输
    '------------------------------------------------------------------------------------------------------------------
    If int来源 = 2 And Not IsNull(rsPati!险类) And gblnInsure Then
        If gclsInsure.GetCapability(support记帐上传, lng病人id, rsPati!险类) And gclsInsure.GetCapability(support记帐完成后上传, lng病人id, rsPati!险类) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strMsg, rsPati!险类) Then
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, gstrSysName
                Else
                    MsgBox "单据""" & strNO & """的数据向医保传送失败,该单据已保存！", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
        
    Screen.MousePointer = 0
    
    MakeChargeBill = strNO
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If blnTran Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "Trans", adTinyInt                   '1表示开始;2表示结束
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = ParamInfo.系统名称
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
        
            strSQL = CStr(rs("SQL").Value)
            Call zlDatabase.ExecuteProcedure(strSQL, strTitle)
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
End Function

Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption)
        
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        
    End With
    
    Set NewCommandBar = objControl
    
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.ID = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.STYLE = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars, Optional ByVal blnEnableCustomization As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    
    cbsMain.VisualTheme = xtpThemeOffice2003
        
    With cbsMain.Options
        .ShowExpandButtonAlways = blnEnableCustomization
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization blnEnableCustomization

    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    cbsMain.Options.LargeIcons = False
    
    CommandBarInit = True
    
End Function

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet              '打印设置
    
        Call zlPrintSet
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel               '打印数据,预览数据,输出到Excel
        
        If objPrnVsf Is Nothing Then Exit Function
        
        Call SearchPrintData(objPrnVsf, frmPubResource.msfPrint)
        
        '调用打印部件处理
        Set objPrint.Body = frmPubResource.msfPrint
        objPrint.Title.Text = strPrintTitle
        Set objAppRow = New zlTabAppRow
        Call objAppRow.Add("")
        Call objAppRow.Add("打印时间:" & Now())
        Call objPrint.BelowAppRows.Add(objAppRow)

        Select Case Control.ID
        Case conMenu_File_Print
            bytMode = zlPrintAsk(objPrint)
            If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
        Case conMenu_File_Preview
            zlPrintOrView1Grd objPrint, 2
        Case conMenu_File_Excel
            zlPrintOrView1Grd objPrint, 3
        End Select
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Text      '按钮文字
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        frmMain.cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Size      '大图标
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_StatusBar         '状态栏
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Help              '帮助主题
    
        Call ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((ParamInfo.系统号) / 100))
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Web_Home          'Web上的中联
        
        Call zlHomePage(frmMain.hWnd)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Web_Forum         'Web上的论坛
    
        Call zlWebForum(frmMain.hWnd)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Web_Mail          '发送反馈
        
        Call zlMailTo(frmMain.hWnd)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_About             '关于
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Exit              '退出
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button            '工具栏
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              '图标文字
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '大图标
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function CopyMenu(ByVal cbsMain As Object, Optional ByVal intNo As Integer = 2) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '弹出菜单处理
    
    On Error GoTo errHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Call cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.ID, cbrControl2.Caption)
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(ParamInfo.系统号, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

'################################################################################################################
'## 功能：  将指定的LOB字段复制为临时文件
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  存放内容的文件名，失败则返回零长度""
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If gobjFSO.FileExists(strZipPath & "TMP.RTF") Then gobjFSO.DeleteFile strZipPath & "TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
End Function

Public Sub ShowDocument(ByRef edt As Object, ByVal lngRecordId As Long, Optional ByVal blnPrivacyProtect As Boolean)
    '******************************************************************************************************************
    '功能：刷新病历显示内容；
    '参数：lngRecordId：电子病历记录ID；blnPrivacyProtect：是否启用隐私保护
    '******************************************************************************************************************
    
    Dim mstrPrivs As String
    Dim blnPrivacy As Boolean
    Dim Elements As New cEPRElements
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    
    If blnPrivacyProtect = True Then
        mstrPrivs = ";" & GetPrivFunc(ParamInfo.系统号, 1070) & ";"
        blnPrivacy = InStr(mstrPrivs, ";忽略隐私保护;") = 0     '保护隐私项目
    End If
    
    Dim strTemp As String
    Dim strZipFile As String

'    mlngRecordId = lngRecordId
    edt.Freeze
    edt.ReadOnly = False
    edt.NewDoc
    strZipFile = zlBlobRead(5, lngRecordId)
    If gobjFSO.FileExists(strZipFile) Then
        strTemp = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strTemp) Then
            '打开文件
            edt.OpenDoc strTemp
            '设置替换项目
            If blnPrivacy Then
                '读取所有的要素
                gstrSQL = "Select A.ID,A.对象标记 From 电子病历内容 A, 隐私保护项目 B,诊治所见项目 C " & _
                    "Where A.对象类型 = 4 And A.替换域 = 1 And A.文件id = [1] And A.对象序号 > 0 and B.项目id = C.ID And A.要素名称 =C.中文名 And C.替换域 = 1 "
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngRecordId)
                If Not rs.EOF Then
                    Do While Not rs.EOF
                        lngKey = Elements.Add(zlCommFun.NVL(rs("对象标记"), 0))
                        Elements("K" & lngKey).GetElementFromDB cprET_单病历编辑, rs("ID"), True, "电子病历内容"
                        '替换要素内容
                        Elements("K" & lngKey).内容文本 = String(Len(Elements("K" & lngKey).内容文本), "*")
                        Elements("K" & lngKey).Refresh edt
                        rs.MoveNext
                    Loop
                End If
                rs.Close
            End If
            gobjFSO.DeleteFile strTemp, True
        End If
        gobjFSO.DeleteFile strZipFile, True
        edt.SelStart = 0
    End If
    
    If lngRecordId > 0 Then
        '设置页面格式
        Dim mEPRFileInfo As New cEPRFileDefineInfo
        gstrSQL = "Select c.ID, a.格式 From   病历页面格式 a, 病历文件列表 b, 电子病历记录 c " & _
                " Where  c.文件id = b.id And a.种类 = b.种类 And a.编号 = b.页面 And c.ID = [1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngRecordId)
        If Not rs.EOF Then
            mEPRFileInfo.格式 = zlCommFun.NVL(rs("格式").Value)
            mEPRFileInfo.SetFormat edt, mEPRFileInfo.格式
            edt.ResetWYSIWYG
        End If
        Set mEPRFileInfo = Nothing
    End If
    edt.UnFreeze
    edt.RefreshTargetDC
    edt.ReadOnly = True
End Sub

Public Function GetDefaultDept(ByVal str类别 As String, ByVal int病人来源 As Integer) As Long
    Dim strTmp As String
    
    strTmp = ""
    Select Case str类别
    Case "4"
'                strTmp = IIf(mint病人来源 = 1, "门诊缺省西药房", "住院缺省西药房")
    Case "5"
        strTmp = IIf(int病人来源 = 1, "门诊缺省西药房", "住院缺省西药房")
    Case "6"
        strTmp = IIf(int病人来源 = 1, "门诊缺省成药房", "住院缺省成药房")
    Case "7"
        strTmp = IIf(int病人来源 = 1, "门诊缺省中药房", "住院缺省中药房")
    End Select
    
    If strTmp <> "" Then
        GetDefaultDept = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, strTmp, "0"))
    End If
    
End Function
