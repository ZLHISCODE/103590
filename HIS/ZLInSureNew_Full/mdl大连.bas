Attribute VB_Name = "mdl大连"
Option Explicit
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--开发区接口
    '参数说明:
    '   msgType-业务请求类型,见以下的参数表
    '   packageType-数据解析格式类型，系统重组数据时使用,见以下的参数表
    '   packageLength-数据串的长度,见以下的参数表
    '   str-数据串,调用时，通过数据串传入参数；函数返回时，数据串中包含返回的数据
    '   strCom:数据请求串口（根据读卡器插口位置，本参数可以取值：'com1','com2')
    '返回:
    '   I.  当函数返回值等于0时，表示成功，字符串中包含了业务处理后返回的数据
    '   II. 当函数返回值不等于0时，参见错误代码一览表，应用需要分析错误代码然后进行适当的处理

'读卡器驱动程序
Private Declare Function IC_Read_Base Lib "ICCNII32.DLL" (ByVal szData As String) As Long
Private Declare Function IC_Read_Plus Lib "ICCNII32.DLL" (nSequence As Long, ByVal szData As String) As Long

'Private Declare Function KfqTransData Lib "OltpTransKfq03.dll" ( _
    ByVal msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
    ByVal str As String, ByVal strCom As String) As Long
    
'2005-08-02 周海全
'开发区医保升级
Private Declare Function KfqTransData Lib "OltpTransKfq05.dll" ( _
    ByVal msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
    ByVal str As String, ByVal strCom As String) As Long
    
'--普通接口
Private Declare Function OltpTransData Lib "OltpTransIc04.dll" ( _
    ByVal msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
    ByVal str As String, ByVal strCom As String) As Long
    '以下为开发区的参数表
''业务请求类型    数据解析格式类型    数据串最小长度         说明
''------------    ----------------    --------------         -----------------------------------------------
''1001            101                 95                     实时验卡（读卡、验卡）
''1002            12                  420                    实时结算
''1003            7                   297                    实时医疗明细数据提交
''1004            9                   136                    实时住院登记数据提交
''1006            12                  420                    实时结算预算
''1008            101                 95                     实时查询（直接查询中心数据）

'2005-08-02 周海全
'开发区升级修改
'业务请求类型    数据解析格式类型    数据串最小长度         说明
'------------    ----------------    --------------         -----------------------------------------------
'1001            101                 96                     实时验卡（读卡、验卡）
'1002            12                  509                    实时结算
'1003            7                   240                    实时医疗明细数据提交
'1004            9                   210                    实时住院登记数据提交
'1006            12                  509                    实时结算预算
'1008            101                 96                     实时查询（直接查询中心数据）

'以下为大连市的参数表
'业务请求类型    数据解析格式类型    数据串最小长度         说明
'------------    ----------------    --------------         -----------------------------------------------
'1001            101                 94                     实时验卡（读卡、验卡）
'1002            12                  424                    实时结算
'1003            7                   230                    实时医疗明细数据提交
'1004            9                   206                    实时住院登记数据提交
'1006            12                  424                    实时结算预算
'1008            101                 94                     实时查询（直接查询中心数据）
'1005            8                   274                    实时医嘱传输
'1007            2                   55                     慢病帐户查询
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public gblnKFQCom_大连  As Boolean   'true-开发区接口,False-普通接口

Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum

Public g病人身份_大连 As 病人身份
Private Type 病人身份
    个人编号            As String
    姓名                As String
    性别                As String
    出生日期            As String
    年龄                As Integer
    身份证号            As String
    IC卡号              As Long
    治疗序号            As Long
    职工就医类别        As String
    基本个人帐户余额    As Double
    补助个人帐户余额    As Double
    当前状态            As Double
    统筹累计            As Double
    月缴费基数          As Double
    帐户状态            As String
    参保类别1           As String
    参保类别2           As String
    参保类别3           As String
    参保类别4           As String
    参保类别5           As String
    
    转诊单号            As String           '身份验证时输入
    医保中心            As Long             '身份验证时选择,保存的序号
    就诊分类            As Long             '身份验证时选择,保存的是结算方式代码
    支付金额            As Double           '
    诊断编码            As String           '诊断编码时输入,门诊有效
    诊断名称            As String           '诊断名称时输入,门诊有效
    
    补助帐户原始值      As Double          '慢病查询获取
    补助帐户当前值      As Double          '慢病查询获取
    慢病帐户状态        As Double          '慢病查询获取
    起付线              As Double
    
    当前个人帐户余额    As Double
    当前补助帐户余额    As Double
    当前统筹累计            As Double
    结算开始            As Boolean             '主要应用于门诊多单据,需处理当前的余额
    历次结算            As Boolean             '历次结算
    
End Type

Public gbln模拟接口   As Boolean      '模拟接口数据

Public gstr医院编码_大连 As String        '医院编码,只能为4位
Public gintComPort_大连 As Integer
Public gbln门诊明细时实上传 As Boolean
Public gbln住院明细时实上传 As Boolean
Public gbln单病种D As Boolean           '大连市出院单病种提示
Public gbln单病种K As Boolean           '开发区出院单病种提示
Public gblnDebug As Boolean             '调试用
Private mblnInit As Boolean     '是否被初始化

Private Function Read模拟数据(ByVal lng中心代码 As Long, _
        msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
        str As String)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:通过该功能读取模拟数据,以例测试
    '--入参数:
    '--出参数:
    '--返  回:字串
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strArr
    Dim strArr1
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    
    strFile = App.Path & "\大连医保模拟提交串数据" & lng中心代码 & ".txt"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    
    objText.WriteLine msgType & Space(10) & packageType & Space(10) & packageLength & "|| " & str
    objText.Close
    
    If Dir(Left(App.Path, 18) & "\医保资料\大连医保\大连医保模拟数据" & lng中心代码 & ".txt") <> "" Then
            Set objText = objFile.OpenTextFile(Left(App.Path, 18) & "\医保资料\大连医保\大连医保模拟数据" & lng中心代码 & ".txt")
            Do While Not objText.AtEndOfStream
                strTemp = Trim(objText.ReadLine)
                strArr = Split(strTemp, "||")
                strArr1 = Split(strArr(0), "|")
                If Val(strArr1(0)) = msgType Then
                     str = strArr(1)
                     Exit Do
                End If
            Loop
            objText.Close
    End If
    
End Function
Public Function 读取病人身份_大连(ByVal lng中心代码 As Long, ByVal intinsure As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取病人的相关身份,并将信息赋给g病人身份_大连
    '--入参数:lng中心代码(2代表开发区)
    '--出参数:
    '--返  回:读取成功,返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    
    Dim strInfor As String
    Dim rsTempset As New ADODB.Recordset
    Dim lngReturn As Long
    Dim int性别 As Integer
    
    读取病人身份_大连 = False
    Err = 0
    On Error GoTo errHand:
    '周海全调试 2003-12-17
    '如果此处不传入空格值时，程序运行至此处会直接退出
    strInfor = Space(100)
    If gbln模拟接口 Then
        Read模拟数据 lng中心代码, 1001, 101, 94, strInfor
        If strInfor = "" Then Exit Function
    Else
        If lng中心代码 = 2 Then
            '1001    101 95  实时验卡（读卡、验卡）
            lngReturn = KfqTransData(1001, 101, 95, strInfor, "com" & gintComPort_大连)
        Else
            '1001    101 94  实时验卡（读卡、验卡）
            lngReturn = OltpTransData(1001, 101, 94, strInfor, "com" & gintComPort_大连)
        End If
        If lngReturn <> 0 Or strInfor = "" Then
            ShowMsgbox GetErrInfo(CStr(lngReturn), intinsure)
            Exit Function
        End If
    End If
    '取掉控格
    strInfor = Mid(strInfor, 2)
    With g病人身份_大连
        .医保中心 = lng中心代码
        If lng中心代码 = 2 Then
            .个人编号 = Substr(strInfor, 1, 10)         '个人保号    1   10      中心返回
            .姓名 = Trim(Substr(strInfor, 11, 8))       '姓名    11  8       中心返回
            .身份证号 = Substr(strInfor, 19, 18)        '身份证号    19  18      中心返回
            .IC卡号 = Substr(strInfor, 37, 7)           'IC卡号  37  7       中心返回
            .治疗序号 = Val(Substr(strInfor, 44, 4))    '治疗序号    44  4       中心返回
            .职工就医类别 = Substr(strInfor, 48, 1)     '职工就医类别    48  1   A在职、B退休    中心返回
            .基本个人帐户余额 = Val(Substr(strInfor, 49, 10)) '基本个人帐户余额    49  10      中心返回
            .补助个人帐户余额 = Val(Substr(strInfor, 59, 10)) '补助个人帐户余额    59  10      中心返回
            .统筹累计 = Val(Substr(strInfor, 69, 10)) '统筹累计    69  10      中心返回
            .月缴费基数 = Val(Substr(strInfor, 79, 10)) '月缴费基数  79  10  月缴费工资  中心返回
            .帐户状态 = Substr(strInfor, 89, 1) '帐户状态    89  1   A正常、B半止付、C全止付、D销户  中心返回
            .参保类别1 = Substr(strInfor, 90, 1) '参保类别1   90  1   是否享受高额 1 享受 0 不享受    中心返回
            .参保类别2 = Substr(strInfor, 91, 1) '参保类别2   91  1   是否享受补助（商业补助、公务员补助）'0 不享受 1 商业 2 公务员    中心返回
            .参保类别3 = Substr(strInfor, 92, 1) '参保类别3   92  1   0 企保、1 事保、2 被征地人员(开发区升级增加)
            .参保类别4 = Substr(strInfor, 93, 1) '参保类别4   93  1   备用    中心返回
            .参保类别5 = Substr(strInfor, 94, 1) '参保类别5   94  1   备用    中心返回
        Else
            .个人编号 = Substr(strInfor, 1, 8)  '个人编号    CHAR    1   8   医保编号    中心
            
            ' 2004-09-23    周海全
            '由于测试卡的姓名为全数字，需要处理为汉字
            .姓名 = Trim(Substr(strInfor, 9, 8))                            '姓名    CHAR    9   8       中心
            .姓名 = IIf(IsNumeric(.姓名), Trim(.姓名) & "测试", .姓名)      '姓名    CHAR    9   8       中心
            
            .身份证号 = Substr(strInfor, 17, 18)    '身份证号    CHAR    17  18  18位或15位  中心
            .IC卡号 = Substr(strInfor, 35, 7)       'IC卡号  NUM 35  7       中心
            .治疗序号 = Val(Substr(strInfor, 42, 4))    '治疗序号    NUM 42  4       中心
            
            '周海全调试 2003-12-17
            '加入：Q企业公费
            .职工就医类别 = Substr(strInfor, 46, 1)     '职工就医类别    CHAR    46  1   A在职、B退休、L离休、T特诊、Q企业公费  中心
            .基本个人帐户余额 = Val(Substr(strInfor, 47, 10))   '基本个人帐户余额    NUM 47  10      中心
            .补助个人帐户余额 = Val(Substr(strInfor, 57, 10))   '补助个人帐户余额    NUM 57  10  现用于公务员单独列帐    中心
            .统筹累计 = Val(Substr(strInfor, 67, 10))   '统筹累计    NUM 67  10      中心
            .月缴费基数 = Val(Substr(strInfor, 77, 10)) '月缴费基数  NUM 77  10  月缴费工资  中心
            .帐户状态 = Substr(strInfor, 87, 1)         '帐户状态    CHAR    87  1   A正常、B半止付、C全止付、D销户  中心
            .参保类别1 = Substr(strInfor, 88, 1)        '参保类别1   CHAR    88  1   是否享受高额: 0 不享受高额、1 享受高额、2 医疗保险不可用    中心
            .参保类别2 = Substr(strInfor, 89, 1)        '参保类别2   CHAR    89  1   是否享受补助（商业补助、公务员补助）0 不享受 1 商业 2 公务员    中心
            .参保类别3 = Substr(strInfor, 90, 1)        '参保类别3   CHAR    90  1   0 企保、1 事保  中心
            .参保类别4 = Substr(strInfor, 91, 1)        '参保类别4   CHAR    91  1   0生育不可用、1生育可用  中心
            .参保类别5 = Substr(strInfor, 92, 1)        '参保类别5   CHAR    92  1   0工伤不可用、1工伤可用  中心
        End If
        
        '获取保险帐户的当前状态
        gstrSQL = "select 当前状态 from 保险帐户 where 险类=" & Decode(lng中心代码, 2, 83, 1, 82) & " and 医保号='" & g病人身份_大连.个人编号 & "'"
        zlDatabase.OpenRecordset rsTempset, gstrSQL, "获取保险帐户当前状态"
        If Not rsTempset.EOF Then
            g病人身份_大连.当前状态 = rsTempset!当前状态
        End If
        
        int性别 = Val(IIf(Len(.身份证号) = 18, Mid(.身份证号, 17, 1), Right(.身份证号, 1))) Mod 2
        '根据身份证取出相应的性别
        .性别 = IIf(int性别 = 0, "女", "男")
        .出生日期 = zlCommFun.GetIDCardDate(Trim(.身份证号))
        '计算年龄
        If IsDate(.出生日期) And .出生日期 <> "" Then
            '.年龄 = Abs(Int((zlDatabase.Currentdate - CDate(.出生日期)) / 366))
            gstrSQL = "Select Months_between(Trunc(SysDate),To_Date('" & .出生日期 & "','YYYY-MM-DD'))/12 As 年龄 From Dual"
            zlDatabase.OpenRecordset rsTempset, gstrSQL, "获取年龄"
            If rsTempset.RecordCount <= 0 Then
                .年龄 = 0
            Else
                .年龄 = rsTempset!年龄
            End If
        Else
            .年龄 = 0
        End If
        
    End With
    读取病人身份_大连 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    读取病人身份_大连 = False
End Function

Public Function 业务请求_大连( _
            ByVal lng中心代码 As Long, _
            ByVal lngMsgType As Long, _
            strTans As String, _
            ByVal intinsure As Integer _
    ) As Boolean
    
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对相关的业务请求,并返回相应的结果
    '--入参数:lng中心代码(2代表开发区)
    '   lngMsgType-业务请求类型
    '   lngPackageType-数据解析格式类型
    '   lngPackageLength-数据串的长度
    '   strTans-数据串,调用时，通过数据串传入参数；函数返回时，数据串中包含返回的数据
    '返回:
    '   成功-true,否则False
    '-----------------------------------------------------------------------------------------------------------
    Dim lngPackageType As Long
    Dim lngPackageLength As Long
    Dim i As Long
    Dim strTmp As String
    Dim strReg As String
    
    i = lngMsgType
    
    '以下为开发区的参数表
'    '业务请求类型    数据解析格式类型    数据串最小长度         说明
'    '------------    ----------------    --------------         -----------------------------------------------
'    '1001            101                 95                     实时验卡（读卡、验卡）
'    '1002            12                  420                    实时结算
'    '1003            7                   297                    实时医疗明细数据提交
'    '1004            9                   136                    实时住院登记数据提交
'    '1006            12                  420                    实时结算预算
'    '1008            101                 95                     实时查询（直接查询中心数据）

    '2005-08-02 周海全
    '开发区升级修改
    '业务请求类型    数据解析格式类型    数据串最小长度         说明
    '------------    ----------------    --------------         -----------------------------------------------
    '1001            101                 96                     实时验卡（读卡、验卡）
    '1002            12                  509                    实时结算
    '1003            7                   240                    实时医疗明细数据提交
    '1004            9                   210                    实时住院登记数据提交
    '1006            12                  509                    实时结算预算
    '1008            101                 96                     实时查询（直接查询中心数据）
    
    '以下为大连市的参数表
    '业务请求类型    数据解析格式类型    数据串最小长度         说明
    '------------    ----------------    --------------         -----------------------------------------------
    '1001            101                 94                     实时验卡（读卡、验卡）
    '1002            12                  424                    实时结算
    '1003            7                   230                    实时医疗明细数据提交
    '1004            9                   206                    实时住院登记数据提交
    '1006            12                  424                    实时结算预算
    '1008            101                 94                     实时查询（直接查询中心数据）
    '1005            8                   274                    实时医嘱传输
    '1007            2                   55                     慢病帐户查询
    
    Dim strInfor As String
    Dim strSQL As String
    Dim lngReturn As Long
    业务请求_大连 = False
    Err = 0
    On Error Resume Next
    If lng中心代码 = 2 Then
        strTmp = Switch(i = 1001, "101|96", i = 1002, "12|509", i = 1003, "7|240", i = 1004, "9|210", i = 1006, "12|509", _
            i = 1008, "101|96")
        If Err <> 0 Then
            strTmp = "|"
        End If
    Else
            strTmp = Switch(i = 1001, "101|94", i = 1002, "12|475", i = 1003, "7|230", i = 1004, "9|206", i = 1006, "12|475", _
                i = 1008, "101|94", i = 1005, "8|221", i = 1007, "2|55")
        If Err <> 0 Then
            strTmp = "|"
        End If
    End If
    lngPackageType = Val(Split(strTmp, "|")(0))
    lngPackageLength = Val(Split(strTmp, "|")(1))
    
    Err = 0
    On Error GoTo errHand:
    strInfor = strTans
    If gbln模拟接口 Then
        Read模拟数据 lng中心代码, lngMsgType, lngPackageType, lngPackageLength, strInfor
        If strInfor = "" Then
            strTans = strInfor
            Exit Function
        End If
    Else
        '因刘洋说,在所有业务类型的请求中都需前加空格.所以特加入:" " &
        strInfor = " " & strInfor
        
        '----临时为切换数据进行分割数据产生提供给老系统进行计算
        If lngMsgType = 1006 Then
            If Val(GetSetting("ZLSOFT", "公共模块\zl9insure\操作", "分割输出", 0)) = 1 Then
                strSQL = "insert into 大连结算分割(业务类型,卡号,申请时间,信息行) values(" & lngMsgType & ",Substrb(substr('" & strInfor & "',2,1000), 14, 7),sysdate,'" & strInfor & "')"
                gcnOracle.Execute strSQL
                MsgBox "分割数据产生成功,退出预结算!"
                业务请求_大连 = False
                Exit Function
            End If
        End If
        If lng中心代码 = 2 Then
            lngReturn = KfqTransData(lngMsgType, lngPackageType, lngPackageLength, strInfor, "com" & gintComPort_大连)
        Else
            lngReturn = OltpTransData(lngMsgType, lngPackageType, lngPackageLength, strInfor, "com" & gintComPort_大连)
        End If
        If lngReturn <> 0 Or strInfor = "" Then
            ShowMsgbox GetErrInfo(CStr(lngReturn), intinsure)
            strTans = ""
            Exit Function
        End If
    End If
    '取掉控格
    strInfor = Mid(strInfor, 2)
    
    strTans = strInfor
    业务请求_大连 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    strTans = ""
    业务请求_大连 = False
End Function


Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取指定字串的值,字串中可以包含汉字
    '--入参数:strInfor-原串
    '         lngStart-直始位置
    '         lngLen-长度
    '--出参数:
    '--返  回:子串
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    Err = 0
    On Error GoTo errHand:

    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
'    strTmp = Right(Substr, 1)
'    If zlCommFun.ActualLen(strTmp) = 1 Then
'        If asc(strTmp) < 32 Or asc(strTmp) > 126 Then
'            Substr = Left(Substr, Len(Substr) - 1)
'        End If
'    End If
    Exit Function
errHand:
    Substr = ""
End Function

Public Function 医保初始化_大连(ByVal intinsure As Integer) As Boolean

    Dim rsTemp  As New ADODB.Recordset
    Dim strReg As String
    
    '功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
    '返回：初始化成功，返回true；否则，返回false
    
    On Error Resume Next
    Err = 0
    On Error GoTo 0
    
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & intinsure
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "获取医院编码")
    gstr医院编码_大连 = Nvl(rsTemp!医院编码, "")
    
    '设置端口号
    Call GetRegInFor(g公共模块, "操作", "端口号", strReg)

    If Val(strReg) = 0 Then
        gintComPort_大连 = 1
    Else
        gintComPort_大连 = IIf(Val(strReg) > 99, 1, Val(strReg))
    End If
    
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        gbln模拟接口 = True
    Else
        gbln模拟接口 = False
    End If
    Call GetRegInFor(g公共模块, "操作", "开发区", strReg)
    
    If intinsure = TYPE_大连开发区 Then
        gblnKFQCom_大连 = True
    Else
        gblnKFQCom_大连 = False
    End If
    Call GetRegInFor(g公共模块, "操作", "调试", strReg)
    
    gblnDebug = strReg = "1"
    
    '设置上传明细参数
    gstrSQL = "Select * From 保险参数 where 参数名 in ('门诊明细时实上传','住院明细时实上传','单病种出院提示') and 险类=" & intinsure
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取保险参数"
    gbln门诊明细时实上传 = True
    gbln住院明细时实上传 = True
    
    '处理出院诊断单病种提示
    Call GetRegInFor(g公共模块, "操作", "大连市单病种", strReg)
    gbln单病种D = strReg = "1"
    Call GetRegInFor(g公共模块, "操作", "开发区单病种", strReg)
    gbln单病种K = strReg = "1"

    Do While Not rsTemp.EOF
        Select Case Nvl(rsTemp!参数名)
        Case "门诊明细时实上传"
            gbln门诊明细时实上传 = IIf(Val(Nvl(rsTemp!参数值)) = 1, True, False)
        Case "住院明细时实上传"
            gbln住院明细时实上传 = IIf(Val(Nvl(rsTemp!参数值)) = 1, True, False)
'        Case "医嘱明细时实上传"
'            gbln医嘱明细时实上传 = IIf(Val(Nvl(rsTemp!参数值)) = 1, True, False)
        Case "单病种出院提示"
            If intinsure = 82 Then
                gbln单病种D = IIf(Val(Nvl(rsTemp!参数值)) = 1, True, False)
            Else
                gbln单病种K = IIf(Val(Nvl(rsTemp!参数值)) = 1, True, False)
            End If
        End Select
        rsTemp.MoveNext
    Loop
    
    mblnInit = True
    医保初始化_大连 = True
End Function

Public Function 个人余额_大连(ByVal lng病人ID As Long, ByVal intinsure As Integer) As Currency
    '功能: 根据病人id取出余额
    '参数: 病人id
    '返回: 返回个人帐户余额
    Dim rsAcc As New ADODB.Recordset
    
    
    '读卡失败则退出
    gstrSQL = "Select Nvl(帐户余额,0) 帐户余额,退休证号 From 保险帐户 Where 险类=[1] And 病人id=[2]"
    
    Set rsAcc = zlDatabase.OpenSQLRecord(gstrSQL, "读取帐户余额", intinsure, lng病人ID)
    
    With g病人身份_大连
        .基本个人帐户余额 = Nvl(rsAcc!帐户余额, 0)
        .补助个人帐户余额 = Val(Nvl(rsAcc!退休证号))
        个人余额_大连 = .基本个人帐户余额 + .补助个人帐户余额
    End With
'    Call WriteDebugInfor_大连("个人余额_大连 ", lng病人id)
End Function

Public Function 医保设置_大连(ByVal lng险类 As Long, ByVal lng医保中心 As Integer) As Boolean
    医保设置_大连 = frmSet大连.ShowME(lng险类, lng医保中心)
    'Call WriteDebugInfor_大连("医保设置_大连 ", lng病人id)
End Function

Public Function 身份标识_大连(Optional bytType As Byte, Optional lng病人ID As Long, Optional intinsure As Integer) As String
    Dim str备注 As String, RSPATIENT As New ADODB.Recordset
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    
    身份标识_大连 = frmIdentify大连.GetPatient(intinsure, bytType, lng病人ID)
'    Call WriteDebugInfor_大连("身份标识_大连 byttype:" & bytType, lng病人id)
    
End Function
Public Function 身份标识_大连2(ByVal strCard As String, ByVal strPass As String, Optional lng病人ID As Long, Optional intinsure As Integer) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    '/**?
    身份标识_大连2 = frmIdentify大连.GetPatient(intinsure, 3, lng病人ID)
'    Call WriteDebugInfor_大连("身份标识_大连2", lng病人id)
    
End Function

Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = Substr(strCode, 1, lngLen)
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '主要有空格引起的
        strTmp = Substr(strCode, 1, lngLen)
    End If
    '取掉最后半个字符
    Rpad = Replace(strTmp, Chr(0), strChar)
End Function
Private Function Get就诊分类(ByVal byt业务 As Byte, ByVal int分类 As Integer) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取就诊分类标识
    '--入参数:byt业务-(0-结算,1-冲帐)
    '         int分类 门诊:(1-普通门诊,2-急诊门诊,3-门诊大病,4-门诊慢病补助)
    '                 住院:(5-普通住院,6-家庭病床住院,7-生育保险住院,8-工伤保险住院)
    '--出参数:
    '--返  回:医保中心的分类标识
    '-----------------------------------------------------------------------------------------------------------
    '医保中心的就诊分类的对应系统
    
    '1 门诊结算
    'A 门诊结算冲账
    '3 急诊结算
    '7 急诊结算冲账
    '5 门诊大病结算
    'B 门诊大病冲账
    'S 慢病补助结算
    'T 慢病补助冲帐
    
    '2 住院结算
    'D 住院冲冲账
    '9 住院冲补账  此功能暂不做
    '4 家庭病床结算
    'C 家庭病房冲冲账
    '8 家庭病房补账     '此功能暂不做
    'O 生育保险住院结算
    'P 生育保险住院冲帐
    'Q 工伤保险结算
    'R 工伤保险冲帐


    Dim i As Integer
    Dim strTmp As String
    i = int分类
    
    '刘兴宏标注:200404
    '     门诊:1-1,2-3,3-5,4-"S"
    '     住院:5-2,6-4,7-"O",8-"Q"
            
            
    Select Case int分类
        Case 1  '1-普通门诊
            strTmp = Decode(byt业务, 0, "1", "A")
        Case 2  '2-急诊门诊
            strTmp = Decode(byt业务, 0, "3", "7")
        Case 3  '3-门诊大病
            strTmp = Decode(byt业务, 0, "5", "B")
        Case 4  '4-门诊慢病补助
            strTmp = Decode(byt业务, 0, "S", "T")
        Case 5  '5-普通住院,
            strTmp = Decode(byt业务, 0, "2", "D")
        Case 6  '6-家庭病床住院,
            strTmp = Decode(byt业务, 0, "4", "C")
        Case 7  '7-生育保险住院
            strTmp = Decode(byt业务, 0, "O", "P")
        Case 8  '8-工伤保险住院
            strTmp = Decode(byt业务, 0, "Q", "R")
        Case Else
            strTmp = ""
    End Select
    Get就诊分类 = strTmp
End Function

Public Function 门诊虚拟结算_大连(rs明细 As ADODB.Recordset, str结算方式 As String, ByVal intinsure As Integer) As Boolean
    Dim curTotal As Double, cur个人帐户 As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rs大类 As New ADODB.Recordset
    Dim rs收费细目 As New ADODB.Recordset
    
    Dim strInfor As String  '定义中心返回串
    Dim dbl诊察费 As Double, dbl草药费 As Double, dbl成药费 As Double
    
    '开发区升级
    Dim dbl药费自费 As Double
'    Dim dbl结算前疾病统筹累计 As Double
'    Dim dbl结算后疾病统筹累计 As Double
    
    Dim dbl西药费 As Double, dbl检查费 As Double, dbl治疗费 As Double
    Dim dbl大检费 As Double, dbl大检自费 As Double
    Dim dbl特殊治疗费 As Double, dbl特殊治疗自费 As Double
    Dim dbl保险内自费费用 As Double, dbl非保险费用 As Double
    Dim dbl其它费 As Double    '针对大连开发区的
    Dim dbl血费 As Double, dbl血费自费 As Double
    Dim dbl统筹比例 As Double, dbl起付标准 As Double
    Dim lng病人ID As Long
    
    Dim str医师代码 As String
    Dim str操作员代码 As String
    Dim str治愈情况标识 As String
    Dim strTmp As String
    Dim str医生 As String
    Dim str明细 As String       '明细串
    Dim str国家编码 As String
    Dim dbl比例 As Double
    Dim str项目统计分类 As String
    Dim str项目编码 As String
    Dim dbl项目名称 As Double
    Static str结算时间 As String
    Static lng病人id1 As Long
On Error GoTo ErrH
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '明细字段
    '   病人ID,收费类别,收据费目,计算单位,开单人,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保,摘要,是否急诊
    
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '注意：接口规定，门诊明细需结算后上传；住院明细需预结算时上传
    
    '将保险支付大类读在本地,以便计算保费及自费
    gstrSQL = "Select * From 保险支付大类"
    zlDatabase.OpenRecordset rs大类, gstrSQL, "保险支付大类"
    
    Dim rs特准项目 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lng病种ID As Long
    With rs明细
    
        '主要处理多收费单据
        If str结算时间 <> Format(!结算时间, "yyyy-mm-dd HH:MM:SS") Or lng病人id1 <> Nvl(!病人ID, 0) Then
              str结算时间 = Format(!结算时间, "yyyy-mm-dd HH:MM:SS")
              lng病人id1 = Nvl(!病人ID, 0)
              g病人身份_大连.结算开始 = True
        Else
              g病人身份_大连.结算开始 = False
        End If
    
        
        '确定病种
        If Not .EOF Then
            lng病人ID = Nvl(!病人ID, 0)
            gstrSQL = "  select 病种id from 保险帐户 where 病人id=" & lng病人ID & "  and 险类=" & intinsure & "  and 医保号='" & g病人身份_大连.个人编号 & "'"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病种信息"
            If Not rsTemp.EOF Then
                lng病种ID = Nvl(rsTemp!病种ID, 0)
            Else
                lng病种ID = 0
            End If
          '打开特准项目
            gstrSQL = "Select * from 保险特准项目  where 病种ID=  " & lng病种ID
            zlDatabase.OpenRecordset rs特准项目, gstrSQL, "获取病种项目数据"
            
        End If
        
        '取出本次发生费用的金额合计
        Do While Not .EOF
            '---周顺利,对金额是否为负数进行判断,如果为负数不准执行医保收费
            If !实收金额 < 0 Then
                ShowMsgbox "该单据中包含有金额为负数的项目,不能执行医保收费!请检查后重新收费"
                门诊虚拟结算_大连 = False
                Exit Function
            End If
            
            If lng病种ID <> 0 Then
                    '第一步,确定允许的收费细目
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=0 And 性质=1 and 收费细目id=" & Nvl(!收费细目ID, 0)
                    If rs特准项目.EOF Then
                        gstrSQL = "Select 编码,名称 from 收费细目 where id=" & Nvl(!收费细目ID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取收费细目"
                        Err.Raise 9000, gstrSysName, "收费细目为“" & Nvl(rsTemp!名称) & "”的项目不是病种中所设定的项目."
                        Exit Function
                    End If
                    
                    '第二步,确定允许的保险大类
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=1 And 性质=1 and  收费细目id=" & Nvl(!保险支付大类ID, 0)
                    If rs特准项目.EOF Then
                        Err.Raise 9000, gstrSysName, "在结算中存在了结算以外的保险支付大类,不能继续。"
                        Exit Function
                    End If
                    '第三步,'确定禁止的收费细目
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=0 And 性质=2 and 收费细目id=" & Nvl(!收费细目ID, 0)
                    If Not rs特准项目.EOF Then
                        gstrSQL = "Select 编码,名称 from 收费细目 where id=" & Nvl(!收费细目ID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取收费细目"
                        Err.Raise 9000, gstrSysName, "收费细目为“" & Nvl(rsTemp!名称) & "”的项目是被禁止使用的项目." & vbCrLf & "不能继续!"
                        Exit Function
                    End If
                    '第四步,'确定禁止的大类
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=1 And 性质=2 and 收费细目id=" & Nvl(!保险支付大类ID, 0)
                    If Not rs特准项目.EOF Then
                        Err.Raise 9000, gstrSysName, "在结算中存在了禁止使用的保险支付大类,不能继续。"
                    End If
            End If
        
            '先判断是否都设置了医保对应项目编码
            gstrSQL = " Select 项目编码,项目名称 From 保险支付项目" & _
                      " Where 险类=[1] And 收费细目ID=[2]"
                      
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否设置了对应的医保项目", intinsure, CLng(!收费细目ID))
            If rsTemp.EOF = True Then
                Err.Raise 9000, gstrSysName, "有项目未设置医保项目，不能结算。"
                Exit Function
            End If
            If str医生 = "" Then
                str医生 = Nvl(!开单人)
            End If
            
            str项目编码 = Nvl(rsTemp!项目编码)
            dbl项目名称 = Val(Nvl(rsTemp!项目名称)) / 100
            lng病人ID = Nvl(!病人ID, 0)
            
            gstrSQL = "" & _
                " Select b.参数名,b.参数值 from 收费类别 a,保险参数 b " & _
                " Where a.类别=b.参数名 and b.险类=" & intinsure & _
                "        and a.编码='" & Nvl(!收费类别) & "'"
            
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "保费计算"
            
            If rsTemp.EOF Then
                strTmp = ""
            Else
                strTmp = Nvl(rsTemp!参数值)
            End If
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                strTmp = Split(strTmp, ";")(0)
                
                '计算保费
                rs大类.Find "id=" & Nvl(!保险支付大类ID, 0), , adSearchForward, 1
                If Not rs大类.EOF Then
                    '2005-10-17 ZHQ
                    '判断门诊大病或者是门诊企离使用住院比例
                    If (g病人身份_大连.就诊分类 = 3 And IsParaBig(intinsure)) Or _
                        (IsParaQ(intinsure) And intinsure = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q") Then
                        dbl统筹比例 = Nvl(rs大类!住院比额, 0) / 100
                    Else
                        dbl统筹比例 = Nvl(rs大类!统筹比额, 0) / 100
                    End If
                Else
                    dbl统筹比例 = 1
                End If
                
                '中心为:A在职、B退休、L离休、T特诊,Q企业公费,我们默认为1在职、2退休、3离休、4特诊
                If intinsure <> TYPE_大连开发区 And g病人身份_大连.职工就医类别 = "L" _
                    And g病人身份_大连.参保类别3 = "0" And Nvl(!是否医保, 0) = 1 Then  '是企保和离休人员且是医保项目
                    '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
                    '  大连市  企业单位离休医保：不完全执行医保政策，如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
                    dbl统筹比例 = 1
                End If
                
                If intinsure = TYPE_大连市 And (g病人身份_大连.职工就医类别 = "L" Or _
                     g病人身份_大连.职工就医类别 = "T") Then
                    '如果是L离休和T特诊的就按事业比例计算
                    dbl统筹比例 = dbl项目名称
                End If
                
                If intinsure = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                    '如果是Q企业公费,如果比例为100自费,则需放入非保险费用中
                    If dbl统筹比例 = 0 Then
                        '自费100
                        strTmp = ""
                    Else
                        '自费部分放入 保险内自费费用中
                    End If
                End If
                                
                '周海全调试 2003-12-17
                '对于特治项目，只要是标识为“特治”的，不应该再区分类别
                'If NVL(!收费类别) = "治疗" And str项目编码 = "特治" Then
                If str项目编码 = "特治" Then
                    strTmp = "特殊治疗费"
                End If
                If str项目编码 = "大检" Then
                    strTmp = "大检费"
                End If
                '计算扣除自费部分的费用.因为只有床位费用是定额报销,而且在门诊不会发生床位费用
                '一旦门诊发生床位费用,将按照统筹比例=0进行计算
                If Not rsTemp.EOF Then
                    If dbl统筹比例 <> 0 Then
                        Select Case strTmp
                            Case "诊察费"
                                
                                dbl诊察费 = dbl诊察费 + Round(Nvl(!实收金额, 0) * dbl统筹比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                               
                            Case "草药费"
                                
                                dbl草药费 = dbl草药费 + Round(Nvl(!实收金额, 0) * dbl统筹比例, 5)
                                If intinsure = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                                Else    '开发区升级
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                                End If
                            Case "成药费"
                                
                                dbl成药费 = dbl成药费 + Round(Nvl(!实收金额, 0) * dbl统筹比例, 5)
                                If intinsure = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                                Else    '开发区升级
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                                End If
                            Case "西药费"
                                
                                dbl西药费 = dbl西药费 + Round(Nvl(!实收金额, 0) * dbl统筹比例, 5)
                                If intinsure = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                                Else    '开发区升级
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                                End If
                               
                            Case "检查费"
                                
                                dbl检查费 = dbl检查费 + Round(Nvl(!实收金额, 0) * dbl统筹比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                                
                            Case "治疗费"
                                
                                dbl治疗费 = dbl治疗费 + Round(Nvl(!实收金额, 0) * dbl统筹比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)

                            Case "大检费"
                                '大连市和开发区对大检费用处理不同,
                                '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                          
                                dbl大检费 = dbl大检费 + Round(Nvl(!实收金额 * dbl统筹比例, 0), 5)
                                
                                dbl大检自费 = dbl大检自费 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                                
                            Case "血费"
                                '大连市和开发区对大检费用处理不同,
                                '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                          
                                dbl血费 = dbl血费 + Round(Nvl(!实收金额 * dbl统筹比例, 0), 5)
                                dbl血费自费 = dbl血费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                            Case "特殊治疗费"
                                '2004/9/11以前:大连市与开发区计算方式不一致，大连市是总额，开发区仅是统筹部分
                                '2004/9/11以后:大连市与开发区计算方式一致，都是统筹部分的金额
                                'If intinsure = TYPE_大连市 Then
                                '   dbl特殊治疗费 = dbl特殊治疗费 + Round(Nvl(!实收金额, 0), 5)
                                'Else
                                    dbl特殊治疗费 = dbl特殊治疗费 + Round(Nvl(!实收金额, 0) * dbl统筹比例, 5)
                                'End If
                                
                                '特殊治疗费自费的计算方式相同,
                                '大连市接口在处理汇总金额时只把特殊治疗费进行汇总,特殊治疗自费部分不再记入
                                dbl特殊治疗自费 = dbl特殊治疗自费 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                            
                        End Select
                    Else
                        '全部是比例为0的项目(包括包床),分别对大连还是开发区进行判断放在不同的字段
                        If intinsure = TYPE_大连市 Then
                            '大连市放在dbl非保险费用
                            dbl非保险费用 = dbl非保险费用 + Round(!实收金额, 5)
                        Else
                            '开发区放在dbl其它费
                            dbl其它费 = dbl其它费 + Round(!实收金额, 5)
                        End If
                    
                    End If

                End If
            End If
            curTotal = curTotal + Round(Nvl(!实收金额, 0), 5)
            .MoveNext
        Loop
    End With
    
    '计算起付线
    If str医生 <> "" Then
        gstrSQL = "Select 编号 From 人员表  where 姓名='" & str医生 & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取医生编号"
        If Not rsTemp.EOF Then
            str医生 = Nvl(rsTemp!编号)
            If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
                str医生 = Substr(str医生, 1, 6)
            End If
        Else
            str医生 = ""
        End If
    End If
    
    '门诊的起付线为零
    If 大连医保结算(intinsure, 0, lng病人ID, 0, 0, 0, False, True, 0, dbl诊察费, dbl草药费, dbl成药费, dbl西药费, dbl检查费, dbl治疗费, dbl血费, dbl血费自费, dbl大检费, dbl大检自费, dbl特殊治疗费, dbl特殊治疗自费, dbl保险内自费费用, dbl非保险费用, dbl其它费, dbl药费自费, curTotal, str医生, str结算方式) = False Then
        Exit Function
    End If
'    Call WriteDebugInfor_大连("门诊虚拟结算_大连", lng病人id)
    门诊虚拟结算_大连 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function Get结算方式(ByVal strOutput As String, ByVal intinsure As Integer) As String
    '功能:获取结算方式
    '参数:strOutPut-输出串
    '返回:结算方式
    '开发区:
    '    本次基本个人帐户支付    NUM 225 10      中心返回
    '    本次补助个人帐户支付    NUM 235 10      中心返回
    '    本次基本统筹支付    NUM 245 10      中心返回
    '    本次基本统筹自付    NUM 255 10      中心返回
    '    本次补充统筹支付    NUM 265 10      中心返回
    '    本次补充统筹自付    NUM 275 10      中心返回
    '    本次基本补助保险支付    NUM 285 10  公务员补助该字段包括门槛费补助部分和基本统筹自付部分的公务员补助支付 中心返回
    '    本次非基本补助保险支付  NUM 295 10  公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付，该部分（超过基本统筹最高限额部分）除去公务员补助支付后，全部打入"本次保险范围外自付"部分  中心返回
    '    本次保险范围外自付  NUM 305 10  限额以外＋门槛费自付部分（个人帐户充抵后）＋各种自费去掉补助部分    中心返回
 
    Dim i As Long
    Dim str结算方式 As String
    
    If intinsure = TYPE_大连开发区 Then
    
        '刘洋 2005-08-16 开发区医保升级
        'i = 225 - 10
         i = 275 - 10
        '确定本次结算方式
        str结算方式 = "个人帐户;" & Format(Val(Substr(strOutput, i + 10, 10)) + Val(Substr(strOutput, i + 20, 10)), "###0.00;-###0.00;0;0") & ";0" '本次基本个人帐户支付,不充许修改
        str结算方式 = str结算方式 & "|" & "基本统筹;" & Format(Val(Substr(strOutput, i + 30, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
        str结算方式 = str结算方式 & "|" & "补充统筹;" & Format(Val(Substr(strOutput, i + 50, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
        
        '刘洋 2005-08-16 开发区医保升级,取消了补助保险和非补助保险,增加了公务员补助和商业保险补助
        '公务员补助:=本次公务员起付标准补助支付+本次公务员基本补助保险支付+本次公务员非基本补助保险支付
        'Val (Substr(strOutPut, i + 70, 10)) + Val(Substr(strOutPut, i + 80, 10)) + Val(Substr(strOutPut, i + 90, 10))
        'str结算方式 = str结算方式 & "|" & "补助保险;" & Format(Val(Substr(strOutPut, i + 70, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
        'str结算方式 = str结算方式 & "|" & "非补助保险;" & Format(Val(Substr(strOutPut, i + 80, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
        str结算方式 = str结算方式 & "|" & "公务员补助;" & Format(Val(Substr(strOutput, i + 70, 10)) + Val(Substr(strOutput, i + 80, 10)) + Val(Substr(strOutput, i + 90, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
        str结算方式 = str结算方式 & "|" & "商业保险补助;" & Format(Val(Substr(strOutput, i + 100, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改

        Get结算方式 = str结算方式
        Exit Function
    End If
    
   '大连市:
    '   3.0.2版
    '    本次基本个人帐户支付    NUM 211 10  如果是慢病结算，表示慢病帐户支付    中心
    '    本次补助个人帐户支付    NUM 221 10  如果是慢病结算返回0 中心
    '    本次基本统筹支付    NUM 231 10      中心
    '    本次基本统筹自付    NUM 241 10      中心
    '    本次补充统筹支付    NUM 251 10  如果是生育结算，本字段用于存放生育保险支付  中心
    '    本次补充统筹自付    NUM 261 10      中心
    '    本次基本补助保险支付    NUM 271 10  1． 如果是商业保险该字段包括基本统筹自付部分的商业保险支付 2． 如果是公务员补助该字段包括门槛费补助部分、基本统筹自付部分的公务员补助支付；基本统筹最高限额内公务员补助支付后剩余款项计入"本次保险范围外自付"部分  中心
    '    本次非基本补助保险支付  NUM 281 10  1． 如果是商业保险该字段是补充统筹自付部分的商业保险支付   2． 如果是公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付；超过基本统筹最高限额公务员补助支付后，剩余款项计入"本次保险范围外自付"部分    中心
    '    本次保险范围外自付  NUM 291 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋保险内自费费用＋非保险费用+大检自费   中心
    '    门诊预结时将个人帐户和补助个人帐户合计后作为预结算时帐户支付金额的判断值
    '   4.0.1版
    '   本次基本个人帐户支付    NUM 241 10  如果是慢病结算，表示慢病帐户支付
    '   本次补助个人帐户支付    NUM 251 10  如果是慢病结算返回0
    '   本次基本统筹支付    NUM 261 10
    '   本次基本统筹自付    NUM 271 10
    '   本次补充统筹支付    NUM 281 10  如果是生育结算，本字段用于存放生育保险支付
    '   本次补充统筹自付    NUM 291 10
    '   本次公务员起付标准补助支付  NUM 301 10
    '   本次公务员基本补助保险支付  NUM 311 10
    '   本次公务员非基本补助保险支付    NUM 321 10
    '   本次商业保险补助支付    NUM 331 10
    '   本次保险内自付  NUM 341 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋（保险内自费费用+大检自费+血费自费+特治自费）的个人账户冲抵后的费用
    '   本次非保险自付  NUM 351 10
    i = 241 - 10
    Dim dblMoney As Double
    With g病人身份_大连
        .当前个人帐户余额 = .当前个人帐户余额 - Format(Val(Substr(strOutput, i + 10, 10)), "###0.00;-###0.00;0;0")
        .当前补助帐户余额 = .当前补助帐户余额 - Format(Val(Substr(strOutput, i + 20, 10)), "###0.00;-###0.00;0;0")
        .当前统筹累计 = .当前统筹累计 + Format(Val(Substr(strOutput, i + 30, 10)), "###0.00;-###0.00;0;0") + Format(Val(Substr(strOutput, i + 50, 10)), "###0.00;-###0.00;0;0")
    End With
    
    '确定本次结算方式
    str结算方式 = "个人帐户;" & Format(Val(Substr(strOutput, i + 10, 10)) + Val(Substr(strOutput, i + 20, 10)), "###0.00;-###0.00;0;0") & ";0" '本次基本个人帐户支付,不充许修改
    str结算方式 = str结算方式 & "|" & "基本统筹;" & Format(Val(Substr(strOutput, i + 30, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    str结算方式 = str结算方式 & "|" & "补充统筹;" & Format(Val(Substr(strOutput, i + 50, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    '公务员补助:=本次公务员起付标准补助支付+本次公务员基本补助保险支付+本次公务员非基本补助保险支付
    dblMoney = Val(Substr(strOutput, i + 70, 10)) + Val(Substr(strOutput, i + 80, 10)) + Val(Substr(strOutput, i + 90, 10))
    
    '2004/09/11刘兴宏,取消了补助保险和非补助保险,增加了公务员补助和商业保险补助
    str结算方式 = str结算方式 & "|" & "公务员补助;" & Format(dblMoney, "###0.00;-###0.00;0;0") & ";0" '不充许修改
    str结算方式 = str结算方式 & "|" & "商业保险补助;" & Format(Val(Substr(strOutput, i + 100, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    
    'str结算方式 = str结算方式 & "|" & "补助保险;" & Format(Val(Substr(strOutPut, i + 70, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    'str结算方式 = str结算方式 & "|" & "非补助保险;" & Format(Val(Substr(strOutPut, i + 80, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    '新增企业离休病人结算方式为'离休统筹',并且对离休病人的离休统筹=非保险费用
    If g病人身份_大连.职工就医类别 = "Q" Then
        'str结算方式 = str结算方式 & "|" & "离休拨付;" & Format(Round(dbl诊察费 + dbl草药费 + dbl成药费 + dbl西药费 + dbl检查费 + dbl治疗费 + dbl大检费 + dbl特殊治疗费 + dbl大检自费 + dbl保险内自费费用, 2), "###0.00;-###0.00;0;0") & ";0" '不充许修改
        '除了"本次非保险自付"以外的所有以"自付"为结尾的字段之合即为医院应该替患者负担的费用
        '离休拨付=本次基本统筹自付+本次补充统筹自付+本次保险内自付
        dblMoney = Val(Substr(strOutput, i + 40, 10)) + Val(Substr(strOutput, i + 60, 10)) + Val(Substr(strOutput, i + 110, 10))
        str结算方式 = str结算方式 & "|" & "离休拨付;" & Format(dblMoney, "###0.00;-###0.00;0;0") & ";0" '不充许修改
    End If
    Get结算方式 = str结算方式
    '现金=本次非保险自付
End Function
Public Function 门诊记帐虚拟结算_大连(rs明细 As ADODB.Recordset, str结算方式 As String, ByVal intinsure As Integer) As Boolean
    '参数：rsExse     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '明细字段
    '记录性质,记录状态,NO,序号,门诊标志,病人ID,主页ID,婴儿费,
    '医保项目编码,保险大类ID,收费类别,收费细目ID,收费名称,
    '计算单位,开单部门,规格,产地,数量,Decode(A.数量,0,0,Round(A.金额/A.数量,4)) as 价格,
    'A.金额,A.医生,A.发生时间 as 登记时间,是否上传 ,是否急诊,保险项目否,摘要
    
    
    Dim curTotal As Double, cur个人帐户 As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rs大类 As New ADODB.Recordset
    Dim rs收费细目 As New ADODB.Recordset
    
    Dim strInfor As String  '定义中心返回串
    Dim dbl诊察费 As Double, dbl草药费 As Double, dbl成药费 As Double, dbl西药费 As Double
    
    '2005-08-02开发区升级
    Dim dbl药费自费 As Double
    
    Dim dbl检查费 As Double, dbl治疗费 As Double, dbl大检费 As Double, dbl大检自费 As Double
    Dim dbl特殊治疗费 As Double, dbl特殊治疗自费 As Double
    Dim dbl保险内自费费用 As Double, dbl非保险费用 As Double, dbl统筹比例 As Double
    Dim dbl血费 As Double, dbl血费自费 As Double
    Dim dbl其它费 As Double, dbl起付标准 As Double
    Dim lng病人ID As Long
    
    Dim str诊断编码 As String, str医师代码 As String, str操作员代码 As String
    Dim str诊断名称 As String, str治愈情况标识 As String
    Dim strTmp As String, str医生 As String, str明细 As String      '明细串
    Dim str国家编码 As String, str项目统计分类 As String, str项目编码 As String, dbl项目名称 As Double
    '----------------对rsExse中部分不正确的字段重新赋值
    Dim int大类id As Integer
    Dim int保险项目否 As Integer
    '--------------------------------------------------
    
    Dim intMouse As Integer
      
    intMouse = Screen.MousePointer
    int大类id = 0
    int保险项目否 = 0
    
    '在虚拟结算前需验证身分
    Screen.MousePointer = 1
    If 身份标识_大连(0, lng病人ID, intinsure) = "" Then
        Screen.MousePointer = intMouse
        门诊记帐虚拟结算_大连 = False
        MsgBox "病人身份身份验证失败,不能进行结算"
        Exit Function
    End If
    Screen.MousePointer = intMouse

    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '注意：接口规定，门诊明细需结算后上传；住院明细需预结算时上传
    
    '将保险支付大类读在本地,以便计算保费及自费
    gstrSQL = "Select * From 保险支付大类 "
    zlDatabase.OpenRecordset rs大类, gstrSQL, "保险支付大类"
    
    Dim rs特准项目 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lng病种ID As Long
    With rs明细
        '确定病种
        If Not .EOF Then
            lng病人ID = Nvl(!病人ID, 0)
            gstrSQL = "  select 病种id from 保险帐户 where 病人id=" & lng病人ID & "  and 险类=" & intinsure & "  and 医保号='" & g病人身份_大连.个人编号 & "'"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病种信息"
            If Not rsTemp.EOF Then
                lng病种ID = Nvl(rsTemp!病种ID, 0)
            Else
                lng病种ID = 0
            End If
          '打开特准项目
            gstrSQL = "Select * from 保险特准项目  where 病种ID=  " & lng病种ID
            zlDatabase.OpenRecordset rs特准项目, gstrSQL, "获取病种项目数据"
            
        End If
        
        '取出本次发生费用的金额合计
        Do While Not .EOF
            If lng病种ID <> 0 Then
                    '第一步,确定允许的收费细目
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=0 And 性质=1 and 收费细目id=" & Nvl(!收费细目ID, 0)
                    If rs特准项目.EOF Then
                        gstrSQL = "Select 编码,名称 from 收费细目 where id=" & Nvl(!收费细目ID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取收费细目"
                        ShowMsgbox "收费细目为“" & Nvl(rsTemp!名称) & "”的项目不是病种中所设定的项目."
                        Exit Function
                    End If
                    
                    '第二步,确定允许的保险大类
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=1 And 性质=1 and  收费细目id=" & Nvl(!保险支付大类ID, 0)
                    If rs特准项目.EOF Then
                        ShowMsgbox "在结算中存在了结算以外的保险支付大类,不能继续。"
                        Exit Function
                    End If
                    '第三步,'确定禁止的收费细目
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=0 And 性质=2 and 收费细目id=" & Nvl(!收费细目ID, 0)
                    If Not rs特准项目.EOF Then
                        gstrSQL = "Select 编码,名称 from 收费细目 where id=" & Nvl(!收费细目ID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取收费细目"
                        ShowMsgbox "收费细目为“" & Nvl(rsTemp!名称) & "”的项目是被禁止使用的项目." & vbCrLf & "不能继续!"
                        Exit Function
                    End If
                    '第四步,'确定禁止的大类
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=1 And 性质=2 and 收费细目id=" & Nvl(!保险支付大类ID, 0)
                    If Not rs特准项目.EOF Then
                        ShowMsgbox "在结算中存在了禁止使用的保险支付大类,不能继续。"
                    End If
            End If
        
            '先判断是否都设置了医保对应项目编码
            gstrSQL = " Select 项目编码,项目名称,大类id,是否医保 From 保险支付项目" & _
                      " Where 险类=[1] And 收费细目ID=[2]"
                      
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否设置了对应的医保项目", intinsure, CLng(!收费细目ID))
            
            If rsTemp.EOF = True Then
                MsgBox "有项目未设置医保项目，不能结算。", vbInformation, gstrSysName
                Exit Function
            Else
                '由于门诊记帐项目上对费用记录没有填写保险大类id和保险项目否的字段,
                '因此提取出来的这两个字段不能正确反映真实值,需要在本次的选择中完成重新赋值
                int大类id = Nvl(rsTemp!大类id, 0)
                int保险项目否 = Nvl(rsTemp!是否医保, 0)
            End If
            
            If str医生 = "" Then
                str医生 = Nvl(!医生)
            End If
            
            str项目编码 = Nvl(rsTemp!项目编码)
            dbl项目名称 = Val(Nvl(rsTemp!项目名称)) / 100
            
            lng病人ID = Nvl(!病人ID, 0)
            gstrSQL = "" & _
                " Select b.参数名,b.参数值 from 收费类别 a,保险参数 b " & _
                " Where a.类别=b.参数名 and b.险类=" & intinsure & _
                "        and a.编码='" & Nvl(!收费类别) & "'"
            
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "保费计算"
            
            If rsTemp.EOF Then
                strTmp = ""
            Else
                strTmp = Nvl(rsTemp!参数值)
            End If
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                strTmp = Split(strTmp, ";")(0)
                
                '计算保费
                rs大类.Find "id=" & int大类id, , adSearchForward, 1
                If Not rs大类.EOF Then
                    dbl统筹比例 = Nvl(rs大类!统筹比额, 0) / 100
                Else
                    dbl统筹比例 = 1
                End If
                '中心为:A在职、B退休、L离休、T特诊,Q企业公费,我们默认为1在职、2退休、3离休、4特诊
                If intinsure <> TYPE_大连开发区 And g病人身份_大连.职工就医类别 = "L" _
                    And g病人身份_大连.参保类别3 = "0" And Nvl(int保险项目否, 0) = 1 Then  '是企保和离休人员且是医保项目
                    '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
                    '  大连市  企业单位离休医保：不完全执行医保政策，如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
                    dbl统筹比例 = 1
                End If
                
                If intinsure = TYPE_大连市 And (g病人身份_大连.职工就医类别 = "L" Or _
                     g病人身份_大连.职工就医类别 = "T") Then
                    '如果是L离休和T特诊的就按事业比例计算
                    dbl统筹比例 = dbl项目名称
                End If
                                                
                '周海全调试 2003-12-17
                '对于特治项目，只要是标识为“特治”的，不应该再区分类别
                'If NVL(!收费类别) = "治疗" And str项目编码 = "特治" Then
                If str项目编码 = "特治" Then
                    strTmp = "特殊治疗费"
                End If
                If str项目编码 = "大检" Then
                    strTmp = "大检费"
                End If
                '计算扣除自费部分的费用.因为只有床位费用是定额报销,而且在门诊不会发生床位费用
                '一旦门诊发生床位费用,将按照统筹比例=0进行计算
                If Not rsTemp.EOF Then
                    If dbl统筹比例 <> 0 Then
                        Select Case strTmp
                            Case "诊察费"
                                
                                dbl诊察费 = dbl诊察费 + Round(Nvl(!金额, 0) * dbl统筹比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)
                               
                            Case "草药费"
                                
                                dbl草药费 = dbl草药费 + Round(Nvl(!金额, 0) * dbl统筹比例, 5)
                                
                                '2005-08-02开发区升级
                                If intinsure = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)
                                Else
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)
                                End If
                                
                            Case "成药费"
                                
                                dbl成药费 = dbl成药费 + Round(Nvl(!金额, 0) * dbl统筹比例, 5)
                                '2005-08-02开发区升级
                                If intinsure = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)
                                Else
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)
                                End If
                                
                            Case "西药费"
                                
                                dbl西药费 = dbl西药费 + Round(Nvl(!金额, 0) * dbl统筹比例, 5)
                                '2005-08-02开发区升级
                                If intinsure = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)
                                Else
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)
                                End If
                                
                            Case "检查费"
                                
                                dbl检查费 = dbl检查费 + Round(Nvl(!金额, 0) * dbl统筹比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)
                                
                            Case "治疗费"
                                
                                dbl治疗费 = dbl治疗费 + Round(Nvl(!金额, 0) * dbl统筹比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)

                                '周海全调试 2003-12-17
                                '大检费在医保参数设置中无法对应项目，这里是如何取得的？
                            Case "大检费"
                                '大连市和开发区对大检费用处理不同,
                                '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                          
                                dbl大检费 = dbl大检费 + Round(Nvl(!金额 * dbl统筹比例, 0), 5)
                                
                                dbl大检自费 = dbl大检自费 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)
                            Case "血费"
                                '大连市和开发区对大检费用处理不同,
                                '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                          
                                dbl血费 = dbl血费 + Round(Nvl(!实收金额 * dbl统筹比例, 0), 5)
                                dbl血费自费 = dbl血费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl统筹比例), 5)
                            Case "特殊治疗费"
                                '2004/9/11以前:大连市与开发区计算方式不一致，大连市是总额，开发区仅是统筹部分
                                '2004/9/11以后:大连市与开发区计算方式一致，都是统筹部分的金额
                                'If intinsure = TYPE_大连市 Then
                                '   dbl特殊治疗费 = dbl特殊治疗费 + Round(Nvl(!实收金额, 0), 5)
                                'Else
                                    dbl特殊治疗费 = dbl特殊治疗费 + Round(Nvl(!实收金额, 0) * dbl统筹比例, 5)
                                'End If
                                
                                '特殊治疗费自费的计算方式相同,
                                '大连市接口在处理汇总金额时只把特殊治疗费进行汇总,特殊治疗自费部分不再记入
                                dbl特殊治疗自费 = dbl特殊治疗自费 + Round(Nvl(!金额, 0) * (1 - dbl统筹比例), 5)
                                
                        End Select
                    Else
                
                        '全部是比例为0的项目(包括包床),分别对大连还是开发区进行判断放在不同的字段
                        If intinsure = TYPE_大连市 Then
                            '大连市放在dbl非保险费用
                            dbl非保险费用 = dbl非保险费用 + Round(!金额, 5)
                        Else
                            '开发区放在dbl其它费
                            dbl其它费 = dbl其它费 + Round(!金额, 5)
                        End If
                    
                    End If

                End If
            End If
            curTotal = curTotal + Round(Nvl(!金额, 0), 5)
            .MoveNext
        Loop
    End With
  
    
    If str医生 <> "" Then
        gstrSQL = "Select 编号 From 人员表  where 姓名='" & str医生 & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取医生编号"
        If Not rsTemp.EOF Then
            str医生 = Nvl(rsTemp!编号)
            If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
                str医生 = Substr(str医生, 1, 6)
            End If
        Else
            str医生 = ""
        End If
    End If
    
    '门诊的起付线为零
    dbl起付标准 = g病人身份_大连.起付线
    If 大连医保结算(intinsure, 0, lng病人ID, 0, 0, 0, False, True, dbl起付标准, dbl诊察费, dbl草药费, dbl成药费, dbl西药费, _
        dbl检查费, dbl治疗费, dbl血费, dbl血费自费, dbl大检费, dbl大检自费, dbl特殊治疗费, dbl特殊治疗自费, dbl保险内自费费用, _
        dbl非保险费用, dbl其它费, dbl药费自费, curTotal, str医生, str结算方式) = False Then
        Exit Function
    End If
    
    门诊记帐虚拟结算_大连 = True
End Function

Private Function 门诊记帐结算及冲帐_大连(ByVal bln冲销 As Boolean, ByVal lng病人ID As Long, ByVal lng结帐ID As Long, ByVal 原结帐id As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean

  '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；
    '      cur个人帐户   从个人帐户中支出的金额
    
    Dim curTotal As Double
    Dim rsTemp As New ADODB.Recordset, rs明细 As New ADODB.Recordset
    Dim strInfor As String  '定义中心返回串
    Dim dbl诊察费 As Double, dbl草药费 As Double, dbl成药费 As Double, dbl西药费 As Double
    
    '2005-08-02开发区升级
    Dim dbl药费自费 As Double
    
    Dim dbl检查费 As Double, dbl治疗费 As Double
    Dim dbl大检费 As Double, dbl大检自费 As Double, dbl特殊治疗费 As Double, dbl特殊治疗自费 As Double, dbl保险内自费费用 As Double
    Dim dbl非保险费用 As Double, dbl其它费 As Double    '针对大连开发区的
    Dim dbl血费 As Double, dbl血费自费 As Double
    Dim dbl起付标准 As Double, dbl比例 As Double
    Dim str医生 As String, str明细 As String      '明细串
    Dim str国家编码 As String, str项目编码 As String, str项目统计分类 As String, strTmp As String
    Dim int业务 As Integer, lng冲销ID As Long
    Dim strNO As String, lng记录性质 As Long
    
    Dim lng治疗序号 As Long, dbl个人帐户余额 As Double
    Dim dbl统筹支付累计 As Double, dbl个人帐户支付 As Double
    Dim dbl补助帐户支付 As Double
    Dim dbl基本统筹支付 As Double
    Dim dbl基本统筹自付 As Double
    Dim dbl补充统筹支付 As Double
    Dim dbl补充统筹自付 As Double
    Dim dbl补助保险支付 As Double
    Dim dbl非补助保险支付 As Double
    Dim dbl保险范围外自付 As Double
    
    Dim dbl结算前基本帐户余额  As Double
    Dim dbl结算前补助账户余额  As Double
    Dim dbl结算前统筹累计  As Double
    Dim lngTmp As Long
    Dim rs特准项目 As New ADODB.Recordset
    Dim lng病种ID As Long
    
    int业务 = IIf(bln冲销, 1, 0)
     门诊记帐结算及冲帐_大连 = False
   
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '注意：接口规定，门诊明细需结算后上传；住院明细需预结算时上传，如果卡内金额不足，可以使用圈存接口，即将卡外的钱，调到卡内，以增加卡内金额
    '卡内余额需要通过卡操作函数读取，可圈存金额是接口返回，需要修改
    
    On Error GoTo errHand
    
    '重新读卡
    If 读取病人身份_大连(IIf(intinsure = TYPE_大连开发区, 2, 1), intinsure) = False Then
        Exit Function
    End If
    
    If bln冲销 Then
        lng冲销ID = 原结帐id
        '验证是否为该病人的IC卡
        gstrSQL = "Select * From  保险帐户 where 病人id=" & lng病人ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人的医保号"
        If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "该病人在保险帐户中无记录!"
            Exit Function
        End If
        
        If g病人身份_大连.IC卡号 <> Nvl(rsTemp!卡号) Then
            Err.Raise 9000, gstrSysName, "该病人的IC卡插入错误,可能是插入了其他人的IC卡!"
            Exit Function
        End If
        '确定就诊分类,转诊单号,诊断编码,诊断名称
        ' 支付顺序号_IN(就诊分类;转诊单号;诊断编码),备注(诊断名称_IN)
        gstrSQL = "Select 支付顺序号,备注 from 保险结算记录  where 记录ID=" & lng冲销ID
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取就诊分类"
        If rsTemp.RecordCount = 0 Then
            Err.Raise 9000, gstrSysName, "在结算记录中无结算记录!"
            Exit Function
        End If
        Dim strArr
        strArr = Split(Nvl(rsTemp!支付顺序号), ";")
        
        '就诊分类;转诊单号;诊断编码
        '1-普通门诊("1", "A"),2-急诊门诊("3", "7")
        '3-门诊大病("5", "B"),4-门诊慢病补助("S", "T")
        If UBound(strArr) >= 2 Then
            g病人身份_大连.就诊分类 = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g病人身份_大连.转诊单号 = strArr(1)
            g病人身份_大连.诊断编码 = strArr(2)
        ElseIf UBound(strArr) = 1 Then
            g病人身份_大连.就诊分类 = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g病人身份_大连.转诊单号 = strArr(1)
        Else
            g病人身份_大连.就诊分类 = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
        End If
        g病人身份_大连.诊断名称 = Nvl(rsTemp!备注)
        
        
        '确定退费记录
        '退费
        gstrSQL = "select id from 门诊费用记录 where 结帐id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊退费", lng冲销ID)
        If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "不存在病人费用冲销记录!"
            Exit Function
        End If
         
    End If
    '打开本次结算明细记录 '--国家编码应该是标识主码+子码
    gstrSQL = " " & _
        "  Select Rownum 标识号,A.ID,A.病人ID,A.收费细目id,A.NO,A.序号,A.记录性质,A.记录状态,A.登记时间,A.开单人 as 医生,H.编号 as 医生编号, " & _
        "      A.数次*A.付数 as 数量,A.计算单位,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额 as 实收金额,F.参数值,G.id as 大类id,G.统筹比额, " & _
        "      a.医嘱序号, A.收费类别,B.编码 as 项目编码,B.名称 as 项目名称,Nvl(J.标识码,Nvl(B.标识主码||B.标识子码,B.编码)) as 国家编码, " & _
        "      D.项目编码 医保编码,D.项目名称 as 医保名称,J.名称 as 剂型,D.是否医保,C.名称 开单部门,E.名称 受单部门, " & _
        "      L.险类,L.中心,L.卡号,L.医保号,L.人员身份,L.单位编码,L.顺序号,L.退休证号,L.帐户余额,L.当前状态,L.病种ID,L.在职,L.年龄段,L.灰度级,L.就诊时间 " & _
        "  From (Select * From 门诊费用记录 Where 记录状态<>0 and 结帐ID=" & IIf(bln冲销, lng冲销ID, lng结帐ID) & " and  Nvl(附加标志,0)<>9 ) A,收费细目 B,部门表 C,保险支付项目 D,部门表 E,  " & _
        "       (Select U.*,K.参数值 From 收费类别 U,保险参数 K where U.类别=K.参数名 and K.险类=" & intinsure & "  ) F, " & _
        "       (Select distinct Q.药品id,Q.标识码,T.名称 From 药品目录 Q,药品信息 R,药品剂型 T  Where  Q.药名id=R.药名id and R.剂型=T.编码 ) J, " & _
        "       保险支付大类 G,人员表 H,保险帐户 L" & _
        "  Where A.收费细目ID=B.ID And A.开单部门ID=C.ID(+) and A.病人id=L.病人id and L.险类=" & intinsure & " and A.收费类别=F.编码(+)  and d.大类id=G.id and a.收费细目id=J.药品id(+) " & _
        "        And A.执行部门ID=E.ID(+) And A.收费细目ID=D.收费细目ID And D.险类= " & intinsure & " and a.开单人=H.姓名(+) " & _
        "  Order by A.ID"
        
    '上传费用明细记录
    zlDatabase.OpenRecordset rs明细, gstrSQL, "读取本次结帐费用明细"
    
    With rs明细
        If Not .EOF Then
            lng病人ID = Nvl(!病人ID, 0)
            str医生 = Nvl(!医生编号)
            If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
                str医生 = Substr(str医生, 1, 6)
            End If
            lng病种ID = Nvl(!病种ID, 0)
            '打开特准项目
            gstrSQL = "Select * from 保险特准项目  where 病种ID=  " & lng病种ID
            zlDatabase.OpenRecordset rs特准项目, gstrSQL, "获取病种项目数据"
        End If
        
        Do While Not .EOF
            If lng病种ID <> 0 And bln冲销 = False Then
                '第一步,确定允许的收费细目
                rs特准项目.Filter = 0
                rs特准项目.Filter = "大类=0 And 性质=1 and 收费细目id=" & Nvl(!收费细目ID, 0)
                If rs特准项目.EOF Then
                    Err.Raise 9000, gstrSysName, "收费细目为“" & Nvl(!项目名称) & "”的项目不是病种中所设定的项目."
                    Exit Function
                End If
                '第二步,确定允许的保险大类
                rs特准项目.Filter = 0
                rs特准项目.Filter = "大类=1 And 性质=1 and  收费细目id=" & Nvl(!大类id, 0)
                If rs特准项目.EOF Then
                    Err.Raise 9000, gstrSysName, "在结算中存在了结算以外的保险支付大类,不能继续。"
                    Exit Function
                End If
                '第三步,'确定禁止的收费细目
                rs特准项目.Filter = 0
                rs特准项目.Filter = "大类=0 And 性质=2 and 收费细目id=" & Nvl(!收费细目ID, 0)
                If Not rs特准项目.EOF Then
                    Err.Raise 9000, gstrSysName, "收费细目为“" & Nvl(!项目名称) & "”的项目是被禁止使用的项目." & vbCrLf & "不能继续!"
                    Exit Function
                End If
                '第四步,'确定禁止的大类
                rs特准项目.Filter = 0
                rs特准项目.Filter = "大类=1 And 性质=2 and 收费细目id=" & Nvl(!大类id, 0)
                If Not rs特准项目.EOF Then
                    Err.Raise 9000, gstrSysName, "在结算中存在了禁止使用的保险支付大类,不能继续。"
                End If
            End If
            strTmp = Nvl(!参数值)
            lng病人ID = Nvl(!病人ID, 0)
            '确定相关数据
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                If Split(strTmp, ";")(1) = "" Then
                    str项目统计分类 = ""
                Else
                    str项目统计分类 = Mid(Split(strTmp, ";")(1), 1, 1)
                End If
                
                strTmp = Split(strTmp, ";")(0)
                '比例
                '中心为:A在职、B退休、L离休、T特诊,我们默认为1在职、2退休、3离休、4特诊
                    
                If Nvl(!险类, 0) <> TYPE_大连开发区 And Val(Nvl(!单位编码, "99")) = 0 And Nvl(!在职, 0) = 3 And Nvl(!是否医保, 0) = 1 Then   '是企保和离休人员且是医保项目
                    '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
                    '大连市    企业单位离休医保：不完全执行医保政策，（如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
                    dbl比例 = 1
                Else
                    dbl比例 = Nvl(!统筹比额, 0) / 100
                End If
                
                If Nvl(!险类, 0) = TYPE_大连市 And (g病人身份_大连.职工就医类别 = "L" Or _
                     g病人身份_大连.职工就医类别 = "T") Then
                    '如果是L离休和T特诊的就按事业比例计算
                    dbl比例 = Val(Nvl(!医保名称)) / 100
                End If
                
                If Nvl(!医保编码) = "特治" Then
                    strTmp = "特殊治疗费"
                End If
                If Nvl(!医保编码) = "大检" Then
                    strTmp = "大检费"
                End If

                If dbl比例 <> 0 Then
                    
                    Select Case strTmp
                            Case "诊察费"
                                dbl诊察费 = dbl诊察费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                            
                            Case "草药费"
                                dbl草药费 = dbl草药费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                
                                '2005-08-02开发区升级
                                If intinsure = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                Else
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                End If
                                
                            Case "成药费"
                                dbl成药费 = dbl成药费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                
                                '2005-08-02开发区升级
                                If intinsure = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                Else
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                End If
                                
                            Case "西药费"
                                dbl西药费 = dbl西药费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                
                                '2005-08-02开发区升级
                                If intinsure = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                Else
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                End If
                                
                            Case "检查费"
                                dbl检查费 = dbl检查费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                
                            Case "治疗费"
                                dbl治疗费 = dbl治疗费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                '周海全调试 2003-12-17
                                '大检费在医保参数设置中无法对应项目，这里是如何取得的？
                            Case "大检费"
                                '大连市和开发区对大检费用处理不同,
                                '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                          
                                dbl大检费 = dbl大检费 + Round(Nvl(!实收金额 * dbl比例, 0), 5)
                                
                                dbl大检自费 = dbl大检自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                            Case "血费"
                                '大连市和开发区对大检费用处理不同,
                                '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                          
                                dbl血费 = dbl血费 + Round(Nvl(!实收金额 * dbl比例, 0), 5)
                                dbl血费自费 = dbl血费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                            Case "特殊治疗费"
                                '2004/9/11以前:大连市与开发区计算方式不一致，大连市是总额，开发区仅是统筹部分
                                '2004/9/11以后:大连市与开发区计算方式一致，都是统筹部分的金额
                                'If intinsure = TYPE_大连市 Then
                                '   dbl特殊治疗费 = dbl特殊治疗费 + Round(Nvl(!实收金额, 0), 5)
                                'Else
                                    dbl特殊治疗费 = dbl特殊治疗费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                'End If
                                
                                '特殊治疗费自费的计算方式相同,
                                '大连市接口在处理汇总金额时只把特殊治疗费进行汇总,特殊治疗自费部分不再记入
                                dbl特殊治疗自费 = dbl特殊治疗自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                
                        End Select
                    Else
                
                        '全部是比例为0的项目(包括包床),分别对大连还是开发区进行判断放在不同的字段
                        If intinsure = TYPE_大连市 Then
                            '大连市放在dbl非保险费用
                            dbl非保险费用 = dbl非保险费用 + Round(!实收金额, 5)
                        Else
                            '开发区放在dbl其它费
                            dbl其它费 = dbl其它费 + Round(!实收金额, 5)
                        End If
                    
                    End If
            Else
                dbl比例 = 1
                str项目统计分类 = ""
            End If

            '上传明细记录,实时医疗明细数据
            '参数控制明细上传
            If gbln门诊明细时实上传 Then
                
                    If Nvl(!险类, 0) = TYPE_大连开发区 Then '开发区
                        str明细 = Lpad(gstr医院编码_大连, 6)     '医院代号    CHAR    1   6       院端填写
                        str明细 = str明细 & Lpad(Nvl(!医保号), 10)  '保险编号    CHAR    7   10      院端填写
                    Else
                        str明细 = Lpad(gstr医院编码_大连, 4)     '医院代码    CHAR    1   4       院端
                        str明细 = str明细 & Lpad(Nvl(!医保号), 8)   '个人编号    CHAR    5   8       院端
                    End If
                
                    str明细 = str明细 & Space(10)   '病志号  CHAR    13  10  门诊明细以空格补位,住院是住院号  院端
                    str明细 = str明细 & Lpad(g病人身份_大连.治疗序号, 4)     '治疗序号
                    
                    'Modified By 朱玉宝 2004-07-29 原因：处理NO号
                    str明细 = str明细 & Lpad(Mid(Nvl(!NO, "00000000"), 2, 7), 10)      '处方号  NUM 27  10      院端
                    str明细 = str明细 & Lpad(!序号, 10)       '处方项目序号    NUM 37  10  对应处方号的记价项目序号    院端
                    
                    '开发区为单据号  CHAR    41  10  医嘱号，    院端填写
                    str明细 = str明细 & Space(10)       '医嘱号  CHAR    47  10  处方对应医嘱的医嘱记录号，门诊明细或没有医嘱的医院以空格补位    院端
                    str明细 = str明细 & Get就诊分类(int业务, Nvl(!灰度级, 0))         '就诊分类    CHAR    57  1   取值详见"就诊分类"说明  院端
                    str明细 = str明细 & Rpad(Format(!登记时间, "yyyymmddHHmmss"), 16)      '处方生成时间（投药时间）    DATETIME    58  16  精确到秒格式为：yyyymmddhhmiss后面以空格补位    院端
                    str明细 = str明细 & Lpad(Nvl(!国家编码), 20)      '项目代码    CHAR    74  20  计价项目代码    院端
                    str明细 = str明细 & Lpad(Nvl(!项目名称), 20)      '项目名称    CHAR    94  20      院端
        
                    If !是否医保 = 1 Then
                        str明细 = str明细 & Lpad(1 - dbl比例, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                    Else
                        str明细 = str明细 & Lpad(1, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                    End If
                    str明细 = str明细 & Lpad(str项目统计分类, 1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                    
                    '2005-08-02开发区升级
                    If Nvl(!险类, 0) = TYPE_大连开发区 Then '开发区
                        str明细 = str明细 & Lpad(Nvl(!数量), 10)  '数量    NUM 121 10   冲方划价为负值  院端
                        str明细 = str明细 & Lpad(Abs(Nvl(!实际价格)), 10) '单价    NUM 127 10   不允许出现负值  院端
                    Else
                        str明细 = str明细 & Lpad(Nvl(!数量), 6)  '数量    NUM 121 6   冲方划价为负值  院端
                        str明细 = str明细 & Lpad(Abs(Nvl(!实际价格)), 8) '单价    NUM 127 8   不允许出现负值  院端
                    End If
                    str明细 = str明细 & Lpad(Nvl(!计算单位), 4) '单位    CHAR    135 4       院端
                    str明细 = str明细 & Lpad(Nvl(!剂型), 20)      '剂型    CHAR    139 20  针剂、片剂…    院端
                    str明细 = str明细 & Lpad(Nvl(!医生), 8)      '医师姓名    CHAR    159 8       院端
                    str明细 = str明细 & Lpad(g病人身份_大连.诊断编码, 16)      '诊断编码    CHAR    167 16      院端
                    str明细 = str明细 & Lpad(Substr(g病人身份_大连.诊断名称, 1, 28), 30)   '诊断名称    CHAR    183 30      院端
                    str明细 = str明细 & Space(16)     '传输时间    DATETIME    213 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，院端空格补位  中心
                
                '上传明细
                '1003    7   230 实时医疗明细数据提交
                门诊记帐结算及冲帐_大连 = 业务请求_大连(IIf(Nvl(!险类, 0) = TYPE_大连开发区, 2, 1), 1003, str明细, intinsure)
                If 门诊记帐结算及冲帐_大连 = False Then
                    Err.Raise 9000, gstrSysName, "门诊记帐结算时医疗明细数据提交失败,不能继续!"
                    Exit Function
                End If
                '上传医嘱明细
                If Nvl(!医嘱序号, 0) <> 0 Then
                    If 医嘱明细数据提交(!医嘱序号, "", str项目统计分类, intinsure) = False Then
                        Err.Raise 9000, gstrSysName, "医嘱明细数据提交失败,不能继续!"
                        Exit Function
                    End If
                End If
                '为病人费用记录打上标记，以便随时上传
                'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            End If
            '计算总额,待用
            curTotal = curTotal + Round(Nvl(!实收金额, 0), 5)
            .MoveNext
        Loop
    End With
    门诊记帐结算及冲帐_大连 = False
    
    '计算起付线
    dbl起付标准 = g病人身份_大连.起付线
    
    If 大连医保结算(intinsure, 2, lng病人ID, 0, IIf(bln冲销, lng冲销ID, lng结帐ID), 原结帐id, bln冲销, False, 0, _
        dbl诊察费, dbl草药费, dbl成药费, dbl西药费, dbl检查费, dbl治疗费, dbl血费, dbl血费自费, dbl大检费, dbl大检自费, _
        dbl特殊治疗费, dbl特殊治疗自费, dbl保险内自费费用, dbl非保险费用, dbl其它费, dbl药费自费, curTotal, str医生, strInfor) = False Then
        Exit Function
    End If
    门诊记帐结算及冲帐_大连 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function 大连医保结算(ByVal intinsure As Integer, ByVal bytType As Byte, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng结帐ID As Long, ByVal lng原结帐ID As Long, ByVal bln冲销 As Boolean, ByVal bln虚拟结算 As Boolean, dbl起付标准 As Double, _
        dbl诊察费 As Double, dbl草药费 As Double, dbl成药费 As Double, dbl西药费 As Double, _
        dbl检查费 As Double, dbl治疗费 As Double, dbl血费 As Double, dbl血费自费 As Double, dbl大检费 As Double, dbl大检自费, _
        dbl特殊治疗费 As Double, dbl特殊治疗自费 As Double, dbl保险内自费费用 As Double, dbl非保险费用 As Double, _
        dbl其它费 As Double, dbl药费自费 As Double, curTotal As Double, str医生 As String, str结算方式 As String, Optional str住院号 As String = "", Optional str出院日期 As String = "") As Boolean
        
        '功能:进行结算
        '参数:bytType-0门诊,1住院,2门诊记帐
        '   lng结帐id-结帐id值
        '   bln虚拟结算=是否虚拟结算
        '   bln冲销 =冲销
        '   dbl开头的传相关费用
        ' 出参:
        '   str结算方式
        '返回:成功,返回true,否则返回False
        
        
        Dim rsTemp As New ADODB.Recordset
        Dim dbl个人帐户余额 As Double, dbl统筹支付累计 As Double, dbl个人帐户支付 As Double
        Dim dbl补助帐户支付 As Double, dbl基本统筹支付 As Double, dbl基本统筹自付 As Double
        Dim dbl补充统筹支付 As Double, dbl补充统筹自付 As Double, dbl补助保险支付 As Double
        Dim dbl非补助保险支付 As Double, dbl保险范围外自付 As Double
        Dim dbl公务员起付标准补助 As Double, dbl公务员基本补助 As Double, dbl公务员非基本补助 As Double
        Dim dbl商业保险补助 As Double, dbl非保险自付 As Double
        Dim int业务 As Integer
        Dim dbl结算前基本帐户余额  As Double
        Dim dbl结算前补助账户余额  As Double
        Dim dbl结算前统筹累计  As Double
        Dim dbl比例 As Double
        
        '2005-08-02 开发区升级
        Dim dbl结算前疾病统筹累计 As Double
        Dim dbl结算后疾病统筹累计 As Double
        
        Dim strInputString As String
        
        int业务 = IIf(bln冲销, 1, 0)
        
        '20040727:同四舍五入造成对不起帐,所以先汇总后再四舍五入
        dbl诊察费 = Round(dbl诊察费, 2)
        dbl草药费 = Round(dbl草药费, 2)
        dbl成药费 = Round(dbl成药费, 2)
        dbl西药费 = Round(dbl西药费, 2)
        dbl检查费 = Round(dbl检查费, 2)
        dbl治疗费 = Round(dbl治疗费, 2)
        dbl大检费 = Round(dbl大检费, 2)
        dbl大检自费 = Round(dbl大检自费, 2)
    
        '2004/9/11 刘兴宏新增血费
        dbl血费 = Round(dbl血费, 2)
        dbl血费自费 = Round(dbl血费自费, 2)
        
        dbl特殊治疗费 = Round(dbl特殊治疗费, 2)
        dbl特殊治疗自费 = Round(dbl特殊治疗自费, 2)
        dbl保险内自费费用 = Round(dbl保险内自费费用, 2)
        dbl非保险费用 = Round(dbl非保险费用, 2)
        dbl其它费 = Round(dbl其它费, 2)
        
        '开发区升级
        dbl药费自费 = Round(dbl药费自费, 2)
        curTotal = Round(curTotal, 2)
    
            
        Err = 0
        On Error GoTo errHand:
        
        大连医保结算 = False
        If bln冲销 Then
            gstrSQL = "" & _
                "   Select *  " & _
                "   From 保险结算记录 " & _
                "   Where 记录id=" & lng原结帐ID
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "提取中心收费时返回的数据"
            If rsTemp.RecordCount = 0 Then
                ShowMsgbox "不存在上次收费的结算记录!"
                Exit Function
            End If
            dbl个人帐户余额 = Round(Nvl(rsTemp!帐户累计增加, 0), 2)
            dbl统筹支付累计 = Round(Nvl(rsTemp!帐户累计支出, 0), 2)
            dbl补助保险支付 = Round(Nvl(rsTemp!累计进入统筹, 0), 2)
            dbl补助帐户支付 = Round(Nvl(rsTemp!累计统筹报销, 0), 2)
            dbl起付标准 = Round(Nvl(rsTemp!起付线, 0), 2)
            dbl保险范围外自付 = Round(Nvl(rsTemp!封顶线, 0), 2)
            dbl基本统筹支付 = Round(Nvl(rsTemp!全自付金额, 0), 2)
            dbl基本统筹自付 = Round(Nvl(rsTemp!首先自付金额, 0), 2)
            dbl补充统筹支付 = Round(Nvl(rsTemp!进入统筹金额, 0), 2)
            dbl补充统筹自付 = Round(Nvl(rsTemp!统筹报销金额, 0), 2)
            dbl非补助保险支付 = Round(Nvl(rsTemp!大病自付金额, 0), 2)
            dbl个人帐户支付 = Round(Nvl(rsTemp!个人帐户支付, 0), 2)
            dbl结算前基本帐户余额 = Round(Nvl(rsTemp!结算前基本帐户余额, 0), 2)
            dbl结算前补助账户余额 = Round(Nvl(rsTemp!结算前补助账户余额, 0), 2)
            dbl结算前统筹累计 = Round(Nvl(rsTemp!结算前统筹累计, 0), 2)
            
            dbl公务员起付标准补助 = Round(Nvl(rsTemp!公务员起付标准补助, 0), 2)
            dbl公务员基本补助 = Round(Nvl(rsTemp!公务员基本补助, 0), 2)
            dbl公务员非基本补助 = Round(Nvl(rsTemp!公务员非基本补助, 0), 2)
            dbl商业保险补助 = Round(Nvl(rsTemp!商业保险补助, 0), 2)
            dbl非保险自付 = Round(Nvl(rsTemp!非保险自付, 0), 2)
            '需判断冲销这种情况,如果冲销,则传入上次结算的费用额
            
            dbl诊察费 = Round(Nvl(rsTemp!诊察费, 0), 2)
            dbl草药费 = Round(Nvl(rsTemp!草药费, 0), 2)
            dbl成药费 = Round(Nvl(rsTemp!成药费, 0), 2)
            dbl西药费 = Round(Nvl(rsTemp!西药费, 0), 2)
            dbl检查费 = Round(Nvl(rsTemp!检查费, 0), 2)
            dbl治疗费 = Round(Nvl(rsTemp!治疗费, 0), 2)
            dbl大检费 = Round(Nvl(rsTemp!大检费, 0), 2)
            dbl大检自费 = Round(Nvl(rsTemp!大检自费, 0), 2)
            
            '2004/9/11 刘兴宏新增血费
            dbl血费 = Round(Nvl(rsTemp!血费, 0), 2)
            dbl血费自费 = Round(Nvl(rsTemp!血费自费, 0), 2)
            
            dbl特殊治疗费 = Round(Nvl(rsTemp!特殊治疗费, 0), 2)
            dbl特殊治疗自费 = Round(Nvl(rsTemp!特殊治疗自费, 0), 2)
            dbl保险内自费费用 = Round(Nvl(rsTemp!保险内自费费用, 0), 2)
            dbl非保险费用 = Round(Nvl(rsTemp!非保险费用, 0), 2)
            dbl其它费 = Round(Nvl(rsTemp!其它费, 0), 2)
            
            '2005-08-02 开发区升级
            dbl药费自费 = Round(Nvl(rsTemp!药费自费, 0), 2)
            dbl结算前疾病统筹累计 = Round(Nvl(rsTemp!结算前疾病统筹累计, 0), 2)
            dbl结算后疾病统筹累计 = Round(Nvl(rsTemp!结算后疾病统筹累计, 0), 2)
            
            curTotal = Round(Nvl(rsTemp!发生费用金额, 0), 2)
        End If
        
    With g病人身份_大连
        If intinsure = TYPE_大连开发区 Then    '开发区
            strInputString = Lpad(gstr医院编码_大连, 6)       '医院代码
        Else
            strInputString = Lpad(gstr医院编码_大连, 4)       '医院代码
        End If
        strInputString = strInputString & " "      '子门诊标识
        If intinsure = TYPE_大连开发区 Then   '开发区
            strInputString = strInputString & Lpad(.个人编号, 10)         '个人编号
        Else
            strInputString = strInputString & Lpad(.个人编号, 8)      '个人编号
        End If
        strInputString = strInputString & Lpad(.IC卡号, 7)       'IC卡号
        If bln虚拟结算 Then
            strInputString = strInputString & Lpad(.治疗序号 + 1, 4)     '治疗序号
        Else
            .治疗序号 = .治疗序号 + 1
            strInputString = strInputString & Lpad(.治疗序号, 4)       '治疗序号
        End If
        
        strInputString = strInputString & Rpad(Format(zlDatabase.Currentdate, "yyyymmddHHmmss"), 16)      '结算时间
        If bytType = 1 Then
            '住院是住院号
            strInputString = strInputString & Lpad(str住院号, 10)  '病志号
                
        Else
            strInputString = strInputString & String(10, " ") '病志号
        End If
        
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl诊察费, 2))), 10) '诊察费
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl草药费, 2))), 10) '草药费
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl成药费, 2))), 10) '成药费
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl西药费, 2))), 10)  '西药费
        If intinsure = TYPE_大连市 Then
        Else
            '2005-08-02开发区升级
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl药费自费, 2))), 10)  '药费自费
        End If
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl检查费, 2))), 10)  '检查费
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl治疗费, 2))), 10)   '治疗费
        
        If intinsure = TYPE_大连市 Then
            '2004/9/11 刘兴宏新增血费,同时改变了顺序
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl血费, 2))), 10)  '血费
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl血费自费, 2))), 10)   '血费自费
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl大检费, 2))), 10)   '大检费
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl大检自费, 2))), 10)   '大检自费
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl特殊治疗费, 2))), 10)   '特殊治疗费
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl特殊治疗自费, 2))), 10)    '特治自费    NUM 145 10      院端填写
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl保险内自费费用, 2))), 10)    '保险内自费费用
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl非保险费用, 2))), 10)    '非保险费用
        Else
            '2005-08-02开发区升级
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl血费, 2))), 10)  '血费
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl血费自费, 2))), 10)   '血费自费
            
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl大检费, 2))), 10)   '大检费
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl大检自费, 2))), 10)   '大检自费
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl特殊治疗费, 2))), 10)   '特殊治疗费
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl特殊治疗自费, 2))), 10)    '特治自费    NUM 145 10      院端填写
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl保险内自费费用, 2))), 10)    '保险内自费费用
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl其它费, 2))), 10)    '保险外自费  NUM 165 10  非医保用药自费部分  院端填写

        End If
        
        If bln冲销 Then
            strInputString = strInputString & Lpad(dbl个人帐户余额, 10)
            strInputString = strInputString & Lpad(dbl统筹支付累计, 10)
            If intinsure = TYPE_大连市 Then
            Else
                '2005-08-02开发区升级
                strInputString = strInputString & Lpad(dbl结算后疾病统筹累计, 10)
            End If
        Else
            strInputString = strInputString & String(10, " ") 'Lpad(dbl个人帐户余额, 10)
            strInputString = strInputString & String(10, " ") 'Lpad(dbl统筹支付累计, 10)
            If intinsure = TYPE_大连市 Then
            Else
                '2005-08-02开发区升级
                strInputString = strInputString & String(10, " ") 'Lpad(dbl结算后疾病统筹累计, 10)
            End If
        End If
        '结算前基本帐户余额（根据验卡返回结果，如果是慢病结算应根据慢病查询的慢病帐户余额填写）
        Dim dbl结算前余额(1 To 3) As Double '1-结算前基本帐户余额,2-结算前补助账户余额,3-结算前统筹支付累计
        
        If bytType = 0 And bln冲销 = False And .结算开始 Then
            .当前个人帐户余额 = Format(.基本个人帐户余额, "####0.00;-####0.00;0;0")
            .当前补助帐户余额 = Format(.补助个人帐户余额, "####0.00;-####0.00;0;0")
            .当前统筹累计 = Format(.统筹累计, "####0.00;-####0.00;0;0")
        End If
        
        If bytType = 0 And bln冲销 = False Then
            dbl结算前余额(1) = Format(.当前个人帐户余额, "####0.00;-####0.00;0;0")
            dbl结算前余额(2) = Format(.当前补助帐户余额, "####0.00;-####0.00;0;0")
            dbl结算前余额(3) = Format(.当前统筹累计, "####0.00;-####0.00;0;0")
        Else
            dbl结算前余额(1) = Format(.基本个人帐户余额, "####0.00;-####0.00;0;0")
            dbl结算前余额(2) = Format(.补助个人帐户余额, "####0.00;-####0.00;0;0")
            dbl结算前余额(3) = Format(.统筹累计, "####0.00;-####0.00;0;0")
        End If
        
        '结算前基本帐户余额（根据验卡返回结果，如果是慢病结算应根据慢病查询的慢病帐户余额填写）
        If bln冲销 Then
                strInputString = strInputString & Lpad(dbl结算前基本帐户余额, 10)   '结算前基本帐户余额
                strInputString = strInputString & Lpad(dbl结算前补助账户余额, 10)    '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
                strInputString = strInputString & Lpad(dbl结算前统筹累计, 10)     '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
        Else
            If intinsure <> TYPE_大连开发区 And Get就诊分类(0, .就诊分类) = "S" Then
                '① 如果是基本医疗结算表示: 基本个人帐户余额 补助个人帐户余额
                '② 如果是慢病结算表示: 慢病帐户余额
                strInputString = strInputString & Lpad(.补助帐户当前值, 10)   '结算前基本帐户余额
                strInputString = strInputString & Lpad("0", 10)   '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
                strInputString = strInputString & Lpad("0", 10)   '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
                dbl结算前余额(1) = .补助帐户当前值
                dbl结算前余额(2) = 0
                dbl结算前余额(3) = 0
            Else
                
                strInputString = strInputString & Lpad(dbl结算前余额(1), 10)   '结算前基本帐户余额
                strInputString = strInputString & Lpad(Trim(CStr(dbl结算前余额(2))), 10)    '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
                strInputString = strInputString & Lpad(Trim(CStr(dbl结算前余额(3))), 10)    '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
            End If
        End If
              
        If bln冲销 Then
            '需上传相应的值
            If intinsure = TYPE_大连市 Then
            Else
                '2005-08-02开发区升级
                strInputString = strInputString & Lpad(dbl结算前疾病统筹累计, 10)
            End If
            strInputString = strInputString & Lpad(dbl个人帐户支付, 10) '中心返回:本次基本个人帐户支付(如果是慢病结算，表示慢病帐户支付)
            strInputString = strInputString & Lpad(dbl补助帐户支付, 10) '中心返回:本次补助个人帐户支付(如果是慢病结算返回0)
            strInputString = strInputString & Lpad(dbl基本统筹支付, 10) '中心返回:本次基本统筹支付
            strInputString = strInputString & Lpad(dbl基本统筹自付, 10) '中心返回:本次基本统筹自付
            strInputString = strInputString & Lpad(dbl补充统筹支付, 10)  '中心返回:本次补充统筹支付
            strInputString = strInputString & Lpad(dbl补充统筹自付, 10) '中心返回:本次补充统筹自付
            
            If intinsure = TYPE_大连市 Then
                '2004/9/11 返回值变动了原来的公务员补助和商业补助共用一组字段改为公务员补助项目和商业补助单独纪录，分别在第33、34、35、36号字段里体现
                strInputString = strInputString & Lpad(dbl公务员起付标准补助, 10)   '33  本次公务员起付标准补助支付  NUM 301 10      中心
                strInputString = strInputString & Lpad(dbl公务员基本补助, 10)       '34  本次公务员基本补助保险支付  NUM 311 10      中心
                strInputString = strInputString & Lpad(dbl公务员非基本补助, 10)     '35  本次公务员非基本补助保险支付    NUM 321 10      中心
                strInputString = strInputString & Lpad(dbl商业保险补助, 10)         '36  本次商业保险补助支付    NUM 331 10      中心
                strInputString = strInputString & Lpad(dbl保险范围外自付, 10)       '37  本次保险内自付  NUM 341 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋（保险内自费费用+大检自费+血费自费+特治自费）的个人账户冲抵后的费用   中心
                strInputString = strInputString & Lpad(dbl非保险自付, 10)           '38  本次非保险自付  NUM 351 10      中心
            Else
                '2005-08-02开发区升级
                strInputString = strInputString & Lpad(dbl公务员起付标准补助, 10)   '36  本次公务员起付标准补助支付  NUM 335 10      中心
                strInputString = strInputString & Lpad(dbl公务员基本补助, 10)       '37  本次公务员基本补助保险支付  NUM 345 10      中心
                strInputString = strInputString & Lpad(dbl公务员非基本补助, 10)     '38  本次公务员非基本补助保险支付    NUM 355 10      中心
                strInputString = strInputString & Lpad(dbl商业保险补助, 10)         '39  本次商业保险补助支付    NUM 365 10      中心
                strInputString = strInputString & Lpad(dbl保险范围外自付, 10)       '40  本次保险内自付  NUM 375 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋（保险内自费费用+大检自费+血费自费+特治自费）的个人账户冲抵后的费用   中心
                strInputString = strInputString & Lpad(dbl非保险自付, 10)           '41  本次非保险自付  NUM 385 10      中心
                
'                strInputString = strInputString & Lpad(dbl补助保险支付, 10) '中心返回:本次基本补助保险支付 ；开发区:公务员补助该字段包括门槛费补助部分和基本统筹自付部分的公务员补助支付 中心返回
'                strInputString = strInputString & Lpad(dbl非补助保险支付, 10)   '中心返回:本次非基本补助保险支付；开发区:公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付，该部分（超过基本统筹最高限额部分）除去公务员补助支付后，全部打入"本次保险范围外自付"部分  中心返回
'                strInputString = strInputString & Lpad(dbl保险范围外自付, 10)   '中心返回:本次保险范围外自付；开发区:限额以外＋门槛费自付部分（个人帐户充抵后）＋各种自费去掉补助部分    中心返回
            End If
        Else
            If intinsure = TYPE_大连市 Then
            Else
                '2005-08-02开发区升级
                strInputString = strInputString & String(10, " ")   'Lpad(dbl结算前疾病统筹累计, 10)
            End If
            strInputString = strInputString & String(10, " ")    '中心返回:本次基本个人帐户支付(如果是慢病结算，表示慢病帐户支付)
            strInputString = strInputString & String(10, " ")    '中心返回:本次补助个人帐户支付(如果是慢病结算返回0)
            strInputString = strInputString & String(10, " ")    '中心返回:本次基本统筹支付
            strInputString = strInputString & String(10, " ")    '中心返回:本次基本统筹自付
            strInputString = strInputString & String(10, " ")    '中心返回:本次补充统筹支付
            strInputString = strInputString & String(10, " ")    '中心返回:本次补充统筹自付
            
            If intinsure = TYPE_大连市 Then
                '2004/9/11 返回值变动了原来的公务员补助和商业补助共用一组字段改为公务员补助项目和商业补助单独纪录，分别在第33、34、35、36号字段里体现
                strInputString = strInputString & String(10, " ")   '33  本次公务员起付标准补助支付  NUM 301 10      中心
                strInputString = strInputString & String(10, " ")   '34  本次公务员基本补助保险支付  NUM 311 10      中心
                strInputString = strInputString & String(10, " ")   '35  本次公务员非基本补助保险支付    NUM 321 10      中心
                strInputString = strInputString & String(10, " ")   '36  本次商业保险补助支付    NUM 331 10      中心
                strInputString = strInputString & String(10, " ")   '37  本次保险内自付  NUM 341 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋（保险内自费费用+大检自费+血费自费+特治自费）的个人账户冲抵后的费用   中心
                strInputString = strInputString & String(10, " ")   '38  本次非保险自付  NUM 351 10      中心
            Else
                '2005-08-02开发区升级
                strInputString = strInputString & String(10, " ")   '36  本次公务员起付标准补助支付  NUM 335 10      中心
                strInputString = strInputString & String(10, " ")   '37  本次公务员基本补助保险支付  NUM 345 10      中心
                strInputString = strInputString & String(10, " ")   '38  本次公务员非基本补助保险支付    NUM 355 10      中心
                strInputString = strInputString & String(10, " ")   '39  本次商业保险补助支付    NUM 365 10      中心
                strInputString = strInputString & String(10, " ")   '40  本次保险内自付  NUM 375 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋（保险内自费费用+大检自费+血费自费+特治自费）的个人账户冲抵后的费用   中心
                strInputString = strInputString & String(10, " ")   '41  本次非保险自付  NUM 385 10      中心
                
'                strInputString = strInputString & String(10, " ")    '开发区:公务员补助该字段包括门槛费补助部分和基本统筹自付部分的公务员补助支付 中心返回
'                strInputString = strInputString & String(10, " ")    '开发区:公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付，该部分（超过基本统筹最高限额部分）除去公务员补助支付后，全部打入"本次保险范围外自付"部分  中心返回
'                strInputString = strInputString & String(10, " ")    '开发区:限额以外＋门槛费自付部分（个人帐户充抵后）＋各种自费去掉补助部分    中心返回
            End If
        End If
                
        '门诊应为零
        strInputString = strInputString & Lpad(Trim(CStr(dbl起付标准)), 10)    '起付标准；开发区:本次住院门槛费  NUM 315 10      院端填写
        strInputString = strInputString & Lpad(.转诊单号, 6)     '转诊单号
        strInputString = strInputString & Lpad(Get就诊分类(int业务, .就诊分类), 1)      '就诊分类
        
       '刘洋 2005-08-16 此处不用判断大连市和开发区
'        If intInsure <> TYPE_大连开发区 Then
            '2004/9/11 增加参保类别2
            strInputString = strInputString & Lpad(.参保类别2, 1)    ''42  参保类别2   CHAR    378 1   0 不享受 1 商业 2 公务员根据验卡结果    院端
            strInputString = strInputString & Lpad(.参保类别3, 1)    '参保类别3:0 企保、1 事保，根据验卡结果
'        End If
        strInputString = strInputString & Lpad(.职工就医类别, 1)       '职工就医类别
        
        strInputString = strInputString & Lpad(.诊断编码, 16)    '诊断编码
        
        strInputString = strInputString & Lpad(str医生, 6)    '医师代码
        strInputString = strInputString & Lpad(UserInfo.编号, 6)    '操作员代码
        strInputString = strInputString & Lpad(Substr(.诊断名称, 1, 28), 30)  '诊断名称
        
        '??49  治愈情况标识    CHAR    439 1   1治愈、2好转、3未愈、4死亡、5其他，住院必添 院端
        'A-治愈、B-好转、C-未愈、D-死亡、E-其他
        If bytType = 1 Then
            strInputString = strInputString & Lpad(Get治渝情况_大连(lng病人ID, lng主页ID), 1)
            strInputString = strInputString & Lpad(str出院日期, 8)      '出院日期
        Else
            strInputString = strInputString & "1"    '治愈情况标识
            strInputString = strInputString & String(8, " ")      '出院日期
        End If
        
        '刘洋 2005-08-16,此处不用判断
'        If intInsure = TYPE_大连开发区 Then       '开发区
'        Else
            strInputString = strInputString & String(16, " ")      '传输时间
'        End If
        strInputString = strInputString & String(10, " ")      '错误代码
    End With
    
    Dim blnReturn As Boolean
    
    '业务请求
    If bln虚拟结算 Then
        blnReturn = 业务请求_大连(IIf(intinsure = TYPE_大连开发区, 2, 1), 1006, strInputString, intinsure)
        
    Else
        blnReturn = 业务请求_大连(IIf(intinsure = TYPE_大连开发区, 2, 1), 1002, strInputString, intinsure)
    End If
    
    If blnReturn = False Then Exit Function
    
    If bln虚拟结算 = True Then
        str结算方式 = Get结算方式(strInputString, intinsure)
        大连医保结算 = True
        Exit Function
    End If
    Dim i As Long
    If intinsure = TYPE_大连开发区 Then
'        i = 225 - 10
        '2005-08-02开发区升级
        i = 275 - 10
    Else
        i = 241 - 10
    End If
    
    If intinsure = TYPE_大连开发区 Then
        '2005-08-02开发区升级
        dbl个人帐户余额 = Val(Substr(strInputString, i - 60, 10))
        dbl统筹支付累计 = Val(Substr(strInputString, i - 50, 10))  '结算后统筹支付累计=基本统筹累计＋补充统筹累计
        dbl结算前疾病统筹累计 = Val(Substr(strInputString, i, 10))
        dbl结算后疾病统筹累计 = Val(Substr(strInputString, i - 40, 10))
    Else
        dbl个人帐户余额 = Val(Substr(strInputString, i - 40, 10))
        dbl统筹支付累计 = Val(Substr(strInputString, i - 30, 10))  '结算后统筹支付累计=基本统筹累计＋补充统筹累计
    End If
    
    dbl个人帐户支付 = Val(Substr(strInputString, i + 10, 10))   '本次基本个人帐户支付=如果是慢病结算，表示慢病帐户支付
    dbl补助帐户支付 = Val(Substr(strInputString, i + 20, 10))   '本次补助个人帐户支付    NUM 221 10  如果是慢病结算返回0
    dbl基本统筹支付 = Val(Substr(strInputString, i + 30, 10))   '本次基本统筹支付    NUM 231 10      中心
    dbl基本统筹自付 = Val(Substr(strInputString, i + 40, 10))   '本次基本统筹自付    NUM 241 10      中心
    dbl补充统筹支付 = Val(Substr(strInputString, i + 50, 10))   '本次补充统筹支付    NUM 251 10      中心
    dbl补充统筹自付 = Val(Substr(strInputString, i + 60, 10))   '本次补充统筹自付    NUM 261 10      中心
    
    With g病人身份_大连
        .当前个人帐户余额 = .当前个人帐户余额 - dbl个人帐户支付
        .当前补助帐户余额 = .当前补助帐户余额 - dbl补助帐户支付
        .当前统筹累计 = .当前统筹累计 + dbl基本统筹支付 + dbl补充统筹支付
    End With
    
    If intinsure = TYPE_大连市 Then
        '2004/9/11 返回值变动了原来的公务员补助和商业补助共用一组字段改为公务员补助项目和商业补助单独纪录，分别在第33、34、35、36号字段里体现
        dbl公务员起付标准补助 = Val(Substr(strInputString, i + 70, 10)) '33  本次公务员起付标准补助支付  NUM 301 10      中心
        dbl公务员基本补助 = Val(Substr(strInputString, i + 80, 10))     '34  本次公务员基本补助保险支付  NUM 311 10      中心
        dbl公务员非基本补助 = Val(Substr(strInputString, i + 90, 10))   '35  本次公务员非基本补助保险支付    NUM 321 10      中心
        dbl商业保险补助 = Val(Substr(strInputString, i + 100, 10))      '36  本次商业保险补助支付    NUM 331 10      中心
        dbl保险范围外自付 = Val(Substr(strInputString, i + 110, 10))    '37  本次保险内自付  NUM 341 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋（保险内自费费用+大检自费+血费自费+特治自费）的个人账户冲抵后的费用   中心
        dbl非保险自付 = Val(Substr(strInputString, i + 120, 10))        '38  本次非保险自付  NUM 351 10      中心
    Else
        '2005-08-02开发区升级
        dbl公务员起付标准补助 = Val(Substr(strInputString, i + 70, 10)) '36  本次公务员起付标准补助支付  NUM 335 10      中心
        dbl公务员基本补助 = Val(Substr(strInputString, i + 80, 10))     '37  本次公务员基本补助保险支付  NUM 345 10      中心
        dbl公务员非基本补助 = Val(Substr(strInputString, i + 90, 10))   '38  本次公务员非基本补助保险支付    NUM 355 10      中心
        dbl商业保险补助 = Val(Substr(strInputString, i + 100, 10))      '39  本次商业保险补助支付    NUM 365 10      中心
        dbl保险范围外自付 = Val(Substr(strInputString, i + 110, 10))    '40  本次保险内自付  NUM 375 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋（保险内自费费用+大检自费+血费自费+特治自费）的个人账户冲抵后的费用   中心
        dbl非保险自付 = Val(Substr(strInputString, i + 120, 10))        '41  本次非保险自付  NUM 385 10      中心
        
'        dbl补助保险支付 = Val(Substr(strInputString, i + 70, 10))     '本次基本补助保险支付    NUM 271 10  1． 如果是商业保险该字段包括基本统筹自付部分的商业保险支付2．   如果是公务员补助该字段包括门槛费补助部分、基本统筹自付部分的公务员补助支付；基本统筹最高限额内公务员补助支付后剩余款项计入"本次保险范围外自付"部分  中心
'        dbl非补助保险支付 = Val(Substr(strInputString, i + 80, 10))     '本次非基本补助保险支付  NUM 281 10  1． 如果是商业保险该字段是补充统筹自付部分的商业保险支付2． 如果是公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付；超过基本统筹最高限额公务员补助支付后，剩余款项计入"本次保险范围外自付"部分
'        dbl保险范围外自付 = Val(Substr(strInputString, i + 90, 10))     '本次保险范围外自付  NUM 291 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋保险内自费费用＋非保险费用+大检自费   中心
    End If
    
    '过程参数:
    '   性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,
    '   帐户累计增加_IN(个人帐户余额),帐户累计支出_IN(统筹支付累计),累计进入统筹_IN(补助保险支付),累计统筹报销_IN(补助帐户支付),住院次数_IN(治疗序号),起付线_IN(起付标准),封顶线_IN(保险范围外自付),实际起付线_IN(起付标准),
    '   发生费用金额_IN(费用总额),全自付金额_IN(基本统筹支付),首先自付金额_IN(基本统筹自付),进入统筹金额_IN(补充统筹支付),统筹报销金额_IN(补充统筹自付),大病自付金额_IN(非补助保险支付),超限自付金额_IN(无),
    '   个人帐户支付_IN(个人帐户支付),支付顺序号_IN(就诊分类;转诊单号;诊断编码),主页ID_IN(主页id),中途结帐_IN(Null),备注_IN(诊断名称),
    '   诊察费_IN,草药费_IN,成药费_IN,西药费_IN,检查费_IN,治疗费_IN,大检费_IN,大检自费_IN,特殊治疗费_IN,特殊治疗自费_IN,
    '   保险内自费费用_IN,非保险费用_IN,统筹比例_IN,其它费_IN,血费_IN,血费自费_IN,结算前基本帐户余额_IN,结算前补助账户余额_IN,结算前统筹累计_IN,
    '   公务员起付标准补助_IN , 公务员基本补助_IN, 公务员非基本补助_IN, 商业保险补助_IN, 非保险自付_IN
'2005-08-02开发区升级
    '   ,药费自费_IN,结算前疾病统筹累计_IN,结算后疾病统筹累计_IN
    
    '周海全调试 2004-10-20
    '冲销时，如果值应该为负数:
    '   1)诊察费/草药费/成药费/西药费/检查费/治疗费/血费/血费自费/大检费/大检自费/特殊治疗费/特殊治疗自费/保险内自费费用/非医疗保险费用/药费自费
    '   2)发生费用总额也应该同时为负

    If bln冲销 Then
        gstrSQL = "zl_保险结算记录_insert(" & IIf(bytType = 0, 1, 2) & "," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
            dbl个人帐户余额 & "," & dbl统筹支付累计 & "," & dbl补助保险支付 & "," & dbl补助帐户支付 & "," & g病人身份_大连.治疗序号 & "," & dbl起付标准 & "," & dbl保险范围外自付 & "," & dbl起付标准 & "," & _
            -curTotal & "," & dbl基本统筹支付 & "," & dbl基本统筹自付 & "," & dbl补充统筹支付 & "," & dbl补充统筹自付 & "," & dbl非补助保险支付 & ",Null," & _
            dbl个人帐户支付 & ",'" & Get就诊分类(int业务, g病人身份_大连.就诊分类) & ";" & g病人身份_大连.转诊单号 & ";" & g病人身份_大连.诊断编码 & "'," & lng主页ID & ",null,'" & Lpad(Substr(g病人身份_大连.诊断名称, 1, 28), 30) & "'," & _
            -dbl诊察费 & "," & -dbl草药费 & "," & -dbl成药费 & "," & -dbl西药费 & "," & -dbl检查费 & "," & -dbl治疗费 & "," & -dbl大检费 & "," & -dbl大检自费 & "," & -dbl特殊治疗费 & "," & -dbl特殊治疗自费 & "," & _
            -dbl保险内自费费用 & "," & -Abs(dbl非保险费用) & "," & dbl比例 & "," & -dbl其它费 & "," & -dbl血费 & "," & -dbl血费自费 & "," & dbl结算前余额(1) & "," & dbl结算前余额(2) & "," & dbl结算前余额(3) & "," & _
            dbl公务员起付标准补助 & "," & dbl公务员基本补助 & "," & dbl公务员非基本补助 & "," & dbl商业保险补助 & "," & -Abs(dbl非保险自付) & "," & _
             -Abs(dbl药费自费) & "," & dbl结算前疾病统筹累计 & "," & dbl结算后疾病统筹累计 & " )"
    Else
        gstrSQL = "zl_保险结算记录_insert(" & IIf(bytType = 0, 1, 2) & "," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
            dbl个人帐户余额 & "," & dbl统筹支付累计 & "," & dbl补助保险支付 & "," & dbl补助帐户支付 & "," & g病人身份_大连.治疗序号 & "," & dbl起付标准 & "," & dbl保险范围外自付 & "," & dbl起付标准 & "," & _
            curTotal & "," & dbl基本统筹支付 & "," & dbl基本统筹自付 & "," & dbl补充统筹支付 & "," & dbl补充统筹自付 & "," & dbl非补助保险支付 & ",Null," & _
            dbl个人帐户支付 & ",'" & Get就诊分类(int业务, g病人身份_大连.就诊分类) & ";" & g病人身份_大连.转诊单号 & ";" & g病人身份_大连.诊断编码 & "'," & lng主页ID & ",null,'" & Lpad(Substr(g病人身份_大连.诊断名称, 1, 28), 30) & "'," & _
            dbl诊察费 & "," & dbl草药费 & "," & dbl成药费 & "," & dbl西药费 & "," & dbl检查费 & "," & dbl治疗费 & "," & dbl大检费 & "," & dbl大检自费 & "," & dbl特殊治疗费 & "," & dbl特殊治疗自费 & "," & _
            dbl保险内自费费用 & "," & dbl非保险费用 & "," & dbl比例 & "," & dbl其它费 & "," & dbl血费 & "," & dbl血费自费 & "," & dbl结算前余额(1) & "," & dbl结算前余额(2) & "," & dbl结算前余额(3) & "," & _
            dbl公务员起付标准补助 & "," & dbl公务员基本补助 & "," & dbl公务员非基本补助 & "," & dbl商业保险补助 & "," & dbl非保险自付 & "," & _
            dbl药费自费 & "," & dbl结算前疾病统筹累计 & "," & dbl结算后疾病统筹累计 & " )"
    End If
    zlDatabase.ExecuteProcedure gstrSQL, "保存" & IIf(bytType = 1, "住院", "门诊") & "收费数据"
    大连医保结算 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_大连(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, ByVal intinsure As Integer) As Boolean
    Dim lng病人ID As Long
    门诊结算_大连 = Set门诊结算或冲销(False, lng结帐ID, cur个人帐户, lng病人ID, strSelfNo, intinsure)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    门诊结算_大连 = False
End Function
Private Function Set门诊结算或冲销(ByVal bln冲销 As Boolean, lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, strSelfNo As String, ByVal intinsure As Integer) As Boolean
  '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；
    '      cur个人帐户   从个人帐户中支出的金额
    
    Dim curTotal As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    
    Dim strInfor As String  '定义中心返回串
    Dim dbl诊察费 As Double, dbl草药费 As Double, dbl成药费 As Double, dbl西药费 As Double
    
    '2005-08-02开发区升级
    Dim dbl药费自费 As Double, dbl结算前疾病统筹累计 As Double, dbl结算后疾病统筹累计 As Double
    
    Dim dbl检查费 As Double, dbl治疗费 As Double, dbl大检费 As Double, dbl大检自费 As Double
    Dim dbl特殊治疗费 As Double, dbl特殊治疗自费 As Double, dbl保险内自费费用 As Double
    Dim dbl非保险费用 As Double, dbl血费 As Double, dbl血费自费 As Double
    Dim dbl其它费 As Double     '针对大连开发区的
    Dim dbl比例 As Double
    Dim str医生 As String, str明细 As String      '明细串
    Dim str国家编码 As String, str项目编码 As String
    Dim str项目统计分类 As String, strTmp As String
    Dim int业务 As Integer, lng冲销ID As Long
    Dim strNO As String, lng记录性质 As Long
    
    Dim lng治疗序号 As Long
    Dim dbl个人帐户余额 As Double, dbl统筹支付累计 As Double, dbl个人帐户支付 As Double
    Dim dbl补助帐户支付 As Double, dbl基本统筹支付 As Double, dbl基本统筹自付 As Double
    Dim dbl补充统筹支付 As Double, dbl补充统筹自付 As Double, dbl补助保险支付 As Double
    Dim dbl非补助保险支付 As Double, dbl保险范围外自付 As Double
    
    Dim dbl结算前基本帐户余额  As Double, dbl结算前补助账户余额 As Double, dbl结算前统筹累计 As Double
    Dim rs特准项目 As New ADODB.Recordset
    Dim lngTmp As Long, lng病种ID As Long
    Dim strInsertSQL As String
    Static str结算时间 As String
    Static lng病人id1 As Long
    
    int业务 = IIf(bln冲销, 1, 0)
     Set门诊结算或冲销 = False
   
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '注意：接口规定，门诊明细需结算后上传；住院明细需预结算时上传，如果卡内金额不足，可以使用圈存接口，即将卡外的钱，调到卡内，以增加卡内金额
    '卡内余额需要通过卡操作函数读取，可圈存金额是接口返回，需要修改
    
    On Error GoTo errHand
    
    '重新读卡
    If 读取病人身份_大连(IIf(intinsure = TYPE_大连开发区, 2, 1), intinsure) = False Then
        Exit Function
    End If
    
    If bln冲销 Then
        '验证是否为该病人的IC卡
        gstrSQL = "Select 病人ID,卡号 From  保险帐户 where 病人id=" & lng病人ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人的医保号"
        If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "该病人在保险帐户中无记录!"
            Exit Function
        End If
        
        If g病人身份_大连.IC卡号 <> Nvl(rsTemp!卡号) Then
            ShowMsgbox "该病人的IC卡插入错误,可能是插入了其他人的IC卡!"
            Exit Function
        End If
        '确定就诊分类,转诊单号,诊断编码,诊断名称
        ' 支付顺序号_IN(就诊分类;转诊单号;诊断编码),备注(诊断名称_IN)
        gstrSQL = "Select 支付顺序号,备注 from 保险结算记录  where 记录ID=" & lng结帐ID
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取就诊分类"
        If rsTemp.RecordCount = 0 Then
            Err.Raise 9000, gstrSysName, "在结算记录中无结算记录!"
            Exit Function
        End If
        Dim strArr
        strArr = Split(Nvl(rsTemp!支付顺序号), ";")
        
        '就诊分类;转诊单号;诊断编码
        '1-普通门诊("1", "A"),2-急诊门诊("3", "7")
        '3-门诊大病("5", "B"),4-门诊慢病补助("S", "T")
        If UBound(strArr) >= 2 Then
            g病人身份_大连.就诊分类 = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g病人身份_大连.转诊单号 = strArr(1)
            g病人身份_大连.诊断编码 = strArr(2)
        ElseIf UBound(strArr) = 1 Then
            g病人身份_大连.就诊分类 = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g病人身份_大连.转诊单号 = strArr(1)
        Else
            g病人身份_大连.就诊分类 = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
        End If
        g病人身份_大连.诊断名称 = Nvl(rsTemp!备注)
        
        
        '确定退费记录
        '退费
          gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
                    " where A.NO=B.NO and A.记录性质=B.记录性质  and A.记录状态=2 and B.结帐ID=[1]"
          Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊退费", lng结帐ID)
          If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "不存在病人费用冲销记录!"
            Exit Function
          Else
            lng冲销ID = rsTemp("结帐ID")
          End If
          
    End If
    '打开本次结算明细记录 '--国家编码应该是标识主码+子码
    gstrSQL = " " & _
        "  Select Rownum 标识号,A.ID,A.病人ID,A.收费细目id,A.NO,A.序号,A.记录性质,A.记录状态,A.登记时间,A.开单人 as 医生,H.编号 as 医生编号, " & _
        "      A.数次*A.付数 as 数量,A.计算单位,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额 as 实收金额,F.参数值,G.id as 大类id,G.统筹比额,G.住院比额, " & _
        "      A.医嘱序号,A.收费类别,B.编码 as 项目编码,B.名称 as 项目名称,Nvl(J.标识码,Nvl(B.标识主码||B.标识子码,B.编码)) as 国家编码, " & _
        "      D.项目编码 医保编码,D.项目名称 as 医保名称,J.名称 as 剂型,D.是否医保,C.名称 开单部门,E.名称 受单部门, " & _
        "      L.险类,L.中心,L.卡号,L.医保号,L.人员身份,L.单位编码,L.顺序号,L.退休证号,L.帐户余额,L.当前状态,L.病种ID,L.在职,L.年龄段,L.灰度级,L.就诊时间 " & _
        "  From (Select * From 门诊费用记录 Where 记录状态<>0 and 结帐ID=" & IIf(bln冲销, lng冲销ID, lng结帐ID) & " and  Nvl(附加标志,0)<>9 ) A,收费细目 B,部门表 C,保险支付项目 D,部门表 E,  " & _
        "       (Select U.*,K.参数值 From 收费类别 U,保险参数 K where U.类别=K.参数名 and K.险类=" & intinsure & "  ) F, " & _
        "       (Select distinct Q.药品id,Q.标识码,T.名称 From 药品目录 Q,药品信息 R,药品剂型 T  Where  Q.药名id=R.药名id and R.剂型=T.编码 ) J, " & _
        "       保险支付大类 G,人员表 H,保险帐户 L" & _
        "  Where A.收费细目ID=B.ID And A.开单部门ID=C.ID(+) and A.病人id=L.病人id and L.险类=" & intinsure & " and A.收费类别=F.编码(+)  and d.大类id=G.id and a.收费细目id=J.药品id(+) " & _
        "        And A.执行部门ID=E.ID(+) And A.收费细目ID=D.收费细目ID And D.险类= " & intinsure & " and a.开单人=H.姓名(+) " & _
        "  Order by A.ID"
        
    '上传费用明细记录
    zlDatabase.OpenRecordset rs明细, gstrSQL, "读取本次结帐费用明细"
    
    With rs明细
        '主要处理多收费单据
        If str结算时间 <> Format(!登记时间, "yyyy-mm-dd HH:MM:SS") Or lng病人id1 <> Nvl(!病人ID, 0) Then
              str结算时间 = Format(!登记时间, "yyyy-mm-dd HH:MM:SS")
              lng病人id1 = Nvl(!病人ID, 0)
              g病人身份_大连.结算开始 = True
        Else
              g病人身份_大连.结算开始 = False
        End If
    
        If Not .EOF Then
            lng病人ID = Nvl(!病人ID, 0)
            str医生 = Nvl(!医生编号)
            If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
                str医生 = Substr(str医生, 1, 6)
            End If
            lng病种ID = Nvl(!病种ID, 0)
            '打开特准项目
            gstrSQL = "Select * from 保险特准项目  where 病种ID=  " & lng病种ID
            zlDatabase.OpenRecordset rs特准项目, gstrSQL, "获取病种项目数据"
        End If
        Do While Not .EOF
        
            If lng病种ID <> 0 And bln冲销 = False Then
                '第一步,确定允许的收费细目
                rs特准项目.Filter = 0
                rs特准项目.Filter = "大类=0 And 性质=1 and 收费细目id=" & Nvl(!收费细目ID, 0)
                If rs特准项目.EOF Then
                    Err.Raise 9000, gstrSysName, "收费细目为“" & Nvl(!项目名称) & "”的项目不是病种中所设定的项目."
                    Exit Function
                End If
                '第二步,确定允许的保险大类
                rs特准项目.Filter = 0
                rs特准项目.Filter = "大类=1 And 性质=1 and  收费细目id=" & Nvl(!大类id, 0)
                If rs特准项目.EOF Then
                    Err.Raise 9000, gstrSysName, "在结算中存在了结算以外的保险支付大类,不能继续。"
                    Exit Function
                End If
                '第三步,'确定禁止的收费细目
                rs特准项目.Filter = 0
                rs特准项目.Filter = "大类=0 And 性质=2 and 收费细目id=" & Nvl(!收费细目ID, 0)
                If Not rs特准项目.EOF Then
                    Err.Raise 9000, gstrSysName, "收费细目为“" & Nvl(!项目名称) & "”的项目是被禁止使用的项目." & vbCrLf & "不能继续!"
                    Exit Function
                End If
                '第四步,'确定禁止的大类
                rs特准项目.Filter = 0
                rs特准项目.Filter = "大类=1 And 性质=2 and 收费细目id=" & Nvl(!大类id, 0)
                If Not rs特准项目.EOF Then
                    Err.Raise 9000, gstrSysName, "在结算中存在了禁止使用的保险支付大类,不能继续。"
                End If
            End If
            strTmp = Nvl(!参数值)
            lng病人ID = Nvl(!病人ID, 0)
            '确定相关数据
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                If Split(strTmp, ";")(1) = "" Then
                    str项目统计分类 = ""
                Else
                    str项目统计分类 = Mid(Split(strTmp, ";")(1), 1, 1)
                End If
                
                strTmp = Split(strTmp, ";")(0)
                '比例
                '中心为:A在职、B退休、L离休、T特诊,我们默认为1在职、2退休、3离休、4特诊
                    
                If Nvl(!险类, 0) <> TYPE_大连开发区 And Val(Nvl(!单位编码, "99")) = 0 And Nvl(!在职, 0) = 3 And Nvl(!是否医保, 0) = 1 Then   '是企保和离休人员且是医保项目
                    '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
                    '大连市    企业单位离休医保：不完全执行医保政策，（如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
                    dbl比例 = 1
                Else
                    '2005-10-14 ZHQ
                    '门诊大病或是企离是否按住院比额计算
                    If (g病人身份_大连.就诊分类 = 3 And IsParaBig(intinsure)) Or _
                        (IsParaQ(intinsure) And intinsure = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q") Then
                        dbl比例 = Nvl(!住院比额, 0) / 100
                    Else
                        dbl比例 = Nvl(!统筹比额, 0) / 100
                    End If
                End If
                
                If Nvl(!险类, 0) = TYPE_大连市 And (g病人身份_大连.职工就医类别 = "L" Or _
                     g病人身份_大连.职工就医类别 = "T") Then
                    '如果是L离休和T特诊的就按事业比例计算
                    dbl比例 = Val(Nvl(!医保名称)) / 100
                End If
                
                If Nvl(!险类, 0) = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                    '如果是Q企业公费,如果比例为100自费,则需放入非保险费用中
                    If dbl比例 = 0 Then
                        '自费100
                        strTmp = ""
                    Else
                        '自费部分放入 保险内自费费用中
                    End If
                End If
                
                If Nvl(!医保编码) = "特治" Then
                    strTmp = "特殊治疗费"
                End If
                If Nvl(!医保编码) = "大检" Then
                    strTmp = "大检费"
                End If

'-----------------------------------------------------调整开始-------------------------------------------------------
                If dbl比例 <> 0 Then
                    Select Case strTmp
                            Case "诊察费"
                                dbl诊察费 = dbl诊察费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                            
                            Case "草药费"
                                dbl草药费 = dbl草药费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                '2005-08-02开发区升级
                                If Nvl(!险类, 0) = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                Else
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                End If
                                
                            Case "成药费"
                                dbl成药费 = dbl成药费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                '2005-08-02开发区升级
                                If Nvl(!险类, 0) = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                Else
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                End If
                                
                            Case "西药费"
                                dbl西药费 = dbl西药费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                '2005-08-02开发区升级
                                If Nvl(!险类, 0) = TYPE_大连市 Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                Else
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                End If
                                
                            Case "检查费"
                                dbl检查费 = dbl检查费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                
                            Case "治疗费"
                                dbl治疗费 = dbl治疗费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                            Case "大检费"
                                '大连市和开发区对大检费用处理不同,
                                '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                          
                                dbl大检费 = dbl大检费 + Round(Nvl(!实收金额 * dbl比例, 0), 5)
                                dbl大检自费 = dbl大检自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                            
                            Case "血费"
                                '2004/9/11 刘兴宏新增血费
                                dbl血费 = dbl血费 + Round(Nvl(!实收金额 * dbl比例, 0), 5)
                                dbl血费自费 = dbl血费自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                          
                            Case "特殊治疗费"
                                '2004/9/11以前:大连市与开发区计算方式不一致，大连市是总额，开发区仅是统筹部分
                                '2004/9/11以后:大连市与开发区计算方式一致，都是统筹部分的金额
                                'If intinsure = TYPE_大连市 Then
                                '    dbl特殊治疗费 = dbl特殊治疗费 + Round(Nvl(!实收金额, 0), 5)
                                'Else
                                    dbl特殊治疗费 = dbl特殊治疗费 + Round(Nvl(!实收金额, 0) * dbl比例, 5)
                                'End If
                                
                                '特殊治疗费自费的计算方式相同,
                                '大连市接口在处理汇总金额时只把特殊治疗费进行汇总,特殊治疗自费部分不再记入
                                dbl特殊治疗自费 = dbl特殊治疗自费 + Round(Nvl(!实收金额, 0) * (1 - dbl比例), 5)
                                
                        End Select
                    Else
                
                        '全部是比例为0的项目(包括包床),分别对大连还是开发区进行判断放在不同的字段
                        If intinsure = TYPE_大连市 Then
                            '大连市放在dbl非保险费用
                            dbl非保险费用 = dbl非保险费用 + Round(!实收金额, 5)
                        Else
                            '开发区放在dbl其它费
                            dbl其它费 = dbl其它费 + Round(!实收金额, 5)
                        End If
                    End If
            Else
                dbl比例 = 1
                str项目统计分类 = ""
            End If

            '上传明细记录,实时医疗明细数据
            '参数控制明细上传
            If gbln门诊明细时实上传 Then
                
                    If Nvl(!险类, 0) = TYPE_大连开发区 Then '开发区
                        str明细 = Lpad(gstr医院编码_大连, 6)     '医院代号    CHAR    1   6       院端填写
                        str明细 = str明细 & Lpad(Nvl(!医保号), 10)  '保险编号    CHAR    7   10      院端填写
                    Else
                        str明细 = Lpad(gstr医院编码_大连, 4)     '医院代码    CHAR    1   4       院端
                        str明细 = str明细 & Lpad(Nvl(!医保号), 8)   '个人编号    CHAR    5   8       院端
                    End If
                    
                    str明细 = str明细 & Space(10)   '病志号  CHAR    13  10  门诊明细以空格补位,住院是住院号  院端
                    str明细 = str明细 & Lpad(g病人身份_大连.治疗序号, 4)     '治疗序号
                    
                    'Modified By 朱玉宝 2004-07-29 原因：处理NO号
                    str明细 = str明细 & Lpad(Mid(Nvl(!NO, "00000000"), 2, 7), 10)   '处方号  NUM 27  10      院端
                    str明细 = str明细 & Lpad(CStr(.AbsolutePosition), 10)           '处方项目序号    NUM 37  10  对应处方号的记价项目序号    院端
                    
                    str明细 = str明细 & Space(10)           '医嘱号  CHAR    47  10  处方对应医嘱的医嘱记录号，门诊明细或没有医嘱的医院以空格补位    院端
                    str明细 = str明细 & Get就诊分类(int业务, Nvl(!灰度级, 0))         '就诊分类    CHAR    57  1   取值详见"就诊分类"说明  院端
                    
                    str明细 = str明细 & Rpad(Format(!登记时间, "yyyymmddHHmmss"), 16)      '处方生成时间（投药时间）    DATETIME    58  16  精确到秒格式为：yyyymmddhhmiss后面以空格补位    院端
                    
                    str明细 = str明细 & Lpad(Nvl(!国家编码), 20)      '项目代码    CHAR    74  20  计价项目代码    院端
                    str明细 = str明细 & Lpad(Nvl(!项目名称), 20)      '项目名称    CHAR    94  20      院端
        
                    If !是否医保 = 1 Then
                        str明细 = str明细 & Lpad(1 - dbl比例, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                    Else
                        str明细 = str明细 & Lpad(1, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                    End If
                    str明细 = str明细 & Lpad(str项目统计分类, 1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                    
                    If Nvl(!险类, 0) = TYPE_大连开发区 Then
                        '2005-08-02开发区升级
                        str明细 = str明细 & Lpad(Nvl(!数量), 10)  '数量    NUM 121 6   冲方划价为负值  院端
                        str明细 = str明细 & Lpad(Nvl(!实际价格), 10) '单价    NUM 127 8   不允许出现负值  院端
                    Else
                        str明细 = str明细 & Lpad(Nvl(!数量), 6)  '数量    NUM 121 6   冲方划价为负值  院端
                        str明细 = str明细 & Lpad(Nvl(!实际价格), 8) '单价    NUM 127 8   不允许出现负值  院端
                    End If
                    str明细 = str明细 & Lpad(Nvl(!计算单位), 4) '单位    CHAR    135 4       院端
                    str明细 = str明细 & Lpad(Nvl(!剂型), 20)      '剂型    CHAR    139 20  针剂、片剂…    院端
                
                    str明细 = str明细 & Lpad(Nvl(!医生), 8)      '医师姓名    CHAR    159 8       院端
                    str明细 = str明细 & Lpad(g病人身份_大连.诊断编码, 16)      '诊断编码    CHAR    167 16      院端
                    str明细 = str明细 & Lpad(Substr(g病人身份_大连.诊断名称, 1, 28), 30)   '诊断名称    CHAR    183 30      院端
                    str明细 = str明细 & Space(16)     '传输时间    DATETIME    213 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，院端空格补位  中心
                                        
                    '上传明细
                    '1003    7   230 实时医疗明细数据提交
                    Set门诊结算或冲销 = 业务请求_大连(IIf(Nvl(!险类, 0) = TYPE_大连开发区, 2, 1), 1003, str明细, intinsure)
                    If Set门诊结算或冲销 = False Then
                        ShowMsgbox "门诊结算时医疗明细数据提交失败,不能继续!"
                        Exit Function
                    End If
                    
                    '上传医嘱明细
                    If Nvl(!医嘱序号, 0) <> 0 Then
                        If 医嘱明细数据提交(!医嘱序号, "", str项目统计分类, intinsure) = False Then
                            ShowMsgbox "医嘱明细数据提交失败,不能继续!"
                            Exit Function
                        End If
                    End If
                    
                    '为病人费用记录打上标记，以便随时上传
                    'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                    gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                    zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            End If
            '计算总额,待用
            curTotal = curTotal + Round(Nvl(!实收金额, 0), 5)
            .MoveNext
        Loop
    End With
    
    Set门诊结算或冲销 = False
    
    '门诊的起付线为零
    If 大连医保结算(intinsure, 0, lng病人ID, 0, IIf(bln冲销, lng冲销ID, lng结帐ID), lng结帐ID, bln冲销, False, 0, _
        dbl诊察费, dbl草药费, dbl成药费, dbl西药费, dbl检查费, dbl治疗费, dbl血费, dbl血费自费, dbl大检费, dbl大检自费, _
        dbl特殊治疗费, dbl特殊治疗自费, dbl保险内自费费用, dbl非保险费用, dbl其它费, dbl药费自费, curTotal, str医生, strInfor) = False Then
        Exit Function
    End If
    
    Set门诊结算或冲销 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_大连(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, ByVal intinsure As Integer) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Err = 0
    On Error GoTo errHand:
    门诊结算冲销_大连 = Set门诊结算或冲销(True, lng结帐ID, cur个人帐户, lng病人ID, "", intinsure)
    Exit Function
errHand:
    门诊结算冲销_大连 = False
End Function

Public Function 入院登记_大连(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String, ByVal intinsure As Integer) As Boolean
    Dim str入院经办时间 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str就诊分类 As String
    Dim str入院科室 As String
    Dim str床位号 As String
    Dim str转诊单号 As String
    Dim lng中心 As Long
    Dim int险类 As Long
    
    
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    
    On Error GoTo errHand
    
    '读取病人的相关保险信息

    gstrSQL = "select 险类,病人ID,人员身份,医保号,顺序号,灰度级 From 保险帐户 where  险类=" & intinsure & "  and 病人id=" & lng病人ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "入院读取保险帐户信息"
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "在保险帐户中无该病人的保险信息!"
        Exit Function
    End If
    int险类 = Nvl(rsTemp!险类)
    str转诊单号 = Nvl(rsTemp!人员身份)
    lng中心 = IIf(intinsure = 83, 2, 1)
    If lng中心 = 2 Then
        strInfor = Lpad(gstr医院编码_大连, 6)                   '医院代码    CHAR    1   6      Y   院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!医保号), 10)      '保险编号    CHAR    7   10      院端填写
    Else
        strInfor = Lpad(gstr医院编码_大连, 4)                   '医院代码    CHAR    1   4       Y   院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!医保号), 8)       '保险编号    CHAR    5   8       Y   院端
    End If
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!顺序号, 1), 4)        '治疗序号    NUM 13  4   必须等于入院时治疗序号  Y   院端
    
    
    '内部标识:5-普通住院,6-家庭病床住院,7-生育保险住院,8-工伤保险住院
    '医保标识:2-住院结算,4-家庭病床结算,O-生育保险住院结算,Q-工伤保险结算
    
    str就诊分类 = Decode(Val(Nvl(rsTemp!灰度级, 0)), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '读取病人信息
    gstrSQL = "Select C.住院号,C.当前病区id,C.当前床号,A.登记人 经办人,B.名称 入院科室,to_char(A.登记时间,'yyyyMMddhh24miss') 入院经办时间," & _
            " to_char(A.入院日期,'yyyyMMdd') 入院日期" & _
            " From 病案主页 A,部门表 B,病人信息 C" & _
            " Where A.病人id=C.病人id and C.病人id=[1]" & _
            "       and A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院信息", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "在病案主页中无此病人!"
        Exit Function
    End If
    
    str入院科室 = Nvl(rsTemp!入院科室)
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!住院号, 0), 10)       '病志号  CHAR    17  10      Y   院端门诊暂定为空，住院就为住院号
    strInfor = strInfor & Lpad(Nvl(rsTemp!入院日期), 8)         '入院日期 Date 27  8   患者实际入院日期，格式为yyyymmdd    Y   院端
    strInfor = strInfor & Rpad(Nvl(rsTemp!入院经办时间), 16)    '登记时间    DATETIME    35  16  精确到秒，数据返回后格式为yyyymmddhhmiss后面以空格补位  Y   院端
    strInfor = strInfor & Lpad(str就诊分类, 1)                  '就诊分类    CHAR    51  1   2住院、4家床、O生育、   Y   院端

    gstrSQL = "Select 病区ID,床号,房间号 From 床位状况记录 D where 病区ID=[1] And 床号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取床位信息", CLng(Nvl(rsTemp!当前病区ID, 0)), CLng(Nvl(rsTemp!当前床号, 0)))
    If rsTemp.EOF Then
        str床位号 = Space(10)
    Else
        str床位号 = Trim(Nvl(rsTemp!房间号)) & "室" & Trim(Nvl(rsTemp!床号)) & "床"
        str床位号 = Lpad(str床位号, 10)
        str床位号 = Substr(str床位号, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.诊断类型,1,b.编码||'~^||'||b.名称,null)) as 入院诊断,  " & _
         "        max(decode(A.诊断类型,1,null,b.编码||'~^||'||b.名称)) as 确诊诊断 " & _
         " from 诊断情况 A,疾病编码目录 b " & _
         " where a.疾病id=b.id and  a.诊断类型 in(1,2) and a.诊断次序=1 and a.病人id=" & lng病人ID & " and a.主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定诊断编码和名称"
    Dim str入院诊断编码 As String
    Dim str入院诊断名称  As String
    Dim str确诊诊断编码 As String
    Dim str确诊诊断名称  As String
    
    If rsTemp.EOF Then
        str入院诊断编码 = ""
        str入院诊断名称 = ""
        str确诊诊断编码 = ""
        str确诊诊断名称 = ""
    Else
        str入院诊断名称 = Nvl(rsTemp!入院诊断)
        str确诊诊断名称 = Nvl(rsTemp!确诊诊断)
        If InStr(1, str入院诊断名称, "~^||") <> 0 Then
            str入院诊断编码 = Split(str入院诊断名称, "~^||")(0)
            str入院诊断名称 = Split(str入院诊断名称, "~^||")(1)
        Else
            str入院诊断编码 = ""
            str入院诊断名称 = ""
        End If
        If InStr(1, str确诊诊断名称, "~^||") <> 0 Then
            str确诊诊断编码 = Split(str确诊诊断名称, "~^||")(0)
            str确诊诊断名称 = Split(str确诊诊断名称, "~^||")(1)
        Else
            str确诊诊断编码 = ""
            str确诊诊断名称 = ""
        End If
    End If
    '2006-02-20 ZHQ Modify
    '由于入院诊断默认门诊诊断，在门诊诊断不为标准诊断（在门诊输入诊断，而非疾病）时不允许入院
    If Len(Trim(str入院诊断编码)) = 0 Then
        MsgBox "入院诊断非标准ICD-10诊断，医保不允许办理入院！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    strInfor = strInfor & Lpad(str入院诊断编码, 16)     '入院诊断编码    CHAR    52  16      Y   院端
    strInfor = strInfor & Lpad(Substr(str入院诊断名称, 1, 28), 30) '入院诊断名称    CHAR    68  30      y 院端
    strInfor = strInfor & Lpad(str确诊诊断编码, 16)     '确诊诊断编码    CHAR    98  16      N   院端
    strInfor = strInfor & Lpad(Substr(str确诊诊断名称, 1, 28), 30) '确诊诊断名称    CHAR    114 30      N   院端
    strInfor = strInfor & Lpad(str入院科室, 20)         '科别名称    CHAR    144 20  如：内科    Y   院端
    strInfor = strInfor & Lpad(Substr(str床位号, 1, 10), 10)             '床位号  CHAR    164 10  如：2003室12床  N   院端
    strInfor = strInfor & Lpad(str转诊单号, 6)          '转诊单号    CHAR    174 6       N   院端
    strInfor = strInfor & Space(8)                      '出院时间    DATE    180 8   系统利用患者结算数据的出院时间自动生成，医院端用空格补位即可。  N   无
    strInfor = strInfor & "A"                           '传输标志    CHAR    188 1   A 入院登记，M 修改在院状态，C取消入院登记   Y   院端
    strInfor = strInfor & Space(16)                     '传输时间    DATATIME    189 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，用于记录数据到达医保中心的时间，院端空格补位  N   中心
    
    '1004    9   206 实时住院登记数据提交
    入院登记_大连 = 业务请求_大连(lng中心, 1004, strInfor, intinsure)
    If 入院登记_大连 = False Then
        Err.Raise 9000, gstrSysName, "实时住院登记数据提交失败!"
        Exit Function
    End If
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & int险类 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    入院登记_大连 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 入院登记撤销_大连(lng病人ID As Long, lng主页ID As Long, ByVal intinsure As Integer) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
                
    Dim str入院经办时间 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str就诊分类 As String
    Dim str入院科室 As String
    Dim str床位号 As String
    Dim str转诊单号 As String
    Dim lng中心 As Long
    Dim int险类 As Integer
        
    On Error GoTo errHand
    
    '读取病人的相关保险信息

    gstrSQL = "select 险类,病人ID,人员身份,医保号,顺序号,灰度级 From 保险帐户 where 病人id=" & lng病人ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "撤消入院读取保险帐户信息"
    If rsTemp.EOF Then
        ShowMsgbox "在保险帐户中无该病人的保险信息!"
        Exit Function
    End If
    int险类 = rsTemp!险类
    str转诊单号 = Nvl(rsTemp!人员身份)
    lng中心 = IIf(rsTemp!险类 = 83, 2, 1)

    If lng中心 = 2 Then
        strInfor = Lpad(gstr医院编码_大连, 6) '医院代码    CHAR    1   6      Y   院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!医保号), 10)     '保险编号    CHAR    7   10      院端填写
    Else
        strInfor = Lpad(gstr医院编码_大连, 4) '医院代码    CHAR    1   4       Y   院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!医保号), 8)     '保险编号    CHAR    5   8       Y   院端
    End If
    strInfor = strInfor & Lpad(Nvl(rsTemp!顺序号, 1), 4)      '治疗序号    NUM 13  4   必须等于入院时治疗序号  Y   院端
    
    '内部标识:5-普通住院,6-家庭病床住院,7-生育保险住院,8-工伤保险住院
    '医保标识:2-住院结算,4-家庭病床结算,O-生育保险住院结算,Q-工伤保险结算
    
    str就诊分类 = Decode(Val(Nvl(rsTemp!灰度级, 0)), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '读取病人信息
    gstrSQL = "Select C.住院号,C.当前病区id,C.当前床号,A.登记人 经办人,B.名称 入院科室,to_char(A.登记时间,'yyyyMMddhh24miss') 入院经办时间," & _
            " to_char(A.登记时间,'yyyyMMdd') 入院日期" & _
            " From 病案主页 A,部门表 B,病人信息 C" & _
            " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
            "       and A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院信息", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    str入院科室 = Nvl(rsTemp!入院科室)
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!住院号, 0), 10)      '病志号  CHAR    17  10      Y   院端门诊暂定为空，住院就为住院号
    strInfor = strInfor & Lpad(Nvl(rsTemp!入院日期), 8)      '入院日期 Date 27  8   患者实际入院日期，格式为yyyymmdd    Y   院端
    strInfor = strInfor & Rpad(Nvl(rsTemp!入院经办时间), 16)      '登记时间    DATETIME    35  16  精确到秒，数据返回后格式为yyyymmddhhmiss后面以空格补位  Y   院端
    
    strInfor = strInfor & Lpad(str就诊分类, 1)                  '就诊分类    CHAR    51  1   2住院、4家床、O生育、   Y   院端
    
    gstrSQL = "Select 病区ID,床号,房间号 From 床位状况记录 D where 病区ID=[1] And 床号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取床位信息", CLng(Nvl(rsTemp!当前病区ID, 0)), CLng(Nvl(rsTemp!当前床号, 0)))
    If rsTemp.EOF Then
        str床位号 = Space(10)
    Else
        str床位号 = Trim(Nvl(rsTemp!房间号)) & "室" & Trim(Nvl(rsTemp!床号)) & "床"
        str床位号 = Lpad(str床位号, 10)
        str床位号 = Substr(str床位号, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.诊断类型,1,b.编码||'~^||'||b.名称,null)) as 入院诊断,  " & _
         "        max(decode(A.诊断类型,1,null,b.编码||'~^||'||b.名称)) as 确诊诊断 " & _
         " from 诊断情况 A,疾病编码目录 b " & _
         " where a.疾病id=b.id and  a.诊断类型 in(1,2) and a.诊断次序=1 and a.病人id=" & lng病人ID & " and a.主页id=" & lng主页ID
         
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定诊断编码和名称"
    Dim str入院诊断编码 As String
    Dim str入院诊断名称  As String
    Dim str确诊诊断编码 As String
    Dim str确诊诊断名称  As String
    
    If rsTemp.EOF Then
        str入院诊断编码 = ""
        str入院诊断名称 = ""
        str确诊诊断编码 = ""
        str确诊诊断名称 = ""
    Else
        str入院诊断名称 = Nvl(rsTemp!入院诊断)
        str确诊诊断名称 = Nvl(rsTemp!确诊诊断)
        If InStr(1, str入院诊断名称, "~^||") <> 0 Then
            str入院诊断编码 = Split(str入院诊断名称, "~^||")(0)
            str入院诊断名称 = Split(str入院诊断名称, "~^||")(1)
        Else
            str入院诊断编码 = ""
            str入院诊断名称 = ""
        End If
        If InStr(1, str确诊诊断名称, "~^||") <> 0 Then
            str确诊诊断编码 = Split(str确诊诊断名称, "~^||")(0)
            str确诊诊断名称 = Split(str确诊诊断名称, "~^||")(1)
        Else
            str确诊诊断编码 = ""
            str确诊诊断名称 = ""
        End If
    End If
    
    strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
    strInfor = strInfor & Lpad(Substr(str入院诊断名称, 1, 28), 30) '入院诊断名称    CHAR    68  30      y 院端
    strInfor = strInfor & Lpad(str确诊诊断编码, 16)  '确诊诊断编码    CHAR    98  16      N   院端
    strInfor = strInfor & Lpad(Substr(str确诊诊断名称, 1, 28), 30) '确诊诊断名称    CHAR    114 30      N   院端
    
    strInfor = strInfor & Lpad(str入院科室, 20)  '科别名称    CHAR    144 20  如：内科    Y   院端
    strInfor = strInfor & str床位号              '床位号  CHAR    164 10  如：2003室12床  N   院端
    
    strInfor = strInfor & Lpad(str转诊单号, 6)   '转诊单号    CHAR    174 6       N   院端
    strInfor = strInfor & Space(8)   '出院时间    DATE    180 8   系统利用患者结算数据的出院时间自动生成，医院端用空格补位即可。  N   无
    strInfor = strInfor & "C"   '传输标志    CHAR    188 1   A 入院登记，M 修改在院状态，C取消入院登记   Y   院端
    strInfor = strInfor & Space(16)   '传输时间    DATATIME    189 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，用于记录数据到达医保中心的时间，院端空格补位  N   中心
    
    '1004    9   206 实时住院登记数据提交
    入院登记撤销_大连 = 业务请求_大连(lng中心, 1004, strInfor, intinsure)
    If 入院登记撤销_大连 = False Then
        ShowMsgbox "实时住院登记撤消数据提交失败!"
        Exit Function
    End If
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & int险类 & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    入院登记撤销_大连 = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_大连(lng病人ID As Long, lng主页ID As Long) As Boolean
    '
    '办理HIS出院
    Dim rsTemp As New ADODB.Recordset
    Dim int险类
    '---现在HIS入出和医保入出是分开执行,因此不能在办理出院的时候改变医保帐户的状态,直接返回为真
    '--周海全 2005-07-21 加入单病种的提示
    gstrSQL = "select Nvl(a.参数值,'0') as 参数值 From 保险参数 a,病案主页 b " & _
            "  where a.险类(+)=b.险类 and a.参数名(+)='单病种出院提示' and b.病人id=" & lng病人ID & " and b.主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取是否单病种提示"
    
    If rsTemp!参数值 = "1" Then
        If MsgBox("医保特别提醒：" & vbCr & vbCr & "    单病种检查，是否重新输入吗？", _
            vbYesNo + vbDefaultButton1 + vbInformation, gstrSysName) = vbYes Then
            出院登记_大连 = False
        Else
            出院登记_大连 = True
        End If
    Else
        出院登记_大连 = True
    End If
    
'    '----
'    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & rsTemp!险类 & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
End Function

Public Function 出院登记撤销_大连(lng病人ID As Long, lng主页ID As Long) As Boolean
    
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    Dim int险类
'    gstrSQL = "select 险类 From 病案主页 where  病人id=" & lng病人ID & " and 主页id=" & lng主页ID
'
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人的参保险类"
'
'    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & rsTemp!险类 & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销出院登记")
    gstrSQL = "select 当前状态 from 保险帐户 where 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人保险帐户的在院状态"
    
    If Not rsTemp.EOF Then
        If rsTemp!当前状态 = 1 Then
            出院登记撤销_大连 = True
        Else
            MsgBox "该病人的保险帐户当前为出院状态,不能进行撤消!必须取消出院结算后才能执行本操作"
            出院登记撤销_大连 = False
        End If
    Else
        出院登记撤销_大连 = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub 获取病人信息_大连(ByVal lng病人ID As Long, ByVal intinsure As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取病人的相关信息,将其值赋给G病人身份
    '--入参数:lng病人id
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    '读取医保病人相关信息，并更新公用结构体
        
    gstrSQL = "" & _
        "   Select *" & _
        "   From 保险帐户" & _
        "   Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医保病人的相关信息", intinsure, lng病人ID)
    
    If Not rsTemp.EOF Then
        With g病人身份_大连
            .IC卡号 = Nvl(rsTemp!卡号, 0)
            .个人编号 = Nvl(rsTemp!医保号)
            .医保中心 = IIf(intinsure = 83, 2, 1) ' NVL(rsTemp!中心, 1)
            .治疗序号 = Nvl(rsTemp!顺序号, 0)
            .转诊单号 = Nvl(rsTemp!人员身份)
            .基本个人帐户余额 = Nvl(rsTemp!帐户余额, 0)
            .补助个人帐户余额 = Val(Nvl(rsTemp!退休证号))
            
            .职工就医类别 = Decode(Nvl(rsTemp!在职, 1), 1, "A", 2, "B", 3, "L", 4, "T", 5, "Q", "E", 6, "")
            .就诊分类 = Nvl(rsTemp!灰度级, 0)
            .参保类别3 = Nvl(rsTemp!单位编码, 0)
            '.起付线 = NVL(rsTemp!统筹报销累计, 0)
        End With
    End If
End Sub
Private Function Get治渝情况_大连(lng病人ID As Long, lng主页ID As Long) As String
    '功能:获取治渝情况标识
    '     A-治愈、B-好转、C-未愈、D-死亡、E-其他
    '??49  治愈情况标识    CHAR    439 1   1治愈、2好转、3未愈、4死亡、5其他，住院必添 院端
    'A-治愈、B-好转、C-未愈、D-死亡、E-其他
    
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.出院情况" & _
             " From 诊断情况 A,疾病编码目录 B " & _
             " Where A.病人ID=[1] And A.疾病ID=B.ID(+) And A.主页ID=[2]" & _
             "       And A.诊断类型 in (2,3)" & _
             " Order by A.诊断类型 Desc"
    Set rsInNote = zlDatabase.OpenSQLRecord(strTmp, "医保接口", lng病人ID, lng主页ID)
    strTmp = ""
    If Not rsInNote.EOF Then
        strTmp = Nvl(rsInNote!出院情况)
    End If
    Get治渝情况_大连 = Decode(strTmp, "治愈", "1", "好转", "2", "未愈", "3", "死亡", "4", "其他", "5", "1")
End Function

Public Function IS当前住院(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '判断住院病人是否是当前的住院病人
    '2004/09/21
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Max(主页id) as 主页id From 病案主页 where 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取最大的主页id"
    IS当前住院 = False
    If rsTemp.EOF Then
        IS当前住院 = True
        Exit Function
    End If
    IS当前住院 = lng主页ID >= Nvl(rsTemp!主页ID)
End Function

Private Function 计算各项费用额(ByVal intinsure As Integer, ByVal lng病人ID As Long, ByVal str参数 As String, ByVal int算法 As Byte, ByVal str收费类别 As String, _
    ByVal int床号 As Integer, ByVal str医保项目编码 As String, ByVal str发生时间 As String, ByVal int保险项目否 As Integer, _
    ByVal dbl数量 As Double, ByVal dbl单价 As Double, ByVal dbl金额 As Double, ByVal dbl婴儿费 As Double, ByVal dbl统筹比例 As Double, ByVal dbl事业比例 As Double, _
    ByVal dbl特准定额 As Double, dblMoney As Variant) As Boolean
        
        '计算各项费用的金额
        '   2004/9/21
        Dim dbl比例 As Double, strTmp As String
        Dim dblTemp(0 To 20) As Double
        Dim rsTemp As New ADODB.Recordset
        Dim i As Long
        Const strTemp = "诊察费;草药费;成药费;西药费;检查费;治疗费;大检费;大检自费;血费;血费自费;特殊治疗费;特殊治疗自费;保险内自费费用;非保险费用;其他费;药费自费"
        Dim strArr
        
        计算各项费用额 = False
    
        strArr = Split(strTemp, ";")
        If str参数 = "" Or InStr(1, str参数, ";") = 0 Then
            ShowMsgbox "未设置收费类别的对应关系,请到保险类别中设置!"
            Exit Function
        End If
        Err = 0
        On Error GoTo errHand:
        
        dbl比例 = dbl统筹比例 / 100
        strTmp = Split(str参数, ";")(0)
        
        '-----------------------------------------------------------------------
        '计算保费
        '---对于算法=2(定额计算)的项目,由于不是按照比例进行计算,因此需要预先设置dbl比例=1,
        '---使其进入报销分割的判断,否则会因为初始值=0导致被作为全自费项目跳过分割
        
        '-------------------------------------------------------------------------
        '中心为:A在职、B退休、L离休、T特诊,我们默认为1在职、2退休、3离休、4特诊
        If g病人身份_大连.医保中心 <> 2 And g病人身份_大连.职工就医类别 = "L" And g病人身份_大连.参保类别3 = "0" And int保险项目否 = 1 Then '是企保和离休人员且是医保项目
            '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
            '  大连市  企业单位离休医保：不完全执行医保政策，（如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
            dbl比例 = 1
        End If
        '-------------------------------------------------------------------------
        If str医保项目编码 = "特治" Then
            strTmp = "特殊治疗费"
        End If
        If str医保项目编码 = "大检" Then
            strTmp = "大检费"
        End If
                
        '---增加一个对特殊治疗费和大检费用报销比例的检查
        If (strTmp = "特殊治疗费" Or strTmp = "大检费") And dbl比例 = 0 Then
            MsgBox "特殊治疗费用或大检费用的报销比例为0,请先检查设置是否正确"
            Exit Function
        End If
                
        If g病人身份_大连.医保中心 <> 2 And (g病人身份_大连.职工就医类别 = "L" Or _
             g病人身份_大连.职工就医类别 = "T") Then
            '如果是L离休和T特诊的就按事业比例计算
            dbl比例 = dbl事业比例 / 100
        End If
                
        '如果是床位,则需按如下方式处理,主包床按统筹比例计算,被包床为100的自费,不分开发区和大连市
        If str收费类别 = "J" Then
            
            gstrSQL = "" & _
                "   Select 附加床位 From 病人变动记录 " & _
                "   Where 床号=" & int床号 & " and 病人id=" & lng病人ID & _
                "         And ( (to_date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600 between 开始时间 and 终止时间) or" & _
                "               ( 终止时间 is null  and 开始时间<=to_date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600)) " & _
                "         And 床号 is not null"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定是否为包床!"
            
            If rsTemp.RecordCount >= 1 Then
            '--如果床号不是正常变动记录中的床(手工记费),则本处选择可能为NOTHING,需要将dbl比例=0
            '--如果有记录,则判断是否为附加床位,是的话dbl比例=0
                If rsTemp!附加床位 = 1 Then
                    '表示被包床位,为全自费
                    dbl比例 = 0
                Else
                    If int算法 = 2 Then
                         dbl比例 = 1
                    End If
                End If
            End If
        End If
                
        '--婴儿费直接归入全自费部分
        If dbl婴儿费 <> 0 Then
            dbl比例 = 0
        End If
        '------------------------------------------------------重整语句-----------------------------------------------------
        '---因为只有床位费用存在限额报销,所以只对床位进行限额判断并计算
        '---是否是正确的附加床位,需要根据变动记录进行判断 床号\时间,才能确认是本人的自动计算床位费
        '---(但是对手工录入的床位费,判断条件还有变化)
        '-------------------------------------------------------------------------------------------------------------------
        If dbl比例 <> 0 Then
            For i = 0 To UBound(strArr)
                If strTmp = strArr(i) Then
                    '"诊察费;草药费;成药费;西药费;检查费;治疗费;大检费;大检自费;血费;血费自费;特殊治疗费;特殊治疗自费;保险内自费费用;非保险费用;其他费;药费自费"
                    Select Case strTmp
                        Case "诊察费", "检查费"
                            dblTemp(i) = dblTemp(i) + Round(dbl金额 * dbl比例, 5)
                            '需算保险内自费
                            '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                            '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                            dblTemp(12) = dblTemp(12) + Round(dbl金额 * (1 - dbl比例), 5)
                        Case "草药费", "成药费", "西药费"
                            dblTemp(i) = dblTemp(i) + Round(dbl金额 * dbl比例, 5)
                            '2005-08-02开发区升级
                            '药费自费放在特定字段中
                            If intinsure = TYPE_大连市 Then
                                dblTemp(12) = dblTemp(12) + Round(dbl金额 * (1 - dbl比例), 5)
                                dblTemp(15) = 0
                            Else
                                dblTemp(15) = dblTemp(15) + Round(dbl金额 * (1 - dbl比例), 5)
                            End If
                        Case "治疗费"
                            If int算法 = 2 Then
                                '由于可能存在那种销帐后再结算,对是否超过限额必须按照绝对值进行判断
                                If dbl特准定额 <= Abs(dbl单价) Then
                                    '如果限额<单价,则允许报销部分为限额*数量
                                    dblTemp(i) = dblTemp(i) + dbl特准定额 * dbl数量 * Sgn(dbl金额)
                                    '超限部分=(单价-限额)*数量,  正负由金额的符号决定,大连市记入dbl保险内自费费用
                                    dblTemp(12) = dblTemp(12) + Round((Abs(dbl单价) - dbl特准定额) * Abs(dbl数量) * Sgn(dbl金额), 5)
                                Else
                                    '如果限额>=单价,则全部记入治疗费用
                                    dblTemp(12) = dblTemp(12) + Round(dbl金额, 5)
                                End If
                            Else
                                dblTemp(i) = dblTemp(i) + Round(dbl金额 * dbl比例, 5)
                                dblTemp(12) = dblTemp(12) + Round(dbl金额 * (1 - dbl比例), 5)
                            End If
                        Case "大检费", "血费", "特殊治疗费"
                                dblTemp(i) = dblTemp(i) + Round(dbl金额 * dbl比例, 5)
                                dblTemp(i + 1) = dblTemp(i + 1) + Round(dbl金额 * (1 - dbl比例), 5)
                        End Select
                        Exit For
                End If
            Next
        Else
            '全部是比例为0的项目(包括包床),分别对大连还是开发区进行判断放在不同的字段
            If intinsure = TYPE_大连市 Then
                '大连市放在dbl非保险费用
                dblTemp(12) = dblTemp(12) + Round(dbl金额, 5)
            Else
                '开发区放在dbl其它费
                dblTemp(14) = dblTemp(14) + Round(dbl金额, 5)
            End If
        End If
        dblMoney = dblTemp
        计算各项费用额 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    Exit Function
End Function

Private Function Get获取历次病人结帐信息(ByVal lng病人ID As Long, lng主页ID As Long, ByVal intinsure As Integer) As Boolean
    '获取历次病人结算时的信息
    Dim rsTemp As New ADODB.Recordset
    Dim strArr
    Get获取历次病人结帐信息 = False
    
    '读取医保病人相关信息，并更新公用结构体
    gstrSQL = "" & _
            "   Select A.卡号,A.医保号,A.在职,A.参保类别1,A.参保类别2,a.参保类别3,A.参保类别4,A.参保类别5," & _
            "          b.姓名,B.性别,B.出生日期,(sysdate-b.出生日期)/365 as 年龄 ,b.身份证号," & _
            "          c.住院次数,c.实际起付线 起付标准,C.支付顺序号,c.帐户累计支出,c.结算前基本帐户余额,c.结算前补助账户余额,c.结算前统筹累计" & _
            "   From 保险帐户 A,病人信息 B,保险结算记录 C " & _
            "   Where A.险类=[1] And A.病人ID=[2]" & _
            "         and A.病人id=B.病人id and A.病人id=C.病人id and C.性质=2 and C.主页ID=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医保病人的相关信息", intinsure, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        ShowMsgbox "没有相关医保病人信息！"
        Exit Function
    End If
    
    strArr = Split(Nvl(rsTemp!支付顺序号, ";;") & ";;", ";")
    
    With g病人身份_大连
        .IC卡号 = Nvl(rsTemp!卡号, 0)
        .个人编号 = Nvl(rsTemp!医保号)
        .医保中心 = IIf(intinsure = 83, 2, 1) ' NVL(rsTemp!中心, 1)
        
        .姓名 = Nvl(rsTemp!姓名)
        .性别 = Nvl(rsTemp!性别)
        .出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
        .年龄 = Nvl(rsTemp!年龄, 0)
        .身份证号 = Nvl(rsTemp!身份证号)
        .治疗序号 = Nvl(rsTemp!住院次数, 1) - 1
        
        .职工就医类别 = Decode(Nvl(rsTemp!在职, 1), 1, "A", 2, "B", 3, "L", 4, "T", 5, "Q", "E", 6, "")
        .基本个人帐户余额 = Nvl(rsTemp!结算前基本帐户余额, 0)
        .补助个人帐户余额 = Nvl(rsTemp!结算前补助账户余额)
        .当前状态 = 1
        .统筹累计 = Nvl(rsTemp!结算前统筹累计, 0)
        .月缴费基数 = 0
        .参保类别1 = Nvl(rsTemp!参保类别1)
        .参保类别2 = Nvl(rsTemp!参保类别2)
        .参保类别3 = Nvl(rsTemp!参保类别3)
        .参保类别4 = Nvl(rsTemp!参保类别4)
        .参保类别5 = Nvl(rsTemp!参保类别5)
        .帐户状态 = 0
         'C.支付顺序号 （ 就诊分类;转诊单号;诊断编码）
        .转诊单号 = strArr(1)
        .就诊分类 = strArr(0)
        .起付线 = Nvl(rsTemp!起付标准, 0)
    End With
    Get获取历次病人结帐信息 = True
    Exit Function
errHand:
        If ErrCenter = 1 Then Resume
End Function

Private Function 历次费用结算预处理(rsExse As Recordset, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As String
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '      字段:记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID, _
    '           收费类别,收费细目ID,收费名称,开单部门,规格,产地,数量,价格,金额,医生,登记时间, _
    '           是否上传,是否急诊,保险项目否,摘要
    
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    '接口返回的报销额减去本次住院期间以往报销额的汇总金额后，才是本次的实际报销额
    'rsExse记录集中的字段清单
    '记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
    
    Dim rsTemp As New ADODB.Recordset, rs费用 As New ADODB.Recordset, rs帐户状态 As New ADODB.Recordset
    '数组对应关系统
    Dim dblMoney(0 To 20) As Double
    '诊察费;草药费;成药费;西药费;检查费;治疗费;大检费;大检自费;血费;血费自费;特殊治疗费;特殊治疗自费;保险内自费费用;非保险费用;其他费;药费自费
    Dim curTotal As Double
    Dim strInfor As String  '定义中心返回串
    Dim str结算方式  As String
    Dim str医生 As String
    Dim str出院日期 As String
    Dim str住院号 As String, strTmp As String
    Dim intMouse As Integer
    
    On Error GoTo errHand
    intMouse = Screen.MousePointer
    
    '在虚拟结算前需验证身分
    Screen.MousePointer = 1
    
    
    '获取历次病人信息
    If Get获取历次病人结帐信息(lng病人ID, lng主页ID, intinsure) = False Then
        Screen.MousePointer = intMouse
        历次费用结算预处理 = ""
        Exit Function
    End If
    
    Screen.MousePointer = intMouse
    
    gstrSQL = " Select B.住院次数 主页ID,to_char(A.入院日期,'yyyy') 入院年份,A.出院日期,B.住院号" & _
              " From 病案主页 A,病人信息 B" & _
              " Where B.病人ID=[1] And A.主页ID=[2] And A.病人ID=B.病人ID"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人入院时间", lng病人ID, lng主页ID)
    
    str出院日期 = Format(rsTemp!出院日期, "yyyymmdd")
    str住院号 = Nvl(rsTemp!住院号)
    
    
    '需确定诊断情况 刘兴宏2004/06/15,因为出现结算时无诊断情况
    If rsTemp.EOF Then
        strTmp = Get入院诊断(lng病人ID, lng主页ID, , True)
    Else
        If IsNull(rsTemp!出院日期) Then
            strTmp = Get入院诊断(lng病人ID, lng主页ID, , True)
        Else
            strTmp = 获取入出院诊断(lng病人ID, lng主页ID, False, , True)
        End If
    End If
    If InStr(1, strTmp, "|") <> 0 Then
        g病人身份_大连.诊断编码 = Split(strTmp, "|")(1)
        g病人身份_大连.诊断名称 = Split(strTmp, "|")(0)
    End If
    strTmp = ""

    '重新获取记录
    Set rs费用 = Get住院虚拟记录(lng病人ID, lng主页ID, intinsure)
    If rs费用.RecordCount <= 0 Then
        ShowMsgbox "存在未设置的医保项目，不能结算!"
        Exit Function
    End If
    
    
    With rs费用
        '上传费用明细
        curTotal = 0
        str医生 = Nvl(!医生编号)
        If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
            str医生 = Substr(str医生, 1, 6)
        End If
        
        Do While Not .EOF
            lng病人ID = Nvl(!病人ID, 0)
            
            If 计算各项费用额(intinsure, lng病人ID, Nvl(!参数值), Nvl(!算法, 0), Nvl(!收费类别), _
                Nvl(!床号, 0), Nvl(!医保项目编码), Format(!发生时间, "yyyy-mm-dd HH:MM:SS"), _
                Nvl(!保险项目否, 0), Nvl(!数量, 0), Nvl(!价格, 0), Nvl(!金额, 0), Nvl(!婴儿费, 0), _
                Nvl(!住院比额, 0), Val(Nvl(!医保项目名称, 0)), Nvl(!特准定额, 0), dblMoney) = False Then
                Exit Function
            End If
            
            curTotal = curTotal + Nvl(!金额, 0)
            .MoveNext
        Loop
        
        '诊察费;草药费;成药费;西药费;检查费;治疗费;大检费;大检自费;血费;血费自费;特殊治疗费;特殊治疗自费;保险内自费费用;非保险费用;其他费;药费自费
        If 大连医保结算(intinsure, 1, lng病人ID, lng主页ID, 0, 0, False, True, g病人身份_大连.起付线, _
            dblMoney(0), dblMoney(1), dblMoney(2), dblMoney(3), dblMoney(4), dblMoney(5), dblMoney(8), _
            dblMoney(9), dblMoney(6), dblMoney(7), dblMoney(10), dblMoney(11), dblMoney(12), dblMoney(13), _
            dblMoney(14), dblMoney(15), curTotal, str医生, str结算方式, str住院号, str出院日期) = False Then
            Exit Function
        End If
        
        g病人身份_大连.支付金额 = curTotal
    End With
    
    历次费用结算预处理 = str结算方式
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function 住院虚拟结算_大连(rsExse As Recordset, ByVal lng病人ID As Long, ByVal intinsure As Integer) As String
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '      字段:记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID, _
    '           收费类别,收费细目ID,收费名称,开单部门,规格,产地,数量,价格,金额,医生,登记时间, _
    '           是否上传,是否急诊,保险项目否,摘要
    
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    '接口返回的报销额减去本次住院期间以往报销额的汇总金额后，才是本次的实际报销额
    'rsExse记录集中的字段清单
    '记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
    Dim rsTemp As New ADODB.Recordset, rs费用 As New ADODB.Recordset, rs帐户状态 As New ADODB.Recordset
    Dim curTotal As Double
    Dim lng主页ID As Long
    Dim cur个人自付 As Currency, cur个人帐户 As Currency
    Dim str入院年份 As String, str结算年份 As String, str结算时间 As String, str经办时间 As String
    Dim strInfor As String  '定义中心返回串
    Dim dbl诊察费 As Double, dbl草药费 As Double, dbl成药费 As Double, dbl西药费 As Double
    
    '2005-08-02开发区升级
    Dim dbl药费自费 As Double
    
    Dim dbl检查费 As Double, dbl治疗费 As Double, dbl大检费 As Double, dbl大检自费 As Double
    Dim dbl特殊治疗费 As Double, dbl特殊治疗自费 As Double, dbl保险内自费费用 As Double
    Dim dbl非保险费用 As Double, dbl其它费 As Double    '针对大连开发区的
    Dim dbl血费 As Double, dbl血费自费 As Double
       
    Dim dbl比例 As Double, dbl起付标准 As Double
    Dim str结算方式  As String
    
    Dim str医师代码 As String, str操作员代码 As String, str治愈情况标识 As String, strTmp As String
    Dim str医生 As String, str明细 As String, str国家编码 As String, str项目统计分类 As String
    Dim str出院日期 As String, dbl项目名称 As Double
    Dim str住院号 As String, str门诊记帐结算方式 As String
    Dim intMouse As Integer
    
    On Error GoTo errHand
    intMouse = Screen.MousePointer
    
    
    If rsExse.EOF Then
        ShowMsgbox "当前没有明细记录!"
        Exit Function
    End If
    
    lng主页ID = Nvl(rsExse!主页ID, 0)
    rsExse.MoveLast
    If Nvl(rsExse!主页ID, 0) <> lng主页ID Then
        ShowMsgbox "不准对多次住院的病人进行一次结帐"
         Exit Function
    End If
    rsExse.MoveFirst
    g病人身份_大连.历次结算 = IIf(IS当前住院(lng病人ID, lng主页ID) = False, True, False)
    
    '--判断是否能够作为门诊记帐进行结算
    gstrSQL = "select 当前状态 from 保险帐户 where 病人id=[1]"
    Set rs帐户状态 = zlDatabase.OpenSQLRecord(gstrSQL, "读取保险帐户状态", lng病人ID)
    
    If Nvl(rs帐户状态!当前状态, 0) = 0 And g病人身份_大连.历次结算 = False Then
        '历次结算不包含在内
        '如果帐户状态=0,则判断待结费用中是否包含住院费用,包含则不能结算
        Do While Not rsExse.EOF
            '发现有住院费用就不准用门诊方式结算
            If rsExse!门诊标志 = 2 Then
                MsgBox "该病人保险帐户未登记在院,但是待结费用中包含住院费用,请检查病人的帐户状态后重新结算"
                住院虚拟结算_大连 = ""
                Exit Function
            End If
            rsExse.MoveNext
        Loop
        
        If MsgBox("该病人待结费用中只有门诊记帐费用,是否按照门诊进行结算?", vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then
            rsExse.MoveFirst
            '选择按门诊方式结算后,预结成功返回结算值
            If 门诊记帐虚拟结算_大连(rsExse, str门诊记帐结算方式, intinsure) Then
                住院虚拟结算_大连 = str门诊记帐结算方式
                Exit Function
            Else
            '门诊预结失败返回空串
                住院虚拟结算_大连 = ""
                Exit Function
            End If
        Else
            '不按照门诊方式结算就直接返回空串
            住院虚拟结算_大连 = ""
            Exit Function
        End If
    Else
        rsExse.Filter = 0
        rsExse.Filter = "门诊标志=1"
        '如果病人的帐户状态为在院,则必须按照住院方式结算;
        '发现有门诊记帐费用,提示是否按照住院方式结算,
        '如果不同意结算,直接取消处理数据后再重新结算
        If Not rsExse.EOF Then
            If MsgBox("待结费用包含门诊记帐费用,是否按照住院方式对这些费用进行结算?", vbQuestion + vbOKCancel + vbDefaultButton2) = vbCancel Then
                住院虚拟结算_大连 = ""
                Exit Function
            End If
        End If
    End If
    '----------------------完成门诊记帐费用结算方式判断------------
    
    
    '在虚拟结算前需验证身分
    Screen.MousePointer = 1
    If 身份标识_大连(4, lng病人ID, intinsure) = "" Then
        Screen.MousePointer = intMouse
'        Call WriteDebugInfor_大连("住院虚拟结算_大连", lng病人id)
        住院虚拟结算_大连 = ""
        Exit Function
    End If
    
    Screen.MousePointer = intMouse
    
    cur个人帐户 = g病人身份_大连.基本个人帐户余额

    gstrSQL = " Select B.住院次数 主页ID,to_char(A.入院日期,'yyyy') 入院年份,A.出院日期,B.住院号" & _
              " From 病案主页 A,病人信息 B" & _
              " Where B.病人ID=[1] And A.主页ID=[2] And A.病人ID=B.病人ID"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人入院时间", lng病人ID, lng主页ID)
    str入院年份 = rsTemp!入院年份
    'lng主页ID = rsTemp!主页ID
    str出院日期 = Format(rsTemp!出院日期, "yyyymmdd")
    str经办时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str结算时间 = str经办时间
    str结算年份 = Mid(str经办时间, 1, 4)
    str住院号 = Nvl(rsTemp!住院号)
    
    '需确定诊断情况 刘兴宏2004/06/15,因为出现结算时无诊断情况
    If rsTemp.EOF Then
        strTmp = Get入院诊断(lng病人ID, lng主页ID, , True)
    Else
        If IsNull(rsTemp!出院日期) Then
            strTmp = Get入院诊断(lng病人ID, lng主页ID, , True)
        Else
            strTmp = 获取入出院诊断(lng病人ID, lng主页ID, False, , True)
        End If
    End If
    If InStr(1, strTmp, "|") <> 0 Then
        g病人身份_大连.诊断编码 = Split(strTmp, "|")(1)
        g病人身份_大连.诊断名称 = Split(strTmp, "|")(0)
    End If
    
    
    strTmp = ""

    '重新获取记录
    Set rs费用 = Get住院虚拟记录(lng病人ID, lng主页ID, intinsure)
    If rs费用.RecordCount <= 0 Then
        ShowMsgbox "有项目未设置医保项目，不能结算!"
        Exit Function
    End If
    dbl起付标准 = g病人身份_大连.起付线
    

    With rs费用
        '上传费用明细
        curTotal = 0
        Do While Not .EOF
        
            If str医生 = "" Then
                str医生 = Nvl(!医生编号)
                If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
                    str医生 = Substr(str医生, 1, 6)
                End If
            End If
            curTotal = curTotal + Nvl(!金额, 0)
            
            lng病人ID = Nvl(!病人ID, 0)
            strTmp = Nvl(!参数值)
            
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                strTmp = Split(strTmp, ";")(0)
                
                '-----------------------------------------------------------------------
                '计算保费
                '---对于算法=2(定额计算)的项目,由于不是按照比例进行计算,因此需要预先设置dbl比例=1,
                '---使其进入报销分割的判断,否则会因为初始值=0导致被作为全自费项目跳过分割
                
                dbl比例 = Nvl(!住院比额, 0) / 100

                
                '---周顺利,此处条件怀疑是无效条件,留待清理
                '-------------------------------------------------------------------------
                '中心为:A在职、B退休、L离休、T特诊,我们默认为1在职、2退休、3离休、4特诊
                If g病人身份_大连.医保中心 <> 2 And g病人身份_大连.职工就医类别 = "L" And g病人身份_大连.参保类别3 = "0" And Nvl(!保险项目否, 0) = 1 Then '是企保和离休人员且是医保项目
                    '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
                    '  大连市  企业单位离休医保：不完全执行医保政策，（如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
                    dbl比例 = 1
                End If
                '-------------------------------------------------------------------------
                
                If Nvl(!医保项目编码) = "特治" Then
                    strTmp = "特殊治疗费"
                End If
                If Nvl(!医保项目编码) = "大检" Then
                    strTmp = "大检费"
                End If
                
                '---增加一个对特殊治疗费和大检费用报销比例的检查
                If (strTmp = "特殊治疗费" Or strTmp = "大检费") And dbl比例 = 0 Then
                    MsgBox "特殊治疗费用或大检费用的报销比例为0,请先检查设置是否正确"
                    Exit Function
                End If
                
                
                If g病人身份_大连.医保中心 <> 2 And (g病人身份_大连.职工就医类别 = "L" Or _
                     g病人身份_大连.职工就医类别 = "T") Then
                    '如果是L离休和T特诊的就按事业比例计算
                    dbl比例 = Val(Nvl(!医保项目名称)) / 100
                End If
                
                '如果是床位,则需按如下方式处理,主包床按统筹比例计算,被包床为100的自费,不分开发区和大连市
                If Nvl(!收费类别) = "J" Then
                    
                    gstrSQL = "" & _
                        "   Select 附加床位 From 病人变动记录 " & _
                        "   Where 床号=" & Nvl(!床号, 0) & " and 病人id=" & lng病人ID & _
                        "         And ( (to_date('" & Format(!发生时间, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600 between 开始时间 and 终止时间) or" & _
                        "               ( 终止时间 is null  and 开始时间<=to_date('" & Format(!发生时间, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600)) " & _
                        "         And 床号 is not null"
                    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定是否为包床!"
                    If rsTemp.RecordCount >= 1 Then
                    '--如果床号不是正常变动记录中的床(手工记费),则本处选择可能为NOTHING,需要将dbl比例=0
                    '--如果有记录,则判断是否为附加床位,是的话dbl比例=0
                        If rsTemp!附加床位 = 1 Then
                            '表示被包床位,为全自费
                            dbl比例 = 0
                        Else
                            If !算法 = 2 Then
                                 dbl比例 = 1
                            End If
                        End If
                    End If
                End If
                
                '--婴儿费直接归入全自费部分
                If !婴儿费 <> 0 Then
                    dbl比例 = 0
                End If
                
                
            End If
'------------------------------------------------------重整语句-----------------------------------------------------
'---因为只有床位费用存在限额报销,所以只对床位进行限额判断并计算
'---是否是正确的附加床位,需要根据变动记录进行判断 床号\时间,才能确认是本人的自动计算床位费
'---(但是对手工录入的床位费,判断条件还有变化)
'-------------------------------------------------------------------------------------------------------------------
                If dbl比例 <> 0 Then
                    Select Case strTmp
                        Case "诊察费"
                                
                                '--扣除自费部分放入本处
                                dbl诊察费 = dbl诊察费 + Round(Nvl(!金额, 0) * dbl比例, 5)
                                
                                If intinsure = TYPE_大连市 Then
                                
                                    '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                Else
                                    '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                End If

                        Case "草药费"
                                    
                                '--扣除自费部分放入本处
                                dbl草药费 = dbl草药费 + Round(Nvl(!金额, 0) * dbl比例, 5)
                                
                                If intinsure = TYPE_大连市 Then
                                    '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                Else
                                    '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                End If
                                
                        Case "成药费"
                                
                                '--扣除自费部分放入本处
                                dbl成药费 = dbl成药费 + Round(Nvl(!金额, 0) * dbl比例, 5)

                                If intinsure = TYPE_大连市 Then
                                    '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                Else
                                    '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                End If

                        Case "西药费"
                                
                                '--扣除自费部分放入本处
                                dbl西药费 = dbl西药费 + Round(Nvl(!金额, 0) * dbl比例, 5)

                                If intinsure = TYPE_大连市 Then
                                    '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                Else
                                    '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                End If

                        Case "检查费"
                                
                                '--扣除自费部分放入本处
                                dbl检查费 = dbl检查费 + Round(Nvl(!金额, 0) * dbl比例, 5)

                                If intinsure = TYPE_大连市 Then
                                    '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                Else
                                    '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                End If

                        Case "治疗费"
                                '---定额床位费用在数据中是作为治疗费用处理,因此治疗费专门做算法判断
                                
                                If Nvl(!算法, 0) = 2 Then
                                    '由于可能存在那种销帐后再结算,对是否超过限额必须按照绝对值进行判断
                                    If Nvl(!特准定额, 0) <= Nvl(Abs(!价格), 0) Then
                                    
                                        '如果限额<单价,则允许报销部分为限额*数量
                                        dbl治疗费 = dbl治疗费 + Round(Nvl(!特准定额, 0), 5) * Nvl(Abs(!数量), 0) * Sgn(!金额)
                                        
                                        If intinsure = TYPE_大连市 Then
                                            '超限部分=(单价-限额)*数量,  正负由金额的符号决定,大连市记入dbl保险内自费费用
'                                            dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(Abs(!价格) - !特准定额, 0) * Abs(!数量) * Sgn(!金额), 5)
                                            If g病人身份_大连.职工就医类别 = "Q" Then
                                                '由于企业离休针对超限部分需要自费，所以特殊加入
                                                '周海全 2005-02-17
                                                dbl非保险费用 = dbl非保险费用 + Round(Nvl(Abs(!价格) - !特准定额, 0) * Abs(!数量) * Sgn(!金额), 5)
                                            Else
                                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(Abs(!价格) - !特准定额, 0) * Abs(!数量) * Sgn(!金额), 5)
                                            End If
                                        Else
                                            '超限部分=(单价-限额)*数量,  正负由金额的符号决定,开发区记入dbl其它费
                                            'dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(Abs(!价格) - !特准定额, 0) * Abs(!数量) * Sgn(!金额), 5)
                                            
                                            '2005-08-02开发区升级
                                            dbl其它费 = dbl其它费 + Round(Nvl(Abs(!价格) - !特准定额, 0) * Abs(!数量) * Sgn(!金额), 5)
                                        End If
                                        
                                    Else
                                        '如果限额>=单价,则全部记入治疗费用
                                        dbl治疗费 = dbl治疗费 + Round(Nvl(!金额, 0), 5)
                                    End If
                                Else

                                    '--扣除自费部分放入本处
                                    dbl治疗费 = dbl治疗费 + Round(Nvl(!金额, 0) * dbl比例, 5)
    
                                    If intinsure = TYPE_大连市 Then
                                        '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                    Else
                                        '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                                    End If
                                End If

                        Case "大检费"
                                
                                '---大连市和开发区在大检费用上的处理完全一致
                                '---大检项目金额的报销部分记入大检费
                                dbl大检费 = dbl大检费 + Round(Nvl(!金额, 0) * dbl比例, 5)
                                '---大检项目金额的自费部分记入大检自费
                                dbl大检自费 = dbl大检自费 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                        Case "血费"
                            '大连市和开发区对大检费用处理不同,
                            '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                      
                            dbl血费 = dbl血费 + Round(Nvl(!金额 * dbl比例, 0), 5)
                            dbl血费自费 = dbl血费自费 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                        Case "特殊治疗费"
                                '---大连市和开发区在特殊治疗费上的处理有所不同
                                '2004/9/11以前:大连市与开发区计算方式不一致，大连市是总额，开发区仅是统筹部分
                                '2004/9/11以后:大连市与开发区计算方式一致，都是统筹部分的金额
                                dbl特殊治疗费 = dbl特殊治疗费 + Round(Nvl(!金额, 0) * dbl比例, 5)
                                dbl特殊治疗自费 = dbl特殊治疗自费 + Round(Nvl(!金额, 0) * (1 - dbl比例), 5)
                    End Select
                Else
                
                    '全部是比例为0的项目(包括包床),分别对大连还是开发区进行判断放在不同的字段
                    If intinsure = TYPE_大连市 Then
                        '大连市放在dbl非保险费用
                        dbl非保险费用 = dbl非保险费用 + Round(!金额, 5)
                    Else
                        '开发区放在dbl其它费
                        dbl其它费 = dbl其它费 + Round(!金额, 5)
                    End If
                    
                End If
            
            .MoveNext
        Loop

        If 大连医保结算(intinsure, 1, lng病人ID, lng主页ID, 0, 0, False, True, dbl起付标准, dbl诊察费, dbl草药费, dbl成药费, dbl西药费, _
            dbl检查费, dbl治疗费, dbl血费, dbl血费自费, dbl大检费, dbl大检自费, dbl特殊治疗费, dbl特殊治疗自费, dbl保险内自费费用, _
            dbl非保险费用, dbl其它费, dbl药费自费, curTotal, str医生, str结算方式, str住院号, str出院日期) = False Then
            Exit Function
        End If
        g病人身份_大连.支付金额 = curTotal
    End With
    
    住院虚拟结算_大连 = str结算方式
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function Get慢病帐户余额_大连(ByVal intinsure As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取慢病帐户余额
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    '医院代码    CHAR    1   4       院端
    '个人编号    CHAR    5   8       院端
    '补助病种    CHAR    13  16  目前为: WZMB    院端
    '治疗序号    NUM 29  4       中心
    '补助帐户原始值  NUM 33  10  每次补助的累计值    中心
    '补助帐户当前值  NUM 43  10      中心
    '帐户状态    CHAR    53  1   A正常、C止付    中心

    
    Dim strTmp As String
    Err = 0
    On Error GoTo errHand:
    With g病人身份_大连
        strTmp = Lpad(gstr医院编码_大连, 4)      '医院代码    CHAR    1   4       院端
        strTmp = strTmp & Lpad(.个人编号, 8) '个人编号    CHAR    5   8       院端
        strTmp = strTmp & Lpad("WZMB", 16)  '补助病种    CHAR    13  16  目前为: WZMB    院端
        strTmp = strTmp & Space(4)  '治疗序号    NUM 29  4       中心
        strTmp = strTmp & Space(10)  '补助帐户原始值  NUM 33  10  每次补助的累计值    中心
        strTmp = strTmp & Space(10)  '补助帐户当前值  NUM 43  10      中心
        strTmp = strTmp & Space(1)   '帐户状态    CHAR    53  1   A正常、C止付    中心
        '向医保中心查询慢病
        '   1007    2   55  慢病帐户查询
        Get慢病帐户余额_大连 = 业务请求_大连(.医保中心, 1007, strTmp, intinsure)
        If Get慢病帐户余额_大连 = False Then
            .补助帐户原始值 = 0
            .补助帐户当前值 = 0
            Exit Function
        End If
        .补助帐户原始值 = Val(Substr(strTmp, 33, 10))
        .补助帐户当前值 = Val(Substr(strTmp, 43, 10))
    End With
    
    Exit Function
errHand:
End Function

Private Function 住院结算及冲帐_大连(ByVal bln冲销 As Boolean, ByVal lng病人ID As Long, ByVal lng结帐ID As Long, ByVal 原结帐id As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean

    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim strTmp As String, str住院号 As String, str项目统计分类 As String, strInsertSQL As String
    Dim strInfor As String  '定义中心返回串
    Dim curTotal As Double
    Dim dbl诊察费 As Double, dbl草药费 As Double, dbl成药费 As Double, dbl西药费 As Double
    
    '2005-08-02开发区升级
    Dim dbl药费自费 As Double
    
    Dim dbl检查费 As Double, dbl治疗费 As Double, dbl大检费 As Double, dbl大检自费 As Double
    Dim dbl特殊治疗费 As Double, dbl特殊治疗自费 As Double, dbl保险内自费费用 As Double, dbl非保险费用 As Double
    Dim dbl血费 As Double, dbl血费自费 As Double
    Dim dbl其它费 As Double     '针对大连开发区的
    Dim dbl比例 As Double, dbl起付标准 As Double
    
    Dim str医生 As String, str明细 As String, str国家编码 As String, str出院日期 As String
    Dim int业务 As Integer
    
    int业务 = IIf(bln冲销, 1, 0)
    
    Err = 0
    On Error GoTo errHand:
    
    '住院应用保险支付大类中的住院比额
    '4-26,周顺利增加婴儿费判断,排除婴儿费不参与社保结算  '--国家编码=主码+子码
    gstrSQL = " " & _
        "        select a.实收金额,a.id,a.记录性质,a.主页id,a.记录状态,a.发生时间,a.登记时间,a.no,a.病人病区id,a.床号,a.序号,a.标识号 as 住院号,a.病人科室id,a.病人id,a.收费类别,b.类别,a.计算单位, " & _
        "               A.计算单位,A.数次*Nvl(A.付数,1) 数量,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额 ,a.开单人 as 医生,c.编号 as 医生编号, " & _
        "               a.医嘱序号,nvl(a.婴儿费,0) as 婴儿费, A.实收金额,nvl(A.是否上传,0) as 是否上传, " & _
        "               F.参数值,D.编码 as 项目编码,D.名称 as 项目名称,Nvl(J.标识码,Nvl(D.标识主码||D.标识子码,D.编码)) as 国家编码, " & _
        "               E.项目编码 as 医保编码,E.项目名称 as 医保名称,e.是否医保,e.大类id,G.住院比额 as 统筹比额,G.特准定额,G.算法,H.名称 as 开单部门,J.名称 as 剂型, " & _
        "               L.险类,l.中心 , l.卡号, l.医保号, l.人员身份, l.单位编码, l.顺序号, l.退休证号, l.帐户余额, l.当前状态, l.病种ID, l.在职, l.年龄段, l.灰度级, l.就诊时间 " & _
        "        from 住院费用记录 a,收费类别 b,人员表 c,收费细目 D,保险支付项目 E,保险支付大类 G,保险帐户 L,部门表 H, " & _
        "             (Select U.*,K.参数值 From 收费类别 U,保险参数 K where U.类别=K.参数名 and K.险类=" & intinsure & "  ) F ," & _
        "             (Select distinct Q.药品id,Q.标识码,T.名称 From 药品目录 Q,药品信息 R,药品剂型 T  Where  Q.药名id=R.药名id and R.剂型=T.编码 ) J " & _
        "        where a.记录状态<>0 and  a.收费类别=b.编码 and a.收费细目id=J.药品id(+)   and  Nvl(a.附加标志,0)<>9 and a.收费细目id=D.id and a.开单人=c.姓名(+) and a.收费类别=F.编码(+) and " & _
        "              a.收费细目id=E.收费细目ID and E.大类id=G.id and a.病人id=L.病人ID and a.开单部门id=h.id  and " & _
        "              a.病人ID = " & lng病人ID & " And a.结帐ID = " & lng结帐ID & " And E.险类 = " & intinsure
        
    zlDatabase.OpenRecordset rs明细, gstrSQL, "提取住院结帐明细"
    
    
    '确定该病人是否已经出院
    gstrSQL = "Select 病人ID,主页ID,出院日期 From 病案主页 where 病人id=" & lng病人ID & " and 主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人是否出院"
    str出院日期 = ""
    If rsTemp.EOF Then
        strTmp = Get入院诊断(lng病人ID, lng主页ID, False, True)
    Else
        If IsNull(rsTemp!出院日期) Then
            strTmp = Get入院诊断(lng病人ID, lng主页ID, False, True)
        Else
            strTmp = 获取入出院诊断(lng病人ID, lng主页ID, False, , True)
            str出院日期 = Format(rsTemp!出院日期, "yyyymmdd")
        End If
    End If
    
    If InStr(1, strTmp, "|") <> 0 Then
        g病人身份_大连.诊断编码 = Split(strTmp, "|")(1)
        g病人身份_大连.诊断名称 = Split(strTmp, "|")(0)
    End If
    
    With rs明细
        If Not .EOF Then
            str住院号 = Nvl(!住院号)
        End If
        Do While Not .EOF
            strTmp = Nvl(!参数值)
            lng病人ID = Nvl(!病人ID, 0)
            If str医生 = "" Then
                str医生 = Nvl(!医生编号)
                If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
                    str医生 = Substr(str医生, 1, 6)
                End If
            End If
            '确定相关数据
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                If Split(strTmp, ";")(1) = "" Then
                    str项目统计分类 = ""
                Else
                    str项目统计分类 = Mid(Split(strTmp, ";")(1), 1, 1)
                End If
                
                strTmp = Split(strTmp, ";")(0)
                
                '比例
                 dbl比例 = Nvl(!统筹比额, 0) / 100
                
                 '---周顺利,此处条件怀疑是无效条件,留待清理
                '-------------------------------------------------------------------------
                If Nvl(!险类, 0) <> TYPE_大连开发区 And Val(Nvl(!单位编码, "99")) = 0 And Nvl(!在职, 0) = 3 And Nvl(!是否医保, 0) = 1 Then '是企保和离休人员且是医保项目
                    '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
                    '    企业单位离休医保：不完全执行医保政策，（如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
                    dbl比例 = 1
                End If
                '--------------------------------------------------------------------------
                
                If Nvl(!医保编码) = "特治" Then
                    strTmp = "特殊治疗费"
                End If
                If Nvl(!医保编码) = "大检" Then
                    strTmp = "大检费"
                End If
                If Nvl(!险类, 0) = TYPE_大连市 And (g病人身份_大连.职工就医类别 = "L" Or _
                     g病人身份_大连.职工就医类别 = "T") Then
                    '如果是L离休和T特诊的就按事业比例计算
                    dbl比例 = Val(Nvl(!医保名称)) / 100
                End If
                
                 '------------------------------此处屏蔽掉---------------------------------------------
                '如果是床位,则需按如下方式处理,主包床按统筹比例计算,被包床为100的自费,不分开发区和大连市
                If Nvl(!收费类别) = "J" Then
                    
                    gstrSQL = "" & _
                        "   Select 附加床位 From 病人变动记录 " & _
                        "   Where 床号=" & Nvl(!床号, 0) & " and 病人id=" & lng病人ID & _
                        "         And ( (to_date('" & Format(!发生时间, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600 between 开始时间 and 终止时间) or" & _
                        "               ( 终止时间 is null  and 开始时间<=to_date('" & Format(!发生时间, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600)) " & _
                        "         And 床号 is not null"
                    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定是否为包床!"
                    If rsTemp.RecordCount >= 1 Then
                       If rsTemp!附加床位 = 1 Then
                            '表示被包床位,为全自费
                            dbl比例 = 0
                       Else
                            If !算法 = 2 Then
                            dbl比例 = 1
                            End If
                       End If
                    End If
                End If
                
                '--增加对婴儿费的判断,婴儿费用在任何情况下都不予报销,全部归入自费部分
                If !婴儿费 <> 0 Then
                    dbl比例 = 0
                End If
                
            '-----------------------------------------------算法整理----------------------------------------------------
            '---按照接口文档进行费用分割
            '-----------------------------------------------------------------------------------------------------------
                If dbl比例 <> 0 Then
                    
                    Select Case strTmp
                        Case "诊察费"
                                
                                '--扣除自费部分放入本处
                                dbl诊察费 = dbl诊察费 + Round(Nvl(!结帐金额, 0) * dbl比例, 5)
                                
                                If intinsure = TYPE_大连市 Then
                                
                                    '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                Else
                                    '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                End If

                        Case "草药费"
                                    
                                '--扣除自费部分放入本处
                                dbl草药费 = dbl草药费 + Round(Nvl(!结帐金额, 0) * dbl比例, 5)
                                
                                If intinsure = TYPE_大连市 Then
                                    '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                Else
                                    '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                End If
                                
                        Case "成药费"
                                
                                '--扣除自费部分放入本处
                                dbl成药费 = dbl成药费 + Round(Nvl(!结帐金额, 0) * dbl比例, 5)

                                If intinsure = TYPE_大连市 Then
                                    '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                Else
                                    '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                End If

                        Case "西药费"
                                
                                '--扣除自费部分放入本处
                                dbl西药费 = dbl西药费 + Round(Nvl(!结帐金额, 0) * dbl比例, 5)

                                If intinsure = TYPE_大连市 Then
                                    '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                Else
                                    '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                    dbl药费自费 = dbl药费自费 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                End If

                        Case "检查费"
                                
                                '--扣除自费部分放入本处
                                dbl检查费 = dbl检查费 + Round(Nvl(!结帐金额, 0) * dbl比例, 5)

                                If intinsure = TYPE_大连市 Then
                                    '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                Else
                                    '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                End If

                        Case "治疗费"
                                '---定额床位费用在数据中是作为治疗费用处理,因此治疗费专门做算法判断
                                
                                If Nvl(!算法, 0) = 2 Then
                                    '由于可能存在那种销帐后再结算,对是否超过限额必须按照绝对值进行判断
                                    If Nvl(!特准定额, 0) <= Nvl(Abs(!实际价格), 0) Then
                                    
                                        '如果限额<单价,则允许报销部分为限额*数量
                                        dbl治疗费 = dbl治疗费 + Round(Nvl(!特准定额, 0), 5) * Nvl(Abs(!数量), 0) * Sgn(!结帐金额)
                                        
                                        If intinsure = TYPE_大连市 Then
                                            '超限部分=(单价-限额)*数量,  正负由金额的符号决定,大连市记入dbl保险内自费费用
                                            If g病人身份_大连.职工就医类别 = "Q" Then
                                                '由于企业离休针对超限部分需要自费，所以特殊加入
                                                '周海全 2005-02-17
                                                dbl非保险费用 = dbl非保险费用 + Round(Nvl(Abs(!实际价格) - !特准定额, 0) * Abs(!数量) * Sgn(!结帐金额), 5)
                                            Else
                                                dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(Abs(!实际价格) - !特准定额, 0) * Abs(!数量) * Sgn(!结帐金额), 5)
                                            End If
                                        Else
                                            '超限部分=(单价-限额)*数量,  正负由金额的符号决定,开发区记入dbl其它费
                                            dbl其它费 = dbl其它费 + Round(Nvl(Abs(!实际价格) - !特准定额, 0) * Abs(!数量) * Sgn(!结帐金额), 5)
                                        End If
                                        
                                    Else
                                        '如果限额>=单价,则全部记入治疗费用
                                        dbl治疗费 = dbl治疗费 + Round(Nvl(!结帐金额, 0), 5)
                                    End If
                                Else

                                    '--扣除自费部分放入本处
                                    dbl治疗费 = dbl治疗费 + Round(Nvl(!结帐金额, 0) * dbl比例, 5)
    
                                    If intinsure = TYPE_大连市 Then
                                        '---大连市大检、特殊治疗外的自费部分记入dbl保险内自费费用
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                    Else
                                        '---开发区大检、特殊治疗外的自费部分记入dbl其它费
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                                    End If
                                End If

                        Case "大检费"
                                
                                '---大连市和开发区在大检费用上的处理完全一致
                                '---大检项目金额的报销部分记入大检费
                                dbl大检费 = dbl大检费 + Round(Nvl(!结帐金额, 0) * dbl比例, 5)
                                '---大检项目金额的自费部分记入大检自费
                                dbl大检自费 = dbl大检自费 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                        Case "血费"
                            '大连市和开发区对大检费用处理不同,
                            '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                      
                            dbl血费 = dbl血费 + Round(Nvl(!结帐金额 * dbl比例, 0), 5)
                            dbl血费自费 = dbl血费自费 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                        Case "特殊治疗费"
                                '---大连市和开发区在特殊治疗费上的处理有所不同
                                '2004/9/11以前:大连市与开发区计算方式不一致，大连市是总额，开发区仅是统筹部分
                                '2004/9/11以后:大连市与开发区计算方式一致，都是统筹部分的金额
                                dbl特殊治疗费 = dbl特殊治疗费 + Round(Nvl(!结帐金额, 0) * dbl比例, 5)
                                dbl特殊治疗自费 = dbl特殊治疗自费 + Round(Nvl(!结帐金额, 0) * (1 - dbl比例), 5)
                    End Select
                Else
                
                    '全部是比例为0的项目(包括包床),分别对大连还是开发区进行判断放在不同的字段
                    If intinsure = TYPE_大连市 Then
                        '大连市放在dbl非保险费用
                        dbl非保险费用 = dbl非保险费用 + Round(!结帐金额, 5)
                    Else
                        '开发区放在dbl其它费
                        dbl其它费 = dbl其它费 + Round(!结帐金额, 5)
                    End If
                    
                End If

            Else
                dbl比例 = 1
                str项目统计分类 = ""
            End If

 
            '上传明细记录,实时医疗明细数据
            If gbln住院明细时实上传 And bln冲销 = False And Nvl(!是否上传, 0) = 0 And Nvl(!结帐金额, 0) <> 0 And Nvl(!实收金额, 0) <> 0 Then
                    If Nvl(!险类, 0) = TYPE_大连开发区 Then '开发区
                        str明细 = Lpad(gstr医院编码_大连, 6)     '医院代号    CHAR    1   6       院端填写
                        str明细 = str明细 & Lpad(Nvl(!医保号), 10)  '保险编号    CHAR    7   10      院端填写
                    Else
                        str明细 = Lpad(gstr医院编码_大连, 4)     '医院代码    CHAR    1   4       院端
                        str明细 = str明细 & Lpad(Nvl(!医保号), 8)   '个人编号    CHAR    5   8       院端
                    End If
                    
                    str明细 = str明细 & Lpad(Nvl(!住院号, 0), 10) '病志号  CHAR    13  10  门诊明细以空格补位,住院是住院号  院端
                    str明细 = str明细 & Lpad(Nvl(!顺序号, 0), 4)   '治疗序号    NUM 23  4   住院明细：必须等于入院登记时治疗序号门诊明细:                         必须等于本次结算治疗序号 院端
                    
                    'Modified By 朱玉宝 2004-07-29 原因：处理NO号
                    str明细 = str明细 & Lpad(Mid(Nvl(!NO, "00000000"), 2, 7), 10)     '处方号  NUM 27  10      院端

                    str明细 = str明细 & Lpad(Nvl(!ID, 0), 10)      '处方项目序号    NUM 37  10  对应处方号的记价项目序号    院端
                    
                    '开发区为单据号  CHAR    41  10  医嘱号，    院端填写
                    str明细 = str明细 & Lpad(Nvl(!医嘱序号, " "), 10)     '医嘱号  CHAR    47  10  处方对应医嘱的医嘱记录号，门诊明细或没有医嘱的医院以空格补位    院端
                    g病人身份_大连.就诊分类 = Nvl(!灰度级, 0)
                    
                    str明细 = str明细 & Get就诊分类(int业务, Nvl(!灰度级, 0))         '就诊分类    CHAR    57  1   取值详见"就诊分类"说明  院端
                    str明细 = str明细 & Rpad(Format(!登记时间, "yyyymmddHHmmss"), 16)      '处方生成时间（投药时间）    DATETIME    58  16  精确到秒格式为：yyyymmddhhmiss后面以空格补位    院端
                    
                    str明细 = str明细 & Lpad(Nvl(!国家编码), 20)      '项目代码    CHAR    74  20  计价项目代码    院端
                    str明细 = str明细 & Lpad(Nvl(!项目名称), 20)      '项目名称    CHAR    94  20      院端
        
                    If !是否医保 = 1 Then
                        str明细 = str明细 & Lpad(1 - dbl比例, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                    Else
                        str明细 = str明细 & Lpad(1, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                    End If
                    str明细 = str明细 & Lpad(str项目统计分类, 1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                    
                    '2005-08-02开发区升级
                    If intinsure = TYPE_大连开发区 Then
                        str明细 = str明细 & Lpad(Abs(Nvl(!数量)) * Sgn(!结帐金额), 10) '数量    NUM 121 6   冲方划价为负值  院端
                        str明细 = str明细 & Lpad(Abs(Nvl(!实际价格)), 10) '单价    NUM 127 8   不允许出现负值  院端
                    Else
                        str明细 = str明细 & Lpad(Abs(Nvl(!数量)) * Sgn(!结帐金额), 6) '数量    NUM 121 6   冲方划价为负值  院端
                        str明细 = str明细 & Lpad(Abs(Nvl(!实际价格)), 8) '单价    NUM 127 8   不允许出现负值  院端
                    End If
                    str明细 = str明细 & Lpad(Nvl(!计算单位), 4) '单位    CHAR    135 4       院端
                    str明细 = str明细 & Lpad(Nvl(!剂型), 20)      '剂型    CHAR    139 20  针剂、片剂…    院端
                    str明细 = str明细 & Lpad(Nvl(!医生), 8)      '医师姓名    CHAR    159 8       院端
                    str明细 = str明细 & Lpad(g病人身份_大连.诊断编码, 16)      '诊断编码    CHAR    167 16      院端
                    str明细 = str明细 & Lpad(Substr(g病人身份_大连.诊断名称, 1, 28), 30)   '诊断名称    CHAR    183 30      院端
                    str明细 = str明细 & Space(16)     '传输时间    DATETIME    213 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，院端空格补位  中心
                 
                    '上传明细
                    '1003    7   230 实时医疗明细数据提交
                    住院结算及冲帐_大连 = 业务请求_大连(IIf(Nvl(!险类, 0) = TYPE_大连开发区, 2, 1), 1003, str明细, intinsure)
                    If 住院结算及冲帐_大连 = False Then
                        ShowMsgbox "住院结算或冲帐明细数据提交失败,不能继续!"
                        Exit Function
                    End If
                    
                    '上传医嘱明细
                    If Nvl(!医嘱序号, 0) <> 0 Then
                        If 医嘱明细数据提交(!医嘱序号, Nvl(!住院号), str项目统计分类, intinsure) = False Then
                            ShowMsgbox "医嘱明细数据提交失败,不能继续!"
                            Exit Function
                        End If
                    End If

                    '为病人费用记录打上标记，以便随时上传
                    'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                    gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                    zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            End If
            '计算总额,待用
            curTotal = curTotal + Round(Nvl(!结帐金额, 0), 5)
            .MoveNext
        Loop
    End With
    
 
    '填写结算记录
    '计算起付线
    dbl起付标准 = g病人身份_大连.起付线
        
    If 大连医保结算(intinsure, 1, lng病人ID, lng主页ID, lng结帐ID, 原结帐id, bln冲销, False, dbl起付标准, dbl诊察费, dbl草药费, dbl成药费, dbl西药费, _
        dbl检查费, dbl治疗费, dbl血费, dbl血费自费, dbl大检费, dbl大检自费, dbl特殊治疗费, dbl特殊治疗自费, dbl保险内自费费用, dbl非保险费用, _
        dbl其它费, dbl药费自费, curTotal, str医生, strInfor, str住院号, str出院日期) = False Then
        Exit Function
    End If
    '如果是历次结算,就不管入出了
    If g病人身份_大连.历次结算 Then
    Else
        If bln冲销 Then
            gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & intinsure & ")"
            zlDatabase.ExecuteProcedure gstrSQL, "医保帐户入院"
        Else
            gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & intinsure & ")"
            zlDatabase.ExecuteProcedure gstrSQL, "医保帐户出院"
        End If
    End If
    住院结算及冲帐_大连 = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_大连(lng结帐ID As Long, ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean

    '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    '虚拟结算（返回的数据减去历次结算数据，就等于本次的真实结算数据）
    'Dim cur个人帐户 As Currency
    Dim lng主页ID As Long
    Dim int主页id As Long
    Dim blnError As Boolean
    'Dim str入院年份 As String, str结算年份 As String
    'Dim str经办时间 As String, str结算时间 As String
    'Dim str就诊编号 As String
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
    
    '=================================================================================
    '门诊记帐医保结算
    '根据传入的结帐id判断本次结算是否门诊结算，如果结算上保留有主页id,则认为是住院结算
    '否则调用门诊结算接口对病人费用记录结算
'    gstrSQL = "select nvl(主页id,0) as 主页id from 病人预交记录 " & _
'            "where mod(记录性质,10)=2 and 病人id=" & lng病人ID & " and 结帐id=" & lng结帐ID
    gstrSQL = "select Nvl(主页id,0) as 主页id from 住院费用记录 " & _
            "where rownum=1 and 病人id=[1] and 结帐id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查病人是否门诊记帐结算", lng病人ID, lng结帐ID)
    
    lng主页ID = Nvl(rsTemp!主页ID, 0)
    
    Do While Not rsTemp.EOF
        int主页id = int主页id + rsTemp!主页ID
        rsTemp.MoveNext
    Loop
    
    '检查是否含有主页id<>0的情况,即是否有住院费用,
    '如有则跳过门诊记帐结算,直接按照住院方式结算。否则执行门诊结算
    If int主页id = 0 Then
        If 门诊记帐结算及冲帐_大连(False, lng病人ID, lng结帐ID, lng结帐ID, 0, intinsure) = True Then
            '门诊结算成功则向HIS返回成功标志
            住院结算_大连 = True
            Exit Function
        Else
            住院结算_大连 = False
            Exit Function
        End If
    End If
    
    '=================================================================================
    '20041021改
    If g病人身份_大连.历次结算 Then
    Else
        Call 个人余额_大连(lng病人ID, intinsure)
    End If
    住院结算_大连 = 住院结算及冲帐_大连(False, lng病人ID, lng结帐ID, lng结帐ID, lng主页ID, intinsure)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function Get起付线(ByVal str职工就医类别 As String, ByVal lng年龄 As Long, ByVal intinsure As Integer) As Double
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取起付线
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------

       Dim strCaption As String
       Dim rsTmp As New ADODB.Recordset
       
       '20040911加入退老部分
       strCaption = Decode(str职工就医类别, "A", "在职", "B", "退休", "L", "离休", "T", "特诊", "Q", "企业公费", "E", "退老", "在职")
    
        gstrSQL = "" & _
            "   Select d.金额*a.比例/100 as 起付线" & _
            "   From 保险支付比例 a,保险人群 b, " & _
            "      (Select * From 保险年龄段  " & _
            "       where ((" & lng年龄 & ">=下限 and " & lng年龄 & "<=上限) or (" & lng年龄 & ">下限 and 上限=0) ) and 险类=" & intinsure & _
            "       ) c,保险支付限额 d " & _
            " where a.险类=" & intinsure & " and b.险类 =a.险类 and a.在职=b.序号 and b.名称='" & strCaption & "' and  " & _
            "       a.年龄段=c.年龄段 and a.在职=c.在职 and a.险类=d.险类 and d.年度='" & Format(zlDatabase.Currentdate, "yyyy") & "' and d.性质='1'"
    
       Err = 0
       On Error GoTo errHand:
       zlDatabase.OpenRecordset rsTmp, gstrSQL, "计算起付线"
       If Not rsTmp.EOF Then
            Get起付线 = Nvl(rsTmp!起付线, 0)
       Else
            Get起付线 = 0
       End If
       Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
       Get起付线 = 0
   
End Function

Public Function 住院结算冲销_大连(lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
    Dim lng冲销ID As Long
    Dim str退单编号 As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    On Error GoTo errHand
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "大连医保", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID

    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "Select 记录ID,病人ID,主页ID,支付顺序号,备注 " & _
              "  From 保险结算记录 Where 性质=2 and 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "大连医保", lng结帐ID)
    If rsTemp.EOF Then
        ShowMsgbox "在保险结算记录中无该结算记录!"
        Exit Function
    End If
         
    lng病人ID = Nvl(rsTemp!病人ID, 0)
    lng主页ID = Nvl(rsTemp!主页ID, 0)
        
    '如果结算记录中主页id填写为空,则表示是门诊记帐结算记录,需要按照门诊方式进行冲帐
    If lng主页ID = 0 Then
        If 门诊记帐结算及冲帐_大连(True, lng病人ID, lng冲销ID, lng结帐ID, 0, intinsure) Then
            住院结算冲销_大连 = True
            Exit Function
        Else
            住院结算冲销_大连 = False
            Exit Function
        End If
    End If
    '------------------------------门诊记帐结算冲销完成
        
    '重新读卡
    If 读取病人身份_大连(IIf(intinsure = TYPE_大连开发区, 2, 1), intinsure) = False Then
        Exit Function
    End If
    
    Dim strArr
    strArr = Split(Nvl(rsTemp!支付顺序号), ";")
    
    '就诊分类;转诊单号;诊断编码
    '5-普通住院("2", "D"),6-家庭病床住院("4", "C")
    '7-生育保险住院("O", "P"),8-工伤保险住院("Q", "R")
    With rsTemp
        If UBound(strArr) >= 2 Then
            g病人身份_大连.就诊分类 = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
            g病人身份_大连.转诊单号 = strArr(1)
            g病人身份_大连.诊断编码 = strArr(2)
        ElseIf UBound(strArr) = 1 Then
            g病人身份_大连.就诊分类 = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
            g病人身份_大连.转诊单号 = strArr(1)
        Else
            g病人身份_大连.就诊分类 = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
        End If
        g病人身份_大连.诊断名称 = Nvl(rsTemp!备注)
    End With
    
    '验证是否为该病人的IC卡
    gstrSQL = "Select 险类,病人ID,卡号 From 保险帐户 where 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人的医保号"
    If rsTemp.EOF Then
        ShowMsgbox "该病人在保险帐户中无记录!"
        Exit Function
    End If
    
    If g病人身份_大连.IC卡号 <> Nvl(rsTemp!卡号) Then
        ShowMsgbox "该病人的IC卡插入错误,可能是插入了其他人的IC卡!"
        Exit Function
    End If
    
    '--------------------------------------------
    '调用撤销结算接口
    住院结算冲销_大连 = 住院结算及冲帐_大连(True, lng病人ID, lng冲销ID, lng结帐ID, lng主页ID, intinsure)
    If 住院结算冲销_大连 = False Then Exit Function
    
    住院结算冲销_大连 = False
    
    '----------------------------------------------
    '查询产生撤消出院处理的信息串
    '----------------------------------------------
    Dim str入院经办时间 As String
    Dim strInfor As String
    Dim str就诊分类 As String
    Dim str入院科室 As String
    Dim str床位号 As String
    Dim str转诊单号 As String
    Dim lng中心 As Long
    
    '需生新验卡
     lng中心 = IIf(intinsure = 83, 2, 1)
    
     If 读取病人身份_大连(lng中心, intinsure) = False Then Exit Function
    
    '存在未结费用的病人才允许撤销HIS出院；否则认为已办理医保出院，不允许再办理HIS出院
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        MsgBox "医保已出院的病人不允许撤销出院！", vbInformation, gstrSysName
        Exit Function
    End If
               
    On Error GoTo errHand
    
    '读取病人的相关保险信息
    gstrSQL = "select 险类,病人ID,人员身份,医保号,灰度级 From 保险帐户 where  险类=" & intinsure & "  and 病人id=" & lng病人ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "撤消入院读取保险帐户信息"
    If rsTemp.EOF Then
        ShowMsgbox "在保险帐户中无该病人的保险信息!"
        Exit Function
    End If
    
    str转诊单号 = Nvl(rsTemp!人员身份)
    
    If lng中心 = 2 Then
        strInfor = Lpad(gstr医院编码_大连, 6) '医院代码    CHAR    1   6      Y   院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!医保号), 10)     '保险编号    CHAR    7   10      院端填写
    Else
        strInfor = Lpad(gstr医院编码_大连, 4) '医院代码    CHAR    1   4       Y   院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!医保号), 8)     '保险编号    CHAR    5   8       Y   院端
    End If
    
    strInfor = strInfor & Lpad(g病人身份_大连.治疗序号, 4)       '治疗序号    NUM 13  4   必须等于入院时治疗序号  Y   院端
    
    
    '内部标识:5-普通住院,6-家庭病床住院,7-生育保险住院,8-工伤保险住院
    '医保标识:2-住院结算,4-家庭病床结算,O-生育保险住院结算,Q-工伤保险结算
    If intinsure = TYPE_大连市 Then
        str就诊分类 = Decode(Nvl(rsTemp!灰度级, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    Else
        '医保标识:5-普通住院
        '医保标识:2-住院结算
        str就诊分类 = Decode(Nvl(rsTemp!灰度级, 0), 5, "2", 6, "4", "2")
    End If
    '读取病人信息
    gstrSQL = "Select C.住院号,C.当前病区id,C.当前床号,A.登记人 经办人,B.名称 入院科室,to_char(A.登记时间,'yyyyMMddhh24miss') 入院经办时间," & _
            " to_char(A.登记时间,'yyyyMMdd') 入院日期" & _
            " From 病案主页 A,部门表 B,病人信息 C" & _
            " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
            "       and A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院信息", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    str入院科室 = Nvl(rsTemp!入院科室)
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!住院号, 0), 10)       '病志号  CHAR    17  10      Y   院端门诊暂定为空，住院就为住院号
    strInfor = strInfor & Lpad(Nvl(rsTemp!入院日期), 8)         '入院日期 Date 27  8   患者实际入院日期，格式为yyyymmdd    Y   院端
    strInfor = strInfor & Rpad(Nvl(rsTemp!入院经办时间), 16)    '登记时间    DATETIME    35  16  精确到秒，数据返回后格式为yyyymmddhhmiss后面以空格补位  Y   院端
    
    strInfor = strInfor & Lpad(str就诊分类, 1)                  '就诊分类    CHAR    51  1   2住院、4家床、O生育、   Y   院端
    
    gstrSQL = "Select 病区ID,床号,房间号 From 床位状况记录 D where 病区ID=[1] And 床号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取床位信息", CLng(Nvl(rsTemp!当前病区ID, 0)), CLng(Nvl(rsTemp!当前床号, 0)))
    If rsTemp.EOF Then
        str床位号 = Space(10)
    Else
        str床位号 = Trim(Nvl(rsTemp!房间号)) & "室" & Trim(Nvl(rsTemp!床号)) & "床"
        str床位号 = Lpad(str床位号, 10)
        str床位号 = Substr(str床位号, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.诊断类型,1,b.编码||'~^||'||b.名称,null)) as 入院诊断,  " & _
         "        max(decode(A.诊断类型,1,null,b.编码||'~^||'||b.名称)) as 确诊诊断 " & _
         " from 诊断情况 A,疾病编码目录 b " & _
         " where a.疾病id=b.id and  a.诊断类型 in(1,2) and a.诊断次序=1 and a.病人id=" & lng病人ID & " and a.主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定诊断编码和名称"
    
    Dim str入院诊断编码 As String
    Dim str入院诊断名称  As String
    Dim str确诊诊断编码 As String
    Dim str确诊诊断名称  As String
    
    If rsTemp.EOF Then
        str入院诊断编码 = ""
        str入院诊断名称 = ""
        str确诊诊断编码 = ""
        str确诊诊断名称 = ""
    Else
        str入院诊断名称 = Nvl(rsTemp!入院诊断)
        str确诊诊断名称 = Nvl(rsTemp!确诊诊断)
        If InStr(1, str入院诊断名称, "~^||") <> 0 Then
            str入院诊断编码 = Split(str入院诊断名称, "~^||")(0)
            str入院诊断名称 = Split(str入院诊断名称, "~^||")(1)
        Else
            str入院诊断编码 = ""
            str入院诊断名称 = ""
        End If
        If InStr(1, str确诊诊断名称, "~^||") <> 0 Then
            str确诊诊断编码 = Split(str确诊诊断名称, "~^||")(0)
            str确诊诊断名称 = Split(str确诊诊断名称, "~^||")(1)
        Else
            str确诊诊断编码 = ""
            str确诊诊断名称 = ""
        End If
    End If
        
    strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
    strInfor = strInfor & Lpad(Substr(str入院诊断名称, 1, 28), 30) '入院诊断名称    CHAR    68  30      y 院端
    strInfor = strInfor & Lpad(str确诊诊断编码, 16)  '确诊诊断编码    CHAR    98  16      N   院端
    strInfor = strInfor & Lpad(Substr(str确诊诊断名称, 1, 28), 30) '确诊诊断名称    CHAR    114 30      N   院端
    strInfor = strInfor & Lpad(str入院科室, 20)  '科别名称    CHAR    144 20  如：内科    Y   院端
    strInfor = strInfor & str床位号              '床位号  CHAR    164 10  如：2003室12床  N   院端
    strInfor = strInfor & Lpad(str转诊单号, 6)   '转诊单号    CHAR    174 6       N   院端
    strInfor = strInfor & Space(8)   '出院时间    DATE    180 8   系统利用患者结算数据的出院时间自动生成，医院端用空格补位即可。  N   无
    
    strInfor = strInfor & "A"   '传输标志    CHAR    188 1   A 入院登记，M 修改在院状态，C取消入院登记   Y   院端
    strInfor = strInfor & Space(16)   '传输时间    DATATIME    189 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，用于记录数据到达医保中心的时间，院端空格补位  N   中心
    
    '--------------------------------------------
    '调用医保接口操作
    '1004    9   206 实时住院登记数据提交
    住院结算冲销_大连 = 业务请求_大连(lng中心, 1004, strInfor, intinsure)
    If 住院结算冲销_大连 = False Then
        'Modify by ZHQ 2005-11-30
        '冲销后入院登记失败后可以继续，允许其补充登记即可
        ShowMsgbox "实时住院登记失败,请在医保帐户管理中进行补充登记!"
        住院结算冲销_大连 = True
        Exit Function
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 医保终止_大连() As Boolean
    mblnInit = False
    医保终止_大连 = True
End Function

Public Function 处方登记_大连(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String, ByVal intinsure As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp   As ADODB.Recordset
    '先写入单据头，再写入单据体
    '记录状态（1-新增;否则为删除），费用处单据只能整张单据删除后，再产生新单据
    On Error GoTo errHand
    处方登记_大连 = False
    If gbln住院明细时实上传 = False Then
        处方登记_大连 = True
        Exit Function
    End If
    gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
    If Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34 Then
        gstrSQL = " " & _
            " Select A.id,A.病人ID,F.住院号,A.NO,A.序号,A.医嘱序号,A.记录性质,A.记录状态,A.收费类别,D.类别,to_char(A.登记时间,'yyyyMMddhh24miss') 登记时间, " & _
            "        A.开单人 医生,V.编号 AS 医生编号,B.名称 开单部门,A.收费细目ID,A.计算单位,Round(A.实收金额/(A.数次*A.付数),2) as 实际价格,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传, " & _
            "        C.项目编码 医保项目编码 ,C.是否医保,decode(a.门诊标志,2,C.住院比额,C.统筹比额) as 统筹比额,F.住院次数 AS 主页id, " & _
            "        Nvl(K.标识码,Nvl(G.标识主码||G.标识子码,G.编码)) AS 国家编码,G.名称 AS 项目名称,K.名称 AS 剂型, " & _
            "        E.险类,E.中心,E.卡号,E.医保号,E.密码,E.人员身份,E.单位编码,E.顺序号,E.退休证号,E.帐户余额,E.当前状态, " & _
            "        E.病种ID,E.在职,E.年龄段,E.灰度级,to_char(E.就诊时间,'yyyyMMddhh24miss') 就诊时间 " & _
            " From 住院费用记录 A,部门表 B,收费类别 D,保险帐户 E,病人信息 F,病案主页 F1,收费细目 G,人员表 V," & _
            "       (Select J.名称,O.药品id,O.标识码 From 药品目录 O, 药品信息 H,药品剂型 J WHERE O.药名id=H.药名id and H.剂型=J.编码) K, " & _
            "       (Select M.项目编码,M.项目名称,M.是否医保,M.收费细目id,Q.统筹比额,Q.住院比额  From 保险支付项目 M,保险支付大类 Q Where M.险类=" & TYPE_大连市 & " and M.大类ID=Q.id) C " & _
            " Where   a.记录状态<>0 and   a.病人id=E.病人ID AND a.病人id=F.病人ID AND A.病人id=F1.病人id and F1.险类=82 AND F.主页id= F1.主页id  AND  a.开单人=V.姓名(+) AND a.收费细目id=k.药品id(+) AND a.收费细目id=G.id AND E.险类=" & TYPE_大连市 & "   AND A.收费类别=D.编码 AND  " & _
            "           A.记录性质=" & lng记录性质 & " and  A.记录状态=" & lng记录状态 & " And A.NO='" & str单据号 & "'" & _
            "           And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 "
        
        gstrSQL = gstrSQL & " Union all " & _
            " Select A.id,A.病人ID,F.住院号,A.NO,A.序号,A.医嘱序号,A.记录性质,A.记录状态,A.收费类别,D.类别,to_char(A.登记时间,'yyyyMMddhh24miss') 登记时间, " & _
            "        A.开单人 医生,V.编号 AS 医生编号,B.名称 开单部门,A.收费细目ID,A.计算单位,Round(A.实收金额/(A.数次*A.付数),2) as 实际价格,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传, " & _
            "        C.项目编码 医保项目编码 ,C.是否医保,decode(a.门诊标志,2,C.住院比额,C.统筹比额) as 统筹比额,F.住院次数 AS 主页id, " & _
            "        Nvl(K.标识码,Nvl(G.标识主码||G.标识子码,G.编码)) AS 国家编码,G.名称 AS 项目名称,K.名称 AS 剂型, " & _
            "        E.险类,E.中心,E.卡号,E.医保号,E.密码,E.人员身份,E.单位编码,E.顺序号,E.退休证号,E.帐户余额,E.当前状态, " & _
            "        E.病种ID,E.在职,E.年龄段,E.灰度级,to_char(E.就诊时间,'yyyyMMddhh24miss') 就诊时间 " & _
            " From 住院费用记录 A,部门表 B,收费类别 D,保险帐户 E,病人信息 F,病案主页 F1,收费细目 G,人员表 V," & _
            "       (Select J.名称,O.药品id,O.标识码 From 药品目录 O, 药品信息 H,药品剂型 J WHERE O.药名id=H.药名id and H.剂型=J.编码) K, " & _
            "       (Select M.项目编码,M.项目名称,M.是否医保,M.收费细目id,Q.统筹比额,Q.住院比额  From 保险支付项目 M,保险支付大类 Q Where M.险类=" & TYPE_大连开发区 & " and M.大类ID=Q.id) C " & _
            " Where  a.记录状态<>0 and a.病人id=E.病人ID AND a.病人id=F.病人ID AND A.病人id=F1.病人id and F1.险类=83 AND F.主页id= F1.主页id  AND  a.开单人=V.姓名(+) AND a.收费细目id=k.药品id(+) AND a.收费细目id=G.id AND E.险类=" & TYPE_大连开发区 & "   AND A.收费类别=D.编码 AND  " & _
            "           A.记录性质=[1] and  A.记录状态=[2] And A.NO=[3]" & _
            "           And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 " & _
            " Order by 病人ID"
    Else
        gstrSQL = " " & _
            " Select A.id,A.病人ID,F.住院号,A.NO,A.序号,A.医嘱序号,A.记录性质,A.记录状态,A.收费类别,D.类别,to_char(A.登记时间,'yyyyMMddhh24miss') 登记时间, " & _
            "        A.开单人 医生,V.编号 AS 医生编号,B.名称 开单部门,A.收费细目ID,A.计算单位,Round(A.实收金额/(A.数次*A.付数),2) as 实际价格,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传, " & _
            "        C.项目编码 医保项目编码 ,C.是否医保,decode(a.门诊标志,2,C.住院比额,C.统筹比额) as 统筹比额,F.住院次数 AS 主页id, " & _
            "        Nvl(K.标识码,Nvl(G.标识主码||G.标识子码,G.编码)) AS 国家编码,G.名称 AS 项目名称,K.名称 AS 剂型, " & _
            "        E.险类,E.中心,E.卡号,E.医保号,E.密码,E.人员身份,E.单位编码,E.顺序号,E.退休证号,E.帐户余额,E.当前状态, " & _
            "        E.病种ID,E.在职,E.年龄段,E.灰度级,to_char(E.就诊时间,'yyyyMMddhh24miss') 就诊时间 " & _
            " From 住院费用记录 A,部门表 B,收费类别 D,保险帐户 E,病人信息 F,病案主页 F1,收费细目 G,人员表 V," & _
            "       (Select J.名称,O.药品id,O.标识码 From 药品目录 O, 药品信息 H,药品剂型 J WHERE O.药名id=H.药名id and H.剂型=J.编码) K, " & _
            "       (Select M.项目编码,M.项目名称,M.是否医保,M.收费细目id,Q.统筹比额,Q.住院比额  From 保险支付项目 M,保险支付大类 Q Where M.险类=" & TYPE_大连市 & " and M.大类ID=Q.id) C " & _
            " Where   a.记录状态<>0 and   a.病人id=E.病人ID AND a.病人id=F.病人ID AND A.病人id=F1.病人id and F1.险类=82 AND F.住院次数= F1.主页id  AND  a.开单人=V.姓名(+) AND a.收费细目id=k.药品id(+) AND a.收费细目id=G.id AND E.险类=" & TYPE_大连市 & "   AND A.收费类别=D.编码 AND  " & _
            "           A.记录性质=" & lng记录性质 & " and  A.记录状态=" & lng记录状态 & " And A.NO='" & str单据号 & "'" & _
            "           And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 "
        
        gstrSQL = gstrSQL & " Union all " & _
            " Select A.id,A.病人ID,F.住院号,A.NO,A.序号,A.医嘱序号,A.记录性质,A.记录状态,A.收费类别,D.类别,to_char(A.登记时间,'yyyyMMddhh24miss') 登记时间, " & _
            "        A.开单人 医生,V.编号 AS 医生编号,B.名称 开单部门,A.收费细目ID,A.计算单位,Round(A.实收金额/(A.数次*A.付数),2) as 实际价格,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传, " & _
            "        C.项目编码 医保项目编码 ,C.是否医保,decode(a.门诊标志,2,C.住院比额,C.统筹比额) as 统筹比额,F.住院次数 AS 主页id, " & _
            "        Nvl(K.标识码,Nvl(G.标识主码||G.标识子码,G.编码)) AS 国家编码,G.名称 AS 项目名称,K.名称 AS 剂型, " & _
            "        E.险类,E.中心,E.卡号,E.医保号,E.密码,E.人员身份,E.单位编码,E.顺序号,E.退休证号,E.帐户余额,E.当前状态, " & _
            "        E.病种ID,E.在职,E.年龄段,E.灰度级,to_char(E.就诊时间,'yyyyMMddhh24miss') 就诊时间 " & _
            " From 住院费用记录 A,部门表 B,收费类别 D,保险帐户 E,病人信息 F,病案主页 F1,收费细目 G,人员表 V," & _
            "       (Select J.名称,O.药品id,O.标识码 From 药品目录 O, 药品信息 H,药品剂型 J WHERE O.药名id=H.药名id and H.剂型=J.编码) K, " & _
            "       (Select M.项目编码,M.项目名称,M.是否医保,M.收费细目id,Q.统筹比额,Q.住院比额  From 保险支付项目 M,保险支付大类 Q Where M.险类=" & TYPE_大连开发区 & " and M.大类ID=Q.id) C " & _
            " Where  a.记录状态<>0 and a.病人id=E.病人ID AND a.病人id=F.病人ID AND A.病人id=F1.病人id and F1.险类=83 AND F.住院次数= F1.主页id  AND  a.开单人=V.姓名(+) AND a.收费细目id=k.药品id(+) AND a.收费细目id=G.id AND E.险类=" & TYPE_大连开发区 & "   AND A.收费类别=D.编码 AND  " & _
            "           A.记录性质=[1] and  A.记录状态=[2] And A.NO=[3]" & _
            "           And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 " & _
            " Order by 病人ID"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "处方登记", lng记录性质, lng记录状态, str单据号)
    If rsTemp.RecordCount = 0 Then
        MsgBox "未找到处方记录，向医保服务器传输数据失败！[处方登记]", vbInformation, gstrSysName
        Exit Function
    End If
    处方登记_大连 = 上传处方_大连(rsTemp, intinsure)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function 上传处方_大连(ByVal rsExse As ADODB.Recordset, ByVal intinsure As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:rsExse-明细数据
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------


    Dim lng病人ID As Long
    Dim curTotal As Currency
    Dim blnUpload As Boolean
    Dim rsPara As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim str明细 As String
    Dim str项目统计分类 As String
    Dim str诊断编码 As String
    Dim str诊断名称 As String
    Dim strInsertSQL As String
    
    
    Dim strTmp As String
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "select * from 保险参数 where 险类 in (82,83)"
    zlDatabase.OpenRecordset rsPara, gstrSQL, "上传处方读取参数"
    With rsExse
        Do While Not .EOF
            lng病人ID = Nvl(!病人ID, 0)
            '确定相关数据
            '上传明细记录,实时医疗明细数据
                
            If Nvl(!险类, 0) = TYPE_大连开发区 Then '开发区
                str明细 = Lpad(gstr医院编码_大连, 6)     '医院代号    CHAR    1   6       院端填写
                str明细 = str明细 & Lpad(Nvl(!医保号), 10)  '保险编号    CHAR    7   10      院端填写
            Else
                str明细 = Lpad(gstr医院编码_大连, 4)     '医院代码    CHAR    1   4       院端
                str明细 = str明细 & Lpad(Nvl(!医保号), 8)   '个人编号    CHAR    5   8       院端
            End If
            
            str明细 = str明细 & Lpad(Nvl(!住院号, 0), 10) '病志号  CHAR    13  10  门诊明细以空格补位,住院是住院号  院端
            str明细 = str明细 & Lpad(Nvl(!顺序号, 0), 4)   '治疗序号    NUM 23  4   住院明细：必须等于入院登记时治疗序号门诊明细:                         必须等于本次结算治疗序号 院端
            
            'Modified By 朱玉宝 2004-07-29 原因：处理NO号
            str明细 = str明细 & Lpad(Mid(Nvl(!NO, "00000000"), 2, 7), 10)     '处方号  NUM 27  10      院端
            str明细 = str明细 & Lpad(CStr(Nvl(!序号, 0)), 10)      '处方项目序号    NUM 37  10  对应处方号的记价项目序号    院端
            
            '开发区为单据号  CHAR    41  10  医嘱号，    院端填写
            str明细 = str明细 & Lpad(Nvl(!医嘱序号, " "), 10)     '医嘱号  CHAR    47  10  处方对应医嘱的医嘱记录号，门诊明细或没有医嘱的医院以空格补位    院端
            str明细 = str明细 & Get就诊分类(0, Nvl(!灰度级))      '就诊分类    CHAR    57  1   取值详见"就诊分类"说明  院端
            str明细 = str明细 & Rpad(Nvl(!登记时间), 16)      '处方生成时间（投药时间）    DATETIME    58  16  精确到秒格式为：yyyymmddhhmiss后面以空格补位    院端
            str明细 = str明细 & Lpad(Nvl(!国家编码), 20)      '项目代码    CHAR    74  20  计价项目代码    院端
            str明细 = str明细 & Lpad(Nvl(!项目名称), 20)      '项目名称    CHAR    94  20      院端

            If !是否医保 = 1 Then
                str明细 = str明细 & Lpad(1 - Nvl(!统筹比额, 0) / 100, 6) '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
            Else
                str明细 = str明细 & Lpad(1, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
            End If
            rsPara.Filter = 0
            rsPara.Filter = " 参数名='" & Nvl(!类别) & "' and 险类=" & Nvl(!险类, 0)
            str项目统计分类 = ""
            If Not rsPara.EOF Then
                strTmp = Nvl(rsPara!参数值)
                If InStr(1, strTmp, ";") <> 0 And strTmp <> ";" Then
                    strTmp = Split(strTmp, ";")(1)
                    If strTmp <> "" Then
                        str项目统计分类 = Substr(strTmp, 1, 1)
                        str明细 = str明细 & Substr(strTmp, 1, 1)   '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                    Else
                        str明细 = str明细 & Space(1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                    End If
                Else
                    str明细 = str明细 & Space(1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                End If
            Else
                    str明细 = str明细 & Space(1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
            End If
            
            '2005-08-02开发区升级
            If Nvl(!险类, 0) = TYPE_大连开发区 Then
                str明细 = str明细 & Lpad(Nvl(!数量), 10)  '数量    NUM 121 6   冲方划价为负值  院端
                str明细 = str明细 & Lpad(Nvl(!实际价格), 10) '单价    NUM 127 8   不允许出现负值  院端
            Else
                str明细 = str明细 & Lpad(Nvl(!数量), 6)  '数量    NUM 121 6   冲方划价为负值  院端
                str明细 = str明细 & Lpad(Nvl(!实际价格), 8) '单价    NUM 127 8   不允许出现负值  院端
            End If
            str明细 = str明细 & Lpad(Nvl(!计算单位), 4) '单位    CHAR    135 4       院端
            str明细 = str明细 & Lpad(Nvl(!剂型), 20)      '剂型    CHAR    139 20  针剂、片剂…    院端
            str明细 = str明细 & Lpad(Nvl(!医生), 8)      '医师姓名    CHAR    159 8       院端
            '确定诊断情况
            
            strTmp = Get入院诊断(Nvl(!病人ID), Nvl(!主页ID, 0), False, True)
            If InStr(1, strTmp, "|") <> 0 Then
                
                str明细 = str明细 & Lpad(Split(strTmp, "|")(1), 16)     '诊断编码    CHAR    167 16      院端
                strTmp = Split(strTmp, "|")(0)
                strTmp = Lpad(strTmp, 30)
                str诊断编码 = Split(strTmp, "|")(1)
                str诊断名称 = Split(strTmp, "|")(0)
                    
                str明细 = str明细 & strTmp     '诊断名称    CHAR    183 30      院端
            Else
                str明细 = str明细 & Space(16)      '诊断编码    CHAR    167 16      院端
                str明细 = str明细 & Space(30)     '诊断名称    CHAR    183 30      院端
                str诊断编码 = ""
                str诊断名称 = ""
            End If
            
            str明细 = str明细 & Space(16)     '传输时间    DATETIME    213 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，院端空格补位  中心
     
            '上传明细
            '1003    7   230 实时医疗明细数据提交
            上传处方_大连 = 业务请求_大连(IIf(Nvl(!险类, 0) = TYPE_大连开发区, 2, 1), 1003, str明细, intinsure)
            If 上传处方_大连 = False Then
                ShowMsgbox "门诊结算时医疗明细数据提交失败,不能继续!"
                Exit Function
            End If

            '为病人费用记录打上标记，以便随时上传
            'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
            zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            .MoveNext
        Loop
    End With
    上传处方_大连 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo errHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDbUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
errHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo errHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDbUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
errHand:
End Sub

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '功能：提示消息框
    '参数：strMsgInfor-提示信息
    '     blnYesNo-是否提供YES或NO按钮
    '返回：blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Private Function 医嘱明细数据提交(ByVal lng医嘱ID As Long, ByVal str住院号 As String, ByVal str项目统计分类 As String, ByVal intinsure As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:提取医嘱明细
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    
    '5.  实时医嘱数据提交接口
    
    '--由于目前大连市医嘱上传是否强制要求还未确定,医嘱明细提交暂时屏蔽,直接返回为成功状态标志
    医嘱明细数据提交 = True
    Exit Function

    '开发区无医嘱接口
    If intinsure = TYPE_大连开发区 Then
        医嘱明细数据提交 = True
        Exit Function
    End If
    
    
    gstrSQL = " " & _
         " select A.ID,A.相关id as 分组号,decode(A.医嘱期效,1,1,0) as  医嘱类型,A.医嘱内容," & _
         "          A.开嘱医生 as 下医嘱医生,to_char(A.开始执行时间,'yyyymmddhh24miss') as 开始执行时间,A.校对护士 as 执行医嘱护士姓名," & _
         "          A.停嘱医生 as 停医嘱医生,to_char(A.停嘱时间,'yyyymmddhh24miss') as 停医嘱时间,A.医生嘱托 as 附加说明, " & _
         "          Decode(B.类别,'Z',decode(B.操作类型,'5','0000','6','0000',a.医嘱内容),a.医嘱内容) as 医嘱描述," & _
         "          A.单次用量 as 药品单量,B.计算单位 as 剂量单位," & _
         "        A.执行频次,A.频率次数, " & _
         "        A.开嘱时间 as 下医嘱时间" & _
         "        " & _
         " from 病人医嘱记录 A,诊疗项目目录 B  " & _
         " Where A.诊疗项目id=B.id and A.id=" & lng医嘱ID
         
    Err = 0
    On Error GoTo errHand:
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取医嘱明细记录"
    
    If rsTemp.EOF Then
        ShowMsgbox "无对应的医嘱记录!"
        Exit Function
    End If
    
    With g病人身份_大连
        strInfor = Lpad(gstr医院编码_大连, 4)   '1   医院代码    CHAR    1   4       院端
        strInfor = strInfor & Lpad(.个人编号, 8) '2   个人编号    CHAR    5   8       院端
        strInfor = strInfor & Lpad(.治疗序号, 4)     '3   治疗序号    NUM 13  4   必须等于入院时治疗序号  院端
        strInfor = strInfor & Lpad(str住院号, 10)    '4   病志号  CHAR    17  10  要跟明细数据“病志号”对应  院端
        strInfor = strInfor & Lpad(lng医嘱ID, 10)    '5   医嘱号  CHAR    27  10  对应明细数据的医嘱号    院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!医嘱类型, 0), 1)   '6   医嘱类型    CHAR    37  1   1 长嘱，0 临时医嘱
        strInfor = strInfor & Substr(Lpad(Nvl(rsTemp!医嘱描述), 80), 1, 80) '7   医嘱内容描述??  CHAR    38  80  如果描述出院信息请用‘0000’    院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!下医嘱医生), 8)   '8   下医嘱医师姓名  CHAR    118 8       院端
        strInfor = strInfor & Rpad(Nvl(rsTemp!开始执行时间), 16) '9   开始执行时间    DATATIME    126 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，此项必添  院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!执行医嘱护士姓名), 8)  '10  执行医嘱护士姓名    CHAR    142 8       院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!停医嘱医生), 8)  '11  终止医嘱医师姓名    CHAR    150 8       院端
        strInfor = strInfor & Rpad(Nvl(rsTemp!停医嘱时间), 16) '12  终止医嘱时间    DATATIME    158 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，对于长期医嘱此项必添对于临时医嘱可以空格补位  院端
        strInfor = strInfor & Substr(Lpad(Nvl(rsTemp!附加说明), 30), 1, 30) '13  备注    CHAR    174 30  用于临时医嘱执行反馈或者其他描述    院端
        strInfor = strInfor & Space(16)                  '14  传输时间    DATATIME    204 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，用于记录数据到达医保中心的时间，院端空格补位  中心
    End With
    
    '1005    8   274 实时医嘱传输
    医嘱明细数据提交 = 业务请求_大连(g病人身份_大连.医保中心, 1005, strInfor, intinsure)
    Exit Function
errHand:
    '如果没装医嘱就不执行
    医嘱明细数据提交 = True
End Function

Private Function Get病人变动记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取病人的变动情况
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "" & _
        "   Select  床号,附加床位,开始时间,终止时间,床位等级id " & _
        "   From 病人变动记录  " & _
        "   Where  病人id=" & lng病人ID & " and 主页id=" & lng主页ID & " and 床号 is not null"
    Err = 0
    On Error GoTo errHand:
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人变动情况"
    Set Get病人变动记录 = rsTemp
'    Call WriteDebugInfor_大连("Get病人变动记录", lng病人id)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Set Get病人变动记录 = Nothing
    Exit Function
End Function
Private Function Get住院虚拟记录(ByVal lng病人ID As Long, Optional lng主页ID As Long = 0, Optional ByVal intinsure As Integer) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取本次虚拟未结记录
    '--入参数:
    '--出参数:
    '--返  回:未结费用
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset


    '--赵元礼增加,主要排除由于冲销结帐单产生的重复明细记录
    '--4-26,周顺利增加对婴儿费的判断,排除婴儿费不参与计算
    If lng主页ID <> 0 Then
        '主要是对历次的结算进行虚拟结算
        strSQL = _
            "   Select  A.记录性质,A.NO,A.序号,A.床号," & _
            "           A.病人ID,A.主页ID,nvl(A.婴儿费,0) as 婴儿费," & _
            "           A.保险大类ID,A.收费类别,A.收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门," & _
            "           Decode(Sign(Instr(B.规格,'┆')),0,B.规格,Substr(B.规格,1,Instr(B.规格,'┆')-1)) as 规格," & _
            "           Decode(Sign(Instr(B.规格,'┆')),0,B.规格,Substr(B.规格,Instr(B.规格,'┆')+1)) as 产地," & _
            "           A.数量,Decode(A.数量,0,0,Round(A.金额/A.数量,4)) as 价格,A.金额,A.医生,w.编号 as 医生编号,A.发生时间,A.登记时间," & _
            "           A.是否急诊,A.保险项目否,A.摘要,C.项目编码 as 医保项目编码," & _
            "           C.项目名称 as 医保项目名称,Q.参数值,Q.参数名,J.统筹比额,J.住院比额,J.特准定额,J.算法" & _
            "   From (" & _
            "           Select  C.险类,Mod(A.记录性质,10) as 记录性质,A.床号,A.NO,Nvl(A.价格父号,序号) as 序号,A.病人ID,A.主页ID,Nvl(A.婴儿费,0) as 婴儿费," & _
            "                   A.开单人 as 医生,A.开单部门ID,A.收费类别,A.收费细目ID,Nvl(A.保险大类ID,0) as 保险大类ID,Avg(Nvl(A.付数,1)*A.数次) as 数量," & _
            "                   Sum(A.标准单价) as 标准单价,Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 金额,A.发生时间,A.登记时间,Nvl(A.是否急诊,0) as 是否急诊,Nvl(A.保险项目否,0) as 保险项目否,A.摘要" & _
            "           From 住院费用记录 A,收入项目 B,保险帐户 C" & _
            "           Where a.记录状态<>0 and A.主页id=" & lng主页ID & " and  A.记帐费用=1 and A.病人id=C.病人id And A.收入项目ID=B.ID And nvl(A.婴儿费,0)=0 And A.病人ID=" & lng病人ID & " and C.险类=" & intinsure & _
            "           Group by    C.险类,Mod(A.记录性质,10),A.NO,Nvl(A.价格父号,序号),A.病人ID,A.主页ID,A.床号,Nvl(A.婴儿费,0),A.开单人," & _
            "                       A.开单部门ID,A.收费类别,A.收费细目ID,Nvl(A.保险大类ID,0),A.发生时间,A.登记时间,Nvl(A.是否急诊,0),Nvl(A.保险项目否,0),A.摘要" & _
            "           Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0) A,收费细目 B,部门表 X," & _
            "           (Select * From 保险支付项目 Where 险类=" & intinsure & ") C," & _
            "           (Select M.编码, L.参数名,L.参数值 from 收费类别 M,保险参数 L  Where M.类别=L.参数名 and L.险类=" & intinsure & ")  Q," & _
            "           (Select * from 保险支付大类  Where 险类=" & intinsure & ")  J,人员表 W" & _
            "   Where     A.收费细目ID=B.ID and a.医生=w.姓名(+) and C.大类id=J.ID and a.收费类别=Q.编码(+) And A.收费细目ID=C.收费细目ID And A.开单部门ID=X.ID"
    Else
        strSQL = _
            "   Select  A.记录性质,A.NO,A.序号,A.床号," & _
            "           A.病人ID,A.主页ID,nvl(A.婴儿费,0) as 婴儿费," & _
            "           A.保险大类ID,A.收费类别,A.收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门," & _
            "           Decode(Sign(Instr(B.规格,'┆')),0,B.规格,Substr(B.规格,1,Instr(B.规格,'┆')-1)) as 规格," & _
            "           Decode(Sign(Instr(B.规格,'┆')),0,B.规格,Substr(B.规格,Instr(B.规格,'┆')+1)) as 产地," & _
            "           A.数量,Decode(A.数量,0,0,Round(A.金额/A.数量,4)) as 价格,A.金额,A.医生,w.编号 as 医生编号,A.发生时间,A.登记时间," & _
            "           A.是否急诊,A.保险项目否,A.摘要,C.项目编码 as 医保项目编码," & _
            "           C.项目名称 as 医保项目名称,Q.参数值,Q.参数名,J.统筹比额,J.住院比额,J.特准定额,J.算法" & _
            "   From (" & _
            "           Select  C.险类,Mod(A.记录性质,10) as 记录性质,A.床号,A.NO,Nvl(A.价格父号,序号) as 序号,A.病人ID,A.主页ID,Nvl(A.婴儿费,0) as 婴儿费," & _
            "                   A.开单人 as 医生,A.开单部门ID,A.收费类别,A.收费细目ID,Nvl(A.保险大类ID,0) as 保险大类ID,Avg(Nvl(A.付数,1)*A.数次) as 数量," & _
            "                   Sum(A.标准单价) as 标准单价,Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 金额,A.发生时间,A.登记时间,Nvl(A.是否急诊,0) as 是否急诊,Nvl(A.保险项目否,0) as 保险项目否,A.摘要" & _
            "           From 住院费用记录 A,收入项目 B,保险帐户 C" & _
            "           Where a.记录状态<>0 and  A.记帐费用=1 and A.病人id=C.病人id And A.收入项目ID=B.ID And nvl(A.婴儿费,0)=0 And A.病人ID=" & lng病人ID & " and C.险类=" & intinsure & _
            "           Group by    C.险类,Mod(A.记录性质,10),A.NO,Nvl(A.价格父号,序号),A.病人ID,A.主页ID,A.床号,Nvl(A.婴儿费,0),A.开单人," & _
            "                       A.开单部门ID,A.收费类别,A.收费细目ID,Nvl(A.保险大类ID,0),A.发生时间,A.登记时间,Nvl(A.是否急诊,0),Nvl(A.保险项目否,0),A.摘要" & _
            "           Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0) A,收费细目 B,部门表 X," & _
            "           (Select * From 保险支付项目 Where 险类=" & intinsure & ") C," & _
            "           (Select M.编码, L.参数名,L.参数值 from 收费类别 M,保险参数 L  Where M.类别=L.参数名 and L.险类=" & intinsure & ")  Q," & _
            "           (Select * from 保险支付大类  Where 险类=" & intinsure & ")  J,人员表 W" & _
            "   Where     A.收费细目ID=B.ID and a.医生=w.姓名(+) and C.大类id=J.ID and a.收费类别=Q.编码(+) And A.收费细目ID=C.收费细目ID And A.开单部门ID=X.ID"
    End If
    Err = 0
    On Error GoTo errHand:
    zlDatabase.OpenRecordset rsTmp, strSQL, "获取本次医保未结费用"
    Set Get住院虚拟记录 = rsTmp
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    Set Get住院虚拟记录 = Nothing
    Exit Function
End Function


Private Function Set门诊挂号结算或冲销(ByVal bln冲销 As Boolean, lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, strSelfNo As String) As Boolean
  '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；
    '      cur个人帐户   从个人帐户中支出的金额
    
    Set门诊挂号结算或冲销 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 挂号结算_大连(ByVal lng结帐ID As Long) As Boolean
     挂号结算_大连 = Set门诊挂号结算或冲销(False, lng结帐ID, 0, 0, 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 挂号冲销_大连(ByVal lng结帐ID As Long) As Boolean
    挂号冲销_大连 = Set门诊挂号结算或冲销(False, lng结帐ID, 0, 0, 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function 在院病人信息_大连(lng病人ID As Long, lng主页ID As Long, ByVal intinsure As Integer) As Boolean
    Dim str入院经办时间 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str就诊分类 As String
    Dim str入院科室 As String
    Dim str床位号 As String
    Dim str转诊单号 As String
    Dim lng中心 As Long
    
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    
    On Error GoTo errHand
    
    '读取病人的相关保险信息

    gstrSQL = "select 险类,病人ID,人员身份,医保号,顺序号,灰度级 From 保险帐户 where  险类=" & intinsure & "  and 病人id=" & lng病人ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "入院读取保险帐户信息"
    If rsTemp.EOF Then
        ShowMsgbox "在保险帐户中无该病人的保险信息!"
        Exit Function
    End If
    str转诊单号 = Nvl(rsTemp!人员身份)
    lng中心 = IIf(intinsure = 83, 2, 1)
    If lng中心 = 2 Then
        strInfor = Lpad(gstr医院编码_大连, 6) '医院代码    CHAR    1   6      Y   院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!医保号), 10)     '保险编号    CHAR    7   10      院端填写
    Else
        strInfor = Lpad(gstr医院编码_大连, 4) '医院代码    CHAR    1   4       Y   院端
        strInfor = strInfor & Lpad(Nvl(rsTemp!医保号), 8)     '保险编号    CHAR    5   8       Y   院端
    End If
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!顺序号, 1), 4)      '治疗序号    NUM 13  4   必须等于入院时治疗序号  Y   院端
    
    '内部标识:5-普通住院,6-家庭病床住院,7-生育保险住院,8-工伤保险住院
    '医保标识:2-住院结算,4-家庭病床结算,O-生育保险住院结算,Q-工伤保险结算
    
    str就诊分类 = Decode(Nvl(rsTemp!灰度级, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '读取病人信息
    gstrSQL = "Select C.住院号,C.当前病区id,C.当前床号,A.登记人 经办人,B.名称 入院科室,to_char(A.登记时间,'yyyyMMddhh24miss') 入院经办时间," & _
            " to_char(A.登记时间,'yyyyMMdd') 入院日期" & _
            " From 病案主页 A,部门表 B,病人信息 C" & _
            " Where A.病人id=C.病人id and C.病人id=[1]" & _
            "       and A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院信息", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    str入院科室 = Nvl(rsTemp!入院科室)
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!住院号, 0), 10)      '病志号  CHAR    17  10      Y   院端门诊暂定为空，住院就为住院号
    strInfor = strInfor & Lpad(Nvl(rsTemp!入院日期), 8)      '入院日期 Date 27  8   患者实际入院日期，格式为yyyymmdd    Y   院端
    strInfor = strInfor & Rpad(Nvl(rsTemp!入院经办时间), 16)     '登记时间    DATETIME    35  16  精确到秒，数据返回后格式为yyyymmddhhmiss后面以空格补位  Y   院端
    strInfor = strInfor & Lpad(str就诊分类, 1)                  '就诊分类    CHAR    51  1   2住院、4家床、O生育、   Y   院端

    gstrSQL = "Select 病区ID,床号,房间号 From 床位状况记录 D where 病区ID=[1] And 床号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取床位信息", CLng(Nvl(rsTemp!当前病区ID, 0)), CLng(Nvl(rsTemp!当前床号, 0)))
    If rsTemp.EOF Then
        str床位号 = Space(10)
    Else
        str床位号 = Trim(Nvl(rsTemp!房间号)) & "室" & Trim(Nvl(rsTemp!床号)) & "床"
        str床位号 = Lpad(str床位号, 10)
        str床位号 = Substr(str床位号, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.诊断类型,1,b.编码||'~^||'||b.名称,null)) as 入院诊断,  " & _
         "        max(decode(A.诊断类型,1,null,b.编码||'~^||'||b.名称)) as 确诊诊断 " & _
         " from 诊断情况 A,疾病编码目录 b " & _
         " where a.疾病id=b.id and  a.诊断类型 in(1,2) and a.诊断次序=1 and a.病人id=" & lng病人ID & "and a.主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定诊断编码和名称"
    Dim str入院诊断编码 As String
    Dim str入院诊断名称  As String
    Dim str确诊诊断编码 As String
    Dim str确诊诊断名称  As String
    
    If rsTemp.EOF Then
        str入院诊断编码 = ""
        str入院诊断名称 = ""
        str确诊诊断编码 = ""
        str确诊诊断名称 = ""
    Else
        str入院诊断名称 = Nvl(rsTemp!入院诊断)
        str确诊诊断名称 = Nvl(rsTemp!确诊诊断)
        If InStr(1, str入院诊断名称, "~^||") <> 0 Then
            str入院诊断编码 = Split(str入院诊断名称, "~^||")(0)
            str入院诊断名称 = Split(str入院诊断名称, "~^||")(1)
        Else
            str入院诊断编码 = ""
            str入院诊断名称 = ""
        End If
        If InStr(1, str确诊诊断名称, "~^||") <> 0 Then
            str确诊诊断编码 = Split(str确诊诊断名称, "~^||")(0)
            str确诊诊断名称 = Split(str确诊诊断名称, "~^||")(1)
        Else
            str确诊诊断编码 = ""
            str确诊诊断名称 = ""
        End If
    End If
    strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
    strInfor = strInfor & Lpad(Substr(str入院诊断名称, 1, 28), 30) '入院诊断名称    CHAR    68  30      y 院端
    strInfor = strInfor & Lpad(str确诊诊断编码, 16)  '确诊诊断编码    CHAR    98  16      N   院端
    strInfor = strInfor & Lpad(Substr(str确诊诊断名称, 1, 28), 30) '确诊诊断名称    CHAR    114 30      N   院端
    strInfor = strInfor & Lpad(str入院科室, 20)  '科别名称    CHAR    144 20  如：内科    Y   院端
    strInfor = strInfor & Lpad(Substr(str床位号, 1, 10), 10)         '床位号  CHAR    164 10  如：2003室12床  N   院端
    strInfor = strInfor & Lpad(str转诊单号, 6)   '转诊单号    CHAR    174 6       N   院端
    strInfor = strInfor & Space(8)   '出院时间    DATE    180 8   系统利用患者结算数据的出院时间自动生成，医院端用空格补位即可。  N   无
    strInfor = strInfor & "M"   '传输标志    CHAR    188 1   A 入院登记，M 修改在院状态，C取消入院登记   Y   院端
    strInfor = strInfor & Space(16)   '传输时间    DATATIME    189 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，用于记录数据到达医保中心的时间，院端空格补位  N   中心
    
    '1004    9   206 实时住院登记数据提交
    在院病人信息_大连 = 业务请求_大连(lng中心, 1004, strInfor, intinsure)
    If 在院病人信息_大连 = False Then
        ShowMsgbox "实时住院登记数据提交失败!"
        Exit Function
    End If
    在院病人信息_大连 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetItemInfo_大连(ByVal lngPatiID As Long, ByVal lngItemID As Long, ByVal intinsure As Integer) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取大连病人的相关提示信息
    '--入参数:
    '--出参数:
    '--返  回:提示串
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim str医疗付款方式 As String
    Dim int险类 As Integer
    Dim bln在院 As Boolean
    Dim dbl统筹比例 As Double
    Dim strMsgInfor As String
    
'    '第一步:确定是否医保病人
'    gstrSQL = "Select 病人id,险类,nvl(当前状态,0) as 状态 from 保险帐户  where 病人id=" & lngPatiID & " and 险类=" & intinsure
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "判断是否为医保病人!"
'    If rsTemp.EOF Then
'        rsTemp.Close
'        GetItemInfo_大连 = ""
'        Exit Function
'    End If
'
'    int险类 = NVL(rsTemp!险类, 0)
'    bln在院 = NVL(rsTemp!状态, 0) > 0
    
    '第二步:确定医疗付款方式
    gstrSQL = "Select 医疗付款方式,decode(当前科室id,null,0,1) as 在院状态,nvl(险类,0) as 险类 from 病人信息 where 病人id=" & lngPatiID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取医疗付款式"
    
    str医疗付款方式 = Nvl(rsTemp!医疗付款方式)
    int险类 = rsTemp!险类
    If rsTemp!在院状态 = 1 Then
        bln在院 = True
    Else
        bln在院 = False
    End If
    
        
    '第三步：确定收费细目的相关数据
    gstrSQL = "" & _
        "   Select a.险类,b.编码,b.名称,b.性质,b.算法,a.项目名称,b.统筹比额,b.特准定额,b.住院比额,a.是否医保 " & _
        "   From 保险支付项目 a,保险支付大类 b " & _
        "   where a.大类id=b.id and a.险类=b.险类 and a.收费细目id=" & lngItemID & _
        "     and a.险类=decode('" & str医疗付款方式 & "'," & _
        "    '社会基本医疗保险',decode(" & int险类 & ",0,82," & int险类 & "),82)"
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取保险支付比例"
    strMsgInfor = ""
    If rsTemp.RecordCount = 0 Then
        GetItemInfo_大连 = ""
        ShowMsgbox "未在保险支付项目中设置相关比例或对应关系,请检查!"
        Exit Function
    End If
    If InStr(1, "社会基本医疗保险;企业离休;工伤保险;生育保险;商业保险;", IIf(str医疗付款方式 = "", "D", str医疗付款方式) & ";") <> 0 Then
        '   医疗付款方式为大连市医保、外地医保、企业离休、工伤保险、生育保险、商业保险的，报销比例按照大连市医保接口的医保项目管理中的医保大类定义中的报销比例进行提示
        '   医疗付款方式为开发区医保的，报销比例按照开发区医保接口的医保大类定义中的报销比例进行提示
        If bln在院 Then
            If Nvl(rsTemp!算法, 0) = 2 Then
                 '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算
                 strMsgInfor = "该项目固定报销:每次" & Format(Nvl(rsTemp!特准定额, 0), "#####0.00;-####0.00; ;") & "元"
            Else
                 If rsTemp!住院比额 < 100 Then
                 strMsgInfor = "该项目住院自付比例:" & Format(100 - Nvl(rsTemp!住院比额, 0), "#####0.00;-####0.00; ;") & "%"
                 End If
            End If
        Else
                 If rsTemp!统筹比额 < 100 Then
                 strMsgInfor = "该项目门诊自付比例:" & Format(100 - Nvl(rsTemp!统筹比额, 0), "#####0.00;-####0.00; ;") & "%"
                 End If
        End If
    ElseIf InStr(1, "公费医疗;合同单位;", IIf(str医疗付款方式 = "", "D", str医疗付款方式) & ";") <> 0 Then
        '   医疗付款方式为公费医疗、合同单位的，报销比例按照大连市医保接口的医保项目管理中的事业公费比例定义进行提示。
        If Val(Nvl(rsTemp!项目名称)) < 100 Then
        strMsgInfor = "该项目自付比例:" & Format(100 - Val(Nvl(rsTemp!项目名称)), "#####0.00;-####0.00; ;") & "%"
        End If
    End If
    If strMsgInfor <> "" Then
        ShowMsgbox strMsgInfor
    End If
    GetItemInfo_大连 = ""
End Function

Public Sub WriteDebugInfor_大连(ByVal strCallFunctionName As String, ByVal lng病人ID As Long)
'将调试信息写入文件中
        Dim objFile As New FileSystemObject
        Dim objText As TextStream
        If gblnDebug = False Then Exit Sub
        
        Dim strFile As String
        Dim rsTemp As New ADODB.Recordset
        
        gstrSQL = "Select '病人id:'||a.病人id||'姓名'||a.姓名||'住院号:'||a.住院号 as 信息  From 病人信息 a,保险帐户 b Where a.出院时间 Is Null and a.入院时间 is not null and  b.当前状态=0 And a.病人ID=" & lng病人ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取当前状态"
        If Not rsTemp.EOF Then
            '存在变化,需记录下来
            strFile = App.Path & "\医保病人当前状态变化跟踪.txt"
            
            If Not Dir(strFile) <> "" Then
                objFile.CreateTextFile strFile
            End If
            Set objText = objFile.OpenTextFile(strFile, ForAppending)
            objText.WriteLine strCallFunctionName & Space(10) & Nvl(rsTemp!信息)
            objText.WriteLine Format(Now, "yyyy-mm-dd")
            objText.Close
        End If
End Sub

Public Sub 更新病种_大连(ByVal lngPatiID As Long, ByVal lngPageID As Long, intinsure As Integer)
    '允许修改上一次出院的诊断
    Dim strIcdCode As String, strIcdName As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo ErrName
    
    gstrSQL = "select 描述信息 from 诊断情况 where 诊断类型=3 and 诊断次序=1 and 病人ID=[1] and 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取上次出院诊断", lngPatiID, lngPageID - 1)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    strIcdCode = Mid(rsTemp(0).Value, 2, InStr(1, rsTemp(0).Value, ")") - 2)
    strIcdName = Mid(rsTemp(0).Value, InStr(1, rsTemp(0).Value, ")") + 1)
    Call frm诊断信息.ShowME(lngPatiID, strIcdCode, strIcdName)
    
    If strIcdCode = "" Then Exit Sub
    
    'ZHQ 2006-03-15 modify
    '由于HIS停用“诊断情况”表，改为视图，所以需要修改此表UPDATE语句
'    gstrSQL = "Update 诊断情况 Set 描述信息='(" & strIcdCode & ")" & strIcdName & "'" & _
'            "  Where 诊断类型=3 And 诊断次序=1 And 病人ID=" & lngPatiID & " And 主页ID=" & lngPageID - 1
    gstrSQL = "Update 病人诊断记录 Set 诊断描述='(" & strIcdCode & ")" & strIcdName & "'" & _
            "  Where 记录来源=2 And 诊断类型=3 And 诊断次序=1 And 病人ID=" & lngPatiID & " And 主页ID=" & lngPageID - 1
    
    Call SQLTest(App.ProductName, "更新出院诊断", gstrSQL)
    gcnOracle.Execute gstrSQL
    Call SQLTest
    Exit Sub

ErrName:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function IsParaBig(intinsure) As Boolean
    '检查当前医保是否允许门诊大病走住院比例
    Dim rsTemp As New ADODB.Recordset
    
    If intinsure <> 82 And intinsure <> 83 Then
        Exit Function
    End If
    On Error GoTo ErrName
    gstrSQL = "Select 参数值 From 保险参数 where 参数名='门诊大病使用住院比例' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参数", intinsure)
    
    If rsTemp.RecordCount <= 0 Then
        MsgBox "系统参数添加不完整，请与系统管理员联系！"
        Exit Function
    End If
    If rsTemp!参数值 = 1 Then
        IsParaBig = True
    Else
        IsParaBig = False
    End If
    Exit Function

ErrName:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function IsParaQ(intinsure) As Boolean
    '检查当前医保是否允许门诊企离走住院比例
    Dim rsTemp As New ADODB.Recordset
    
    If intinsure <> 82 Then
        Exit Function
    End If
    On Error GoTo ErrName
    gstrSQL = "Select 参数值 From 保险参数 where 参数名='门诊企离使用住院比例' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参数", intinsure)
    
    If rsTemp.RecordCount <= 0 Then
        MsgBox "系统参数添加不完整，请与系统管理员联系！"
        Exit Function
    End If
    If rsTemp!参数值 = 1 Then
        IsParaQ = True
    Else
        IsParaQ = False
    End If
    Exit Function

ErrName:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function






