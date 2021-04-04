Attribute VB_Name = "mdl北京"
Option Explicit
'----------------------------------------------------------------
'诊疗目录中，以W01打头的，都是服务设施项目，在中间库中，类别=2
'----------------------------------------------------------------
'交易流水号：每次结算都要新产生，规则：医院编码（8位）+就诊时间(YMDHms)+结算时间(YMDHms)
'需要提供小工具：
'1、由于结算时诊断不一定出来了，因此该工具用来在上传前，修改出院诊断信息，引自何学彬：但北京情况不太一样，因为确实病案室在短时间内出诊断比较困难，如果病人等太长时间，医院肯定会遭到投诉的。
'2、允许增加、删除、修改指标体系
'3、修改程序时注意，大部分模块内定义有常量

Private mblnInit As Boolean                 '医保初始化成功标志
Public gcnBJYB As New ADODB.Connection

Private Type ComInfo_北京
    医院编码    As String
    入参目录    As String
    出参目录    As String
    业务类型    As String
    交易流水号  As String
    卡号        As String   '卡号/手册号(手册号以S结尾)
End Type
Public gComInfo_北京 As ComInfo_北京

Private Enum 接口功能
    启动政策机服务              'StartPolicy
    停止政策机服务              'StopPolicy
    获取持卡病人信息            'GetPersonCommInfo
    获取持卡病人待遇信息        'Get_SumInfo
    获取手册病人待遇信息        'Get_SumInfo2
    获取特殊病信息              'Get_SpecInfo
    帐户支付                    'Reg_Account    仅持卡病人存在扣减帐户
    交易确认                    'Reg            仅持卡病人需要执行此函数
    接口版本信息                'Get_Ver        获取接口的版本信息
    获取错误信息                'Get_ErrInfo    获取错误信息
    '费用分解函数
    费用分解_通用1              'Divide
    费用分解_通用2              'Divide2
    费用分解_普通门诊           'Poli_Divide
    费用分解_特殊门诊1          'Spec_Divide
    费用分解_特殊门诊2          'Spec_Divide2
    费用分解_家庭病床1          'Home_Divide
    费用分解_家庭病床2          'Home_Divide2
    费用分解_住院1              'Hosp_Divide
    费用分解_住院2              'Hosp_Divide2
    费用分解_住院3              'Hosp_Divide3
End Enum

Private Enum 交易待遇
    卡号 = 0
    费用年度
    本年门诊医保内
    本年门诊大额支付
    本年统筹支付
    本年大额支付
    本年次数累计
    本周期起始日期
    本周期交易号
    本周期医保内
    本周期统筹支付
    本周期大额支付
    住院家庭病床年度
    本次住院家庭病床标识
    本次住院家庭病床交易号
    本次住院家庭病床起始日期
    本次住院家庭病床医保内
    本次住院家庭病床统筹支付
    本次住院大额支付
End Enum

Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度
Private Const mstrSplit As String = "|"

'接口申明
'--基本函数说明
Public Declare Function BJ_StartPolicy Lib "FYFJ.dll" Alias "StartPolicy" (ByVal ShowProgress As Long) As Long
Public Declare Function BJ_StopPolicy Lib "FYFJ.dll" Alias "StopPolicy" () As Long
Public Declare Function BJ_GetPersonCommInfo Lib "FYFJ.dll" Alias "GetPersonCommInfo" (ByVal strPersonalInfo As String) As Long
Public Declare Function BJ_Get_SumInfo Lib "FYFJ.dll" Alias "Get_SumInfo" (ByVal strSumInfo As String) As Long
Public Declare Function BJ_Get_SumInfo2 Lib "FYFJ.dll" Alias "Get_SumInfo2" (ByVal strPersonalInfo As String, ByVal strSumInfo As String) As Long
Public Declare Function BJ_Get_SpecInfo Lib "FYFJ.dll" Alias "Get_SpecInfo" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Divide Lib "FYFJ.dll" Alias "Divide" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Divide2 Lib "FYFJ.dll" Alias "Divide2" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Reg_Account Lib "FYFJ.dll" Alias "Reg_Account" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Get_Ver Lib "FYFJ.dll" Alias "Get_Ver" (ByVal strDllVer As String, ByVal strDateVer As String) As Long
Public Declare Function BJ_Get_ErrInfo Lib "FYFJ.dll" Alias "Get_ErrInfo" (ByVal strErrMsg As String) As Long
Public Declare Function BJ_Reg Lib "FYFJ.dll" Alias "Reg" (ByVal strIn As String, ByVal strOut As String) As Long
'--以下是费用分解函数
Public Declare Function BJ_Poli_Divide Lib "FYFJ.dll" Alias "Poli_Divide" (ByVal strIn As String) As Long
Public Declare Function BJ_Spec_Divide Lib "FYFJ.dll" Alias "Spec_Divide" (ByVal strIn As String) As Long
Public Declare Function BJ_Spec_Divide2 Lib "FYFJ.dll" Alias "Spec_Divide2" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Home_Divide Lib "FYFJ.dll" Alias "Home_Divide" (ByVal strIn As String) As Long
Public Declare Function BJ_Home_Divide2 Lib "FYFJ.dll" Alias "Home_Divide2" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Hosp_Divide Lib "FYFJ.dll" Alias "Hosp_Divide" (ByVal strIn As String) As Long
Public Declare Function BJ_Hosp_Divide2 Lib "FYFJ.dll" Alias "Hosp_Divide2" (ByVal strIn As String) As Long
Public Declare Function BJ_Hosp_Divide3 Lib "FYFJ.dll" Alias "Hosp_Divide3" (ByVal strIn As String) As Long


'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'以下是功能函数
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
Private Function GetDeal(ByVal lng病人ID As Long, ByVal str日期 As String) As ADODB.Recordset
    '----------------------------------------------------------------
    '功能描述   ：用于就诊登记、结算、结算冲销时，获取指定日期以前的待遇信息
    '编写人     ：朱玉宝
    '编写日期   ：2004-06-28
    '----------------------------------------------------------------
    On Error GoTo errHand
    Dim rsDeal As New ADODB.Recordset
    '获取指定日期以前的待遇信息（从手册消费记录中提取，结算后，要产生一条手册消费记录）
    Call DebugTool("获取" & str日期 & "以前的待遇信息")
    Call WriteBusinessLOG("获取" & str日期 & "以前的待遇信息", "", "")
    gstrSQL = " Select A.医疗类别,to_Char(A.入院日期,'yyyyMMdd'),to_Char(A.出院日期,'yyyyMMdd'),A.入院类型,A.出院类型," & _
              "        A.费用总额,A.统筹支付,A.大额支付,A.个人自付,A.个人自费,A.统筹封顶后医保内" & _
              " From 手册消费记录 A,保险帐户 B" & _
              " Where A.卡号=B.卡号 And B.病人ID=" & lng病人ID & " And A.入院日期<=TO_DATE('" & str日期 & "','yyyy-MM-dd')"
    If rsDeal.State = 1 Then rsDeal.Close
    Call SQLTest(App.Title, "ZL9INSURE\GETDEAL", gstrSQL): rsDeal.Open gstrSQL, gcnBJYB: Call SQLTest
    Call DebugTool("待遇信息记录条数为：" & rsDeal.RecordCount)
    Call WriteBusinessLOG("待遇信息记录条数为：" & rsDeal.RecordCount, "", "")
    Set GetDeal = rsDeal
    Exit Function
errHand:
    Call DebugTool("获取待遇信息时发生错误！" & vbCrLf & "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("获取待遇信息时发生错误！" & vbCrLf & "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
End Function

Private Function MakeFile_Center(ByVal rsData As ADODB.Recordset, ByVal int功能 As 接口功能) As Boolean
    Dim strFile_In As String
    '----------------------------------------------------------------
    '功能描述   ：（主控模块）根据传入的记录集数据，产生相应的文件
    '编写人     ：朱玉宝
    '编写日期   ：2004-07-03
    '补充说明   ：支持的接口：获取手册病人待遇信息、费用分解_普通门诊、费用分解_特殊门诊1、费用分解_家庭病床1、费用分解_住院1、、、、
    '  入参文件 ：目前只有一个入参文件，存于strFile_In
    '  出参文件 ：可能有多个，以"||"分隔，存于strFile_Out
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    Select Case int功能
    Case 接口功能.获取手册病人待遇信息
        strFile_In = "med_note.in"
        '----传入文件说明
'        序号    数据项  类型    最大长度    说明
'        1      医疗类别    C   2   参见标准AKA130
'        2      入院日期    D   8   住院为手册中入院日期（分段起始日期） 普通门急诊、门诊特殊病、家庭病床为手册中就诊日期
'        3      出院日期    D   8   住院为手册中出院日期（分段结束日期） 普通门急诊、门诊特殊病、家庭病床为手册中就诊日期
'        4      入院类型    C   2   住院为：0-普通住院，1-特殊病病人住院，2-器官移植住院，3-精神病住院，4-中医医院针灸科住院 普通门急诊、门诊特殊病、家庭病床为空
'        5      出院类型    C   1   0-正常；1-转出院、普通门急诊、门诊特殊病、家庭病床为空；
'        6      费用总金额  N   8,2 手册对应项录入
'        7      统筹支付金额    N   8,2 手册对应项录入
'        8      大额（公务员补助）支付金额  N   8,2 非公务员为大额费用，公务员为公务员补助金额
'        9      个人自付金额    N   8,2 手册对应项录入
'        10     个人自费金额    N   8,2 手册对应项录入
'        11     统筹封顶后医保内金额    N   8,2 非公务员为0，公务员按手册对应项录入
    Case 接口功能.费用分解_普通门诊
        strFile_In = "Poli_Divide.in"
        '----传入文件说明
'        序号    数据项  类型    最大长度    说明
'        1      文件记录条数    C   4   文本文件包含的记录的数量
'        2      交易流水号  C   20  医疗机构代码（8左对齐）+医院流水号（12右对齐）中间补零。由医院端产生
'        3      医疗类别    C   2   AKA130
'        4      特殊标识    C   2   0-普通，2-器官移植
'        5      本次交易费用总金额  N   8,2
'        第二行起为费用明细信息，具体数据项如下列表：
'        序号    数据项  类型    最大长度    说明
'        1      项目序号    C   9   顺序号
'        2      处方号  C   20  参见标准AKC220，可空
'        3      项目代码    C   20  药品、诊疗项目或服务设施的一保编码
'        4      项目名称    C   100 本医院项目名称
'        5      项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'        6      单价    N   10,4    AKC225
'        7      数量    N   8,2 AKC226
'        8      费用总金额  N   10,4    实际结算金额
'        9      费用发生日期    D   8   费用发生日期
    Case 接口功能.费用分解_特殊门诊1
        strFile_In = "Spec_Divide.in"
        '----传入文件说明
'        序号    数据项  类型    最大长度    说明
'        1   项目序号    C   9   顺序号
'        2   处方号  C   20  参见标准AKC220，可空
'        3   项目代码    C   20  药品、诊疗项目或服务设施编码
'        4   项目名称    C   100 本医院项目名称
'        5   项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'        6   单价    N   10,4    AKC225
'        7   数量    N   8,2 AKC226
'        8   费用总金额  N   10,4    实际结算金额
'        9   费用发生日期    D   8   YYYYMMDD
    Case 接口功能.费用分解_家庭病床1
        strFile_In = "Home_Divide.in"
        '----传入文件说明
'        序号    数据项  类型    最大长度    说明
'        1   项目序号    C   9   顺序号
'        2   处方号  C   20  参见标准AKC220，可空
'        3   项目代码    C   20  药品、诊疗项目或服务设施编码
'        4   项目名称    C   100 本医院项目名称
'        5   项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'        6   单价    N   10,4    AKC225
'        7   数量    N   8,2 AKC226
'        8   费用总金额  N   10,4    实际结算金额
'        9   费用发生日期    D   8   YYYYMMDD
    Case 接口功能.费用分解_住院1
        strFile_In = "Hosp_Divide.in"
        '----传入文件说明
'        序号    数据项  类型    最大长度    说明
'        1   项目序号    C   9   顺序号
'        2   医嘱号  C   20  （可空）
'        3   项目代码    C   20  药品、诊疗项目或服务设施编码
'        4   项目名称    C   100 本医院项目名称
'        5   项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'        6   单价    N   10,4    AKC225
'        7   数量    N   8,2 AKC226
'        8   费用总金额  N   10,4    实际结算金额
'        9   费用发生日期    D   8   YYYYMMDD
    
    Case Else
        Exit Function       '不支持该功能函数
    End Select
    
    MakeFile_Center = MakeFile(rsData, strFile_In)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MakeFile(ByVal rsData As ADODB.Recordset, ByVal strFile As String) As Boolean
    Dim lng记录数 As Long
    Dim dbl费用总额 As Double
    Dim strLine As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim lngCol As Long, lngCols As Long
    
    On Error GoTo errHand
    '----------------------------------------------------------------
    '功能描述   ：根据传入的记录集数据，产生相应的文件
    '编写人     ：朱玉宝
    '编写日期   ：2004-07-03
    '补充说明   ：支持的接口：获取手册病人待遇信息、费用分解_普通门诊、费用分解_特殊门诊1、费用分解_家庭病床1、费用分解_住院1、、、、
    '  入参文件 ：目前只有一个入参文件，存于strFile_In
    '  出参文件 ：可能有多个，以"||"分隔，存于strFile_Out
    '----------------------------------------------------------------
    
    '如果文件存在，删除
    
    Call DebugTool("如果文件存在就删除后重新创建(zl9INSURE\MakeFile)" & vbCrLf & _
        "入参strFile=" & strFile)
    Call WriteBusinessLOG("如果文件存在就删除后重新创建(zl9INSURE\MakeFile)" & vbCrLf & _
        "入参strFile=" & strFile, "", "")
    strFile = gComInfo_北京.入参目录 & "\" & strFile
    If objFileSystem.FileExists(strFile) Then Call objFileSystem.DeleteFile(strFile, True)
    Set objStream = objFileSystem.CreateTextFile(strFile)
    
    '此处为需要单独处理的部分（目前只有普通门诊）
    If strFile = "Poli_Divide.in" Then
        '----传入文件说明
'        序号    数据项  类型    最大长度    说明
'        1      文件记录条数    C   4   文本文件包含的记录的数量
'        2      交易流水号  C   20  医疗机构代码（8左对齐）+医院流水号（12右对齐）中间补零。由医院端产生
'        3      医疗类别    C   2   AKA130
'        4      特殊标识    C   2   0-普通，2-器官移植
'        5      本次交易费用总金额  N   8,2
        With rsData
            lng记录数 = .RecordCount
            Do While Not .EOF
                dbl费用总额 = dbl费用总额 + Nvl(!实收金额)
                .MoveNext
            Loop
        End With
        dbl费用总额 = Format(dbl费用总额, "#####0.00;-#####0.00;0;")
        strLine = lng记录数 & mstrSplit & gComInfo_北京.交易流水号 & mstrSplit & _
            gComInfo_北京.业务类型 & mstrSplit & "0" & mstrSplit & dbl费用总额
        Call objStream.WriteLine(strLine)
    End If
    
    '统一处理：按传入记录集产生文件
    With rsData
        lngCols = .Fields.Count
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strLine = ""
            For lngCol = 0 To lngCols - 1
                strLine = strLine & mstrSplit & Nvl(.Fields(lngCol).Value)
            Next
            strLine = Mid(strLine, Len(mstrSplit) + 1)
            objStream.WriteLine (strLine)
            .MoveNext
        Loop
    End With
    
    objStream.Close
    Set objStream = Nothing
    MakeFile = True
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\MakeFile)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\MakeFile)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
    If ErrCenter = 1 Then Resume
End Function

Private Function AnalyFile_Center(ByVal rsData As ADODB.Recordset, ByVal int功能 As 接口功能, _
    Optional ByVal bln预算 As Boolean = True) As Boolean
    Dim strReturn As String
    '----------------------------------------------------------------
    '功能描述   ：（主控模块）分析接口函数的输出文件，产生为内部记录集，并将数据更新到数据库中
    '编写人     ：朱玉宝
    '编写日期   ：2004-07-03
    '补充说明   ：支持的接口：费用分解_普通门诊、费用分解_特殊门诊1、费用分解_家庭病床1、费用分解_住院1
    '  入参文件 ：目前只有一个入参文件，存于strFile_In
    '  出参文件 ：可能有多个，以"||"分隔，存于strFile_Out
    '  当bln预算为真时，不需要保存分解结果，也可以说，费用明细分析这一步都可以不做了
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    Select Case int功能
    Case 接口功能.费用分解_普通门诊
        strReturn = AnalyFile_普通门诊(bln预算)
    Case 接口功能.费用分解_特殊门诊1
        strReturn = AnalyFile_特殊门诊1(bln预算)
    Case 接口功能.费用分解_家庭病床1
        strReturn = AnalyFile_家庭病床1(bln预算)
    Case 接口功能.费用分解_住院1
        strReturn = AnalyFile_住院1(bln预算)
    Case Else
        Exit Function       '不支持该功能函数
    End Select
    
    AnalyFile_Center = (strReturn <> "")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyFile_普通门诊(Optional ByVal bln预算 As Boolean = True) As String
    '返回汇总行数据
    Dim strTotal As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Const strFile_Out As String = "Poli_Divide.out"
    On Error GoTo errHand
'    ----传出文件说明
'    文件第一行为总的分解结果，具体数据项如下列表：
'    序号    数据项  类型    最大长度    说明
'    1      文件记录条数    C   4   文本文件包含的记录的数量
'    2      交易流水号  C   20  医疗机构代码（8左对齐）+医院流水号（12右对齐）中间补零。由医院端产生
'    3      医疗类别    C   2   参见标准AKA130
'    4      本次交易费用总金额  N   8,2
'    5      医保内总费用    N   8,2
'    6      医保外总费用    N   8,2
'    7      大额/公务员支付金额 N   8,2
'    8      大额/公务员自付金额 N   8,2
'    9      个人应付总金额  N   8,2
'    第二行起为费用明细信息的分解结果，格式为：
'    序号    数据项  类型    最大长度    说明
'    1      项目序号    C   9   顺序号
'    2      处方号  C   20  参见标准AKC220，可空
'    3      项目代码    C   20  药品、诊疗项目或服务设施编码
'    4      项目名称    C   100 本医院项目名称
'    5      项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'    6      单价    N   10,4
'    7      数量    N   8,2
'    8      费用总金额  N   10,4
'    9      费用发生日期    D   8   费用发生日期
'    10     医保内费用  N   10,4
'    11     医保外费用  N   10,4
'    12     分解状态    C   1   0-正常，1-不符合特殊标识，2-医保目录内不存在，3-对照错误
    
    objStream.Close
    Set objStream = Nothing
    AnalyFile_普通门诊 = strTotal
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\MakeFile)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\MakeFile)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
End Function

Private Function AnalyFile_特殊门诊1(Optional ByVal bln预算 As Boolean = True, _
    Optional ByVal lng结帐ID As Long = 0, Optional ByVal bln住院 As Boolean = False) As String        '返回汇总行数据
    Dim lngRow As Long, lngRows As Long     '当前行及总行数
    Dim strTotal As String
    Dim arrRow
    Dim strNO As String                     '门诊单据号
    Dim str本周期交易号 As String           '本周期交易号（待遇信息中获取）
    Dim str医生 As String
    Dim str处方日期 As String               '费用发生时间
    Dim rsTemp As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset     '医保项目信息
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Const strFile_Out As String = "Spec_Divide.out"
    Dim dbl现金 As Double, dbl医保基金 As Double, dbl大额补助 As Double, dbl费用总额 As Double
    '以下是文件字段常量（汇总行）
    Const col记录数 As Integer = 0
    Const col交易流水号 As Integer = 1
    Const col费用总额 As Integer = 2
    Const col普通门诊医保内 As Integer = 3
    Const col普通门诊医保外 As Integer = 4
    Const col统筹支付 As Integer = 5
    Const col统筹自付 As Integer = 6
    Const col大额支付 As Integer = 7
    Const col大额自付 As Integer = 8
    Const col特殊病自付 As Integer = 9
    Const col特殊病医保外 As Integer = 10
    Const col首先自付 As Integer = 11
    Const col统筹封顶后医保内 As Integer = 12
    '以下是文件字段常量（费用明细）
    Const col项目序号 As Integer = 0
    Const col处方号 As Integer = 1
    Const col项目代码 As Integer = 2
    Const col项目名称 As Integer = 3
    Const col项目类别 As Integer = 4
    Const col单价 As Integer = 5
    Const col数量 As Integer = 6
    Const col明细总额 As Integer = 7
    Const col费用发生日期 As Integer = 8
    Const col医保内 As Integer = 9
    Const col医保外 As Integer = 10
    Const col明细_首先自付 As Integer = 11
    Const col分解状态 As Integer = 12
    
    On Error GoTo errHand
'    ----传出文件说明
'    该函数有一个隐含的传出参数，文件Spec_Divide.out，文件的内容为享受特殊病待遇的项目的费用分解，该文件放于HIS客户端和本DLL的交换目录SWAP_PATH（在环境变量中定义）。文件分两行，第一行为总的分解结果，具体数据项如下列表：
'    序号    数据项  类型    最大长度    说明
'    1   文件记录条数    C   4   文本文件包含的记录的数量
'    2   交易流水号  C   20  医疗机构代码（8左对齐）+医院流水号（12右对齐）中间补零。由医院端产生
'    3   本次交易费用总金额  N   8,2
'    4   普通门诊医保内费用  N   8,2
'    5   普通门诊医保外费用  N   8,2
'    6   统筹(特殊病)支付金额    N   8,2
'    7   统筹（特殊病）自付金额  N   8,2
'    8   大额/公务员(特殊病)支付金额 N   8,2
'    9   大额/公务员(特殊病)自付金额 N   8,2
'    10  特殊病个人自付金额  N   8,2 特殊病用药及诊疗医保内个人自付金额
'    11  特殊病医保外金额    N   8,2 特殊病用药及诊疗医保外金额
'    12  个人自付二金额  N   8,2 乙类项目按比例个人负担部分
'    13  本次交易统筹封顶后医保内金额    N   8,2
'    第二行起为费用明细信息的分解结果，具体数据项如下列表：
'    序号    数据项  类型    最大长度    说明
'    1   项目序号    C   9   顺序号
'    2   处方号  C   20  参见标准AKC220，可空
'    3   项目代码    C   20  药品、诊疗项目或服务设施编码
'    4   项目名称    C   100 本医院项目名称
'    5   项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'    6   单价    N   10,4
'    7   数量    N   8,2
'    8   费用总金额  N   10,4    实际结算金额
'    9   费用发生日期    D   8   YYYYMMDD
'    10  医保内费用  N   10,4
'    11  医保外费用  N   10,4
'    12  个人自付二金额  N   8,2 乙类项目按比例个人负担部分
'    13  分解状态    C   1   0-正常，1-不符合特殊标识，2-医保目录内不存在，3-对照错误
    
    '如果是预算，仅分析汇总行数据，且直接返回
    Call DebugTool("进入(zl9Insure\AnalyFile_特殊门诊1)")
    Call WriteBusinessLOG("进入(zl9Insure\AnalyFile_特殊门诊1)", "", "")
    If Not objFileSystem.FileExists(gComInfo_北京.出参目录 & "\" & strFile_Out) Then Exit Function
    Set objStream = objFileSystem.OpenTextFile(gComInfo_北京.出参目录 & "\" & strFile_Out)
    
    Call DebugTool("读取汇总行数据(zl9Insure\AnalyFile_特殊门诊1)")
    Call WriteBusinessLOG("读取汇总行数据(zl9Insure\AnalyFile_特殊门诊1)", "", "")
    '可能每行文本是以换行符结束，而VB认的是回车换行，将所有数据都读出来了，需要SPLIT
    strTotal = objStream.ReadLine
    arrRow = Split(strTotal, vbCr)
    objStream.Close
    Set objStream = Nothing
    
    lngRows = UBound(arrRow)
    '预结算仅需要汇总行数据（各种结算方式支付额）
    If bln预算 Then
        '对费用分解函数返回的费用明细进行检查，如果存在不正常的分解状态，提示并退出
        Call DebugTool("对费用分解函数返回的费用明细的分解状态进行检查(zl9Insure\AnalyFile_特殊门诊1)")
        Call WriteBusinessLOG("对费用分解函数返回的费用明细的分解状态进行检查(zl9Insure\AnalyFile_特殊门诊1)", "", "")
        For lngRow = 1 To lngRows
            If CheckDetail(Split(arrRow(lngRow), mstrSplit)(col项目名称), Val(Split(arrRow(lngRow), mstrSplit)(col分解状态))) Then Exit Function
        Next
        
        AnalyFile_特殊门诊1 = arrRow(0)
        Exit Function
    End If
    
    str处方日期 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    '提取收费单据号
    Call DebugTool("提取收费单据号(zl9Insure\AnalyFile_特殊门诊1")
    Call WriteBusinessLOG("提取收费单据号(zl9Insure\AnalyFile_特殊门诊1", "", "")
    gstrSQL = "" & _
        " SELECT 实际票号,开单人" & _
        " From 门诊费用记录" & _
        " WHERE 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取收费单据号", lng结帐ID)
    strNO = Nvl(rsTemp!实际票号)
    str医生 = Nvl(rsTemp!开单人)
    
    If bln住院 Then
        '提取收费单据号
        Call DebugTool("从结帐单中取发票号(zl9Insure\AnalyFile_特殊门诊1")
        Call WriteBusinessLOG("从结帐单中取发票号(zl9Insure\AnalyFile_特殊门诊1", "", "")
        gstrSQL = "" & _
            " SELECT 实际票号" & _
            " From 病人结帐记录" & _
            " WHERE ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "从结帐单中取发票号", lng结帐ID)
        strNO = Nvl(rsTemp!实际票号)
    End If
    
    '提取本周期交易号
    Call DebugTool("从交易待遇信息中提取本周期交易号(zl9Insure\AnalyFile_特殊门诊1")
    Call WriteBusinessLOG("从交易待遇信息中提取本周期交易号(zl9Insure\AnalyFile_特殊门诊1", "", "")
    gstrSQL = "" & _
        " SELECT 本周期交易号" & _
        " From 交易待遇信息" & _
        " WHERE 交易流水号='" & gComInfo_北京.交易流水号 & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\AnalyFile_特殊门诊1", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    str本周期交易号 = Nvl(rsTemp!本周期交易号)
    
    '提取参保人基本信息（只能在前面使用rsTemp记录集，因为后面直接使用了该记录集的数据）
    Call DebugTool("提取参保人基本信息(zl9Insure\AnalyFile_特殊门诊1")
    Call WriteBusinessLOG("提取参保人基本信息(zl9Insure\AnalyFile_特殊门诊1", "", "")
    gstrSQL = "" & _
        " SELECT 卡号,社保证号,缴费地区代码,业务类型,参保类别,公务员,公务员待遇,病种标识,特殊病截止日期" & _
        " From 保险帐户" & _
        " WHERE 卡号='" & gComInfo_北京.卡号 & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\AnalyFile_特殊门诊1", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    
    '正式结算，需要保存汇总行数据（门诊交易信息）及费用明细数据（门诊费用明细）
    '处理汇总行，参数如下
    '交易流水号,缴费地区代码,社保证号,卡号,终端机编号,收费单据号,业务类型,特殊标识,本周期交易号,交易时间
    '操作员代码(可空),验证通过方式(可空),费用总额,普通门诊医保内,普通门诊医保外,统筹支付,统筹自付,大额支付
    '大额自付,特殊病自付,特殊病医保外,首先自付,个人帐户支付,现金支付,个人帐户消费后余额
    '特殊情况标识(可空),MAC1(可空),上传
    dbl费用总额 = Val(Split(arrRow(0), mstrSplit)(col费用总额))
    dbl医保基金 = Val(Split(arrRow(0), mstrSplit)(col统筹支付))
    dbl大额补助 = Val(Split(arrRow(0), mstrSplit)(col大额支付))
    dbl现金 = dbl费用总额 - dbl大额补助 - dbl医保基金
    gstrSQL = "ZL_门诊交易信息_INSERT(" & _
        "'" & gComInfo_北京.交易流水号 & "','" & rsTemp!缴费地区代码 & "','" & rsTemp!社保证号 & "'," & _
        "'" & rsTemp!卡号 & "',NULL,'" & strNO & "','" & rsTemp!业务类型 & "','0','" & str本周期交易号 & "'," & _
        "To_Date('" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & _
        "NULL,NULL," & Val(Split(arrRow(0), mstrSplit)(col费用总额)) & "," & Val(Split(arrRow(0), mstrSplit)(col普通门诊医保内)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col普通门诊医保外)) & "," & Val(Split(arrRow(0), mstrSplit)(col统筹支付)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col统筹自付)) & "," & Val(Split(arrRow(0), mstrSplit)(col大额支付)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col大额自付)) & "," & Val(Split(arrRow(0), mstrSplit)(col特殊病自付)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col特殊病医保外)) & "," & Val(Split(arrRow(0), mstrSplit)(col首先自付)) & "," & _
        "0," & dbl现金 & ",0,NULL,NULL,0)"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    '处理明细（费用明细），参数如下
    '交易流水号,项目序号,卡号,处方号,医保编码,项目名称,项目类别,剂型,规格,单价,数量
    '处方日期,费用总额,收费类别,医师姓名,医保内费用,医保外费用,首先自付,分解状态,上传
    For lngRow = 1 To lngRows
        '需要取出该医保项目的剂型、规格及收费类别，用于上传（HIS中多个项目对应一个医保项目时，可能返回多条记录）
        gstrSQL = "" & _
            " Select A.规格,C.收费类别,C.剂型" & _
            " From 收费细目 A,保险支付项目 B," & GetUser & ".药品目录 C,门诊费用记录 F" & _
            " Where A.ID=F.收费细目ID ANd A.ID=B.收费细目ID And B.项目编码=C.编码 " & _
            " And F.结帐ID=" & lng结帐ID & " AND F.序号=" & Split(arrRow(lngRow), mstrSplit)(col项目序号) & "" & _
            " Union" & _
            " Select A.规格,C.收费类别,'' AS 剂型" & _
            " From 收费细目 A,保险支付项目 B," & GetUser & ".诊疗目录 C,门诊费用记录 F" & _
            " Where A.ID=F.收费细目ID ANd A.ID=B.收费细目ID And B.项目编码=C.编码 " & _
            " And F.结帐ID=[1] AND F.序号=[2]" & _
            " And A.类别 Not In ('5','6','7')"
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "获取项目信息", lng结帐ID, CLng(Split(arrRow(lngRow), mstrSplit)(col项目序号)))
        
        '如果为空，说明是非医保项目，传固定的值
        If rsDetail.RecordCount = 0 Then
            gstrSQL = " Select A.规格,A.标识码 As 收费类别,'' AS 剂型 " & _
                      " From 药品目录 A,门诊费用记录 F" & _
                      " Where A.药品ID=F.收费细目ID And F.结帐ID=" & lng结帐ID & " AND F.序号=[2]"
            gstrSQL = gstrSQL & " UNION " & _
                      " Select A.规格,A.标识主码 As 收费类别,'' AS 剂型 " & _
                      " From 收费细目 A,门诊费用记录 F" & _
                      " Where A.ID=F.收费细目ID And F.结帐ID=[1] AND F.序号=[2]" & _
                      " And A.类别 Not In ('5','6','7')"
            Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "获取非医保项目的信息", lng结帐ID, CLng(Split(arrRow(lngRow), mstrSplit)(col项目序号)))
        End If
        
        gstrSQL = "ZL_门诊费用明细_INSERT(" & _
            "'" & gComInfo_北京.交易流水号 & "','" & Val(Split(arrRow(lngRow), mstrSplit)(col项目序号)) & "'," & _
            "'" & gComInfo_北京.卡号 & "','" & strNO & "','" & Split(arrRow(lngRow), mstrSplit)(col项目代码) & "'," & _
            "'" & Split(arrRow(lngRow), mstrSplit)(col项目名称) & "','" & Split(arrRow(lngRow), mstrSplit)(col项目类别) & "'," & _
            "'" & Nvl(rsDetail!剂型) & "','" & ToVarchar(Nvl(rsDetail!规格), 40) & "'," & Val(Split(arrRow(lngRow), mstrSplit)(col单价)) & "," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col数量)) & ",TO_Date('" & str处方日期 & "','yyyy-MM-dd hh24:mi:ss')," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col明细总额)) & ",'" & Nvl(rsDetail!收费类别) & "','" & str医生 & "'," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col医保内)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col医保外)) & "," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col明细_首先自付)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col分解状态)) & ",0" & _
            ")"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    Next
    
    AnalyFile_特殊门诊1 = arrRow(0)
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\MakeFile)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\MakeFile)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
End Function

Private Function AnalyFile_家庭病床1(Optional ByVal bln预算 As Boolean = True) As String
    '返回汇总行数据
    Dim strTotal As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Const strFile_Out As String = "Home_Divide.out"
    On Error GoTo errHand
'    ----传出文件说明
'    该函数有一个隐含的传出参数，文件Home_Divide.out，包含费用明细信息，该文件放于HIS客户端和本DLL的交换目录SWAP_PATH（在环境变量中定义）。文件分两行，第一行为总的分解结果，具体数据项如下列表：
'    序号    数据项  类型    最大长度    说明
'    1   文件记录条数    C   4
'    2   交易流水号  C   20  医疗机构代码（8左对齐）+医院流水号（12右对齐）中间补零。由医院端产生
'    3   本次交易费用总金额  N   8,2
'    4   本次交易医保内费用总金额    N   8,2
'    5   统筹支付金额    N   8,2
'    6   统筹自付金额    N   8,2
'    7   大额/公务员支付金额 N   8,2
'    8   大额/公务员自付金额 N   8,2
'    9   个人应付总金额  N   8,2
'    10  个人自付二金额  N   8,2 乙类项目按比例个人负担部分
'    11  本次交易统筹封顶后医保内金额    N   8,2
'    第二行起为费用明细信息分解结果，格式为：
'    序号    数据项  类型    最大长度    说明
'    1   项目序号    C   9   顺序号
'    2   处方号  C   20  参见标准AKC220，可空
'    3   项目代码    C   20  药品、诊疗项目或服务设施编码
'    4   项目名称    C   100 本医院项目名称
'    5   项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'    6   单价    N   10,4
'    7   数量    N   8,2
'    8   费用总金额  N   10,4    实际结算金额
'    9   费用发生日期    D   8   YYYYMMDD
'    10  医保内费用  N   10,4
'    11  医保外费用  N   10,4
'    12  个人自付二金额  N   8,2 乙类项目按比例个人负担部分
'    13  分解状态    C   1   0-正常，1-不符合特殊标识，2-医保目录内不存在，3-对照错误
    
    objStream.Close
    Set objStream = Nothing
    AnalyFile_家庭病床1 = strTotal
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\AnalyFile_家庭病床)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\AnalyFile_家庭病床)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
End Function

Private Function AnalyFile_住院1(Optional ByVal bln预算 As Boolean = True, Optional ByVal lng结帐ID As Long = 0) As String
    '返回汇总行数据
    Dim strTotal As String
    Dim strUser As String
    Dim strNO As String
    Dim str出院类型 As String
    Dim str开始日期 As String, str分段开始日期 As String, str分段结束日期 As String
    Dim arrRow
    Dim lngRow As Long, lngRows As Long
    Dim lng病人ID As Long, lng主页ID As Long
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim rs费用发生日期 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset
    Dim dbl费用总额 As Double, dbl医保基金 As Double, dbl大额补助 As Double, dbl现金 As Double
    Const strFile_Out As String = "Hosp_Divide.out"
    
    '字段常量说明
    Const col汇总_费用分段数 As Integer = 0
    Const col汇总_费用明细数 As Integer = 1
    Const col汇总_交易流水号 As Integer = 2
    Const col汇总_费用总额 As Integer = 3
    Const col汇总_医保内 As Integer = 4
    Const col汇总_统筹支付 As Integer = 5
    Const col汇总_统筹自付 As Integer = 6
    Const col汇总_大额支付 As Integer = 7
    Const col汇总_大额自付 As Integer = 8
    Const col汇总_个人应付 As Integer = 9
    Const col汇总_首先自付 As Integer = 10
    Const col汇总_统筹封顶后医保内 As Integer = 11
    Const col分段_周期序号 As Integer = 0
    Const col分段_费用年度 As Integer = 1
    Const col分段_起始日期 As Integer = 2
    Const col分段_截止日期 As Integer = 3
    Const col分段_费用总额 As Integer = 4
    Const col分段_医保内 As Integer = 5
    Const col分段_统筹支付 As Integer = 6
    Const col分段_统筹自付 As Integer = 7
    Const col分段_大额支付 As Integer = 8
    Const col分段_大额自付 As Integer = 9
    Const col分段_首先自付 As Integer = 10
    Const col分段_个人应付 As Integer = 11
    Const col明细_项目序号 As Integer = 0
    Const col明细_医嘱号 As Integer = 1
    Const col明细_项目代码 As Integer = 2
    Const col明细_项目名称 As Integer = 3
    Const col明细_项目类别 As Integer = 4
    Const col明细_单价 As Integer = 5
    Const col明细_数量 As Integer = 6
    Const col明细_费用总额 As Integer = 7
    Const col明细_费用发生日期 As Integer = 8
    Const col明细_医保内 As Integer = 9
    Const col明细_医保外 As Integer = 10
    Const col明细_首先自付 As Integer = 11
    Const col明细_分解状态 As Integer = 12
    On Error GoTo errHand
'    ----传出文件说明
'    该函数有一个隐含的传出参数，文件Hosp_Divide.out，包含费用分解结果。该文件放于HIS客户端和本DLL的交换目录SWAP_PATH（在环境变量中定义）。文件分为三个部分，第一部分为总的分解结果，长度为一行。具体数据项如下列表：
'    序号    数据项  类型    最大长度    说明
'    1   费用分段分解条数    N   2
'    2   费用明细记录条数    N   4   文本文件包含的明细记录的数量
'    3   交易流水号  C   20  医疗机构代码（8左对齐）+医院流水号（12右对齐）中间补零。限长20位。由医院端预产生，如果交易失败则作废
'    4   本次交易费用总金额  N   8,2
'    5   本次交易医保内总金额    N   8,2
'    6   统筹支付金额    N   8,2
'    7   统筹自付金额    N   8,2
'    8   大额/公务员支付金额 N   8,2
'    9   大额/公务员自付金额 N   8,2
'    10  个人应付总金额  N   8,2
'    11  个人自付二金额  N   8,2 乙类项目按比例个人负担部分
'    12  本次交易统筹封顶后医保内金额    N   8,2
'    第二部分为费用分段分解的分解结果，长度为第一行中的费用分段分解条数。具体数据项如下列表：
'    序号    数据项  类型    最大长度    说明
'    1   费用结算周期序号    N   2
'    2   费用年度    C   4
'    3   本段费用起始日期    D   8
'    4   本段费用截止日期    D   8
'    5   本段费用总金额  N   8,2
'    6   本段费用医保内总金额    N   8,2
'    7   本段费用统筹支付金额    N   8,2
'    8   本段费用统筹自付金额    N   8,2
'    9   本段费用大额/公务员支付金额 N   8,2
'    10  本段费用大额/公务员自付金额 N   8,2
'    11  本段个人自付二金额  N   8,2 乙类项目按比例个人负担部分
'    12  本段费用个人应付总金额  N   8,2
'    第三部分为费用明细信息的分解结果，长度为第一行中的费用明细记录条数。具体数据项如下列表：
'    序号    数据项  类型    最大长度    说明
'    1   项目序号    C   9   顺序号
'    2   医嘱号  C   20  可空
'    3   项目代码    C   20  药品、诊疗项目或服务设施编码
'    4   项目名称    C   100 本医院项目名称
'    5   项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'    6   单价    N   10,4
'    7   数量    N   8,2
'    8   费用总金额  N   10,4
'    9   费用发生日期    D   8   YYYYMMDD
'    10  医保内费用  N   10,4
'    11  医保外费用  N   10,4
'    12  个人自付二金额  N   8,2 乙类项目按比例个人负担部分
'    13  分解状态    C   1   0-正常，1-不符合特殊标识，2-医保目录内不存在，3-对照错误
    
    '如果是预算，仅分析汇总行数据，且直接返回
    Call DebugTool("进入(zl9Insure\AnalyFile_住院1)")
    Call WriteBusinessLOG("进入(zl9Insure\AnalyFile_住院1)", "", "")
    If Not objFileSystem.FileExists(gComInfo_北京.出参目录 & "\" & strFile_Out) Then Exit Function
    Set objStream = objFileSystem.OpenTextFile(gComInfo_北京.出参目录 & "\" & strFile_Out)
    
    Call DebugTool("读取汇总行数据(zl9Insure\AnalyFile_住院1)")
    Call WriteBusinessLOG("读取汇总行数据(zl9Insure\AnalyFile_住院1)", "", "")
    '可能每行文本是以换行符结束，而VB认的是回车换行，将所有数据都读出来了，需要SPLIT
    strTotal = objStream.ReadLine
    arrRow = Split(strTotal, vbCr)
    objStream.Close
    Set objStream = Nothing
    
    '预结算仅需要汇总行数据（各种结算方式支付额）
    If bln预算 Then
        '对费用分解函数返回的费用明细进行检查，如果存在不正常的分解状态，提示并退出
        Call DebugTool("对费用分解函数返回的费用明细的分解状态进行检查(zl9Insure\AnalyFile_住院1)")
        Call WriteBusinessLOG("对费用分解函数返回的费用明细的分解状态进行检查(zl9Insure\AnalyFile_住院1)", "", "")
        For lngRow = (1 + Val(Split(arrRow(0), "|")(col汇总_费用分段数))) To lngRows
            If CheckDetail(Split(arrRow(lngRow), mstrSplit)(col明细_项目名称), Val(Split(arrRow(lngRow), mstrSplit)(col明细_分解状态))) Then Exit Function
        Next
        
        AnalyFile_住院1 = arrRow(0)
        Exit Function
    End If
    
    strUser = GetUser
    '提取收费单据号
    Call DebugTool("提取收费单据号(zl9Insure\AnalyFile_特殊门诊1")
    Call WriteBusinessLOG("提取收费单据号(zl9Insure\AnalyFile_特殊门诊1", "", "")
    gstrSQL = "" & _
        " SELECT 病人ID,主页ID" & _
        " From 住院费用记录" & _
        " WHERE 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取收费单据号", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    lng主页ID = rsTemp!主页ID
    gstrSQL = "" & _
        " SELECT 实际票号" & _
        " From 病人结帐记录" & _
        " WHERE ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取收费单据号", lng结帐ID)
    strNO = Nvl(rsTemp!实际票号)
    
    '读取费用开始日期
    Call DebugTool("提取费用开始日期(zl9Insure\AnalyFile_住院1")
    Call WriteBusinessLOG("提取费用开始日期(zl9Insure\AnalyFile_住院1", "", "")
    gstrSQL = "Select min(登记时间) 开始日期 From 住院费用记录 Where 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取费用开始日期", lng结帐ID)
    str开始日期 = Format(rsTemp!开始日期, "yyyy-MM-dd HH:mm:ss")
    
    '读取出院方式
    Call DebugTool("提取出院日期及出院方式(zl9Insure\AnalyFile_住院1")
    Call WriteBusinessLOG("提取出院日期及出院方式(zl9Insure\AnalyFile_住院1", "", "")
    gstrSQL = "Select 出院日期,出院方式 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取出院日期及出院方式", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "未找到该病人的病案信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '出院类型:0-出院,1-转出院，2-中途结算
    If IsNull(rsTemp!出院日期) Then
        '肯定是中途结算
        str出院类型 = 2
    Else
        '判断出院方式
        str出院类型 = IIf(rsTemp!出院方式 = "转院", 1, 0)
    End If
    
    '提取参保人基本信息（只能在前面使用rsTemp记录集，因为后面直接使用了该记录集的数据）
    Call DebugTool("提取参保人基本信息(zl9Insure\AnalyFile_住院1")
    Call WriteBusinessLOG("提取参保人基本信息(zl9Insure\AnalyFile_住院1", "", "")
    gstrSQL = "" & _
        " SELECT A.卡号,A.社保证号,A.缴费地区代码,A.业务类型,A.参保类别,A.公务员," & _
        "     A.公务员待遇,A.病种标识,A.特殊病截止日期,B.入院登记号,B.入院类型,B.入院方式," & _
        "     to_Char(B.入院日期,'yyyy-MM-dd hh24:mi:ss') AS 最初入院日期,D.出院日期" & _
        " From " & strUser & ".保险帐户 A," & strUser & ".入院信息 B,病人信息 C," & strUser & ".出院诊断信息 D" & _
        " WHERE A.病人ID=B.病人ID And A.病人ID=C.病人ID And B.主页ID=C.住院次数 And B.病人ID=D.病人ID(+) And B.主页ID=D.主页ID(+) And A.卡号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参保人基本信息", gComInfo_北京.卡号)
    
    '正式结算，需要保存汇总行数据（住院交易信息）及费用明细数据（住院费用明细）
    '处理汇总行，参数如下
    '交易流水号,缴费地区代码,社保证号,入院登记号,卡号,终端机编号,入院类型,入院方式,入院日期,最初入院日期,出院类型
    '出院日期,交易日期,收费单据号,操作员代码,验证通过方式,费用总额,医保内金额,医保外金额,统筹支付,统筹自付,大额支付
    '大额自付,首先自付,帐户支付,现金支付,个人帐户消费后余额,统筹定额,大额定额,个人自付公务员补助,个人定额自付,
    '特需费用,特殊情况标识,MAC1,医院端结算方式,上传
    dbl费用总额 = Val(Split(arrRow(0), mstrSplit)(col汇总_费用总额))
    dbl医保基金 = Val(Split(arrRow(0), mstrSplit)(col汇总_统筹支付))
    dbl大额补助 = Val(Split(arrRow(0), mstrSplit)(col汇总_大额支付))
    dbl现金 = dbl费用总额 - dbl大额补助 - dbl医保基金
    
    '医院端结算方式:0：普通;1:单病种;2：总额预付
    gstrSQL = "ZL_住院交易信息_INSERT(" & _
        "'" & gComInfo_北京.交易流水号 & "','" & rsTemp!缴费地区代码 & "','" & rsTemp!社保证号 & "','" & rsTemp!入院登记号 & "'," & _
        "'" & rsTemp!卡号 & "',NULL,'" & rsTemp!入院类型 & "','" & rsTemp!入院方式 & "'," & _
        "to_Date('" & str开始日期 & "','yyyy-MM-dd hh24:mi:ss'),to_Date('" & rsTemp!最初入院日期 & "','yyyy-MM-dd hh24:mi:ss')," & _
        "'" & str出院类型 & "',to_Date('" & IIf(IsNull(rsTemp!出院日期), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"), Format(rsTemp!出院日期, "yyyy-MM-dd HH:mm:ss")) & "','yyyy-MM-dd hh24:mi:ss')" & "," & _
        "to_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'),'" & strNO & "'," & _
        "NULL,NULL," & Val(Split(arrRow(0), mstrSplit)(col汇总_费用总额)) & "," & Val(Split(arrRow(0), mstrSplit)(col汇总_医保内)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col汇总_费用总额)) - Val(Split(arrRow(0), mstrSplit)(col汇总_医保内)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col汇总_统筹支付)) & "," & Val(Split(arrRow(0), mstrSplit)(col汇总_统筹自付)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col汇总_大额支付)) & "," & Val(Split(arrRow(0), mstrSplit)(col汇总_大额自付)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col汇总_首先自付)) & ",0," & dbl现金 & ",0,0,0,0,0,0,NULL,NULL,0,0)"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    '处理分段明细,参数如下
    lngRows = Val(Split(arrRow(0), mstrSplit)(col汇总_费用分段数))
    '交易流水号,费用结算周期序号,费用年度,入院登记号,卡号,起始日期,截止日期,费用总额
    '医保内,统筹支付,统筹自付,大额支付,大额自付,个人应付,首先自付,统筹定额
    '大额定额,个人自付公务员补助,个人定额自付,特需费用,上传
    For lngRow = 1 To lngRows
        str分段开始日期 = Split(arrRow(lngRow), mstrSplit)(col分段_起始日期)
        str分段开始日期 = Mid(str分段开始日期, 1, 4) & "-" & Mid(str分段开始日期, 5, 2) & "-" & Mid(str分段开始日期, 7, 2)
        str分段结束日期 = Split(arrRow(lngRow), mstrSplit)(col分段_截止日期)
        str分段结束日期 = Mid(str分段结束日期, 1, 4) & "-" & Mid(str分段结束日期, 5, 2) & "-" & Mid(str分段结束日期, 7, 2)
        gstrSQL = "ZL_住院费用分段明细_INSERT(" & _
                "'" & gComInfo_北京.交易流水号 & "','" & Split(arrRow(lngRow), mstrSplit)(col分段_周期序号) & "'," & _
                "" & Val(Split(arrRow(lngRow), mstrSplit)(col分段_费用年度)) & ",'" & rsTemp!入院登记号 & "','" & gComInfo_北京.卡号 & "',to_Date('" & str分段开始日期 & "','yyyy-MM-dd hh24:mi:ss')," & _
                "" & "to_Date('" & str分段结束日期 & "','yyyy-MM-dd hh24:mi:ss')," & Val(Split(arrRow(lngRow), mstrSplit)(col分段_费用总额)) & "," & _
                "" & Val(Split(arrRow(lngRow), mstrSplit)(col分段_医保内)) & "," & _
                "" & Val(Split(arrRow(lngRow), mstrSplit)(col分段_统筹支付)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col分段_统筹自付)) & "," & _
                "" & Val(Split(arrRow(lngRow), mstrSplit)(col分段_大额支付)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col分段_大额自付)) & "," & _
                "" & Val(Split(arrRow(lngRow), mstrSplit)(col分段_个人应付)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col分段_首先自付)) & "," & _
                "" & "0,0,0,0,0,0)"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    Next
    
    '处理明细（费用明细），参数如下
    '交易流水号,项目序号,入院登记号,卡号,医嘱号,医保编码,项目名称,项目类别,剂型,规格,单价,数量
    '费用总额,发生日期,收费类别,医保内费用,医保外费用,首先自付,分解状态,特需标识,上传
    Dim str明细_NO As String, str明细_性质 As String, str明细_状态 As String, str明细_序号 As String, str标识 As String
    lngRows = Val(Split(arrRow(0), mstrSplit)(col汇总_费用明细数)) + (Val(Split(arrRow(0), "|")(col汇总_费用分段数)))
    For lngRow = 1 + (Val(Split(arrRow(0), "|")(col汇总_费用分段数))) To lngRows
        'col明细_医嘱号=NO|记录性质|记录状态|序号
        '分解出NO、记录性质、记录状态、序号
        str标识 = Split(arrRow(lngRow), mstrSplit)(col明细_医嘱号)
        str明细_NO = Split(str标识, "*")(0)
        str明细_性质 = Split(str标识, "*")(1)
        str明细_状态 = Split(str标识, "*")(2)
        str明细_序号 = Split(str标识, "*")(3)
        
        '需要取出该医保项目的剂型、规格及收费类别，用于上传（HIS中多个项目对应一个医保项目时，可能返回多条记录）
        gstrSQL = "" & _
            " Select A.规格,C.收费类别,C.剂型" & _
            " From 收费细目 A,保险支付项目 B," & strUser & ".药品目录 C,住院费用记录 F" & _
            " Where A.ID=F.收费细目ID ANd A.ID=B.收费细目ID And B.项目编码=C.编码 " & _
            " And F.NO=[1] And F.记录性质=[2] ANd F.记录状态=[3] ANd F.序号=[4]" & _
            " Union" & _
            " Select A.规格,C.收费类别,'' AS 剂型" & _
            " From 收费细目 A,保险支付项目 B," & strUser & ".诊疗目录 C,住院费用记录 F" & _
            " Where A.ID=F.收费细目ID ANd A.ID=B.收费细目ID And B.项目编码=C.编码 " & _
            " And F.NO=[1] And F.记录性质=[2] ANd F.记录状态=[3] ANd F.序号=[4]" & _
            " And A.类别 Not In ('5','6','7')"
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "获取项目信息", str明细_NO, str明细_性质, str明细_状态, str明细_序号)
        
        '如果为空，说明是非医保项目，传固定的值
        If rsDetail.RecordCount = 0 Then
            gstrSQL = " Select A.规格,A.标识码 As 收费类别,'' AS 剂型 " & _
                      " From 药品目录 A,住院费用记录 F" & _
                      " Where A.药品ID=F.收费细目ID And F.NO=[1] And F.记录性质=[2] ANd F.记录状态=[3] ANd F.序号=[4]"
            gstrSQL = gstrSQL & " UNION " & _
                      " Select A.规格,A.标识主码 As 收费类别,'' AS 剂型 " & _
                      " From 收费细目 A,住院费用记录 F" & _
                      " Where A.ID=F.收费细目ID And F.NO=[1] And F.记录性质=[2] ANd F.记录状态=[3] ANd F.序号=[4]" & _
                      " And A.类别 Not In ('5','6','7')"
            Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "获取非医保项目的信息", str明细_NO, str明细_性质, str明细_状态, str明细_序号)
        End If
        
        gstrSQL = "Select to_Char(发生时间,'yyyy-MM-dd hh24:mi:ss') AS 发生时间 From 住院费用记录 F" & _
            " WHERE F.NO=[1] And F.记录性质=[2] ANd F.记录状态=[3] ANd F.序号=[4]"
        Set rs费用发生日期 = zlDatabase.OpenSQLRecord(gstrSQL, "取该记录的费用发生时间", str明细_NO, str明细_性质, str明细_状态, str明细_序号)
        str开始日期 = rs费用发生日期!发生时间
        
        gstrSQL = "ZL_住院费用明细_INSERT(" & _
            "'" & gComInfo_北京.交易流水号 & "','" & Val(Split(arrRow(lngRow), mstrSplit)(col明细_项目序号)) & "'," & _
            "'" & rsTemp!入院登记号 & "','" & gComInfo_北京.卡号 & "',NULL,'" & Split(arrRow(lngRow), mstrSplit)(col明细_项目代码) & "'," & _
            "'" & Split(arrRow(lngRow), mstrSplit)(col明细_项目名称) & "','" & Split(arrRow(lngRow), mstrSplit)(col明细_项目类别) & "'," & _
            "'" & Nvl(rsDetail!剂型) & "','" & ToVarchar(Nvl(rsDetail!规格), 40) & "'," & Val(Split(arrRow(lngRow), mstrSplit)(col明细_单价)) & "," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col明细_数量)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col明细_费用总额)) & "," & _
            "to_Date('" & str开始日期 & "','yyyy-MM-dd hh24:mi:ss'),'" & Nvl(rsDetail!收费类别) & "'," & Val(Split(arrRow(lngRow), mstrSplit)(col明细_医保内)) & "," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col明细_医保外)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col明细_首先自付)) & "," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col明细_分解状态)) & ",0,0)"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    Next
    
    AnalyFile_住院1 = arrRow(0)
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\AnalyFile_住院1)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\AnalyFile_住院1)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
End Function

Private Function CheckDetail(ByVal str项目名称 As String, ByVal int分解状态 As Integer) As Boolean
    Dim strMsg As String
    '----------------------------------------------------------------
    '功能描述   ：检查费用明细分解的情况，对于不正常的项目，给予提示
    '编写人     ：朱玉宝
    '编写日期   ：2004-07-03
    '----------------------------------------------------------------
    Exit Function
    
    strMsg = "该项目[" & str项目名称 & "]的分解状态不正常，具体信息如下：" & vbCrLf
    Select Case int分解状态
    Case 0
        Exit Function
    Case 1
        strMsg = strMsg & "不符合特殊标识"
    Case 2
        strMsg = strMsg & "医保目录内不存在该项目"
    Case 3
        strMsg = strMsg & "医院项目与医保项目对码错误"
    Case Else
        strMsg = strMsg & "未知错误，可能是医保端接口政策有变动，而HIS端的医保接口未更新"
    End Select
    MsgBox strMsg, vbInformation, gstrSysName
    CheckDetail = True
End Function

Private Function CheckBlockage(ByVal StrInput As String, Optional ByVal bln卡号 As Boolean = True) As Boolean
    '----------------------------------------------------------------
    '功能描述   ：返回最小的封锁日期（使用卡的才判断充值黑名单）
    '编写人     ：朱玉宝
    '编写日期   ：2004-06-28
    '1.  个人黑名单：如病人卡号（手册号）在个人黑名单中， 则提示操作员且该卡（医疗手册）不能使用，如要继续则视为非医保病人；
    '2.  充值黑名单：如病人卡号在充值黑名单中，则提示病人应去充值，交易不能进行；
    '3.  单位黑名单：如病人社保登记证号在单位黑名单中，则提示病人不享受医保待遇，所有费用均为自费，但交易可用个人帐户进行支付；
    '参数说明   ：bln卡号=False:strInput=病人ID;否则,strInput=卡号|社保证号
    '----------------------------------------------------------------
    Dim str卡号 As String, str社保证号 As String
    Dim bln用卡标志 As Boolean
    Dim rsPersonal As New ADODB.Recordset
    Dim rsBlockage As New ADODB.Recordset
    On Error GoTo errHand
    
    '提取病人的卡号及社保证号
    Call DebugTool("提取病人的卡号及社保证号")
    Call WriteBusinessLOG("提取病人的卡号及社保证号", "", "")
    If Not bln卡号 Then
        gstrSQL = "Select 卡号,社保证号 " & _
            " From 保险帐户 " & _
            " Where 病人ID=" & Val(str卡号)
        If rsPersonal.State = 1 Then rsPersonal.Close
        Call SQLTest(App.Title, "ZL9INSURE\CheckBlockage", gstrSQL): rsPersonal.Open gstrSQL, gcnBJYB: Call SQLTest
        If rsPersonal.EOF Then
            Call DebugTool("未找到该参保人的帐户信息[zlBJ.保险帐户]")
            Call WriteBusinessLOG("未找到该参保人的帐户信息[zlBJ.保险帐户]", "", "")
            Exit Function
        End If
        str卡号 = Nvl(rsPersonal!卡号)
        str社保证号 = Nvl(rsPersonal!社保证号)
    Else
        str卡号 = Split(StrInput, mstrSplit)(0)
        str社保证号 = Split(StrInput, mstrSplit)(1)
    End If
    bln用卡标志 = Not (UCase(Right(str卡号, 1)) = "S")
    
    '依次判断个人黑名单、充值黑名单及单位黑名单
    Call DebugTool("判断个人黑名单")
    Call WriteBusinessLOG("判断个人黑名单", "", "")
    gstrSQL = "SELECT 封锁原因 FROM 个人黑名单 WHERE 卡号='" & str卡号 & "'"
    If rsBlockage.State = 1 Then rsBlockage.Close
    Call SQLTest(App.Title, "ZL9INSURE\CheckBlockage", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsBlockage.RecordCount <> 0 Then
        MsgBox "该医疗手册或卡在个人黑名单中，不允许使用，请按普通病人办理！" & vbCrLf & _
            "封锁原因：" & rsBlockage!封锁原因, vbInformation, gstrSysName
        Exit Function
    End If
    
    If bln用卡标志 Then
        Call DebugTool("判断充值黑名单")
        Call WriteBusinessLOG("判断充值黑名单", "", "")
        gstrSQL = "SELECT 封锁原因 FROM 充值黑名单 WHERE 卡号='" & str卡号 & "'"
        If rsBlockage.State = 1 Then rsBlockage.Close
        Call SQLTest(App.Title, "ZL9INSURE\CheckBlockage", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
        If rsBlockage.RecordCount <> 0 Then
            MsgBox "该卡必须充值后才能继续交易！" & vbCrLf & _
                "封锁原因：" & rsBlockage!封锁原因, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Call DebugTool("判断单位黑名单")
    Call WriteBusinessLOG("判断单位黑名单", "", "")
    gstrSQL = "SELECT 封锁原因 FROM 单位黑名单 WHERE 社保证号='" & str社保证号 & "'"
    If rsBlockage.State = 1 Then rsBlockage.Close
    Call SQLTest(App.Title, "ZL9INSURE\CheckBlockage", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsBlockage.RecordCount <> 0 Then
        MsgBox "该医疗手册号在单位黑名单中，如要继续，所有费用均为自费，但交易时可由帐户支付！" & vbCrLf & _
            "封锁原因：" & rsBlockage!封锁原因, vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckBlockage = True
    Exit Function
errHand:
    Call DebugTool("待遇信息检查时发生错误！" & vbCrLf & "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("待遇信息检查时发生错误！" & vbCrLf & "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
End Function

Private Function GetBlockage(ByVal StrInput As String, Optional ByVal bln卡号 As Boolean = True) As String
    '----------------------------------------------------------------
    '功能描述   ：返回最小的封锁日期（使用卡的才判断充值黑名单）
    '编写人     ：朱玉宝
    '编写日期   ：2004-06-28
    '1.  个人黑名单：如病人卡号（手册号）在个人黑名单中， 则提示操作员且该卡（医疗手册）不能使用，如要继续则视为非医保病人；
    '2.  充值黑名单：如病人卡号在充值黑名单中，则提示病人应去充值，交易不能进行；
    '3.  单位黑名单：如病人社保登记证号在单位黑名单中，则提示病人不享受医保待遇，所有费用均为自费，但交易可用个人帐户进行支付；
    '参数说明   ：bln卡号=False:strInput=病人ID;否则,strInput=卡号|社保证号
    '----------------------------------------------------------------
    Dim str卡号 As String, str社保证号 As String
    Dim bln用卡标志 As Boolean
    Dim rsPersonal As New ADODB.Recordset
    Dim rsBlockage As New ADODB.Recordset
    Dim str封锁日期 As String
    Dim str个人黑名单 As String, str充值黑名单 As String, str单位黑名单 As String
    On Error GoTo errHand
    
    '提取病人的卡号及社保证号
    Call DebugTool("提取病人的卡号及社保证号")
    Call WriteBusinessLOG("提取病人的卡号及社保证号", "", "")
    If Not bln卡号 Then
        gstrSQL = "Select 卡号,社保证号 " & _
            " From 保险帐户 " & _
            " Where 病人ID=" & Val(str卡号)
        If rsPersonal.State = 1 Then rsPersonal.Close
        Call SQLTest(App.Title, "ZL9INSURE\GETBLOCKAGE", gstrSQL): rsPersonal.Open gstrSQL, gcnBJYB: Call SQLTest
        If rsPersonal.EOF Then
            Call DebugTool("未找到该参保人的帐户信息[zlBJ.保险帐户]")
            Call WriteBusinessLOG("未找到该参保人的帐户信息[zlBJ.保险帐户]", "", "")
            Exit Function
        End If
        str卡号 = Nvl(rsPersonal!卡号)
        str社保证号 = Nvl(rsPersonal!社保证号)
    Else
        str卡号 = Split(StrInput, mstrSplit)(0)
        str社保证号 = Split(StrInput, mstrSplit)(1)
    End If
    bln用卡标志 = Not (UCase(Right(str卡号, 1)) = "S")
    
    '依次判断个人黑名单、充值黑名单及单位黑名单
    Call DebugTool("判断个人黑名单")
    Call WriteBusinessLOG("判断个人黑名单", "", "")
    gstrSQL = "SELECT to_char(封锁日期,'yyyy-MM-dd') As 封锁日期 FROM 个人黑名单 WHERE 卡号='" & str卡号 & "'"
    If rsBlockage.State = 1 Then rsBlockage.Close
    Call SQLTest(App.Title, "ZL9INSURE\GETBLOCKAGE", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsBlockage.RecordCount <> 0 Then
        str个人黑名单 = Nvl(rsBlockage!封锁日期)
    End If
    
    If bln用卡标志 Then
        Call DebugTool("判断充值黑名单")
        Call WriteBusinessLOG("判断充值黑名单", "", "")
        gstrSQL = "SELECT to_char(封锁日期,'yyyy-MM-dd') As 封锁日期 FROM 充值黑名单 WHERE 卡号='" & str卡号 & "'"
        If rsBlockage.State = 1 Then rsBlockage.Close
        Call SQLTest(App.Title, "ZL9INSURE\GETBLOCKAGE", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
        If rsBlockage.RecordCount <> 0 Then
            str充值黑名单 = Nvl(rsBlockage!封锁日期)
        End If
    End If
    
    Call DebugTool("判断单位黑名单")
    Call WriteBusinessLOG("判断单位黑名单", "", "")
    gstrSQL = "SELECT to_char(封锁日期,'yyyy-MM-dd') As 封锁日期 FROM 单位黑名单 WHERE 社保证号='" & str社保证号 & "'"
    If rsBlockage.State = 1 Then rsBlockage.Close
    Call SQLTest(App.Title, "ZL9INSURE\GETBLOCKAGE", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsBlockage.RecordCount <> 0 Then
        str单位黑名单 = Nvl(rsBlockage!封锁日期)
    End If
    
    If str个人黑名单 <> "" Then str封锁日期 = str个人黑名单
    If str充值黑名单 <> "" Then
        If str封锁日期 <> "" Then
            If str充值黑名单 < str封锁日期 Then
                str封锁日期 = str充值黑名单
            End If
        Else
            str封锁日期 = str充值黑名单
        End If
    End If
    If str单位黑名单 <> "" Then
        If str封锁日期 <> "" Then
            If str单位黑名单 < str封锁日期 Then
                str封锁日期 = str单位黑名单
            End If
        Else
            str封锁日期 = str单位黑名单
        End If
    End If
    
    GetBlockage = str封锁日期
    Exit Function
errHand:
    Call DebugTool("获取黑名单封锁日期时发生错误！" & vbCrLf & "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("获取黑名单封锁日期时发生错误！" & vbCrLf & "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
End Function

Private Function SaveBusinessDeal(ByVal strDeal As String) As Boolean
    '----------------------------------------------------------------
    '功能描述   ：保存交易待遇信息
    '编写人     ：朱玉宝
    '编写日期   ：2004-07-06
    '----------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim str封锁日期 As String           '封锁日期
    Dim int黑名单 As Integer            '是否在单位黑名单中(0-是;1-否)
    Dim arrDeal
    
    On Error GoTo errHand
    
    Call DebugTool("进入(zl9Insure\SaveBusinessDeal")
    Call WriteBusinessLOG("进入(zl9Insure\SaveBusinessDeal", "", "")
    arrDeal = Split(strDeal, mstrSplit)
    
    '判断此刻是否在单位黑名单中
    Call DebugTool("判断此刻是否在单位黑名单中(zl9Insure\SaveBusinessDeal")
    Call WriteBusinessLOG("判断此刻是否在单位黑名单中(zl9Insure\SaveBusinessDeal", "", "")
    gstrSQL = "Select 封锁日期 From 单位黑名单 Where 社保证号=" & _
        " (Select 社保证号 From 保险帐户 Where 卡号='" & gComInfo_北京.卡号 & "')"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\SaveBusinessDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsTemp.EOF Then
        int黑名单 = 1
        str封锁日期 = ""
    Else
        int黑名单 = 0
        If Not IsNull(rsTemp!封锁日期) Then
            str封锁日期 = Format(rsTemp!封锁日期, "yyyy-MM-dd") '不可能为空
        End If
    End If
    
    '提取参保人基本信息
    Call DebugTool("提取参保人基本信息(zl9Insure\SaveBusinessDeal")
    Call WriteBusinessLOG("提取参保人基本信息(zl9Insure\SaveBusinessDeal", "", "")
    gstrSQL = "" & _
        " SELECT 业务类型,参保类别,公务员,公务员待遇,病种标识,特殊病截止日期" & _
        " From 保险帐户" & _
        " WHERE 卡号='" & gComInfo_北京.卡号 & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\SaveBusinessDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    
    '准备保存交易待遇信息
    Call DebugTool("保险交易待遇信息(zl9Insure\SaveBusinessDeal")
    Call WriteBusinessLOG("保险交易待遇信息(zl9Insure\SaveBusinessDeal", "", "")
'    业务类型,卡号,交易流水号,参保类别,公务员待遇标识,公务员待遇,特殊病标识,特殊病截止日期,费用年度,
'    本年门诊医保内,本年门诊大额支付,本年统筹支付,本年统筹大额,本年次数累计,本周期起始日期,本周期交易号,
'    本周期医保内,本周期统筹支付,本周期大额支付,住院家庭病床年度,本次住院家庭病床标识,本次住院家庭病床交易号,
'    本次住院家庭病床起始日期,本次住院家庭病床医保内,本次住院家庭病床统筹支付,
'    本次住院大额支付,交易时是否在单位黑名单中,单位黑名单封锁日期,上传
    gstrSQL = "ZL_交易待遇信息_INSERT(" & _
              "'" & rsTemp!业务类型 & "','" & gComInfo_北京.卡号 & "','" & gComInfo_北京.交易流水号 & "'," & _
              "'" & rsTemp!参保类别 & "','" & rsTemp!公务员 & "','" & IIf(rsTemp!公务员待遇 = -1, "", rsTemp!公务员待遇) & "'," & _
              "" & rsTemp!病种标识 & "," & IIf(IsNull(rsTemp!特殊病截止日期), "NULL", "to_Date('" & rsTemp!特殊病截止日期 & "','yyyy-MM-dd')") & "," & _
              "" & arrDeal(交易待遇.费用年度) & "," & arrDeal(交易待遇.本年门诊医保内) & "," & arrDeal(交易待遇.本年门诊大额支付) & "," & _
              "" & arrDeal(交易待遇.本年统筹支付) & "," & arrDeal(交易待遇.本年大额支付) & "," & arrDeal(交易待遇.本年次数累计) & "," & _
              "" & IIf(Trim(arrDeal(交易待遇.本周期起始日期)) = "", "NULL", "To_Date('" & arrDeal(交易待遇.本周期起始日期) & "','yyyy-MM-dd')") & "," & _
              "'" & arrDeal(交易待遇.本周期交易号) & "'," & arrDeal(交易待遇.本周期医保内) & "," & _
              "" & arrDeal(交易待遇.本周期统筹支付) & "," & arrDeal(交易待遇.本周期大额支付) & "," & arrDeal(交易待遇.住院家庭病床年度) & "," & _
              "'" & arrDeal(交易待遇.本次住院家庭病床标识) & "','" & arrDeal(交易待遇.本次住院家庭病床交易号) & "'," & _
              "" & IIf(Trim(arrDeal(交易待遇.本次住院家庭病床起始日期)) = "", "NULL", "To_Date('" & arrDeal(交易待遇.本次住院家庭病床起始日期) & "','yyyy-MM-dd')") & "," & _
              "" & arrDeal(交易待遇.本次住院家庭病床医保内) & "," & arrDeal(交易待遇.本次住院家庭病床统筹支付) & "," & arrDeal(交易待遇.本次住院大额支付) & "," & _
              "'" & int黑名单 & "'," & IIf(str封锁日期 = "", "NULL", "to_Date('" & str封锁日期 & "','yyyy-MM-dd')") & _
              ")"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    SaveBusinessDeal = True
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\SaveBusinessDeal)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\SaveBusinessDeal)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveBusinessVersion(Optional ByVal bln门诊 As Boolean = True) As Boolean
    Dim strReturn As String
    Dim strDllVersion As String, strDataVersion As String, strHISVersion As String
    Dim strCompanyVersion As String, strPersonalVersion As String, strCostVersion As String
    '----------------------------------------------------------------
    '功能描述   ：保存交易版本信息
    '编写人     ：朱玉宝
    '编写日期   ：2004-07-06
    '----------------------------------------------------------------
    On Error GoTo errHand
    Dim rsVersion As New ADODB.Recordset
    
    Call DebugTool("准备调用接口获取DLL版本号(zl9Insure\SaveBusinessVersion)")
    Call WriteBusinessLOG("准备调用接口获取DLL版本号(zl9Insure\SaveBusinessVersion)", "", "")
    '获取接口部件的组件版本号、数据包版本号
    If Not 调用接口_北京(接口版本信息, strReturn) Then Exit Function
    strDllVersion = Split(strReturn, "||")(0)
    strDataVersion = Split(strReturn, "||")(1)
    
    Call DebugTool("准备提取黑名单版本号(zl9Insure\SaveBusinessVersion)")
    Call WriteBusinessLOG("准备提取黑名单版本号(zl9Insure\SaveBusinessVersion)", "", "")
    '获取单位黑名单、个人黑名单及充值黑名单版本号(05：个人黑名单;06：单位黑名单;07：充值黑名单)
    gstrSQL = "Select 文件编码,版本号 From 版本控制 Where 文件编码 IN ('05','55','06','56','07','57')"
    If rsVersion.State = 1 Then rsVersion.Close
    rsVersion.Open gstrSQL, gcnBJYB
    '不可能没有记录，如果本次下载没有新的黑名单，将保留上次的黑名单，数据库中只保留最新的记录
    '增量与全量的版本号肯定是一致的
    Do While Not rsVersion.EOF
        Select Case rsVersion!文件编码
        Case "05", "55"
            If strPersonalVersion = "" Then strPersonalVersion = Nvl(rsVersion!版本号)
        Case "06", "56"
            If strCompanyVersion = "" Then strCompanyVersion = Nvl(rsVersion!版本号)
        Case "07", "57"
            If strCostVersion = "" Then strCostVersion = Nvl(rsVersion!版本号)
        End Select
        rsVersion.MoveNext
    Loop
    
    '获取HIS部件版本号
    Call DebugTool("准备获取医保部件版本号(zl9Insure\SaveBusinessVersion)")
    Call WriteBusinessLOG("准备获取医保部件版本号(zl9Insure\SaveBusinessVersion)", "", "")
    strHISVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    '产生一条本次交易的交易版本信息
'    性质,交易流水号,卡号,DLL组件,DLL数据包,单位黑名单,个人黑名单,充值黑名单,HIS部件,上传
    Call DebugTool("插入交易版本信息(zl9Insure\SaveBusinessVersion)")
    Call WriteBusinessLOG("插入交易版本信息(zl9Insure\SaveBusinessVersion)", "", "")
    gstrSQL = "ZL_交易版本信息_INSERT(" & IIf(bln门诊, 1, 2) & ",'" & gComInfo_北京.交易流水号 & "'," & _
              "'" & gComInfo_北京.卡号 & "','" & strDllVersion & "','" & strDataVersion & "'," & _
              "'" & strCompanyVersion & "','" & strPersonalVersion & "','" & strCostVersion & "'," & _
              "'" & strHISVersion & "')"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    SaveBusinessVersion = True
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\SaveBusinessVersion)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\SaveBusinessVersion)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then Resume
End Function

Private Function GetUser() As String
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=[1] And 参数名='医保用户名'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取中间库用户名", TYPE_北京)
    GetUser = rsTemp!参数值
End Function

Private Function TransationSpec(ByVal str待遇 As String) As String
    Dim arr待遇
    Dim strReturn As String
    Dim str费用发生日期 As String
    Dim rsTemp As New ADODB.Recordset
    
    '----------------------------------------------------------------
    '功能描述   ：保存交易版本信息
    '编写人     ：朱玉宝
    '编写日期   ：2004-07-06
    '参数说明   ：bln性质=TRUE（门诊）,FALSE（住院）
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    arr待遇 = Split(str待遇, mstrSplit)
    str费用发生日期 = Format(zlDatabase.Currentdate(), "yyyyMMdd")
    '提取参保人基本信息
    Call DebugTool("提取参保人基本信息(zl9Insure\TransationSpec")
    Call WriteBusinessLOG("提取参保人基本信息(zl9Insure\TransationSpec", "", "")
    gstrSQL = "" & _
        " SELECT 业务类型,参保类别,公务员,公务员待遇,病种标识,TO_CHAR(特殊病截止日期,'yyyyMMdd') AS 特殊病截止日期 " & _
        " From 保险帐户" & _
        " WHERE 卡号='" & gComInfo_北京.卡号 & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\TransationSpec", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    
    strReturn = gComInfo_北京.交易流水号 & mstrSplit & gComInfo_北京.卡号 & mstrSplit & _
        rsTemp!参保类别 & mstrSplit & rsTemp!公务员 & mstrSplit & IIf(rsTemp!公务员待遇 = -1, "", rsTemp!公务员待遇) & mstrSplit & _
        rsTemp!病种标识 & mstrSplit & Nvl(rsTemp!特殊病截止日期) & mstrSplit & arr待遇(交易待遇.费用年度) & mstrSplit & _
        arr待遇(交易待遇.本年门诊医保内) & mstrSplit & arr待遇(交易待遇.本年门诊大额支付) & mstrSplit & arr待遇(交易待遇.本年统筹支付) & mstrSplit & _
        arr待遇(交易待遇.本年大额支付) & mstrSplit & arr待遇(交易待遇.本年次数累计) & mstrSplit & arr待遇(交易待遇.本周期起始日期) & mstrSplit & _
        arr待遇(交易待遇.本周期交易号) & mstrSplit & arr待遇(交易待遇.本周期医保内) & mstrSplit & arr待遇(交易待遇.本周期统筹支付) & mstrSplit & _
        arr待遇(交易待遇.本周期大额支付) & mstrSplit & str费用发生日期
        
    TransationSpec = strReturn
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\TransationSpec)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\TransationSpec)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function TransationHosp(ByVal str待遇 As String) As String
    Dim arr待遇
    Dim strReturn As String
    Dim strUser As String
    Dim str入院日期 As String               '都是入院日期
    Dim str出院日期 As String               '中途结算将当前日期做为出院日期
    Dim str封锁日期 As String               '最小的封锁日期
    Dim rsTemp As New ADODB.Recordset
    
    '----------------------------------------------------------------
    '功能描述   ：保存交易版本信息
    '编写人     ：朱玉宝
    '编写日期   ：2004-07-06
    '参数说明   ：bln性质=TRUE（门诊）,FALSE（住院）
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    arr待遇 = Split(str待遇, mstrSplit)
    strUser = GetUser()
    
    '提取参保人本次住院基本信息
    Call DebugTool("提取参保人本次住院基本信息(zl9Insure\TransationHosp")
    Call WriteBusinessLOG("提取参保人本次住院基本信息(zl9Insure\TransationHosp", "", "")
    gstrSQL = "" & _
        " SELECT A.社保证号,A.业务类型,A.参保类别,A.公务员,A.公务员待遇,A.病种标识," & _
        "        TO_CHAR(A.特殊病截止日期,'yyyyMMdd') AS 特殊病截止日期," & _
        "        B.入院类型,B.入院方式,B.入院日期,C.出院日期" & _
        " From " & strUser & ".保险帐户 A," & strUser & ".入院信息 B," & strUser & ".出院诊断信息 C,病人信息 D" & _
        " WHERE A.病人ID=B.病人ID And B.病人ID=C.病人ID(+) and B.主页ID=C.主页ID(+)" & _
        " And A.病人ID=D.病人ID And B.主页ID=D.住院次数 And A.卡号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参保人本次住院基本信息", gComInfo_北京.卡号)
    str入院日期 = Format(rsTemp!入院日期, "yyyyMMdd")
    If Not IsNull(rsTemp!出院日期) Then
        str出院日期 = Format(rsTemp!出院日期, "yyyyMMdd")
    Else
        str出院日期 = Format(zlDatabase.Currentdate, "yyyyMMdd")
    End If
    
    '将封锁日期格式转换成yyyyMMdd
    str封锁日期 = GetBlockage(gComInfo_北京.卡号 & mstrSplit & rsTemp!社保证号)
    If str封锁日期 <> "" Then
        str封锁日期 = Format(str封锁日期, "yyyyMMdd")
    End If
    strReturn = gComInfo_北京.交易流水号 & mstrSplit & gComInfo_北京.卡号 & mstrSplit & _
        rsTemp!参保类别 & mstrSplit & rsTemp!公务员 & mstrSplit & IIf(rsTemp!公务员待遇 = -1, "", rsTemp!公务员待遇) & mstrSplit & _
        rsTemp!入院类型 & mstrSplit & rsTemp!病种标识 & mstrSplit & Nvl(rsTemp!特殊病截止日期) & mstrSplit & _
        rsTemp!入院方式 & mstrSplit & str入院日期 & mstrSplit & str出院日期 & mstrSplit & arr待遇(交易待遇.费用年度) & mstrSplit & _
        arr待遇(交易待遇.本年门诊医保内) & mstrSplit & arr待遇(交易待遇.本年门诊大额支付) & mstrSplit & arr待遇(交易待遇.本年统筹支付) & mstrSplit & _
        arr待遇(交易待遇.本年大额支付) & mstrSplit & arr待遇(交易待遇.本年次数累计) & mstrSplit & arr待遇(交易待遇.本周期起始日期) & mstrSplit & _
        arr待遇(交易待遇.本周期交易号) & mstrSplit & arr待遇(交易待遇.本周期医保内) & mstrSplit & arr待遇(交易待遇.本周期统筹支付) & mstrSplit & _
        arr待遇(交易待遇.本周期大额支付) & mstrSplit & arr待遇(交易待遇.住院家庭病床年度) & mstrSplit & arr待遇(交易待遇.本次住院家庭病床标识) & mstrSplit & _
        arr待遇(交易待遇.本次住院家庭病床交易号) & mstrSplit & arr待遇(交易待遇.本次住院家庭病床起始日期) & mstrSplit & arr待遇(交易待遇.本次住院家庭病床医保内) & mstrSplit & _
        arr待遇(交易待遇.本次住院家庭病床统筹支付) & mstrSplit & arr待遇(交易待遇.本次住院大额支付) & mstrSplit & _
        IIf(str封锁日期 = "", 1, 0) & mstrSplit & IIf(str封锁日期 = "", "", str封锁日期)

    TransationHosp = strReturn
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\TransationHosp)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\TransationHosp)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveDeal(Optional ByVal bln门诊 As Boolean = True, Optional ByVal bln保存 As Boolean = False) As Boolean
    '----------------------------------------------------------------
    '功能描述   ：保存交易版本信息
    '编写人     ：朱玉宝
    '编写日期   ：2004-07-06
    '参数说明   ：bln性质=TRUE（门诊）,FALSE（住院）
    '门诊:
    '    总金额 = 特殊病医保内 + 特殊病医保外
    '    统筹基金支付 = 统筹支付
    '    大额医疗/公务员补助费用 = 大额支付
    '    个人自付：自付1 = 特殊病医保内 - 统筹支付 - 大额支付；自付2 = 自付2
    '    个人自费 = 特殊病医保外
    '    统筹封顶后医保内金额 = 统筹封顶后医保内金额
    '住院:
    '    总金额 = 本段费用总金额
    '    统筹基金支付 = 本段统筹支付
    '    大额医疗/公务员补助费用 = 本段大额支付
    '    个人自付：自付1 = 本段费用总金额 - 本段统筹支付 - 本段大额支付；自付2 = 本段自付2
    '    个人自费 = 本段医保外
    '    统筹封顶后医保内金额 = 本段统筹封顶后医保内金额
    '----------------------------------------------------------------
    Dim rsHead As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Dim str起始日期 As String, str终止日期 As String, str经办日期 As String, str入院类型 As String
    Dim dbl费用总额 As Double, dbl统筹支付 As Double, dbl大额支付 As Double
    Dim dbl个人自付 As Double, dbl个人自费 As Double, dbl首先自付 As Double, dbl统筹封顶后医保内 As Double
    Dim str医院名称 As String, str医院等级 As String
    Dim strFields As String, strValues As String
    Dim strTotal As String
    Dim strFile_Out As String
    
    Dim arrRow
    Dim lngRow As Long, lngRows As Long
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
    
    Const str门诊 As String = "Spec_Divide.out"
    Const str住院 As String = "Hosp_Divide.out"
    Const col汇总行_分段行数 As Integer = 0
    Const col门诊特殊病_记录数 As Integer = 0
    Const col门诊特殊病_交易流水号 As Integer = 1
    Const col门诊特殊病_费用总额 As Integer = 2
    Const col门诊特殊病_普通门诊医保内 As Integer = 3
    Const col门诊特殊病_普通门诊医保外 As Integer = 4
    Const col门诊特殊病_统筹支付 As Integer = 5
    Const col门诊特殊病_统筹自付 As Integer = 6
    Const col门诊特殊病_大额支付 As Integer = 7
    Const col门诊特殊病_大额自付 As Integer = 8
    Const col门诊特殊病_特殊病自付 As Integer = 9
    Const col门诊特殊病_特殊病医保外 As Integer = 10
    Const col门诊特殊病_首先自付 As Integer = 11
    Const col门诊特殊病_统筹封顶后医保内 As Integer = 12
    Const col住院_周期序号 As Integer = 0
    Const col住院_费用年度 As Integer = 1
    Const col住院_起始日期 As Integer = 2
    Const col住院_截止日期 As Integer = 3
    Const col住院_费用总额 As Integer = 4
    Const col住院_医保内 As Integer = 5
    Const col住院_统筹支付 As Integer = 6
    Const col住院_统筹自付 As Integer = 7
    Const col住院_大额支付 As Integer = 8
    Const col住院_大额自付 As Integer = 9
    Const col住院_首先自付 As Integer = 10
    Const col住院_个人应付 As Integer = 11
    On Error GoTo errHand
    
    str经办日期 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    '获取医院名称，用以填写手册消费记录
    Call DebugTool("提取医院名称及医院等级(zl9Insure\SaveDeal")
    gstrSQL = "Select A.医院编码,A.医院名称,B.名称 AS 医院等级 " & _
            " From 医院等级 A,(Select B.编码,B.名称 From 指标主表 A,指标体系对照表 B Where A.类别=B.类别 And A.名称='医院等级') B" & _
            " Where A.医院等级=B.编码 And A.医院编码='" & gComInfo_北京.医院编码 & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\SaveDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsTemp.RecordCount = 0 Then
        MsgBox "医院编码有误，医院等级清单中没有对应的医院信息！", vbInformation, gstrSysName
        Exit Function
    End If
    str医院名称 = rsTemp!医院编码 & vbCrLf & rsTemp!医院名称
    str医院等级 = rsTemp!医院等级
    
    If Not bln门诊 Then
        '获取参保人入院类型
        Call DebugTool("获取参保人入院类型(zl9Insure\SaveDeal")
        gstrSQL = "Select B.名称 AS 入院类别" & _
                " From 保险帐户 A,(Select B.编码,B.名称 From 指标主表 A,指标体系对照表 B Where A.类别=B.类别 And A.名称='入院类别') B" & _
                " Where A.入院类别=B.编码 And A.卡号='" & gComInfo_北京.卡号 & "'"
        If rsTemp.State = 1 Then rsTemp.Close
        Call SQLTest(App.Title, "ZL9INSURE\SaveDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
        str入院类型 = rsTemp!入院类别
    End If
    
    '初始化记录集
    If bln门诊 Then
        strFile_Out = str门诊
        strFields = "医院名称," & adLongVarChar & ",100|" & _
                    "就诊日期," & adLongVarChar & ",100|" & _
                    "医院级别," & adLongVarChar & ",100|" & _
                    "初步诊断," & adLongVarChar & ",100"
        Call Record_Init(rsHead, strFields)
        strFields = "费用总额," & adLongVarChar & ",100|" & _
                    "统筹支付," & adLongVarChar & ",100|" & _
                    "大额支付," & adLongVarChar & ",100|" & _
                    "个人自付," & adLongVarChar & ",100|" & _
                    "首先自付," & adLongVarChar & ",100|" & _
                    "个人自费," & adLongVarChar & ",100|" & _
                    "统筹封顶后医保内金额," & adLongVarChar & ",100|" & _
                    "起始日期," & adLongVarChar & ",100|" & _
                    "终止日期," & adLongVarChar & ",100|" & _
                    "经办日期," & adLongVarChar & ",100"
        Call Record_Init(rsDeal, strFields)
    Else
        strFile_Out = str住院
        strFields = "医院名称," & adLongVarChar & ",100|" & _
                    "入院日期," & adLongVarChar & ",100|" & _
                    "医院级别," & adLongVarChar & ",100|" & _
                    "初步诊断," & adLongVarChar & ",100|" & _
                    "入院类型," & adLongVarChar & ",100|" & _
                    "转出日期," & adLongVarChar & ",100"
        Call Record_Init(rsHead, strFields)
        strFields = "费用总额," & adLongVarChar & ",100|" & _
                    "统筹支付," & adLongVarChar & ",100|" & _
                    "大额支付," & adLongVarChar & ",100|" & _
                    "个人自付," & adLongVarChar & ",100|" & _
                    "首先自付," & adLongVarChar & ",100|" & _
                    "个人自费," & adLongVarChar & ",100|" & _
                    "统筹封顶后医保内金额," & adLongVarChar & ",100|" & _
                    "起始日期," & adLongVarChar & ",100|" & _
                    "终止日期," & adLongVarChar & ",100|" & _
                    "经办日期," & adLongVarChar & ",100"
        Call Record_Init(rsDeal, strFields)
    End If
    
    '准备分析结算文件，产生结算记录集，以便显示结算信息
    Call DebugTool("准备分析结算文件，产生结算记录集，以便显示结算信息(zl9Insure\SaveDeal)")
    Call WriteBusinessLOG("准备分析结算文件，产生结算记录集，以便显示结算信息(zl9Insure\SaveDeal)", "", "")
    If Not objFileSys.FileExists(gComInfo_北京.出参目录 & "\" & strFile_Out) Then Exit Function
    Set objStream = objFileSys.OpenTextFile(gComInfo_北京.出参目录 & "\" & strFile_Out)
    
    Call DebugTool("读取汇总行数据(zl9Insure\SaveDeal)")
    Call WriteBusinessLOG("读取汇总行数据(zl9Insure\SaveDeal)", "", "")
    '可能每行文本是以换行符结束，而VB认的是回车换行，将所有数据都读出来了，需要SPLIT
    strTotal = objStream.ReadLine
    arrRow = Split(strTotal, vbCr)
    objStream.Close
    Set objStream = Nothing
    
    If bln门诊 Then
        
        strFields = "医院名称|就诊日期|医院级别|初步诊断"
        strValues = str医院名称 & "|" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd") & "|" & str医院等级 & "|"
        Call Record_Add(rsHead, strFields, strValues)
        
        '仅处理汇总行即可
        dbl费用总额 = Val(Split(arrRow(0), "|")(col门诊特殊病_费用总额)) - Val(Split(arrRow(0), "|")(col门诊特殊病_普通门诊医保内)) - Val(Split(arrRow(0), "|")(col门诊特殊病_普通门诊医保外))
        dbl大额支付 = Val(Split(arrRow(0), "|")(col门诊特殊病_大额支付))
        dbl统筹支付 = Val(Split(arrRow(0), "|")(col门诊特殊病_统筹支付))
        dbl统筹封顶后医保内 = Val(Split(arrRow(0), "|")(col门诊特殊病_统筹封顶后医保内))
        dbl个人自费 = Val(Split(arrRow(0), "|")(col门诊特殊病_特殊病医保外))
        '个人自付=特殊病医保内-统筹支付-大额支付
        dbl个人自付 = (dbl费用总额 - dbl个人自费) - dbl大额支付 - dbl统筹支付
        dbl首先自付 = Val(Split(arrRow(0), "|")(col门诊特殊病_首先自付))
        
        strFields = "费用总额|统筹支付|大额支付|个人自付|首先自付|个人自费|统筹封顶后医保内金额|起始日期|终止日期|经办日期"
        strValues = dbl费用总额 & "|" & dbl统筹支付 & "|" & dbl大额支付 & "|" & dbl个人自付 & "|" & dbl首先自付 & _
            "|" & dbl个人自费 & "|" & dbl统筹封顶后医保内 & "|" & str经办日期 & "|" & str经办日期 & "|" & str经办日期
        Call Record_Add(rsDeal, strFields, strValues)
    Else
        lngRows = Val(Split(arrRow(0), "|")(col汇总行_分段行数))
        '需要处理分段分解明细行
        For lngRow = 1 To lngRows
            strFields = "医院名称|入院日期|医院级别|初步诊断|入院类型|转出日期"
            str起始日期 = Split(arrRow(lngRow), "|")(col住院_起始日期)
            str终止日期 = Split(arrRow(lngRow), "|")(col住院_截止日期)
            strValues = str医院名称 & "|" & str起始日期 & "|" & str医院等级 & "||" & str入院类型 & "|" & str终止日期
            Call Record_Add(rsHead, strFields, strValues)
            
            dbl费用总额 = Val(Split(arrRow(lngRow), "|")(col住院_费用总额))
            dbl大额支付 = Val(Split(arrRow(lngRow), "|")(col住院_大额支付))
            dbl统筹支付 = Val(Split(arrRow(lngRow), "|")(col住院_统筹支付))
            If lngRow = lngRows Then
                dbl统筹封顶后医保内 = Val(Split(arrRow(0), "|")(11))
            Else
                dbl统筹封顶后医保内 = 0
            End If
            dbl个人自费 = dbl费用总额 - Val(Split(arrRow(lngRow), "|")(col住院_医保内))
            dbl个人自付 = Val(Split(arrRow(lngRow), "|")(col住院_医保内)) - dbl大额支付 - dbl统筹支付
            dbl首先自付 = Val(Split(arrRow(lngRow), "|")(col住院_首先自付))
            
            strFields = "费用总额|统筹支付|大额支付|个人自付|首先自付|个人自费|统筹封顶后医保内金额|起始日期|终止日期|经办日期"
            strValues = dbl费用总额 & "|" & dbl统筹支付 & "|" & dbl大额支付 & "|" & dbl个人自付 & "|" & dbl首先自付 & _
                "|" & dbl个人自费 & "|" & dbl统筹封顶后医保内 & "|" & str起始日期 & "|" & str终止日期 & "|" & str经办日期
            Call Record_Add(rsDeal, strFields, strValues)
        Next
    End If
    
    If Not bln保存 Then
        '显示详细的手册消费信息
        Call frm手册内容.ShowBalance(rsHead, rsDeal, bln门诊)
        SaveDeal = True
        Exit Function
    End If
    
    '获取本次结算基本信息
    If bln门诊 Then
        Call DebugTool("获取本次门诊就诊信息(zl9Insure\SaveDeal")
        Call WriteBusinessLOG("获取本次门诊就诊信息(zl9Insure\SaveDeal", "", "")
        gstrSQL = "" & _
            " SELECT A.社保证号,A.业务类型,A.参保类别,A.公务员,A.公务员待遇,A.病种标识," & _
            "        TO_CHAR(A.特殊病截止日期,'yyyyMMdd') AS 特殊病截止日期," & _
            "        0 AS 入院类型,0 AS 出院类型," & _
            "        to_Char(sysdate,'yyyy-MM-dd hh24:mi:ss') As 入院日期," & _
            "        to_Char(sysdate,'yyyy-MM-dd hh24:mi:ss') AS 出院日期" & _
            " From 保险帐户 A,门诊交易信息 C" & _
            " WHERE A.卡号=C.卡号 And C.交易流水号='" & gComInfo_北京.交易流水号 & "'"
        If rsTemp.State = 1 Then rsTemp.Close
        Call SQLTest(App.Title, "ZL9INSURE\SaveDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    Else
        Call DebugTool("获取本次住院基本信息(zl9Insure\SaveDeal")
        Call WriteBusinessLOG("获取本次住院基本信息(zl9Insure\SaveDeal", "", "")
        gstrSQL = "" & _
            " SELECT A.社保证号,A.业务类型,A.参保类别,A.公务员,A.公务员待遇,A.病种标识," & _
            "        TO_CHAR(A.特殊病截止日期,'yyyyMMdd') AS 特殊病截止日期," & _
            "        B.入院类型,B.入院方式,C.出院类型," & _
            "        to_Char(C.入院日期,'yyyy-MM-dd hh24:mi:ss') As 入院日期," & _
            "        to_Char(C.出院日期,'yyyy-MM-dd hh24:mi:ss') AS 出院日期" & _
            " From 保险帐户 A,入院信息 B,住院交易信息 C" & _
            " WHERE A.卡号=B.卡号 And B.卡号=C.卡号 And C.交易流水号='" & gComInfo_北京.交易流水号 & "'"
        If rsTemp.State = 1 Then rsTemp.Close
        Call SQLTest(App.Title, "ZL9INSURE\SaveDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    End If
    
    '参数如下
    '卡号,医疗机构,医疗类别,入院类型,入院日期,出院类型,出院日期,统筹支付
    '大额支付,个人自付,个人自费,统筹封顶后医保内,交易流水号,清除历史记录
    With rsDeal
        If .RecordCount <> 0 Then .MoveFirst
        
        Do While Not .EOF
            gstrSQL = "ZL_手册消费记录_INSERT(" & _
                      "'" & gComInfo_北京.卡号 & "','" & str医院名称 & "','" & rsTemp!业务类型 & "'," & _
                      "'" & rsTemp!入院类型 & "',TO_DATE('" & !起始日期 & "','yyyy-MM-dd hh24:mi:ss')," & _
                      "'" & IIf(.AbsolutePosition <> .RecordCount, "2", "0") & "',TO_DATE('" & !终止日期 & "','yyyy-MM-dd hh24:mi:ss')," & _
                      "" & Val(!统筹支付) & "," & Val(!大额支付) & "," & Val(!个人自付) & "," & Val(!个人自费) & "," & _
                      "" & Val(!统筹封顶后医保内金额) & ",'" & gComInfo_北京.交易流水号 & "',0)"
            gcnBJYB.Execute gstrSQL, , adCmdStoredProc
            .MoveNext
        Loop
    End With
    
    SaveDeal = True
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\SaveDeal)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\SaveDeal)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, mstrSplit)
    arrValues = Split(strValues, mstrSplit)
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, mstrSplit)
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub










'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'以下是基础接口
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
Public Function CheckTradeName(ByVal lng收费细目ID As Long, ByVal str医保编码 As String) As Boolean
    Dim strUser As String
    Dim rsTemp As New ADODB.Recordset
    '检查商品名或别名是否在医保规定的商品名或别名中，如果是才允许设置对照
    On Error GoTo errHand
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=[1] And 参数名='医保用户名'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取中间库的用户名", TYPE_北京)
    strUser = rsTemp!参数值
    
    gstrSQL = " Select 1 From " & strUser & ".药品别名 " & _
              " Where 编码='" & str医保编码 & "'" & _
              " And 名称 IN ( " & _
              "     Select 名称 From 药品别名 Where 药名ID= " & _
              "         (Select 药名ID From 药品目录 Where 药品ID=[1]))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", lng收费细目ID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "医保中心下发的药品别名中没有此项目的商品名和别名，请重新选择！", vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckTradeName = True
errHand:
    Exit Function
End Function

Private Sub GetSequence_北京(ByVal lng病人ID As Long)
    Dim strText As String
    Dim str就诊时间 As String
    Dim str当前时间 As String
    Dim strSequence As String
    Dim rsTemp As New ADODB.Recordset
    Dim intDO As Integer, intCOUNT As Integer, intPos As Integer
    '将当前时间、就诊时间进行处理，转换为唯一的流水号标识
    '编程思路：将年、月、日、时、分、秒都转换为一个字母的形式表示，因为一共只有20位，其前有8位医院编码做前导字符
    intCOUNT = 6
    intPos = 1
    str当前时间 = Format(zlDatabase.Currentdate, "yyMMddHHmmss")
    
    Call DebugTool("准备产生顺序号(zl9Insure\GetSequenct_北京)")
    Call WriteBusinessLOG("准备产生顺序号(zl9Insure\GetSequenct_北京)", "", "")
    '获取该病人的就诊时间
    gstrSQL = "Select to_char(就诊时间,'yyyy-MM-dd hh24:mi:ss') As 就诊时间 From 保险帐户" & _
        " Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取该病人的就诊时间", TYPE_北京, lng病人ID)
    '肯定不会为空
    str就诊时间 = Format(rsTemp!就诊时间, "yyMMddHHmmss")
    
    For intDO = 1 To intCOUNT
        strText = Mid(str就诊时间, intPos, 2)
        intPos = intPos + 2
        strSequence = strSequence & Chr(asc("0") + Val(strText))
    Next
    intPos = 1
    For intDO = 1 To intCOUNT
        strText = Mid(str当前时间, intPos, 2)
        intPos = intPos + 2
        strSequence = strSequence & Chr(asc("0") + Val(strText))
    Next
    gComInfo_北京.交易流水号 = gComInfo_北京.医院编码 & strSequence
    Call DebugTool("顺序号:" & gComInfo_北京.交易流水号 & "(zl9Insure\GetSequenct_北京)")
    Call WriteBusinessLOG("顺序号:" & gComInfo_北京.交易流水号 & "(zl9Insure\GetSequenct_北京)", "", "")
End Sub

Public Function 待遇审查_北京(ByVal StrInput As String, Optional ByVal bln卡号 As Boolean = True) As Boolean
    待遇审查_北京 = CheckBlockage(StrInput, bln卡号)
End Function

Public Function 身份标识_北京(ByVal bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim strReturn As String
    strReturn = frmIdentify北京.GetIdentify(bytType, lng病人ID)
    身份标识_北京 = strReturn
End Function

Public Function 医保设置_北京() As Boolean
    医保设置_北京 = frmSet北京.参数设置()
End Function

Public Function 医保初始化_北京(Optional ByVal blnTest As Boolean = False) As Boolean
    Dim strUser As String, strServer As String, strPass As String
    Dim rsTemp As New ADODB.Recordset
    Dim strReturn As String
    On Error GoTo errHand
    
    If mblnInit = False Then
        '读出连接医保服务器的配置
        gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", TYPE_北京)
        
        Do Until rsTemp.EOF
            Select Case rsTemp("参数名")
                Case "医保用户名"
                    strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                Case "医保服务器"
                    strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                Case "医保用户密码"
                    strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                Case "入参目录"
                    gComInfo_北京.入参目录 = Nvl(rsTemp!参数值)
                Case "出参目录"
                    gComInfo_北京.出参目录 = Nvl(rsTemp!参数值)
            End Select
            rsTemp.MoveNext
        Loop
        
        If OraDataOpen(gcnBJYB, strServer, strUser, strPass, False) = False Then
            MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If Not blnTest Then
            If Nvl(gComInfo_北京.入参目录) = "" Then
                MsgBox "请为该医保设置入参文件目录！", vbInformation, gstrSysName
                Exit Function
            End If
            If Nvl(gComInfo_北京.出参目录) = "" Then
                MsgBox "请为该医保设置出参文件目录！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '取医院编码
        gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_北京)
        gComInfo_北京.医院编码 = Nvl(rsTemp!医院编码)
        
        '启动政策机服务
        If Not 调用接口_北京(启动政策机服务, strReturn) Then Exit Function
    End If
    
    mblnInit = True
    医保初始化_北京 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保终止_北京() As Boolean
    Dim strReturn As String
    Call 调用接口_北京(停止政策机服务, strReturn)
End Function

Private Function 调用接口_北京(ByVal int功能 As 接口功能, ByRef str参数 As String) As Boolean
    '----------------------------------------------------------------
    '功能描述   ：调用接口函数
    '编写人     ：朱玉宝
    '编写日期   ：2004-07-03
    '参数说明   ：str参数=入参，其值可以为空，如果函数存在多个入参（一般只有两个入参），以"||"分隔，目前无多个入参的函数
    '                     出参，其值可以为空，如果函数存在出参，则返回出参的值，以"||分隔，目前仅获取版本接口要返回多个出参
    '----------------------------------------------------------------
    Dim lngReturn As Long           '函数返回值
    Dim strFunction As String       '当前执行函数名
    Dim strInPara As String
    Dim strOutPara1 As String * 2000      '出参
    Dim strOutPara2 As String * 2000      '出参
    Dim strErrMsg As String * 255
    Dim arrPara                     '入参数组
    On Error GoTo errHand
    
    strInPara = str参数
    Select Case int功能
    Case 接口功能.启动政策机服务
        strFunction = "StartPolicy[INTERFACE]"
        '如果启动失败了，需要调用停止政策机服务
        lngReturn = BJ_StartPolicy(0)
        Call WriteBusinessLOG(strFunction, "", "")
        If lngReturn <> 0 Then
            Call BJ_StopPolicy
            Call WriteBusinessLOG("StopPolicy[INTERFACE]", "", "")
        End If
        lngReturn = 0
    Case 接口功能.停止政策机服务
        strFunction = "StopPolicy[INTERFACE]"
        Call BJ_StopPolicy
        Call WriteBusinessLOG(strFunction, "", "")
    Case 接口功能.获取持卡病人信息
        strFunction = "GetPersonCommInfo[INTERFACE]"
        lngReturn = BJ_GetPersonCommInfo(str参数)
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.获取持卡病人待遇信息
        strFunction = "GetSumInfo[INTERFACE]"
        lngReturn = BJ_Get_SumInfo(str参数)
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.获取手册病人待遇信息
        strFunction = "GetSumInfo2[INTERFACE]"
        lngReturn = BJ_Get_SumInfo2(str参数, strOutPara1)
        str参数 = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.获取特殊病信息
        strFunction = "GetSpecInfo[INTERFACE]"
        lngReturn = BJ_Get_SpecInfo(str参数, strOutPara1)
        str参数 = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.帐户支付
        strFunction = "RegAccount[INTERFACE]"
        lngReturn = BJ_Reg_Account(str参数, strOutPara1)
        str参数 = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.交易确认
        strFunction = "Reg[INTERFACE]"
        lngReturn = BJ_Reg(str参数, strOutPara1)
        str参数 = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.接口版本信息
        strFunction = "GetVer[INTERFACE]"
        lngReturn = BJ_Get_Ver(strOutPara1, strOutPara2)
        str参数 = TrimStr(strOutPara1) & "||" & TrimStr(strOutPara2)
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.费用分解_通用1
        strFunction = "Divide[INTERFACE]"
        lngReturn = BJ_Divide(str参数, strOutPara1)
        str参数 = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.费用分解_通用2
        strFunction = "Divide2[INTERFACE]"
        lngReturn = BJ_Divide2(str参数, strOutPara1)
        str参数 = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.费用分解_普通门诊
        strFunction = "Poli_Divide[INTERFACE]"
        lngReturn = BJ_Poli_Divide(str参数)
        str参数 = ""
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.费用分解_特殊门诊1
        strFunction = "Spec_Divide[INTERFACE]"
        lngReturn = BJ_Spec_Divide(str参数)
        str参数 = ""
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.费用分解_特殊门诊2
        strFunction = "Spec_Divide2[INTERFACE]"
        lngReturn = BJ_Spec_Divide2(str参数, strOutPara1)
        str参数 = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.费用分解_家庭病床1
        strFunction = "Home_Divide1[INTERFACE]"
        lngReturn = BJ_Home_Divide(str参数)
        str参数 = ""
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.费用分解_家庭病床2
        strFunction = "Home_Divide2[INTERFACE]"
        lngReturn = BJ_Home_Divide2(str参数, strOutPara1)
        str参数 = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.费用分解_住院1
        strFunction = "Hosp_Divide1[INTERFACE]"
        lngReturn = BJ_Hosp_Divide(str参数)
        str参数 = ""
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.费用分解_住院2
        strFunction = "Hosp_Divide2[INTERFACE]"
        lngReturn = BJ_Hosp_Divide2(str参数)
        str参数 = ""
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    Case 接口功能.费用分解_住院3
        strFunction = "Hosp_Divide3[INTERFACE]"
        lngReturn = BJ_Hosp_Divide3(str参数)
        str参数 = ""
        Call WriteBusinessLOG(strFunction, strInPara, str参数)
    End Select
    
    '表示出错了
    If lngReturn <> 0 Then
        lngReturn = BJ_Get_ErrInfo(strErrMsg)
        If lngReturn <> 0 Then
            strErrMsg = "执行函数[" & strFunction & "]时发生未知错误！"
        Else
            strErrMsg = "执行函数[" & strFunction & "]时发生以下错误：" & vbCrLf & strErrMsg
        End If
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    str参数 = TrimStr(str参数)
    调用接口_北京 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub WriteBusinessLOG(ByVal strFunc As String, ByVal StrInput As String, ByVal strOutput As String)
    '以下变量用于记录调用接口的入参
    Const strFile As String = "C:\Business_"
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    
    If gintDebug = -1 Then gintDebug = Val(GetSetting("ZLSOFT", "医保", "调试", 0))
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If gintDebug = 0 Then Exit Sub
    strFileName = strFile & Format(Date, "yyyyMMdd") & ".LOG"
    
    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (String(50, "-"))
    objStream.WriteLine ("执行时间:" & strDate)
    objStream.WriteLine ("函数名:" & strFunc)
    objStream.WriteLine ("  入参:" & StrInput)
    objStream.WriteLine ("  出参:" & strOutput)
    objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objStream = Nothing
End Sub

Public Function 门诊虚拟结算_北京(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    Dim int项目类别 As Integer  '0-药品 1-诊疗项目 2-服务设施
    Dim lng病人ID As Long
    Dim str就诊日期 As String, str费用发生日期 As String
    Dim str医保编码 As String, strHIS项目名称 As String, str公务员标识 As String, strReturn As String
    Dim str待遇 As String, str费用分解待遇 As String
    Dim strFields As String, strValues As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim arr结算方式
    Const int医保基金 As Integer = 5
    Const int大病补助 As Integer = 7
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    '补充说明：目前仅支持特殊病门诊
    On Error GoTo errHandle
    lng病人ID = rs明细!病人ID
    str就诊日期 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    str费用发生日期 = Format(zlDatabase.Currentdate(), "yyyyMMdd")
    
    '先获取本次交易流水号
    Call DebugTool("产生流水号(zl9Insure\门诊虚拟结算)")
    Call WriteBusinessLOG("产生流水号(zl9Insure\门诊虚拟结算)", "", "")
    Call GetSequence_北京(lng病人ID)
    
    '获取参保人公务员标识
    Call DebugTool("获取参保人公务员标识(zl9Insure\门诊虚拟结算)")
    Call WriteBusinessLOG("获取参保人公务员标识(zl9Insure\门诊虚拟结算)", "", "")
    gstrSQL = "Select 公务员 From 保险帐户 Where 卡号='" & gComInfo_北京.卡号 & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\门诊虚拟结算_北京", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    str公务员标识 = rsTemp!公务员
    
    '获取此刻的待遇信息（费用分解函数的入参）
    Call DebugTool("获取此刻的历史消费记录（费用分解函数的入参）(zl9Insure\门诊虚拟结算)")
    Call WriteBusinessLOG("获取此刻的历史消费记录（费用分解函数的入参）(zl9Insure\门诊虚拟结算)", "", "")
    Set rsDeal = GetDeal(lng病人ID, str就诊日期)
    '根据录入的历史消费记录，产生入参文件
    Call DebugTool("根据录入的历史消费记录，产生入参文件(zl9Insure\门诊虚拟结算)")
    Call WriteBusinessLOG("根据录入的历史消费记录，产生入参文件(zl9Insure\门诊虚拟结算)", "", "")
    If Not MakeFile_Center(rsDeal, 接口功能.获取手册病人待遇信息) Then Exit Function
    '得到接口返回的待遇串
    strReturn = gComInfo_北京.卡号 & mstrSplit & str公务员标识
    Call DebugTool("调用获取待遇信息接口(zl9Insure\门诊虚拟结算)")
    Call WriteBusinessLOG("调用获取待遇信息接口(zl9Insure\门诊虚拟结算)", "", "")
    If Not 调用接口_北京(接口功能.获取手册病人待遇信息, strReturn) Then Exit Function   'strReturn的内容将用于费用分解，以下不能进行赋值
    str待遇 = strReturn
    
    '将费用明细生成明细文件（也是费用分解函数的入参）
'    ----传入文件说明
'    序号    数据项  类型    最大长度    说明
'    1   项目序号    C   9   顺序号
'    2   处方号  C   20  参见标准AKC220，可空
'    3   项目代码    C   20  药品、诊疗项目或服务设施编码
'    4   项目名称    C   100 本医院项目名称
'    5   项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'    6   单价    N   10,4    AKC225
'    7   数量    N   8,2 AKC226
'    8   费用总金额  N   10,4    实际结算金额
'    9   费用发生日期    D   8   YYYYMMDD
    strFields = "项目序号," & adLongVarChar & ",9" & mstrSplit & _
                "处方号," & adLongVarChar & ",20" & mstrSplit & _
                "项目代码," & adLongVarChar & ",20" & mstrSplit & _
                "项目名称," & adLongVarChar & ",100" & mstrSplit & _
                "项目类别," & adLongVarChar & ",3" & mstrSplit & _
                "单价," & adDouble & ",18" & mstrSplit & _
                "数量," & adDouble & ",18" & mstrSplit & _
                "费用总金额," & adDouble & ",18" & mstrSplit & _
                "费用发生日期," & adLongVarChar & ",20"
    Call Record_Init(rsCharge, strFields)
    
    '得到本次录入的费用明细的相关医保信息（编码以w01打头，表示服务设施类项目）
    Call DebugTool("产生费用明细记录集(zl9Insure\门诊虚拟结算)")
    Call WriteBusinessLOG("产生费用明细记录集(zl9Insure\门诊虚拟结算)", "", "")
    strFields = "项目序号|处方号|项目代码|项目名称|项目类别|单价|数量|费用总金额|费用发生日期"
    If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
    Do Until rs明细.EOF
        '将药品的通用名称或非药品项目的名称取出来
        gstrSQL = "Select A.项目编码 As 医保编码,C.通用名称 AS 项目名称 " & _
            " From 保险支付项目 A,药品目录 B,药品信息 C" & _
            " Where A.险类(+)=[1] And A.收费细目ID(+)=B.药品ID And B.药名ID=C.药名ID " & _
            " AND B.药品ID=[2]" & _
            " UNION " & _
            " Select A.项目编码 As 医保编码,B.名称 AS 项目名称" & _
            " From 保险支付项目 A,收费细目 B" & _
            " Where A.险类(+)=[1] AND B.ID=[2]" & _
            " And A.收费细目ID(+)=B.ID AND B.类别 Not In ('5','6','7')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", TYPE_北京, CLng(rs明细!收费细目ID))
        If rsTemp.EOF Then
            MsgBox "还有项目没有和医保项目设置对照关系！[保险项目]", vbInformation, gstrSysName
            Exit Function
        End If
        str医保编码 = Nvl(rsTemp!医保编码, 0)
        strHIS项目名称 = rsTemp!项目名称
        
        If InStr(1, "5,6,7", rs明细!收费类别) <> 0 Then
            int项目类别 = 0 '0-药品
        Else
            int项目类别 = (IIf(str医保编码 Like "w01*", 2, 1)) '1-诊疗项目 2-服务设施
        End If
        
        '产生费用明细记录集,以供产生入参文件
        strValues = rs明细.AbsolutePosition & mstrSplit & "" & mstrSplit & _
            str医保编码 & mstrSplit & strHIS项目名称 & mstrSplit & _
            int项目类别 & mstrSplit & Format(rs明细!单价, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            Format(rs明细!数量, "#####0.00;-#####0.00;0;") & mstrSplit & Format(rs明细!实收金额, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            str费用发生日期
        Call Record_Add(rsCharge, strFields, strValues)
        
        rs明细.MoveNext
    Loop
    
    '产生费用明细文件
    Call DebugTool("根据费用明细，产生入参文件(zl9Insure\门诊虚拟结算)")
    Call WriteBusinessLOG("根据费用明细，产生入参文件(zl9Insure\门诊虚拟结算)", "", "")
    If Not MakeFile_Center(rsCharge, 费用分解_特殊门诊1) Then Exit Function
    
    '将返回的待遇信息转换为费用分解所需的格式
    str费用分解待遇 = TransationSpec(str待遇)
    
    '调用费用分解函数（strReturn中是待遇信息）
    Call DebugTool("调用费用明细分解函数(zl9Insure\门诊虚拟结算)")
    Call WriteBusinessLOG("调用费用明细分解函数(zl9Insure\门诊虚拟结算)", "", "")
    strReturn = str费用分解待遇
    If Not 调用接口_北京(费用分解_特殊门诊1, strReturn) Then Exit Function
    strReturn = AnalyFile_特殊门诊1(True)
    If strReturn = "" Then Exit Function
    
    If Not SaveDeal(True, False) Then Exit Function
    
    '组织成结算方式串并返回
    Call DebugTool("分析费用分解返回的汇总数据，将结算信息返回给主调程序(zl9Insure\门诊虚拟结算)")
    Call WriteBusinessLOG("分析费用分解返回的汇总数据，将结算信息返回给主调程序(zl9Insure\门诊虚拟结算)", "", "")
    arr结算方式 = Split(strReturn, mstrSplit)
    str结算方式 = mstrSplit & "统筹支付;" & arr结算方式(int医保基金) & ";0"
    str结算方式 = str结算方式 & mstrSplit & "大额支付;" & arr结算方式(int大病补助) & ";0"
    str结算方式 = Mid(str结算方式, 2)
    
    门诊虚拟结算_北京 = True
    Exit Function
errHandle:
    Call DebugTool("(zl9INSURE\门诊虚拟结算_北京)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\门诊虚拟结算_北京)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_北京(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, _
    Optional ByVal bln住院 As Boolean = False) As Boolean
    Dim lng病人ID As Long
    Dim blnTrans As Boolean
    Dim str就诊日期 As String, str费用发生日期 As String
    Dim str公务员标识 As String, strReturn As String
    Dim str待遇 As String, str费用分解待遇 As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim arr结算方式
    Dim dbl费用总额 As Double, dbl医保基金 As Double, dbl大病补助 As Double, dbl现金 As Double
    Dim dbl医保内 As Double, dbl特殊病个人自付 As Double, dbl特殊病医保外 As Double, dbl首先自付 As Double, dbl统筹封顶后医保内 As Double
    Const int费用总额 As Integer = 2
    Const int医保内 As Integer = 3
    Const int医保基金 As Integer = 5
    Const int大病补助 As Integer = 7
    Const int特殊病个人自付 As Integer = 9
    Const int特殊病医保外 As Integer = 10
    Const int首先自付 As Integer = 11
    Const int统筹封顶后医保内 As Integer = 12
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
        '此时所有收费细目必然有对应的医保编码
    '由于进行了门诊虚拟结算，因此所有的入参文件已经产生，此过程不再需要费用明细，直接组织一下即可
    On Error GoTo errHandle
    '提取病人ID
    gstrSQL = "Select 病人ID From 门诊费用记录 Where 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人ID", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    str就诊日期 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    str费用发生日期 = Format(zlDatabase.Currentdate(), "yyyyMMdd")
    
    '获取参保人公务员标识
    Call DebugTool("获取参保人公务员标识(zl9Insure\门诊结算)")
    Call WriteBusinessLOG("获取参保人公务员标识(zl9Insure\门诊结算)", "", "")
    gstrSQL = "Select 卡号,公务员 From 保险帐户 Where 病人ID=" & lng病人ID
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\门诊结算_北京", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    gComInfo_北京.卡号 = rsTemp!卡号
    str公务员标识 = rsTemp!公务员
    
    '获取此刻的待遇信息（费用分解函数的入参）
    Call DebugTool("获取此刻的历史消费记录（费用分解函数的入参）(zl9Insure\门诊结算)")
    Call WriteBusinessLOG("获取此刻的历史消费记录（费用分解函数的入参）(zl9Insure\门诊结算)", "", "")
    Set rsDeal = GetDeal(lng病人ID, str就诊日期)
    '得到接口返回的待遇串（由于预结算时，已产生获取待遇的入参文件，此处不再产生此文件，直接调用接口）
    strReturn = gComInfo_北京.卡号 & mstrSplit & str公务员标识
    Call DebugTool("调用获取待遇信息接口(zl9Insure\门诊结算)")
    Call WriteBusinessLOG("调用获取待遇信息接口(zl9Insure\门诊结算)", "", "")
    If Not 调用接口_北京(接口功能.获取手册病人待遇信息, strReturn) Then  'strReturn的内容将用于费用分解，以下不能进行赋值
        Exit Function
    End If
    str待遇 = strReturn
    '将返回的待遇信息转换为费用分解所需的格式
    str费用分解待遇 = TransationSpec(str待遇)
    '调用费用分解函数（strReturn中是待遇信息）
    Call DebugTool("调用费用明细分解函数(zl9Insure\门诊结算)")
    Call WriteBusinessLOG("调用费用明细分解函数(zl9Insure\门诊结算)", "", "")
    strReturn = str费用分解待遇
    '调用费用分解函数（由于预结算时，已产生费用分解函数的入参文件，此处不再产生此文件，直接调用接口）
    If Not 调用接口_北京(费用分解_特殊门诊1, strReturn) Then
        Exit Function
    End If
    
    '组织成结算方式串并返回
    Call DebugTool("分析费用分解返回的汇总数据，以便保存到保险结算记录中(zl9Insure\门诊结算)")
    Call WriteBusinessLOG("分析费用分解返回的汇总数据，以便保存到保险结算记录中(zl9Insure\门诊结算)", "", "")
    '先按预结算模式，取得汇总行数据
    strReturn = AnalyFile_特殊门诊1(True)
    If strReturn <> "" Then
        arr结算方式 = Split(strReturn, mstrSplit)
        dbl费用总额 = Val(arr结算方式(int费用总额))
        dbl医保基金 = Val(arr结算方式(int医保基金))
        dbl大病补助 = Val(arr结算方式(int大病补助))
        
        dbl医保内 = Val(arr结算方式(int医保内))
        dbl特殊病个人自付 = Val(arr结算方式(int特殊病个人自付))
        dbl特殊病医保外 = Val(arr结算方式(int特殊病医保外))
        dbl首先自付 = Val(arr结算方式(int首先自付))
        dbl统筹封顶后医保内 = Val(arr结算方式(int统筹封顶后医保内))
    End If
    dbl现金 = dbl费用总额 - dbl医保基金 - dbl大病补助
    
    '*****************保存中间库，以备上传*****************
    '**保存本次的交易待遇信息**
    gcnBJYB.BeginTrans
    blnTrans = True
    Call DebugTool("保存本次的交易待遇信息(zl9Insure\门诊结算)")
    Call WriteBusinessLOG("保存本次的交易待遇信息(zl9Insure\门诊结算)", "", "")
    If Not SaveBusinessDeal(str待遇) Then
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '**保存交易版本信息**
    Call DebugTool("保存本次的交易版本信息(zl9Insure\门诊结算)")
    Call WriteBusinessLOG("保存本次的交易版本信息(zl9Insure\门诊结算)", "", "")
    If Not SaveBusinessVersion(True) Then
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '**保存门诊交易信息及门诊费用明细**
    Call DebugTool("保存门诊交易信息及门诊费用明细(zl9Insure\门诊结算)")
    Call WriteBusinessLOG("保存门诊交易信息及门诊费用明细(zl9Insure\门诊结算)", "", "")
    strReturn = AnalyFile_特殊门诊1(False, lng结帐ID, bln住院)
    If strReturn = "" Then
        '保存失败
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    If Not SaveDeal(True, True) Then Exit Function
    
    '**保存保险结算记录**
    '进入统筹金额=公务员补助基金;统筹报销金额=统筹基金;大病自付金额=大病基金
    '支付顺序号=交易流水号,备注=业务类型
    Call DebugTool("保存保险结算记录(zl9Insure\门诊结算)")
    Call WriteBusinessLOG("保存保险结算记录(zl9Insure\门诊结算)", "", "")
    gstrSQL = "zl_保险结算记录_insert(" & IIf(bln住院, 2, 1) & "," & lng结帐ID & "," & TYPE_北京 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        dbl费用总额 & "," & dbl现金 & ",0,0," & dbl医保基金 & "," & dbl大病补助 & "," & _
        0 & ",0,'" & gComInfo_北京.交易流水号 & "',null,null,'" & gComInfo_北京.业务类型 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
    
    gcnBJYB.CommitTrans
    blnTrans = False
    门诊结算_北京 = True
    Exit Function
errHandle:
    Call DebugTool("(zl9Insure\门诊结算_北京)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\门诊结算_北京)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then gcnBJYB.RollbackTrans
End Function

Public Function 门诊结算冲销_北京(ByVal lng结帐ID As Long, ByVal cur个人帐户 As Currency, ByVal lng病人ID As Long) As Boolean
    On Error GoTo errHand
    Dim blnTrans As Boolean
    Dim lng冲销ID As Long
    Dim str卡号 As String, str原交易流水号 As String
    Dim str结算时间 As String, str退费时间 As String
    Dim rsTemp As New ADODB.Recordset
    
    Call GetSequence_北京(lng病人ID)
    
    '取出原始结算记录的结算时间
    Call DebugTool("取出原始结算记录的结算时间(zl9Insure\门诊结算冲销)")
    Call WriteBusinessLOG("取出原始结算记录的结算时间(zl9Insure\门诊结算冲销)", "", "")
    gstrSQL = "Select 卡号,交易流水号,结算时间 From " & GetUser & ".交易待遇信息 " & _
            " Where 交易流水号=(Select 支付顺序号 From 保险结算记录 Where 性质=1 AND 记录ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出原始结算记录的结算时间", lng结帐ID)
    str卡号 = rsTemp!卡号
    str原交易流水号 = rsTemp!交易流水号
    str结算时间 = Format(rsTemp!结算时间, "yyyy-MM-dd HH:mm:ss")
    str退费时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    '检查，如果在该交易后，还发生了其它交易，则不允许进行冲销（把已退费的记录排开）
    Call DebugTool("检查是否从最后一笔开始退费(zl9Insure\门诊结算冲销)")
    Call WriteBusinessLOG("检查是否从最后一笔开始退费(zl9Insure\门诊结算冲销)", "", "")
    gstrSQL = "Select Count(*) AS Records From 交易待遇信息 A,保险帐户 B" & _
            " Where A.卡号=B.卡号 And B.病人ID=" & lng病人ID & _
            " And 结算时间>to_Date('" & str结算时间 & "','yyyy-MM-dd hh24:mi:ss')" & _
            " And 交易流水号 Not In (Select 原交易流水号 From 退费信息 Where 卡号='" & str卡号 & "')"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\门诊结算_北京", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsTemp!Records > 0 Then
        MsgBox "医保接口不允许从中间开始退费，您只能从最后一笔业务开始退费！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '取冲销记录的结帐ID，单据号
    Call DebugTool("取冲销ID(zl9Insure\门诊结算冲销)")
    Call WriteBusinessLOG("取冲销ID(zl9Insure\门诊结算冲销)", "", "")
    gstrSQL = "select distinct A.结帐ID,A.NO from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!结帐ID
    
    '提取原始的保险结算记录
    Call DebugTool("保存保险结算记录(zl9Insure\门诊结算冲销)")
    Call WriteBusinessLOG("保存保险结算记录(zl9Insure\门诊结算冲销)", "", "")
    gstrSQL = "Select * From 保险结算记录 Where 性质=1 AND 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取原始的保险结算记录", lng结帐ID)
    
    '保存保险结算记录
    Call DebugTool("保存保险结算记录(zl9Insure\门诊结算冲销)")
    Call WriteBusinessLOG("保存保险结算记录(zl9Insure\门诊结算冲销)", "", "")
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_北京 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & ",0,0," & -1 * Nvl(rsTemp!统筹报销金额, 0) & _
        "," & -1 * Nvl(rsTemp!大病自付金额, 0) & "," & 0 & ",0,'" & gComInfo_北京.交易流水号 & "',null,null,'" & Nvl(rsTemp!备注) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
    
    blnTrans = True
    gcnBJYB.BeginTrans
    '产生一条退费记录即可，参数如下
    '红冲交易流水号,卡号,原交易流水号,原交易日期,退费日期,操作员姓名,上传
    Call DebugTool("产生红冲退费记录(zl9Insure\门诊结算冲销)")
    Call WriteBusinessLOG("产生红冲退费记录(zl9Insure\门诊结算冲销)", "", "")
    gstrSQL = "ZL_退费信息_INSERT(" & _
            "'" & gComInfo_北京.交易流水号 & "','" & str卡号 & "','" & str原交易流水号 & "'," & _
            "to_Date('" & str结算时间 & "','yyyy-MM-dd hh24:mi:ss')," & _
            "to_Date('" & str退费时间 & "','yyyy-MM-dd hh24:mi:ss')," & _
            "'" & UserInfo.姓名 & "',0)"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
'    **住院结算冲销时，要删除上次的历史就诊记录**
    gstrSQL = "ZL_手册消费记录_DELETE('" & str原交易流水号 & "')"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc

    gcnBJYB.CommitTrans
    blnTrans = False
    门诊结算冲销_北京 = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\门诊结算冲销)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\门诊结算冲销)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnBJYB.RollbackTrans
End Function

Public Function 入院登记_北京(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    On Error GoTo errHand
    Dim str入院日期 As String
    Dim rsTemp As New ADODB.Recordset
    
    Call DebugTool("获取入院登记流水号(zl9Insure\入院登记)")
    Call WriteBusinessLOG("获取入院登记流水号(zl9Insure\入院登记)", "", "")
    Call GetSequence_北京(lng病人ID)
    '因住院流水号只有18位，而流水号是按20位产生，因此去掉年月两个字符
    gComInfo_北京.交易流水号 = Mid(gComInfo_北京.交易流水号, 1, 8) & Mid(gComInfo_北京.交易流水号, 11)
    
    gstrSQL = "Select 入院日期 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取入院日期", lng病人ID, lng主页ID)
    str入院日期 = Format(rsTemp!入院日期, "yyyy-MM-dd HH:mm:ss")
    
    '获取此次入院的方式、类型
    Call DebugTool("获取该病人的入院类型、入院方式及入院日期(zl9Insure\入院登记)")
    Call WriteBusinessLOG("获取该病人的入院类型、入院方式及入院日期(zl9Insure\入院登记)", "", "")
    gstrSQL = "Select 业务类型,卡号,入院类别,入院方式,入院日期 From 保险帐户 Where 卡号='" & gComInfo_北京.卡号 & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\门诊结算_北京", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    
    If Not 门诊特殊病(lng病人ID) Then
        '产生入院记录，参数如入
        '病人ID,主页ID,卡号,入院登记号,入院类型,入院方式,入院日期,上传
        Call DebugTool("保存入院信息(zl9Insure\入院登记)")
        Call WriteBusinessLOG("保存入院信息(zl9Insure\入院登记)", "", "")
        gstrSQL = "ZL_入院信息_INSERT(" & _
            "" & lng病人ID & "," & lng主页ID & ",'" & rsTemp!卡号 & "','" & gComInfo_北京.交易流水号 & "'," & _
            "" & rsTemp!入院类别 & "," & rsTemp!入院方式 & ",to_Date('" & str入院日期 & "','yyyy-MM-dd hh24:mi:ss')" & ",0)"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    '改变病人状态
    Call DebugTool("修改病人的当前状态(zl9Insure\入院登记)")
    Call WriteBusinessLOG("修改病人的当前状态(zl9Insure\入院登记)", "", "")
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_北京 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    入院登记_北京 = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\入院登记)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\入院登记)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_北京(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    On Error GoTo errHand
    Dim str入院登记号 As String
    Dim rsTemp As New ADODB.Recordset
    
    '如果该记录未上传,则允许撤销入院（住院病人只有出院后，才会上传）
    Call DebugTool("检查：已上传则不允许撤销入院(zl9Insure\入院登记撤销)")
    Call WriteBusinessLOG("检查：已上传则不允许撤销入院(zl9Insure\入院登记撤销)", "", "")
    gstrSQL = "Select Nvl(上传,0) 上传,入院登记号 From 入院信息 Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\门诊结算_北京", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsTemp!上传 = 1 Then
        MsgBox "该病人的所有记录都已上传到医保中心，不允许进行撤销入院！", vbInformation, gstrSysName
        Exit Function
    End If
    str入院登记号 = Nvl(rsTemp!入院登记号)
    
    '改变病人状态
    Call DebugTool("修改状态为不在院(zl9Insure\入院登记撤销)")
    Call WriteBusinessLOG("修改状态为不在院(zl9Insure\入院登记撤销)", "", "")
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_北京 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    If Not 门诊特殊病(lng病人ID) Then
        '清除入院记录（同时清除出院诊断信息）
        Call DebugTool("删除本次入院信息(zl9Insure\入院登记撤销)")
        Call WriteBusinessLOG("删除本次入院信息(zl9Insure\入院登记撤销)", "", "")
        gstrSQL = "ZL_入院信息_DELETE('" & str入院登记号 & "')"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    入院登记撤销_北京 = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\入院登记撤销)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\入院登记撤销)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_北京(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    On Error GoTo errHand
    Dim str出院诊断 As String, str出院科室 As String, str出院日期 As String, int出院情况 As Integer
    Dim str入院登记号 As String, str卡号 As String
    Dim rsTemp As New ADODB.Recordset
    
    '如果没有发生任何费用，不允许办理出院
    Call DebugTool("未发生费用，只能撤销入院(zl9Insure\出院登记)")
    Call WriteBusinessLOG("未发生费用，只能撤销入院(zl9Insure\出院登记)", "", "")
    gstrSQL = "Select Count(*) AS Records From 住院费用记录" & _
        " Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "'如果没有发生任何费用，不允许办理出院", lng病人ID, lng主页ID)
    If rsTemp!Records = 0 Then
        MsgBox "不允许无费用出院，请撤销入院登记！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not 门诊特殊病(lng病人ID) Then
        '填写出院诊断信息（部分信息需要用小工具补充，才能上传）
        Call DebugTool("提取出院诊断、出院科室、出院情况(zl9Insure\出院登记)")
        Call WriteBusinessLOG("提取出院诊断、出院科室、出院情况(zl9Insure\出院登记)", "", "")
        str出院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, False, True, True)
        '提取临床部门性质及出院情况
        gstrSQL = " Select B.工作性质,A.出院方式,A.出院日期 " & _
                  " From 病案主页 A,临床部门 B " & _
                  " Where A.出院科室ID=B.部门ID(+) And A.病人ID=[1] And A.主页ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取临床部门性质及出院情况", lng病人ID, lng主页ID)
        str出院科室 = Nvl(rsTemp!工作性质)
        str出院日期 = Format(rsTemp!出院日期, "yyyy-MM-dd HH:mm:ss")
        '治愈-1,好转-2,未愈-3,死亡-4,转院-5,转外-6,其它-9
        Select Case rsTemp!出院方式
        Case "治愈"
            int出院情况 = 1
        Case "好转"
            int出院情况 = 2
        Case "未愈"
            int出院情况 = 3
        Case "死亡"
            int出院情况 = 4
        Case "转院"
            int出院情况 = 5
        Case "转外"
            int出院情况 = 6
        Case Else
            int出院情况 = 9
        End Select
        
        '取出入院登记号及卡号
        Call DebugTool("从入院信息中取出入院登记号及卡号(zl9Insure\出院登记)")
        Call WriteBusinessLOG("从入院信息中取出入院登记号及卡号(zl9Insure\出院登记)", "", "")
        gstrSQL = "Select 入院登记号,卡号 From 入院信息 Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID
        If rsTemp.State = 1 Then rsTemp.Close
        Call SQLTest(App.Title, "zl9Insure\出院登记", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
        str入院登记号 = Nvl(rsTemp!入院登记号)
        str卡号 = Nvl(rsTemp!卡号)
        
        If str出院科室 = "" Then
            MsgBox "请先设置出院科室与医保科室的对照[部门管理]", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_北京 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    If Not 门诊特殊病(lng病人ID) Then
        '参数如下：
        '入院登记号,卡号,出院科别编码,主要诊断,主要疾病编码,手术名称,手术名称编码,出院情况编码,上传
        Call DebugTool("产生出院诊断信息记录(zl9Insure\出院登记)")
        Call WriteBusinessLOG("产生出院诊断信息记录(zl9Insure\出院登记)", "", "")
        gstrSQL = "ZL_出院诊断信息_UPDATE('" & str入院登记号 & "'," & _
                "NULL,'" & str卡号 & "'," & lng病人ID & "," & lng主页ID & ",'" & str出院科室 & "'," & _
                "'" & Replace(Split(str出院诊断, mstrSplit)(0), "(" & Split(str出院诊断, mstrSplit)(1) & ")", "") & "','" & Split(str出院诊断, mstrSplit)(1) & "'," & _
                "NULL,NULL," & int出院情况 & ",to_Date('" & str出院日期 & "','yyyy-MM-dd hh24:mi:ss')" & ",0)"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    出院登记_北京 = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\出院登记)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\出院登记)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_北京(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    
    If Not 门诊特殊病(lng病人ID) Then
        '出院诊断信息已上传的不允许撤销
        gstrSQL = " Select Nvl(上传,0) 上传 From 出院诊断信息 " & _
                  " Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID
        If rsTemp.State = 1 Then rsTemp.Close
        Call SQLTest(App.Title, "zl9Insure\出院登记撤销", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
        If rsTemp!上传 = 1 Then
            MsgBox "该病人本次住院的数据已上传至中心，不允许办理出院登记撤销！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '删除出院诊断信息
        gstrSQL = "zl_出院诊断信息_DELETE(" & lng病人ID & "," & lng主页ID & ")"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_北京 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    出院登记撤销_北京 = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\出院登记撤销)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\出院登记撤销)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_北京(ByVal lng病人ID As Long) As Currency
    个人余额_北京 = 0
End Function

Public Function 住院结算_北京(ByVal lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    '将本次住院发生的所有费用进行费用分解
    Dim int项目类别 As Integer  '0-药品 1-诊疗项目 2-服务设施
    Dim str结算日期 As String
    Dim str医保编码 As String, strHIS项目名称 As String, str公务员标识 As String, strReturn As String
    Dim str待遇 As String, str费用分解待遇 As String, str结算方式 As String
    Dim dbl费用总额 As Double, dbl医保基金 As Double, dbl大病补助 As Double, dbl现金 As Double
    Dim dbl医保内 As Double, dbl个人应付 As Double, dbl首先自付 As Double, dbl统筹封顶后医保内 As Double
    Dim strFields As String, strValues As String
    Dim blnTrans As Boolean
    Dim lng主页ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim arr结算方式
    Const int费用总额 As Integer = 3
    Const int医保内 As Integer = 4
    Const int医保基金 As Integer = 5
    Const int大病补助 As Integer = 7
    Const int个人应付 As Integer = 9
    Const int首先自付 As Integer = 10
    Const int统筹封顶后医保内 As Integer = 11
    On Error GoTo errHandle
    
    If Not 医保病人已经出院(lng病人ID) Then
        MsgBox "不支持中途结算，请为该病人办理出院后再进行出院结算！", vbInformation, gstrSysName
        Exit Function
    End If
    
    str结算日期 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    
    '读取本次费用明细
    gstrSQL = " Select 主页ID From 住院费用记录 Where 结帐ID=[1] ANd Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取本次费用明细", lng结帐ID)
    lng主页ID = rsTemp!主页ID
    
    '获取参保人公务员标识
    Call DebugTool("获取参保人公务员标识(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("获取参保人公务员标识(zl9Insure\住院虚拟结算)", "", "")
    gstrSQL = "Select 业务类型,卡号,公务员 From 保险帐户 Where 病人ID=" & lng病人ID
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\住院虚拟结算_北京", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    str公务员标识 = rsTemp!公务员
    gComInfo_北京.卡号 = rsTemp!卡号
    gComInfo_北京.业务类型 = rsTemp!业务类型
    
    '走门诊流程
    If gComInfo_北京.业务类型 = "12" Then
        住院结算_北京 = 门诊结算_北京(lng结帐ID, 0, "", True)
        Exit Function
    End If
    
    '得到接口返回的待遇串
    strReturn = gComInfo_北京.卡号 & mstrSplit & str公务员标识
    Call DebugTool("调用获取待遇信息接口(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("调用获取待遇信息接口(zl9Insure\住院虚拟结算)", "", "")
    If Not 调用接口_北京(接口功能.获取手册病人待遇信息, strReturn) Then Exit Function   'strReturn的内容将用于费用分解，以下不能进行赋值
    str待遇 = strReturn
    
    '将返回的待遇信息转换为费用分解所需的格式
    str费用分解待遇 = TransationHosp(str待遇)
    
    '调用费用分解函数（strReturn中是待遇信息）
    Call DebugTool("调用费用明细分解函数(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("调用费用明细分解函数(zl9Insure\住院虚拟结算)", "", "")
    strReturn = str费用分解待遇
    If Not 调用接口_北京(费用分解_住院1, strReturn) Then Exit Function
    strReturn = AnalyFile_住院1(True, lng结帐ID)
    If strReturn <> "" Then
        arr结算方式 = Split(strReturn, mstrSplit)
        dbl费用总额 = Val(arr结算方式(int费用总额))
        dbl医保基金 = Val(arr结算方式(int医保基金))
        dbl大病补助 = Val(arr结算方式(int大病补助))
        
        dbl医保内 = Val(arr结算方式(int医保内))
        dbl个人应付 = Val(arr结算方式(int个人应付))
        dbl首先自付 = Val(arr结算方式(int首先自付))
        dbl统筹封顶后医保内 = Val(arr结算方式(int统筹封顶后医保内))
    End If
    dbl现金 = dbl费用总额 - dbl医保基金 - dbl大病补助
    
    
    '*****************保存中间库，以备上传*****************
    blnTrans = True
    gcnBJYB.BeginTrans
    '**保存本次的交易待遇信息**
    Call DebugTool("保存本次的交易待遇信息(zl9Insure\住院结算)")
    Call WriteBusinessLOG("保存本次的交易待遇信息(zl9Insure\住院结算)", "", "")
    If Not SaveBusinessDeal(str待遇) Then
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '**保存交易版本信息**
    Call DebugTool("保存本次的交易版本信息(zl9Insure\住院结算)")
    Call WriteBusinessLOG("保存本次的交易版本信息(zl9Insure\住院结算)", "", "")
    If Not SaveBusinessVersion(False) Then
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '**保存住院交易信息、住院分段及住院费用明细**
    Call DebugTool("保存住院交易信息、住院分段及住院费用明细(zl9Insure\住院结算)")
    Call WriteBusinessLOG("保存住院交易信息、住院分段及住院费用明细(zl9Insure\住院结算)", "", "")
    strReturn = AnalyFile_住院1(False, lng结帐ID)
    If strReturn = "" Then
        '保存失败
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '**产生本次手册消费记录**
    Call DebugTool("产生本次手册消费记录(zl9Insure\住院结算)")
    Call WriteBusinessLOG("产生本次手册消费记录(zl9Insure\住院结算)", "", "")
    If Not SaveDeal(False, True) Then
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '更新出院诊断信息中的交易流水号
    gstrSQL = "Update 出院诊断信息 Set 交易流水号='" & gComInfo_北京.交易流水号 & "' Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID
    gcnBJYB.Execute gstrSQL
    
    '**保存保险结算记录**
    '进入统筹金额=公务员补助基金;统筹报销金额=统筹基金;大病自付金额=大病基金
    '支付顺序号=交易流水号,备注=业务类型
    Call DebugTool("保存保险结算记录(zl9Insure\住院结算)")
    Call WriteBusinessLOG("保存保险结算记录(zl9Insure\住院结算)", "", "")
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_北京 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng主页ID & "," & 0 & "," & 0 & "," & 0 & "," & _
        dbl费用总额 & "," & dbl现金 & ",0,0," & dbl医保基金 & "," & dbl大病补助 & "," & _
        0 & ",0,'" & gComInfo_北京.交易流水号 & "'," & lng主页ID & ",null,'" & gComInfo_北京.业务类型 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
    
    gcnBJYB.CommitTrans
    blnTrans = False
    住院结算_北京 = True
    Exit Function
errHandle:
    Call DebugTool("(zl9INSURE\住院虚拟结算_北京)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\住院虚拟结算_北京)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then gcnBJYB.RollbackTrans
End Function

Public Function 住院虚拟结算_北京(ByVal rs明细 As ADODB.Recordset, ByVal lng病人ID As Long) As String
    '将本次住院发生的所有费用进行费用分解
    Dim int项目类别 As Integer  '0-药品 1-诊疗项目 2-服务设施
    Dim str结算日期 As String
    Dim str医保编码 As String, strHIS项目名称 As String, str公务员标识 As String, strReturn As String, strUser As String
    Dim str待遇 As String, str费用分解待遇 As String, str结算方式 As String
    Dim strFilter As String
    Dim strFields As String, strValues As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim rsLimit As New ADODB.Recordset
    Dim arr结算方式
    Const int医保基金 As Integer = 5
    Const int大病补助 As Integer = 7
    'rs明细记录集中的字段清单
    'ID,记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
    
    On Error GoTo errHandle
    
    strUser = GetUser
    str结算日期 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    
    '先获取本次交易流水号
    Call DebugTool("产生流水号(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("产生流水号(zl9Insure\住院虚拟结算)", "", "")
    Call GetSequence_北京(lng病人ID)
    
    '获取参保人公务员标识
    Call DebugTool("获取参保人公务员标识(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("获取参保人公务员标识(zl9Insure\住院虚拟结算)", "", "")
    gstrSQL = "Select 卡号,公务员 From 保险帐户 Where 病人ID=" & lng病人ID
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\住院虚拟结算_北京", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    str公务员标识 = rsTemp!公务员
    gComInfo_北京.卡号 = rsTemp!卡号
    
    '获取参保人本次就诊的业务类型，如果是特殊病，则进行单独处理
    If 门诊特殊病(lng病人ID) Then
        '走门诊特殊病流程
        If Not 门诊特殊病_北京(rs明细, str结算方式) Then Exit Function
        住院虚拟结算_北京 = str结算方式
        Exit Function
    End If
    
    '获取此刻的待遇信息（费用分解函数的入参）
    Call DebugTool("获取此刻的历史消费记录（费用分解函数的入参）(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("获取此刻的历史消费记录（费用分解函数的入参）(zl9Insure\住院虚拟结算)", "", "")
    Set rsDeal = GetDeal(lng病人ID, str结算日期)
    '根据录入的历史消费记录，产生入参文件
    Call DebugTool("根据录入的历史消费记录，产生入参文件(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("根据录入的历史消费记录，产生入参文件(zl9Insure\住院虚拟结算)", "", "")
    If Not MakeFile_Center(rsDeal, 接口功能.获取手册病人待遇信息) Then Exit Function
    '得到接口返回的待遇串
    strReturn = gComInfo_北京.卡号 & mstrSplit & str公务员标识
    Call DebugTool("调用获取待遇信息接口(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("调用获取待遇信息接口(zl9Insure\住院虚拟结算)", "", "")
    If Not 调用接口_北京(接口功能.获取手册病人待遇信息, strReturn) Then Exit Function   'strReturn的内容将用于费用分解，以下不能进行赋值
    str待遇 = strReturn
    
    '判断是否存在限制使用的项目，如果存在，则需要操作员确定是否使用于医保内
    Call DebugTool("产生限制使用项目记录集，供操作员选择是否属于医保内(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("产生限制使用项目记录集，供操作员选择是否属于医保内(zl9Insure\住院虚拟结算)", "", "")
    strFields = "医保内," & adLongVarChar & ",1" & mstrSplit & _
              "NO," & adLongVarChar & ",8" & mstrSplit & _
              "记录性质," & adLongVarChar & ",3" & mstrSplit & _
              "记录状态," & adLongVarChar & ",3" & mstrSplit & _
              "序号," & adLongVarChar & ",5" & mstrSplit & _
              "HIS项目名称," & adLongVarChar & ",100" & mstrSplit & _
              "医保编码," & adLongVarChar & ",100" & mstrSplit & _
              "数量," & adLongVarChar & ",20" & mstrSplit & _
              "单价," & adLongVarChar & ",20" & mstrSplit & _
              "金额," & adLongVarChar & ",20" & mstrSplit & _
              "限制," & adLongVarChar & ",500" & mstrSplit & _
              "备注," & adLongVarChar & ",500" & mstrSplit & _
              "开单医生," & adLongVarChar & ",20" & mstrSplit & _
              "登记时间," & adLongVarChar & ",20"
    Call Record_Init(rsLimit, strFields)
    strFields = "医保内|NO|记录性质|记录状态|序号|HIS项目名称|医保编码|数量|单价|金额|限制|备注|开单医生|登记时间"
    
    If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
    Do Until rs明细.EOF
        '只有药品项目才存在限制等级
        If rs明细!记录状态 = 1 Or rs明细!记录状态 = 3 Then
            gstrSQL = "Select A.项目编码 As 医保编码,C.通用名称 AS 项目名称,D.名称 AS 限制等级,F.备注 " & _
                " From 保险支付项目 A,药品目录 B,药品信息 C," & strUser & ".药品目录 F," & _
                " (Select B.编码,B.名称 From " & strUser & ".指标主表 A," & strUser & ".指标体系对照表 B Where A.类别=B.类别 And A.名称='使用限制等级') D" & _
                " Where A.险类=[1] And A.收费细目ID=B.药品ID And B.药名ID=C.药名ID And F.编码=A.项目编码 " & _
                " AND B.药品ID=[2] ANd F.使用限制等级=D.编码"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "住院预算", TYPE_北京, CLng(rs明细!收费细目ID))
            If Not rsTemp.EOF Then
                '产生记录集，以便操作员完成医保内外的确定
                strValues = "√" & mstrSplit & rs明细!NO & mstrSplit & rs明细!记录性质 & mstrSplit & rs明细!记录状态 & mstrSplit & rs明细!序号 & mstrSplit & _
                    rsTemp!项目名称 & mstrSplit & rsTemp!医保编码 & mstrSplit & _
                    Format(rs明细!数量, "#####0.00;-#####0.00;0;") & mstrSplit & Format(rs明细!价格, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
                    Format(rs明细!金额, "#####0.0000;-#####0.0000;0;") & mstrSplit & Nvl(rsTemp!限制等级) & mstrSplit & Nvl(rsTemp!备注) & mstrSplit & _
                    Nvl(rs明细!医生) & mstrSplit & Format(rs明细!发生时间, "yyyyMMdd")
                Call Record_Add(rsLimit, strFields, strValues)
            End If
        End If
        
        rs明细.MoveNext
    Loop
    '需要操作员确认这类项目的医保内外
    If rsLimit.RecordCount <> 0 Then Call frm限制用药医保内外划分.ShowEditor(rsLimit)
    
    '将费用明细生成明细文件（也是费用分解函数的入参）
'    ----传入文件说明
'    序号    数据项  类型    最大长度    说明
'    1   项目序号    C   9   顺序号
'    2   处方号  C   20  参见标准AKC220，可空
'    3   项目代码    C   20  药品、诊疗项目或服务设施编码（适应症用药需要操作员确定哪些是医保内用药，非医保内用药传零）
'    4   项目名称    C   100 本医院项目名称
'    5   项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'    6   单价    N   10,4    AKC225
'    7   数量    N   8,2 AKC226
'    8   费用总金额  N   10,4    实际结算金额
'    9   费用发生日期    D   8   YYYYMMDD
    strFields = "项目序号," & adLongVarChar & ",9" & mstrSplit & _
                "医嘱号," & adLongVarChar & ",20" & mstrSplit & _
                "项目代码," & adLongVarChar & ",20" & mstrSplit & _
                "项目名称," & adLongVarChar & ",100" & mstrSplit & _
                "项目类别," & adLongVarChar & ",3" & mstrSplit & _
                "单价," & adDouble & ",18" & mstrSplit & _
                "数量," & adDouble & ",18" & mstrSplit & _
                "费用总金额," & adDouble & ",18" & mstrSplit & _
                "费用发生日期," & adLongVarChar & ",20"
    Call Record_Init(rsCharge, strFields)
    
    '得到本次录入的费用明细的相关医保信息（编码以w01打头，表示服务设施类项目）
    Call DebugTool("产生费用明细记录集(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("产生费用明细记录集(zl9Insure\住院虚拟结算)", "", "")
    strFields = "项目序号|医嘱号|项目代码|项目名称|项目类别|单价|数量|费用总金额|费用发生日期"
    If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
    Do Until rs明细.EOF
        '将药品的通用名称或非药品项目的名称取出来
        gstrSQL = "Select A.项目编码 As 医保编码,C.通用名称 AS 项目名称 " & _
            " From 保险支付项目 A,药品目录 B,药品信息 C" & _
            " Where A.险类(+)=[1] And A.收费细目ID(+)=B.药品ID And B.药名ID=C.药名ID " & _
            " AND B.药品ID=[2]" & _
            " UNION " & _
            " Select A.项目编码 As 医保编码,B.名称 AS 项目名称" & _
            " From 保险支付项目 A,收费细目 B" & _
            " Where A.险类(+)=[1] AND B.ID=[2]" & _
            " And A.收费细目ID(+)=B.ID AND B.类别 Not In ('5','6','7')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "住院预算", TYPE_北京, CLng(rs明细!收费细目ID))
        If rsTemp.EOF Then
            MsgBox "还有项目没有和医保项目设置对照关系！[保险项目]", vbInformation, gstrSysName
            Exit Function
        End If
        
        '先判断是否是限制用药，如果是，根据操作员的选择确定
        strFilter = "NO='" & rs明细!NO & "' And 记录性质=" & rs明细!记录性质 & IIf(rs明细!记录状态 = "2", " And 记录状态=3", " And 记录状态=1") & " And 序号=" & rs明细!序号
        rsLimit.Filter = strFilter
        If rsLimit.RecordCount = 0 Then
            str医保编码 = Nvl(rsTemp!医保编码, 0)
        Else
            If rsLimit!医保内 = 1 Or rsLimit!医保内 = "√" Then
                str医保编码 = Nvl(rsTemp!医保编码, 0)
            Else
                str医保编码 = 0
            End If
        End If
        strHIS项目名称 = rsTemp!项目名称
        
        If InStr(1, "5,6,7", rs明细!收费类别) <> 0 Then
            int项目类别 = 0 '0-药品
        Else
            int项目类别 = (IIf(str医保编码 Like "w01*", 2, 1)) '1-诊疗项目 2-服务设施
        End If
        
        '产生费用明细记录集,以供产生入参文件
        strValues = rs明细.AbsolutePosition & mstrSplit & (rs明细!NO & "*" & rs明细!记录性质 & "*" & rs明细!记录状态 & "*" & rs明细!序号) & mstrSplit & _
            str医保编码 & mstrSplit & strHIS项目名称 & mstrSplit & _
            int项目类别 & mstrSplit & Format(rs明细!价格, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            Format(rs明细!数量, "#####0.00;-#####0.00;0;") & mstrSplit & Format(rs明细!金额, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            Format(rs明细!发生时间, "yyyyMMdd")
        Call Record_Add(rsCharge, strFields, strValues)
        
        rs明细.MoveNext
    Loop
    
    '产生费用明细文件
    Call DebugTool("根据费用明细，产生入参文件(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("根据费用明细，产生入参文件(zl9Insure\住院虚拟结算)", "", "")
    If Not MakeFile_Center(rsCharge, 费用分解_住院1) Then Exit Function
    
    '将返回的待遇信息转换为费用分解所需的格式
    str费用分解待遇 = TransationHosp(str待遇)
    
    '调用费用分解函数（strReturn中是待遇信息）
    Call DebugTool("调用费用明细分解函数(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("调用费用明细分解函数(zl9Insure\住院虚拟结算)", "", "")
    strReturn = str费用分解待遇
    If Not 调用接口_北京(费用分解_住院1, strReturn) Then Exit Function
    strReturn = AnalyFile_住院1(True)
    If strReturn = "" Then Exit Function
    
    If Not SaveDeal(False, False) Then Exit Function
    
    '组织成结算方式串并返回
    Call DebugTool("分析费用分解返回的汇总数据，将结算信息返回给主调程序(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("分析费用分解返回的汇总数据，将结算信息返回给主调程序(zl9Insure\住院虚拟结算)", "", "")
    arr结算方式 = Split(strReturn, mstrSplit)
    str结算方式 = mstrSplit & "统筹支付;" & arr结算方式(int医保基金) & ";0"
    str结算方式 = str结算方式 & mstrSplit & "大额支付;" & arr结算方式(int大病补助) & ";0"
    str结算方式 = Mid(str结算方式, 2)
    
    住院虚拟结算_北京 = str结算方式
    Exit Function
errHandle:
    Call DebugTool("(zl9INSURE\住院虚拟结算_北京)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\住院虚拟结算_北京)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_北京(ByVal lng结帐ID As Long) As Boolean
    '**住院结算冲销时，要删除上次的历史就诊记录**
    On Error GoTo errHand
    Dim lng病人ID As Long
    Dim lng冲销ID As Long
    Dim blnTrans As Boolean
    Dim str卡号 As String, str原交易流水号 As String
    Dim str结算时间 As String, str退费时间 As String
    Dim rsTemp As New ADODB.Recordset
    
    '获取病人ID
    Call DebugTool("获取病人ID(zl9Insure\住院结算冲销)")
    Call WriteBusinessLOG("获取病人ID(zl9Insure\住院结算冲销)", "", "")
    gstrSQL = "Select 病人ID From 住院费用记录 Where 结帐ID=[1] ANd Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人ID", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    
    Call DebugTool("产生交易流水号(zl9Insure\住院结算冲销)")
    Call WriteBusinessLOG("产生交易流水号(zl9Insure\住院结算冲销)", "", "")
    Call GetSequence_北京(lng病人ID)
    
    '取出原始结算记录的结算时间
    Call DebugTool("取出原始结算记录的结算时间(zl9Insure\住院结算冲销)")
    Call WriteBusinessLOG("取出原始结算记录的结算时间(zl9Insure\住院结算冲销)", "", "")
    gstrSQL = "Select 卡号,交易流水号,结算时间 From " & GetUser & ".交易待遇信息 " & _
            " Where 交易流水号=(Select 支付顺序号 From 保险结算记录 Where 性质=2 AND 记录ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出原始结算记录的结算时间", lng结帐ID)
    str卡号 = rsTemp!卡号
    str原交易流水号 = rsTemp!交易流水号
    str结算时间 = Format(rsTemp!结算时间, "yyyy-MM-dd HH:mm:ss")
    str退费时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    '检查，如果在该交易后，还发生了其它交易，则不允许进行冲销（把已退费的记录排开）
    Call DebugTool("检查是否从最后一笔开始退费(zl9Insure\住院结算冲销)")
    Call WriteBusinessLOG("检查是否从最后一笔开始退费(zl9Insure\住院结算冲销)", "", "")
    gstrSQL = "Select Count(*) AS Records From 交易待遇信息 A,保险帐户 B" & _
            " Where A.卡号=B.卡号 And B.病人ID=" & lng病人ID & _
            " And 结算时间>to_Date('" & str结算时间 & "','yyyy-MM-dd hh24:mi:ss')" & _
            " And 交易流水号 Not In (Select 原交易流水号 From 退费信息 Where 卡号='" & str卡号 & "')"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\门诊结算_北京", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsTemp!Records > 0 Then
        MsgBox "医保接口不允许从中间开始退费，您只能从最后一笔业务开始退费！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '取冲销记录的结帐ID，单据号
    Call DebugTool("取冲销ID(zl9Insure\住院结算冲销)")
    Call WriteBusinessLOG("取冲销ID(zl9Insure\住院结算冲销)", "", "")
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!ID
    
    '提取原始的保险结算记录
    Call DebugTool("提取原始的保险结算记录(zl9Insure\住院结算冲销)")
    Call WriteBusinessLOG("提取原始的保险结算记录(zl9Insure\住院结算冲销)", "", "")
    gstrSQL = "Select * From 保险结算记录 Where 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取原始的保险结算记录", lng结帐ID)
    
    '保存保险结算记录
    Call DebugTool("保存保险结算记录(zl9Insure\住院结算冲销)")
    Call WriteBusinessLOG("保存保险结算记录(zl9Insure\住院结算冲销)", "", "")
    gstrSQL = "zl_保险结算记录_insert(" & rsTemp!性质 & "," & lng冲销ID & "," & TYPE_北京 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & Nvl(rsTemp!主页ID, 0) & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & ",0,0," & -1 * Nvl(rsTemp!统筹报销金额, 0) & _
        "," & -1 * Nvl(rsTemp!大病自付金额, 0) & "," & 0 & ",0,'" & gComInfo_北京.交易流水号 & "'," & Nvl(rsTemp!主页ID, 0) & ",null,'" & Nvl(rsTemp!备注) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
    
    blnTrans = True
    gcnBJYB.BeginTrans
    '产生一条退费记录即可，参数如下
    '红冲交易流水号,卡号,原交易流水号,原交易日期,退费日期,操作员姓名,上传
    Call DebugTool("产生红冲退费记录(zl9Insure\住院结算冲销)")
    Call WriteBusinessLOG("产生红冲退费记录(zl9Insure\住院结算冲销)", "", "")
    gstrSQL = "ZL_退费信息_INSERT(" & _
            "'" & gComInfo_北京.交易流水号 & "','" & str卡号 & "','" & str原交易流水号 & "'," & _
            "to_Date('" & str结算时间 & "','yyyy-MM-dd hh24:mi:ss')," & _
            "to_Date('" & str退费时间 & "','yyyy-MM-dd hh24:mi:ss')," & _
            "'" & UserInfo.姓名 & "',0)"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    '**住院结算冲销时，要删除上次的历史就诊记录**
    gstrSQL = "ZL_手册消费记录_DELETE('" & str原交易流水号 & "')"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    gcnBJYB.CommitTrans
    住院结算冲销_北京 = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\住院结算冲销)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\住院结算冲销)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnBJYB.RollbackTrans
End Function

Private Function 门诊特殊病(ByVal lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 业务类型 From 保险帐户 Where 病人ID=" & lng病人ID
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\门诊特殊病", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    门诊特殊病 = (rsTemp!业务类型 = "12")
End Function

Private Function 门诊特殊病_北京(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    Dim int项目类别 As Integer  '0-药品 1-诊疗项目 2-服务设施
    Dim lng病人ID As Long
    Dim str就诊日期 As String, str费用发生日期 As String
    Dim str医保编码 As String, strHIS项目名称 As String, str公务员标识 As String, strReturn As String
    Dim str待遇 As String, str费用分解待遇 As String
    Dim strFields As String, strValues As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim arr结算方式
    Const int医保基金 As Integer = 5
    Const int大病补助 As Integer = 7
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    'rs明细记录集中的字段清单
    'ID,记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
    On Error GoTo errHandle
    lng病人ID = rs明细!病人ID
'    gstrSQL = "Select 入院日期 From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 And B.病人ID=" & lng病人ID
'    Call OpenRecordset(rsTemp, "提取病人入院日期")
'    str就诊日期 = Format(rsTemp!入院日期, "yyyy-MM-dd")
    str就诊日期 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    
    '先获取本次交易流水号
    Call DebugTool("产生流水号(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("产生流水号(zl9Insure\住院虚拟结算)", "", "")
    Call GetSequence_北京(lng病人ID)
    
    '获取参保人公务员标识
    Call DebugTool("获取参保人公务员标识(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("获取参保人公务员标识(zl9Insure\住院虚拟结算)", "", "")
    gstrSQL = "Select 公务员 From 保险帐户 Where 卡号='" & gComInfo_北京.卡号 & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\门诊特殊病_北京", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    str公务员标识 = rsTemp!公务员
    
    '获取此刻的待遇信息（费用分解函数的入参）
    Call DebugTool("获取此刻的历史消费记录（费用分解函数的入参）(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("获取此刻的历史消费记录（费用分解函数的入参）(zl9Insure\住院虚拟结算)", "", "")
    Set rsDeal = GetDeal(lng病人ID, str就诊日期)
    '根据录入的历史消费记录，产生入参文件
    Call DebugTool("根据录入的历史消费记录，产生入参文件(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("根据录入的历史消费记录，产生入参文件(zl9Insure\住院虚拟结算)", "", "")
    If Not MakeFile_Center(rsDeal, 接口功能.获取手册病人待遇信息) Then Exit Function
    '得到接口返回的待遇串
    strReturn = gComInfo_北京.卡号 & mstrSplit & str公务员标识
    Call DebugTool("调用获取待遇信息接口(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("调用获取待遇信息接口(zl9Insure\住院虚拟结算)", "", "")
    If Not 调用接口_北京(接口功能.获取手册病人待遇信息, strReturn) Then Exit Function   'strReturn的内容将用于费用分解，以下不能进行赋值
    str待遇 = strReturn
    
    '将费用明细生成明细文件（也是费用分解函数的入参）
'    ----传入文件说明
'    序号    数据项  类型    最大长度    说明
'    1   项目序号    C   9   顺序号
'    2   处方号  C   20  参见标准AKC220，可空
'    3   项目代码    C   20  药品、诊疗项目或服务设施编码
'    4   项目名称    C   100 本医院项目名称
'    5   项目类别    C   3   0-药品 1-诊疗项目 2-服务设施
'    6   单价    N   10,4    AKC225
'    7   数量    N   8,2 AKC226
'    8   费用总金额  N   10,4    实际结算金额
'    9   费用发生日期    D   8   YYYYMMDD
    strFields = "项目序号," & adLongVarChar & ",9" & mstrSplit & _
                "处方号," & adLongVarChar & ",20" & mstrSplit & _
                "项目代码," & adLongVarChar & ",20" & mstrSplit & _
                "项目名称," & adLongVarChar & ",100" & mstrSplit & _
                "项目类别," & adLongVarChar & ",3" & mstrSplit & _
                "单价," & adDouble & ",18" & mstrSplit & _
                "数量," & adDouble & ",18" & mstrSplit & _
                "费用总金额," & adDouble & ",18" & mstrSplit & _
                "费用发生日期," & adLongVarChar & ",20"
    Call Record_Init(rsCharge, strFields)
    
    '得到本次录入的费用明细的相关医保信息（编码以w01打头，表示服务设施类项目）
    Call DebugTool("产生费用明细记录集(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("产生费用明细记录集(zl9Insure\住院虚拟结算)", "", "")
    strFields = "项目序号|处方号|项目代码|项目名称|项目类别|单价|数量|费用总金额|费用发生日期"
    If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
    Do Until rs明细.EOF
        '将药品的通用名称或非药品项目的名称取出来
        gstrSQL = "Select A.项目编码 As 医保编码,C.通用名称 AS 项目名称 " & _
            " From 保险支付项目 A,药品目录 B,药品信息 C" & _
            " Where A.险类(+)=[1] And A.收费细目ID(+)=B.药品ID And B.药名ID=C.药名ID " & _
            " AND B.药品ID=[2]" & _
            " UNION " & _
            " Select A.项目编码 As 医保编码,B.名称 AS 项目名称" & _
            " From 保险支付项目 A,收费细目 B" & _
            " Where A.险类(+)=[1] AND B.ID=[2]" & _
            " And A.收费细目ID(+)=B.ID AND B.类别 Not In ('5','6','7')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊预算", TYPE_北京, CLng(rs明细!收费细目ID))
        If rsTemp.EOF Then
            MsgBox "还有项目没有和医保项目设置对照关系！[保险项目]", vbInformation, gstrSysName
            Exit Function
        End If
        str医保编码 = Nvl(rsTemp!医保编码, 0)
        strHIS项目名称 = rsTemp!项目名称
        
        If InStr(1, "5,6,7", rs明细!收费类别) <> 0 Then
            int项目类别 = 0 '0-药品
        Else
            int项目类别 = (IIf(str医保编码 Like "w01*", 2, 1)) '1-诊疗项目 2-服务设施
        End If
        
        '产生费用明细记录集,以供产生入参文件
        strValues = rs明细.AbsolutePosition & mstrSplit & "" & mstrSplit & _
            str医保编码 & mstrSplit & strHIS项目名称 & mstrSplit & _
            int项目类别 & mstrSplit & Format(rs明细!价格, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            Format(rs明细!数量, "#####0.00;-#####0.00;0;") & mstrSplit & Format(rs明细!金额, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            Format(rs明细!登记时间, "yyyyMMdd")
        Call Record_Add(rsCharge, strFields, strValues)
        
        rs明细.MoveNext
    Loop
    
    '产生费用明细文件
    Call DebugTool("根据费用明细，产生入参文件(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("根据费用明细，产生入参文件(zl9Insure\住院虚拟结算)", "", "")
    If Not MakeFile_Center(rsCharge, 费用分解_特殊门诊1) Then Exit Function
    
    '将返回的待遇信息转换为费用分解所需的格式
    str费用分解待遇 = TransationSpec(str待遇)
    
    '调用费用分解函数（strReturn中是待遇信息）
    Call DebugTool("调用费用明细分解函数(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("调用费用明细分解函数(zl9Insure\住院虚拟结算)", "", "")
    strReturn = str费用分解待遇
    If Not 调用接口_北京(费用分解_特殊门诊1, strReturn) Then Exit Function
    strReturn = AnalyFile_特殊门诊1(True)
    If strReturn = "" Then Exit Function
    
    If Not SaveDeal(True, False) Then Exit Function
    
    '组织成结算方式串并返回
    Call DebugTool("分析费用分解返回的汇总数据，将结算信息返回给主调程序(zl9Insure\住院虚拟结算)")
    Call WriteBusinessLOG("分析费用分解返回的汇总数据，将结算信息返回给主调程序(zl9Insure\住院虚拟结算)", "", "")
    arr结算方式 = Split(strReturn, mstrSplit)
    str结算方式 = mstrSplit & "统筹支付;" & arr结算方式(int医保基金) & ";0"
    str结算方式 = str结算方式 & mstrSplit & "大额支付;" & arr结算方式(int大病补助) & ";0"
    str结算方式 = Mid(str结算方式, 2)
    
    门诊特殊病_北京 = True
    Exit Function
errHandle:
    Call DebugTool("(zl9INSURE\门诊特殊病_北京)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\门诊特殊病_北京)" & vbCrLf & _
        "错误号:" & Err.Number & "|错误信息:" & Err.Description, "", "")
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
