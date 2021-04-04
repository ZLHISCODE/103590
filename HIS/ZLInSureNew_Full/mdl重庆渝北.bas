Attribute VB_Name = "mdl重庆渝北"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)
    
'-------------------------------------------------------------------------------------------------------------------------------------
'API的医保接口声明

'刘兴宏:测试屏蔽
    Private Type Struct
        lngAppCode  As Long   '标志服务执行状态代码。等于1时表示服务执行正常结束，小于0时表示服务执行异常或错误。
        strErrMsg  As String  '当服务执行状态代码AppCod小于0时，描述服务执行的异常或错误信息。
    End Type
    
    '声明API函数
    '功能:调用远程数据服务，返回远程数据服务的反馈结果
    'Private Declare Function DataUpload Lib "YHMdcrDataUpldSvr.dll" Alias "_DataUpload@12" ( _
         strInputString As String, strOutPutstring As String, AppInfo As Struct) As Boolean
    '新接口
   ' Private Declare Function DataUpload Lib "YHMdcrDataUpldSvr.dll" Alias "_DataUpload@4" (ByVal strInputString As String) As as Boolean
 
 
    Private tmpStruct As Struct

    '解析卡内数据函数
    '   strPerNo-个人编号
    '   strCardNO-卡号
    '   strExInfor-应用执行信息
    Private Declare Sub srd_4428_info Lib "Mwic_32.dll" ( _
         ByVal strPerNO As String, ByVal strCardNO As String, ByVal strExInfor As String)


    
    '下载定点信息
    Private Declare Function ExportKB01 Lib "YHMdcrAsistntSvr.dll" Alias "_ExportKB01@8" (ByVal strFileName As String, ByRef tmpStrut As Struct) As Boolean
    
    '获取就诊编号
    Private Declare Function GetAKC190 Lib "YHMdcrAsistntSvr.dll" Alias "_GetAKC190@12" (ByVal strYab003 As String, ByRef strAkc190 As String, ByRef tmpStrut As Struct) As Boolean
    
    '获取结算编号
    Private Declare Function GetYKA105 Lib "YHMdcrAsistntSvr.dll" Alias "_GetYKA105@12" (ByVal strYab003 As String, ByRef strYka105 As String, ByRef tmpStrut As Struct) As Boolean
    
'-------------------------------------------------------------------------------------------------------------------------------------
'当用的变量声明
Public gcnOracle_CQYB       As New Connection        '连接到oracle数据库(中间库)
Public gobjStruct As Struct
Private Type 结算信息
    dbl医保基金 As Double
    dbl大额补助 As Double
    dbl个人帐户 As Double
    dbl公务员补助 As Double
    dbl单病种超限 As Double
    dbl费用总额 As Double
    bln验证 As Boolean
End Type
Private m虚拟结算信息 As 结算信息

Private Type InitbaseInfor
    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    医院编码 As String                      '初始医院编码
    机构类别 As String                      '医疗机构类别
    经办机构代码 As String
    医保机构名称 As String
    医疗机构名称 As String
    定点状态标识 As String
    单价限价    As Double
    收入项目id As Long
    待遇提醒期限 As Double
    解析卡内数据 As Boolean
End Type
Public InitInfor_重庆渝北 As InitbaseInfor

Public Enum 业务类型_重庆渝北
    获取系统时间 = 0
    身份鉴别
    修改密码
    IC卡帐户支付
    资格审批待遇核定
    就诊信息写入
    处方明细写入
    结算基本信息写入
    结算结果写入
    核对帐户支付信息
    核对就诊信息
    核对处方明细信息
    核对费用结算结果
    核对费用结算基本信息
    导出服务项目目录
    导出ICD_10信息
    导出病种目录
    导出病种就诊结算信息
    导出医保定点信息
    获取客户机标识号
    解析卡内数据
    获取就诊编号
    获取记帐流水号
    获取结算编号
    费用结算
    作废审批记录
End Enum

Public g病人身份_重庆渝北 As 病人身份
Private Type 病人身份
    个人编号            As String
    卡号                As String
    姓名                As String
    密码                As String
    性别                As String
    身份证号            As String
    出生日期            As String
    医疗人员类别        As String
    医疗照顾类别        As String
    医疗补助类别        As String
    社保经办构构代码    As String
    单位编码            As Long
    单位名称            As String
    年龄                As Integer
    累计缴费月数        As Integer
    
    帐户余额            As Double
    
    病种ID              As Long
    病种编码            As String
    病种名称            As String
    
    病情ID              As Long
    病情编码            As String
    病情名称            As String
    
    病情1ID              As Long
    病情编码1            As String
    病情名称1            As String
    
    病情2ID              As Long
    病情编码2            As String
    病情名称2            As String
    
    
    支付类别            As String
    就诊编号            As String
    就诊结算方式        As String       '初始时赋值,主要是当前只有一种就诊结算方式即:0-按项目结算
    结算编号            As String
    
    结算标志            As Integer      '表示当前为采取的结算方式 0-门诊,1-住院,2-挂号,3-住院处方登记
    结帐ID              As Long         '表示当前的结帐ID
    冲销                As Boolean      '表示当前是否为冲销
    费用总额            As Double       '表示当前费用总额
    中途结帐            As Boolean      '示示当前结算为中途结算
    冲销ID              As Long
    虚拟结算            As Boolean      '当前是否为虚拟结算
    
    发票号              As String       '当前结算的发票号码
    就诊时间            As String       '以yyyy-mm-dd格式
    
    lng病人ID           As Long
    lng主页ID           As Long
    
    结算信息            As String       '当前的结算信息,主要是在虚拟结算用
    本次起付线          As Double
 End Type
 
 
 Public Enum CodeType
    医疗人员类别 = 0
    医疗照顾类别
    医疗补助类别
 End Enum

'相关Xml对象定义
Private gobjXMLInPut As MSXML2.DOMDocument
Private gobjXMLOutput As MSXML2.DOMDocument
Private Const gstrXMLRootPart  As String = "XMLBODY"       '根节点
Private gstrAppPath      As String
Private gobj费用结算   As Object
Public gobjYingHaiDll As Object
Private mblnInit As Boolean     '是否初始化

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'常用函数过程声明

Public Function 医保初始化_重庆渝北() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:初始医保的相关变量
    '--入参数:
    '--出参数:
    '--返  回:初始化成功，返回true；否则，返回false
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim strUser As String
    Dim strServer As String
    Dim strPass As String
    Dim rsTemp As New ADODB.Recordset
    
    If mblnInit = True Then
        医保初始化_重庆渝北 = True
        Exit Function
    End If
    
    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_重庆渝北.模拟数据 = True
    Else
        InitInfor_重庆渝北.模拟数据 = False
    End If
    
    Call GetRegInFor(g公共模块, "操作", "解析卡内数据", strReg)
    If Val(strReg) = 1 Then
        InitInfor_重庆渝北.解析卡内数据 = True
    Else
        InitInfor_重庆渝北.解析卡内数据 = False
    End If
    InitInfor_重庆渝北.解析卡内数据 = InitInfor_重庆渝北.解析卡内数据 Or InitInfor_重庆渝北.模拟数据
    
    '取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_重庆渝北)
    InitInfor_重庆渝北.医院编码 = Nvl(rsTemp!医院编码)

    '中间库连接
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "渝北医保", TYPE_重庆渝北)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    
    '限价确定
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名='处方单价限制' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "渝北医保", TYPE_重庆渝北)
    If Not rsTemp.EOF Then
        InitInfor_重庆渝北.单价限价 = Val(Nvl(rsTemp!参数值))
    Else
        InitInfor_重庆渝北.单价限价 = 200
    End If
    '限价确定
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名='个人帐户' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "渝北医保", TYPE_重庆渝北)
    If Not rsTemp.EOF Then
        InitInfor_重庆渝北.收入项目id = Val(Nvl(rsTemp!参数值))
    Else
        InitInfor_重庆渝北.收入项目id = 0
    End If
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名='享受待遇提醒月数' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "渝北医保", TYPE_重庆渝北)
    
    If Not rsTemp.EOF Then
        InitInfor_重庆渝北.待遇提醒期限 = Val(Nvl(rsTemp!参数值))
    Else
        InitInfor_重庆渝北.待遇提醒期限 = 3 '默认为3个月
    End If
    
    
    If OraDataOpen(gcnOracle_CQYB, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '给就诊结算方式赋值：目前主要只有一种方式:即:0-按项目结算
    g病人身份_重庆渝北.就诊结算方式 = "0"
    
    '初始设置医保相关文本文件目录
    gstrAppPath = App.Path & "\医保"
    
    '创建动态的费用结算对象
    Err = 0
    On Error Resume Next
    Set gobj费用结算 = Nothing
    
    Set gobj费用结算 = CreateObject("PB80.n_yhmedicarereckon")
    
    If gobj费用结算 Is Nothing Or Err <> 0 Then
        If InitInfor_重庆渝北.模拟数据 Then
        Else
            ShowMsgbox "费用结算部件有错误,请与医保接口商联系."
            Exit Function
        End If
    End If
    
    Set gobjYingHaiDll = Nothing
    Set gobjYingHaiDll = CreateObject("PB80.n_dll_in")
    If gobjYingHaiDll Is Nothing Then
        If InitInfor_重庆渝北.模拟数据 Then
        Else
            ShowMsgbox "创建医保接口出错，请与医保接口提供商联系!"
            Exit Function
        End If
    End If
    
    
    '下载信息
     Call 下载定点医疗机构
    mblnInit = True
    
    医保初始化_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 身份标识_重庆渝北(Optional bytType As Byte, Optional lng病人ID As Long) As String
    Dim str备注 As String, RSPATIENT As New ADODB.Recordset
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    
    身份标识_重庆渝北 = frmIdentify重庆渝北.GetPatient(bytType, lng病人ID)
    
End Function
Public Function 身份标识_重庆渝北2(ByVal strCard As String, ByVal strPass As String, Optional lng病人ID As Long) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    '/**?
    身份标识_重庆渝北2 = frmIdentify重庆渝北.GetPatient(3, lng病人ID)
    
End Function

Private Function Get病人信息(ByVal lng病人ID As Long)
    Dim rsTemp As New ADODB.Recordset
    '保险帐户目前存的值
    '--病人id, 险类, 中心, 卡号（医保卡号), 医保号(个人编号), 密码(支付类别 ), 人员身份(参保人员所在的社保经办机构代码), 单位编码(单位名称(单位编码)), 顺序号(无),
    '--退休证号(医疗人员类别|医疗照顾类别|医疗补助类别|累计缴费月数), 帐户余额(帐户余额), 当前状态, 病种id（病种ID), 在职(1), 年龄段(年龄), 灰度级, 就诊时间
    Dim strTemp As String
    Dim strArr
    
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "select a.卡号,a.医保号,a.密码,a.人员身份,a.单位编码,b.工作单位,a.顺序号,a.退休证号,a.帐户余额,a.当前状态,a.病种id,a.在职,a.年龄段,a.灰度级,a.就诊时间," & _
             "        b.姓名,decode( b.性别,'男','1','女','2','3') as 性别, b.年龄, b.出生日期, b.身份证号,A.就诊编号,A.结算编号,A.支付类别 " & _
             " from 保险帐户 a,病人信息 b " & _
             " WHERE a.病人id=" & lng病人ID & " AND a.病人id=b.病人id and a.险类=" & TYPE_重庆渝北
 
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人信息"
    
    With g病人身份_重庆渝北
        .卡号 = Nvl(rsTemp!卡号)
        .个人编号 = Nvl(rsTemp!医保号)
        .姓名 = Nvl(rsTemp!姓名)
        .性别 = Nvl(rsTemp!性别)
        .年龄 = Nvl(rsTemp!年龄段, 0)
        .出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
        .单位编码 = Val(Nvl(rsTemp!单位编码))
        
        strTemp = Nvl(rsTemp!工作单位)
        If InStr(1, strTemp, "(") <> 0 Then
            .单位名称 = Mid(strTemp, 1, InStr(1, strTemp, "(") - 1)
        Else
            .单位名称 = strTemp
        End If
        
        .密码 = Nvl(rsTemp!密码)
        .支付类别 = Nvl(rsTemp!支付类别)
        .社保经办构构代码 = Nvl(rsTemp!人员身份)
        strTemp = Nvl(rsTemp!退休证号, "|||")
        strTemp = IIf(strTemp = "", "|||", strTemp)
        strArr = Split(strTemp, "|")
        .医疗人员类别 = strArr(0)
        .医疗照顾类别 = strArr(1)
        .医疗补助类别 = strArr(2)
        .累计缴费月数 = Val(strArr(3))
        .帐户余额 = Nvl(rsTemp!帐户余额, 0)
        
        .身份证号 = Nvl(rsTemp!身份证号)
        .病种ID = Nvl(rsTemp!病种ID, 0)
        .就诊编号 = Nvl(rsTemp!就诊编号)
        .结算编号 = Nvl(rsTemp!结算编号)
        
        If .病种ID <> 0 Then
           gstrSQL = "Select 编码 From 医保病种目录 where id=" & .病种ID
           OpenRecordset_ZLYB rsTemp, "获取病种"
           If rsTemp.EOF Then
                .病种编码 = "00000"
           Else
                .病种编码 = Nvl(rsTemp!编码, "000000")
           End If
        Else
            .病种编码 = "000000"
        End If
    End With
Exit Function
errHand:
        DebugTool "获取病人信息失败" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & " 错误信息:" & Err.Description
End Function
Private Sub OpenRecordset_ZLYB(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_CQYB, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Private Function 下载定点医疗机构() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:下载定点医疗机构类别
    '--入参数:
    '--出参数:
    '--返  回:下载成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strXMLText As String
    Dim objStruct As Struct
    Dim objTest As Object
    Dim blnTrue  As Boolean
    
    strFile = gstrAppPath & "\定点信息.txt"
    
    下载定点医疗机构 = False
    
    Err = 0
    On Error GoTo errHand:
    If Not objFile.FolderExists(gstrAppPath) Then
        '不存在文件夹，需创建
        objFile.CreateFolder gstrAppPath
    End If
    
    objFile.CreateTextFile strFile, True
    
    DebugTool "进入(" & "下载定点医疗机构" & ")"
    objStruct.strErrMsg = Space(5000)
    Err = 0
    
    On Error GoTo errHand:
     '下载定点医疗构机信息
'            gobjYbTest.ExportKB01 strFile, objStruct
   ExportKB01 strFile, objStruct
     
     If objStruct.lngAppCode < 0 Then
        ShowMsgbox "下载定点医疗信息出错"
     End If
     
    Set objText = objFile.OpenTextFile(strFile)
    '存储过程参数:
    '病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 经办机构代码, 医疗人员类别, 医疗照顾类别, 医疗补助类别,
    '年度, 下限金额, 自付金额, 支付金额, 公务员补助, 先行自付金额, 累计缴费月数, 实足年龄, 医疗行政类别, 帐户支付, 分段标准,
    '全自费金额, 挂构自费, 本次自付, 本次支付金额, 公务员统筹支付, 补肋自付累计
    
    Call intXML
    blnTrue = False
    strXMLText = ""
    Do While Not objText.AtEndOfStream
        strXMLText = objText.ReadLine
        blnTrue = True
        Exit Do
    Loop
    If strXMLText = "" Then
        DebugTool "文件无内容(下载定点医疗机构)，文件:" & strFile
        Exit Function
    End If
    If GetXML串(strXMLText, False) = False Then
        DebugTool "XML格式无效，格式:" & strXMLText
        Exit Function
    End If
    With InitInfor_重庆渝北
        .经办机构代码 = GetXMLOutput("YAB003")
        .医保机构名称 = GetXMLOutput("AAB300")
        .医院编码 = GetXMLOutput("AKB020")
        .医疗机构名称 = GetXMLOutput("AKB021")
        .机构类别 = GetXMLOutput("AKB023")
        .定点状态标识 = GetXMLOutput("YKB002")
    End With
    Exit Function
errHand:
    DebugTool "下载定点医疗机构出错(下载定点医疗机构)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
    If InitInfor_重庆渝北.模拟数据 Then
        InitInfor_重庆渝北.经办机构代码 = "1200"
    End If
End Function
Public Function 医保终止_重庆渝北() As Boolean
    mblnInit = False
    If gcnOracle_CQYB.State = 1 Then
        gcnOracle_CQYB.Close
    End If
    Set gobjYingHaiDll = Nothing
    Set gobj费用结算 = Nothing
    
    医保终止_重庆渝北 = True
End Function
Public Function 身份鉴别_重庆渝北() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:远程身份鉴别
    '--入参数:
    '--出参数:
    '--返  回:成功true,否则False
    '-----------------------------------------------------------------------------------------------------------
    Dim strOutput As String
    Dim strXMLText As String
    Dim blnReturn As Boolean
    Err = 0
    On Error GoTo errHand:
    
    身份鉴别_重庆渝北 = False
    If intXML = False Then Exit Function
        
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "03"
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    AppendXMLNode gobjXMLInPut.documentElement, "code", Substr(g病人身份_重庆渝北.卡号, 1, 20)
    AppendXMLNode gobjXMLInPut.documentElement, "ykc005", Substr(g病人身份_重庆渝北.密码, 1, 6)
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_重庆渝北.医院编码, 1, 8)
    
    strXMLText = gobjXMLInPut.documentElement.xml
    '取掉前导XML串
    strXMLText = 取掉XML的前导标识(strXMLText)
        
    '业务请求
    
    blnReturn = 业务请求_重庆渝北(身份鉴别, strXMLText, strOutput)
    If blnReturn = False Then
        Exit Function
    End If
    
    '输出串
    strXMLText = strOutput
    
    '获取子项
    If GetXML串(strXMLText) = False Then
        ShowMsgbox "身份鉴别返回串是错误的XML串,不能继续!"
        Exit Function
    End If
    '给公用变量赋值
    With g病人身份_重庆渝北
        .个人编号 = GetXMLOutput("aac001")
        .身份证号 = GetXMLOutput("aac002")
        .姓名 = GetXMLOutput("aac003")
        .性别 = GetXMLOutput("aac004")
        .出生日期 = GetXMLOutput("aac006")
        .医疗人员类别 = GetXMLOutput("akc021")
        .医疗照顾类别 = GetXMLOutput("ykc120")
        .医疗补助类别 = GetXMLOutput("ykc121")
        .社保经办构构代码 = GetXMLOutput("yab003")
        .单位编码 = Val(GetXMLOutput("aab001"))
        .单位名称 = GetXMLOutput("aab004")
        .累计缴费月数 = Val(GetXMLOutput("ykc021"))
        .年龄 = Val(GetXMLOutput("akc023"))
        .帐户余额 = Val(GetXMLOutput("LastBaseICUsable")) + Val(GetXMLOutput("PastBaseICUsable")) + Val(GetXMLOutput("LastOfficialICUsable")) + Val(GetXMLOutput("PastOfficialICUsable")) + Val(GetXMLOutput("ThisBaseICUsable")) + Val(GetXMLOutput("ThisOfficialICUsable"))
    End With
    
    身份鉴别_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    身份鉴别_重庆渝北 = False
End Function
Private Function 取掉XML的前导标识(ByVal strXMLText As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:取掉XML的前导标识
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strXML As String
    
     strXML = Substr(strXMLText, Len(gstrXMLRootPart) + 3, LenBString(strXMLText) - Len(gstrXMLRootPart) * 2 - 5)
     If Right(strXML, 2) = "</" Then
        strXML = Mid(strXML, 1, Len(strXML) - 2)
     End If
    取掉XML的前导标识 = strXML
End Function
Private Function LenBString(ByVal strTxt As String) As Long
     LenBString = LenB(StrConv(strTxt, vbFromUnicode))
End Function

Private Function 资格审核待遇核定(ByVal lng病人ID As Long, ByVal str开始享用时间 As String, ByVal str结束享用时间 As String, Optional bln保存审核信息 As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:进行待遇核定
    '--入参数:
    '--出参数:
    '--返  回:记录集
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText  As String
    Dim strOutput As String
    
    If intXML = False Then Exit Function
          
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(g病人身份_重庆渝北.社保经办构构代码, 1, 4)
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "06"
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    AppendXMLNode gobjXMLInPut.documentElement, "code", Substr(g病人身份_重庆渝北.卡号, 1, 20)
    AppendXMLNode gobjXMLInPut.documentElement, "ChkCardSymbol", 2
    
    AppendXMLNode gobjXMLInPut.documentElement, "ykc005", Substr(g病人身份_重庆渝北.密码, 1, 6)
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", Substr(g病人身份_重庆渝北.就诊编号, 1, 20)
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_重庆渝北.医院编码, 1, 8)
    AppendXMLNode gobjXMLInPut.documentElement, "aka123", IIf(g病人身份_重庆渝北.病种ID = 0, 0, 1)
    
    AppendXMLNode gobjXMLInPut.documentElement, "yka026", Substr(g病人身份_重庆渝北.病种编码, 1, 20)
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", Substr(g病人身份_重庆渝北.支付类别, 1, 6)
    
    '-目前只有一种就诊结算方式,就是
    AppendXMLNode gobjXMLInPut.documentElement, "yka222", Substr(g病人身份_重庆渝北.就诊结算方式, 1, 6)
    AppendXMLNode gobjXMLInPut.documentElement, "akc192", str开始享用时间
    AppendXMLNode gobjXMLInPut.documentElement, "akc194", str结束享用时间
    AppendXMLNode gobjXMLInPut.documentElement, "SaveSymbol", IIf(bln保存审核信息, 2, 1)
    
    strXMLText = gobjXMLInPut.documentElement.xml
    
    '取掉前导XML串
    strXMLText = 取掉XML的前导标识(strXMLText)
        
       
    '业务请求
    资格审核待遇核定 = 业务请求_重庆渝北(资格审批待遇核定, strXMLText, strOutput)
    If 资格审核待遇核定 = False Then
        Exit Function
    End If
    
    '输出串
    strXMLText = strOutput
    
    资格审核待遇核定 = False
    '验证XML是否正确
    If GetXML串(strXMLText) = False Then
        ShowMsgbox "资格审批待遇核定返回串是错误的XML串,不能继续!"
        Exit Function
    End If
    
    资格审核待遇核定 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function Save审批信息(ByVal lng病人ID As Long, Optional bln虚拟结算 As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存审批信息
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strHead As String
    Dim strXMLText As String
    Dim strFile As String   '保存审批文件
    Dim strText As String
    Dim strTemp As String
    Dim str医疗机构类别 As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim blnYes As Boolean
    
    
    'aac001  Number  15  0   个人编号
    'aae073  Number  15  0   审批编号
    'akb021  String  50      定点医疗服务机构名称
    'akc190  String  20      就诊编号
    'akb020  String  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    'ykb012  String  8       转诊前服务机构编号
    'akb023  String  6       医疗机构类别，见代码表
    'aac002  String  18      身份号码
    'aac003  String  20      姓名
    'aac004  String  1       性别，见代码表
    'aac006  Date        日  出生日期
    'yab003  String  4       参保人员所在的社保经办机构代码，足四位长
    'aab001  Number  15  0   单位编号
    'aab004  String  50      单位名称
    'PastBaseICUsable    Number  14  2   基本医疗历年IC卡可圈存额(磁卡模式下等于账户真实余额)
    'LastBaseICUsable    Number  14  2   基本医疗上年IC卡可圈存额(磁卡模式下等于账户真实余额)
    'ThisBaseICUsable    Number  14  2   基本医疗本年IC卡可圈存额(磁卡模式下等于账户真实余额)
    'NotPastBaseICUsable Number  14  2   基本医疗本年划入非本年账户历年IC卡可圈存额(磁卡模式下等于账户真实余额)
    'NotLastBaseICUsable Number  14  2   基本医疗本年划入非本年账户上年IC卡可圈存额(磁卡模式下等于账户真实余额)
    'NotThisBaseICUsable Number  14  2   基本医疗本年划入非本年账户本年IC卡可圈存额(磁卡模式下等于账户真实余额)
    'PastOfficialICUsable    Number  14  2   公务员历年IC卡可圈存额磁卡模式下等于账户真实余额)
    'LastOfficialICUsable    Number  14  2   公务员上年IC卡可圈存额(磁卡模式下等于账户真实余额)
    'ThisOfficialICUsable    Number  14  2   公务员本年IC卡可圈存额(磁卡模式下等于账户真实余额)
    
    
    'ykc114  Number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
    'ykc007  String  1       是否采用新的结算标准，'0' 不需要，'1' 需要
    'akc021  String  6       医疗人员类别，见代码表
    'ykc021  Number  3       累计缴费月数
    'akc023  Number  3       实足年龄
    'yka114  Number  14  2   起付标准
    'yka115  Number  14  2   本次起付线
    'yka116  Number  14  2   起付线支付累计
    'yka117  Number  14  2   本次门诊可补助限额
    'yka118  Number  14  2   门诊支付累计
    'yka203  Number  14  2   本次基本医疗支付限额标准
    'yka119  Number  14  2   本次基本医疗支付限额
    'yka120  Number  14  2   本次基本医疗进入统筹累计
    'yka204  Number  14  2   本次大额医疗限额标准
    'yka121  Number  14  2   本次大额医疗支付限额
    'yka122  Number  14  2   本次大额医疗进入统筹累计
    'yka123  Number  14  2   本次公务员支付限额
    'yka124  Number  14  2   本次公务员进入统筹累计
    'ykc008  String  4000        可享受信息:待遇享受期,
    '                           取其中OutputString中的ykc008字段的后8位，格式为'yyyymmdd'，无享受信息时值为'xxxxxxxx'，享受期与当前时间比较，在三个月以内则给出提示。
    'ykc022  Number  3   0   计算起付线累计住院次数
    'ykc006  Number  3   0   计算统筹累计住院次数
    'ykc141  Number  14  2   本次结存账户限额
    'ykc142  Number  14  2   本年结存账户支付累计
    'yka125  Number  14  2   统筹可支付类就诊个人自付部分累计
    'yka126  number  14  2   统筹不可支付类就诊个人自付部分累计
    'ykc120  string  6       医疗照顾类别，见代码表
    'ykc121  string  6       享受医疗补助类别，见代码表
    'yka273  number  14  2   本次特殊门诊支付限额标准
    'yka274  number  14  2   本次特殊门诊支付限额
    'yka275  number  14  2   本年收入总额
    'akc315  string  6       医疗待遇行政类别，见代码表
    'ykc054  number  14  2   本年特殊门诊基本医疗支付累计
    
    
    '过程参数:
    '   审批编号, 病人id, 个人编号, 服务机构名称, 就诊编号, 医院编号, 服务机构编号, 医疗机构类别, 经办机构代码, 单位编号, 单位名称,
    '   帐户余额, 审批记录号, 新结算标准, 医疗人员类别, 累计缴费月数, 实足年龄, 起付标准, 本次起付线, 起付线支付累计, 门诊补助限额, 门诊支付累计, 基本支付标准,
    '   基本支付限额, 基本进入累计, 大额限额标准, 大额支付限额, 大额进入累计, 公务员支付限额, 公务员进行累计, 可享受信息, 起付线累计次数, 统筹累计次数,
    '   结存帐户限额, 结存帐户累计, 可支自付累计, 不可支自付累计, 照顾类别, 补助类别, 特殊门诊标准, 特殊门诊限额, 本年收入总额, 待遇行政类别,
    '   特殊门诊支付累计
    
    '公用信息部份
    strFile = gstrAppPath & "\待遇审批信息.txt"
    DebugTool ("进入Save审批信息:" & strFile)
    If Not objFile.FolderExists(gstrAppPath) Then
        '不存在文件夹，需创建
        objFile.CreateFolder gstrAppPath
    End If
    If Not objFile.FileExists(strFile) Then
        objFile.CreateTextFile strFile, True
    End If

    Set objText = objFile.OpenTextFile(strFile, ForWriting)
    
    Err = 0
    On Error GoTo errHand:
    
    g病人身份_重庆渝北.帐户余额 = Val(GetXMLOutput("LastBaseICUsable")) + Val(GetXMLOutput("PastBaseICUsable")) + Val(GetXMLOutput("LastOfficialICUsable")) + Val(GetXMLOutput("PastOfficialICUsable")) + Val(GetXMLOutput("ThisBaseICUsable")) + Val(GetXMLOutput("ThisOfficialICUsable"))
    g病人身份_重庆渝北.就诊编号 = GetXMLOutput("akc190")
    
    
    
    strHead = "ZL_资格审批待遇核定_INSERT("
    'aae073  Number  15  0   审批编号
    strHead = strHead & Val(GetXMLOutput("aae073")) & ","
    strHead = strHead & lng病人ID & ",'"
    'aac001  Number  15  0   个人编号
    strHead = strHead & GetXMLOutput("aac001") & "','"
    'akb021  String  50      定点医疗服务机构名称
    strHead = strHead & GetXMLOutput("akb021") & "','"
    'akc190  String  20      就诊编号
    strHead = strHead & GetXMLOutput("akc190") & "','"
    'akb020  String  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    strHead = strHead & GetXMLOutput("akb020") & "','"
    'ykb012  String  8       转诊前服务机构编号
    strHead = strHead & GetXMLOutput("ykb012") & "','"
    'akb023  String  6       医疗机构类别，见代码表
    strTemp = GetXMLOutput("akb023")
    strHead = strHead & strTemp & "','"
    str医疗机构类别 = strTemp
    'yab003  String  4       参保人员所在的社保经办机构代码，足四位长
    strHead = strHead & GetXMLOutput("yab003") & "','"
    'aab001  Number  15  0   单位编号
    strHead = strHead & GetXMLOutput("aab001") & "','"
    'aab004  String  50      单位名称
    strHead = strHead & GetXMLOutput("aab004") & "',"
    
    strHead = strHead & g病人身份_重庆渝北.帐户余额 & ","
    g病人身份_重庆渝北.病种编码 = GetXMLOutput("yka026")
    
    '表体记录部份
    Dim lngCount As Long
    Dim lngRow As Long
    lngCount = GetOutXMLRows("ykc114")
    For lngRow = 0 To lngCount - 1
        gstrSQL = ""
        strText = ""
        'ykc114  Number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
        strTemp = GetXMLOutput("ykc114", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        strText = strText & g病人身份_重庆渝北.就诊编号 & vbTab
        strText = strText & g病人身份_重庆渝北.支付类别 & vbTab
        strText = strText & GetXMLOutput("akc021", , lngRow) & vbTab
        strText = strText & g病人身份_重庆渝北.医疗照顾类别 & vbTab
        strText = strText & g病人身份_重庆渝北.医疗补助类别 & vbTab
        strText = strText & str医疗机构类别 & vbTab
        '--跨年度标志
        strText = strText & "0" & vbTab
        strText = strText & g病人身份_重庆渝北.病种编码 & vbTab  'g病人身份_重庆渝北.病种编码
         
        'ykc007  String  1       是否采用新的结算标准，'0' 不需要，'1' 需要
        gstrSQL = gstrSQL & Val(GetXMLOutput("ykc007", , lngRow)) & ",'"
        'akc021  String  6       医疗人员类别，见代码表
        gstrSQL = gstrSQL & GetXMLOutput("akc021", , lngRow) & "',"
        'ykc021  Number  3       累计缴费月数
        gstrSQL = gstrSQL & Val(GetXMLOutput("ykc021", , lngRow)) & ","
        'akc023  Number  3       实足年龄
        strTemp = GetXMLOutput("akc023", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        strText = strText & Val(GetXMLOutput("ykc021", , lngRow)) & vbTab   '累计缴费月数
        strText = strText & Val(GetXMLOutput("ykc006", , lngRow)) & vbTab    '计算统筹支付累计住院次数
        
        'yka114  Number  14  2   起付标准
        strTemp = GetXMLOutput("yka114", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        'yka115  Number  14  2   本次起付线

        strTemp = GetXMLOutput("yka115", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        g病人身份_重庆渝北.本次起付线 = Val(strTemp)
        
        'yka116  Number  14  2   起付线支付累计
        strTemp = GetXMLOutput("yka116", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka117  Number  14  2   本次门诊可补助限额
        strTemp = GetXMLOutput("yka117", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka118  Number  14  2   门诊支付累计
        strTemp = GetXMLOutput("yka118", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        strTemp = GetXMLOutput("ykc141", , lngRow)  'ykc141  Number  14  2   本次结存账户限额
        strText = strText & Val(strTemp) & vbTab
        strTemp = GetXMLOutput("ykc142", , lngRow) 'ykc142  Number  14  2   本年结存账户支付累计
        strText = strText & Val(strTemp) & vbTab
        
        'yka203  Number  14  2   本次基本医疗支付限额标准
        strTemp = GetXMLOutput("yka203", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka119  Number  14  2   本次基本医疗支付限额
        strTemp = GetXMLOutput("yka119", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka120  Number  14  2   本次基本医疗进入统筹累计
        strTemp = GetXMLOutput("yka120", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka204  Number  14  2   本次大额医疗限额标准
        strTemp = GetXMLOutput("yka204", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka121  Number  14  2   本次大额医疗支付限额
        strTemp = GetXMLOutput("yka121", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka122  Number  14  2   本次大额医疗进入统筹累计
        strTemp = GetXMLOutput("yka122", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka123  Number  14  2   本次公务员支付限额
        strTemp = GetXMLOutput("yka123", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka124  Number  14  2   本次公务员进入统筹累计
        strTemp = GetXMLOutput("yka124", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ",'"
        strText = strText & Val(strTemp) & vbTab
        
        'yka125  Number  14  2   统筹可支付类就诊个人自付部分累计
        strTemp = GetXMLOutput("yka125", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        
        'yka126  number  14  2   统筹不可支付类就诊个人自付部分累计
        strTemp = GetXMLOutput("yka126", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        
        
        '1       number  14  2   基本医疗本年账户可支付额
         strTemp = GetXMLOutput("ThisBaseICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        '2       number  14  2   基本医疗上年账户可支付额
        '3       number  14  2   基本医疗历年账户可支付额
        strTemp = GetXMLOutput("LastBaseICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        strTemp = GetXMLOutput("PastBaseICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
         
        
        
        '4       number  14  2   基本医疗本年划入非本年账户本年可支付余额
        '5       number  14  2   基本医疗本年划入非本年账户上年可支付余额
        '6       number  14  2   基本医疗本年划入非本年账户历年可支付余额
        strTemp = GetXMLOutput("NotPastBaseICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        strTemp = GetXMLOutput("NotLastBaseICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        strTemp = GetXMLOutput("NotThisBaseICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        
        '7       number  14  2   公务员补助本年账户可支付额
        strTemp = GetXMLOutput("ThisOfficialICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        '8       number  14  2   公务员补助上年账户可支付额
        '9       number  14  2   公务员补助历年账户可支付额
        
        strTemp = GetXMLOutput("LastOfficialICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
        strTemp = GetXMLOutput("PastOfficialICUsable", , lngRow)
        strText = strText & Val(strTemp) & vbTab
 

        strText = strText & g病人身份_重庆渝北.社保经办构构代码 & vbTab

        
        'ykc008  String  4000        可享受信息
        strTemp = GetXMLOutput("ykc008", , lngRow)
        strText = strText & strTemp & vbTab
        gstrSQL = gstrSQL & strTemp & "',"
        
        If blnYes = False Then
            '提醒相关待遇信息
            '刘兴宏:2007/09/04
            strTemp = Right(strTemp, 8)
            strTemp = Mid(strTemp, 1, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7, 2)
            DebugTool ("进行待遇信息检查YKC008:" & strTemp)
            If UCase(strTemp) = UCase("xxxx-xx-xx") Or (Not IsDate(strTemp)) Then
                If MsgBox("该病人没有待遇享受,是否继续?", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
                    Exit Function
                End If
                blnYes = True
            Else
                If zlDatabase.Currentdate > DateAdd("m", -1 * InitInfor_重庆渝北.待遇提醒期限, CDate(strTemp)) Then
                    If MsgBox("该病人享受待遇时间超过了所设定的时间(待遇时间为" & Format(strTemp, "yyyy-mm-dd") & "),是否继续?", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
                blnYes = True
            End If
        End If
        'ykc022  Number  3   0   计算起付线累计住院次数
        
        gstrSQL = gstrSQL & Val(GetXMLOutput("ykc022", , lngRow)) & ","
        'ykc006  Number  3   0   计算统筹累计住院次数
        gstrSQL = gstrSQL & Val(GetXMLOutput("ykc006", , lngRow)) & ","
        
        'ykc141  Number  14  2   本次结存账户限额
        strTemp = GetXMLOutput("ykc141", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        'ykc142  Number  14  2   本年结存账户支付累计
        strTemp = GetXMLOutput("ykc142", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        
        'yka125  Number  14  2   统筹可支付类就诊个人自付部分累计
        strTemp = GetXMLOutput("yka125", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        
        'yka126  number  14  2   统筹不可支付类就诊个人自付部分累计
        strTemp = GetXMLOutput("yka126", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ",'"
        
        
        'ykc120  string  6       医疗照顾类别，见代码表
        gstrSQL = gstrSQL & GetXMLOutput("ykc120", , lngRow) & "','"
        
        'ykc121  string  6       享受医疗补助类别，见代码表
        gstrSQL = gstrSQL & Val(GetXMLOutput("ykc121", , lngRow)) & "',"
        'yka273  number  14  2   本次特殊门诊支付限额标准
        strTemp = GetXMLOutput("yka273", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka274  number  14  2   本次特殊门诊支付限额
        strTemp = GetXMLOutput("yka274", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ","
        strText = strText & Val(strTemp) & vbTab
        
        'yka275  number  14  2   本年收入总额
        strTemp = GetXMLOutput("yka275", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ",'"
        strText = strText & Val(strTemp) & vbTab
        
        'akc315  string  6       医疗待遇行政类别，见代码表
        strTemp = GetXMLOutput("akc315", , lngRow)
        gstrSQL = gstrSQL & strTemp & "',"
        strText = strText & strTemp & vbTab
        
        'ykc054  number  14  2   本年特殊门诊基本医疗支付累计
        strTemp = GetXMLOutput("ykc054", , lngRow)
        gstrSQL = gstrSQL & Val(strTemp) & ")"
        strText = strText & Val(strTemp) & vbTab
        
        '插入表数据
        gstrSQL = strHead & gstrSQL
        If bln虚拟结算 = False Then
            '在过程前用的事务处理
            gcnOracle_CQYB.Execute gstrSQL, , adCmdStoredProc
        End If
        '加入审批文件
        objText.WriteLine strText
    Next
    Save审批信息 = True
    DebugTool ("Save审批信息成功")
    objText.Close
    Exit Function
errHand:
    DebugTool "审批信息保存出错(Save审批信息)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
    objText.Close
End Function
Private Function GetOutXMLRows(ByVal STRNAME As String) As Long
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取XML行数
    '--入参数:
    '--出参数:
    '--返  回:行数
    '-----------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String
    Err = 0
    On Error Resume Next
    GetOutXMLRows = gobjXMLOutput.getElementsByTagName(STRNAME).Length
    If Err <> 0 Then
        strErrMsg = "错误序号:" & vbCrLf & "   " & Err.Description
    End If
    DebugTool "获取XML的记录行数(GetOutXMLRows)《 " & STRNAME & "》" & vbCrLf & strErrMsg
End Function
Private Function IsertInto医保明细(ByVal lng费用ID As Long, ByVal strNO As String, ByVal lng序号 As Long, ByVal lng记录性质 As Long, ByVal strCode As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:插入记录
    '--入参数:
    '--出参数:strCode-诊疗项目编码(主要用于挂号)
    '--返  回:插入记录
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    If g病人身份_重庆渝北.冲销 Then
        gstrSQL = " Select ID From 门诊费用记录 where no=[1] and 记录性质=[2] and 记录状态=3 and 序号=[3]" & _
                  " UNION " & _
                  " Select ID From 住院费用记录 where no=[1] and 记录性质=[2] and 记录状态=3 and 序号=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "退单流水号", strNO, lng记录性质, lng序号)
        If rsTemp.EOF Then
            IsertInto医保明细 = False
            Exit Function
        End If
        gstrSQL = "Select 记帐流水号,审批人,审批人职称,审批标志,项目编码 From 医保明细费用 where 费用id= [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "退单数据插入", CLng(Nvl(rsTemp!ID)))
        '--ZL_医保明细费用_INSERT(
            '费用ID_IN IN 医保明细费用.费用ID%TYPE,
            '审批人_IN IN 医保明细费用.审批人%TYPE,
            '审批人职称_IN IN 医保明细费用.审批人职称%TYPE,
            '审批标志_IN IN 医保明细费用.审批标志%TYPE,
            '就诊编号_IN IN 医保明细费用.就诊编号%TYPE,
            '结算编号_IN IN 医保明细费用.结算编号%TYPE,
            '退单流水号_IN   IN 医保明细费用.退单流水号%type
            ')
        If rsTemp.RecordCount = 0 Then
            ShowMsgbox "原始费用单据不存在,请核查!"
            Exit Function
        End If
        gstrSQL = "ZL_医保明细费用_INSERT(" & _
         lng费用ID & ",'" & _
         Nvl(rsTemp!审批人) & "','" & _
         Nvl(rsTemp!审批人职称) & "'," & _
         Nvl(rsTemp!审批标志, 1) & ",'" & _
         g病人身份_重庆渝北.就诊编号 & "','" & _
         g病人身份_重庆渝北.结算编号 & "'," & _
         Nvl(rsTemp!记帐流水号, 0) & ",'" & _
         Nvl(rsTemp!项目编码) & "')"
         
         
    Else
        gstrSQL = "ZL_医保明细费用_INSERT(" & _
         lng费用ID & ",'" & _
         "" & "','" & _
         "" & "'," & _
         1 & ",'" & _
         g病人身份_重庆渝北.就诊编号 & "','" & _
         g病人身份_重庆渝北.结算编号 & "'," & _
           "NULL" & ",'" & _
           strCode & "')"
    End If
    Err = 0
    On Error GoTo errHand:
    Call SQLTest(App.ProductName, "插入医保明细数据", gstrSQL)
    gcnOracle_CQYB.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
    IsertInto医保明细 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    IsertInto医保明细 = False
End Function
Private Function Save医保明细数据(ByVal rs明细 As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strTemp  As String
    Dim strCode As String
    Err = 0
    On Error GoTo errHand:
    g病人身份_重庆渝北.费用总额 = 0
    With rs明细
         .MoveFirst
         Do While Not .EOF
                 If Nvl(!项目编码) = "" Then
                     ShowMsgbox "存在未构对医保项目,请在保险项目中设置相应的对应关系!"
                     Exit Function
                 End If
                 If g病人身份_重庆渝北.结算标志 = 3 Then
                    '取当前的结算编号
                    g病人身份_重庆渝北.结算编号 = Nvl(!结算号)
                    g病人身份_重庆渝北.就诊编号 = Nvl(!就诊编号)
                 End If
                 If g病人身份_重庆渝北.冲销 Then
                        If IsertInto医保明细(!ID, Nvl(!NO), Nvl(!序号, 0), Nvl(!记录性质, 0), "") = False Then Exit Function
                 Else
                    If g病人身份_重庆渝北.结算标志 = 2 And InitInfor_重庆渝北.收入项目id <> 0 Then
                       If Nvl(!收入项目id, 0) = InitInfor_重庆渝北.收入项目id Then
                          If frm诊疗项目编码选择.ShowCard(strCode) = False Then Exit Function
                          
                       End If
                    End If
                     '贵重药品需确定其价格的审批人
                     If Nvl(!实际价格, 0) > InitInfor_重庆渝北.单价限价 Then
                         strTemp = frm限价输入_渝北.Get审批信息(!ID, strCode)
                         '--刘兴宏:200507月更改,如按取消.需要退出
                         If strTemp = "" Then Exit Function
                     Else
                         IsertInto医保明细 !ID, Nvl(!NO), Nvl(!序号, 0), Nvl(!记录性质, 0), strCode
                     End If
                     
                 End If
             g病人身份_重庆渝北.费用总额 = g病人身份_重庆渝北.费用总额 + Nvl(!实收金额, 0)
             .MoveNext
         Loop
     End With
     Save医保明细数据 = True

    Exit Function
errHand:
  DebugTool "保存医保明细数据写入(Save医保明细数据)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
End Function
Private Function Get审批标志(ByVal lng费用ID As Long) As Long
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取审批标志
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select 记帐流水号,审批人,审批人职称,审批标志 From 医保明细费用 where 费用id= " & lng费用ID
    OpenRecordset_ZLYB rsTemp, "获取审批标志"
    If rsTemp.EOF Then
        Get审批标志 = 0
    Else
        Get审批标志 = Nvl(rsTemp!审批标志, 0)
    End If
End Function
Private Function Get诊疗项目编码(ByVal lng费用ID As Long)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select 项目编码 From 医保明细费用 where 费用id=" & lng费用ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取诊疗项目编码"
    If rsTemp.EOF Then
        Get诊疗项目编码 = ""
    Else
        Get诊疗项目编码 = Nvl(rsTemp!项目编码)
    End If
End Function
Private Function Save费用明细文本文件(ByVal rs明细 As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:产生明细的文本文件
    '--入参数:
    '--出参数:
    '--返  回:产生成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strText As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp明细 As New ADODB.Recordset
    Dim strFile As String
    
    strFile = gstrAppPath & "\结算明细信息.txt"
    
    Save费用明细文本文件 = False
    
    Err = 0
    On Error GoTo errHand:
    If Not objFile.FolderExists(gstrAppPath) Then
        '不存在文件夹，需创建
        objFile.CreateFolder gstrAppPath
    End If
    If Not objFile.FileExists(strFile) Then
        objFile.CreateTextFile strFile, True
    End If
    Set objText = objFile.OpenTextFile(strFile, ForWriting)
    
    
    If rs明细 Is Nothing Then Exit Function
    
    Dim byt审批 As Byte
    Dim lng流水号 As Long
    
    With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
                
                Set rsTemp = Get保险项目(Nvl(!项目编码))
                Set rsTmp明细 = Get医保明细费用(!ID)
                If rsTemp.RecordCount = 0 Then
                    ShowMsgbox "不存医保项目,请核查!"
                    Exit Function
                End If
                '贵重药品需确定其价格的审批人
                strText = Nvl(!就诊编号) & vbTab
                strText = strText & Nvl(rsTmp明细!记帐流水号, 0) & vbTab
                     
                If Nvl(rsTemp!新增标志, 0) = 1 Then
                    strText = strText & "59000000000010000" & vbTab  '为医院自己新增的编码,固定编码
                Else
                    strText = strText & Nvl(rsTemp!医保编码) & vbTab
                End If
                
                If g病人身份_重庆渝北.结算标志 = 2 And Nvl(!收入项目id) = InitInfor_重庆渝北.收入项目id Then
                    
                    strText = strText & Nvl(rsTmp明细!项目编码) & vbTab
                    strText = strText & Nvl(rsTmp明细!项目编码) & vbTab
                Else
                    strText = strText & Nvl(!项目编码) & vbTab
                    If Nvl(rsTemp!新增标志, 0) = 1 Then
                        
                        strText = strText & Nvl(rsTemp!标准编号) & vbTab
                    Else
                        strText = strText & Nvl(!项目编码) & vbTab
                    End If
                End If
                strText = strText & 1 & vbTab       '目前该比例只转值为1
                strText = strText & Nvl(rsTmp明细!结算编号) & vbTab
                strText = strText & Nvl(rsTmp明细!退单流水号) & vbTab
                strText = strText & Nvl(!实际价格, 0) & vbTab
                strText = strText & Nvl(!数量, 0) & vbTab
                strText = strText & Nvl(!实收金额, 0) & vbTab
                strText = strText & Nvl(rsTmp明细!审批标志, 0) & vbTab
                strText = strText & Nvl(!经办机构代码) & vbTab
                strText = strText & Format(!经办时间, "yyyy-mm-dd HH:MM:SS") & vbTab
                If Not rsTemp.EOF Then
                    strText = strText & Nvl(rsTemp!目录分类)
                Else
                    strText = strText & ""
                End If
                objText.WriteLine strText
            .MoveNext
        Loop
    End With
    objText.Close
    Save费用明细文本文件 = True
    Exit Function
errHand:
     DebugTool "明细信息保存出错(Save费用明细文本文件)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
    objText.Close
End Function

Private Function Get保险项目(ByVal str项目编号 As String) As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 目录分类,新增标志,医保编码,剂型,标准编号,商品代码,商品名,自付比例1 as 自付比例  From 医保服务项目目录 where 商品代码='" & str项目编号 & "'"
    With rsTemp
        .Open gstrSQL, gcnOracle_CQYB
    End With
    Set Get保险项目 = rsTemp
End Function
Private Function Get记帐流水号(ByVal lng费用ID As Long) As Long
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取记帐流水号
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 记帐流水号 From 医保明细费用 where 费用ID=" & lng费用ID
    
    Call SQLTest(App.ProductName, "获取记帐流水号", gstrSQL)
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle_CQYB
    Call SQLTest
    If rsTemp.EOF Then
        Get记帐流水号 = 0
    Else
        Get记帐流水号 = Nvl(rsTemp!记帐流水号, 0)
    End If
End Function
Private Function Get医保明细费用(ByVal lng费用ID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取记帐流水号
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    DebugTool "进入(" & "Get医保明细费用" & ")"
    
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "Select * From 医保明细费用 where 费用ID=" & lng费用ID
    
    Call SQLTest(App.ProductName, "获取医保明细费用", gstrSQL)
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle_CQYB
    Call SQLTest
    Set Get医保明细费用 = rsTemp
    
    Exit Function
errHand:
  DebugTool "获取医保明细费用出错(Get医保明细费用)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
End Function

Private Function Save历史费用结算结果文本(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional bln门诊 As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:产生明细的文本文件
    '--入参数:
    '--出参数:
    '--返  回:产生成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strText As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim rsTemp As New ADODB.Recordset
    Dim strFile As String
    
    strFile = gstrAppPath & "\历次结算信息.txt"
    
    Save历史费用结算结果文本 = False
    
    Err = 0
    On Error GoTo errHand:
    If Not objFile.FolderExists(gstrAppPath) Then
        '不存在文件夹，需创建
        objFile.CreateFolder gstrAppPath
    End If
    objFile.CreateTextFile strFile, True
    Set objText = objFile.OpenTextFile(strFile, ForWriting)
     
    If bln门诊 Then
        '门诊只是一个空文件
        Save历史费用结算结果文本 = True
        Exit Function
    End If
    
    
    gstrSQL = "Select * From 费用结算结果 where 就诊编号='" & g病人身份_重庆渝北.就诊编号 & "' and 病人id=" & lng病人ID & " order by 结算编号 "
    
    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open gstrSQL, gcnOracle_CQYB, adOpenStatic
        Do While Not .EOF
                    strText = Nvl(!就诊编号) & vbTab
                    strText = strText & Nvl(!结算编号) & vbTab
                    strText = strText & Nvl(!退单结算号) & vbTab
                    strText = strText & Nvl(!审批记录序号) & vbTab
                    strText = strText & Nvl(!经办机构代码) & vbTab
                    strText = strText & Nvl(!年度) & vbTab
                    strText = strText & Nvl(!分段标准) & vbTab
                    strText = strText & Nvl(!全自费金额, 0) & vbTab
                    strText = strText & Nvl(!挂构自费, 0) & vbTab
                    strText = strText & Nvl(!符合金额, 0) & vbTab
                    strText = strText & Nvl(!本次自付, 0) & vbTab
                    strText = strText & Nvl(!本次支付金额, 0) & vbTab
                    strText = strText & Nvl(!公务员统筹支付, 0) & vbTab
                    strText = strText & Nvl(!补助自付累计, 0) & vbTab
                    objText.WriteLine strText
                .MoveNext
            Loop
    End With
    objText.Close
    Save历史费用结算结果文本 = True
    Exit Function
errHand:
    DebugTool "历史费用结算结果查询出错(Save历史费用结算结果)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
    objText.Close
End Function
Private Function 费用结果分解(ByVal strFile As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:费用结果分解
    '--入参数:
    '--出参数:
    '--返  回:分析成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strText As String
    Dim strXMLText As String
    Dim strOutput As String
    Dim blnFirst As Boolean
    Dim objXMLItem As MSXML2.IXMLDOMElement
    Dim strXMLtext1 As String
    Dim dblTmp(0 To 11) As Double
    Dim dblSumMony(0 To 11) As Double
    Dim dblSumSubMony(0 To 11) As Double
    Dim dblSumSubmony1(0 To 11) As Double
    Dim rsTemp As New ADODB.Recordset
    Dim str经办时间 As String
    Dim i As Long
    
    str经办时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    '0-全自费:decode(分段标准,'02',yka056+yka057,0)
    '1-先自付部份:decode(分段标准,'05',yka057,'06',yka057,0)
    '2-符合金额:decode(分段标准,'03',yka11,'04',yka11,'05',yka11,'06',yka11,'07',yka11,'10',yka11,0)
    '3-费用总额:decode(分段标准,'02','yka056,'02','yka053+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
    '4-进入起付线:decode(分段标准,'03',yka106+yka057,0)
    '5-基本医疗自付:decode(分段标准,'05',yka106,0)
    '6-基本医疗统筹支付:decode(分段标准,'05',yka107,0)
    '7-大额自付:decode(分段标准,'06','yka106,0)
    '8-大额支付:decode(分段标准,'06','yka107,0)
    '门诊
    '   9-公务员补助:decode(分段标准,'05',yka063,'07',yka063,0)
    
    '住院
    '   9-公务员补助:decode(分段标准,'07',yka063,0)
    '10-超限自付:decode(分段标准,'08',yka106+yka057,0)
    '11-单病种超限decode(分段标准,'11','yka056,0)
    
    Dim strTemp As String
    Dim strValues As String
    Dim strvalues1 As String
    Dim strArr
    Dim str基本信息 As String
    Dim str审批记录号  As String
    
    Dim bytType As Byte         '1-基本信息写入失败,2-费用结果写入失败,3-扣个人帐户失败
    
    DebugTool "费用结果分解!文件为:" & strFile
    
    Err = 0
    On Error GoTo errHand:
    
    Set objText = objFile.OpenTextFile(strFile)
    
    '存储过程参数:
    'ID,病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 经办机构代码, 医疗人员类别, 医疗照顾类别, 医疗补助类别,
    '年度, 下限金额, 自付金额, 支付金额, 公务员补助, 先行自付金额, 累计缴费月数, 实足年龄, 医疗行政类别, 帐户支付, 分段标准,
    '全自费金额, 挂构自费, 本次自付, 本次支付金额, 公务员统筹支付, 补肋自付累计
    
    Call intXML
    blnFirst = True
    For i = 0 To 11
        dblSumMony(i) = 0
        dblSumSubMony(i) = 0
        dblSumSubmony1(i) = 0
        dblTmp(i) = 0
    Next
    Dim lngID As Long
    
    Do While Not objText.AtEndOfStream
            
        gstrSQL = "Select 费用结算结果_ID.nextval as ID from dual"
        OpenRecordset_ZLYB rsTemp, "获取结算编号"
        lngID = Nvl(rsTemp!ID, 0)
        
        strTemp = Trim(objText.ReadLine)
        strArr = Split(strTemp, vbTab)
        
        strSQL = "ZL_费用结算结果_INSERT("
        
        strSQL = strSQL & lngID & ","
        strSQL = strSQL & lng病人ID & ","
        strSQL = strSQL & IIf(lng主页ID = 0, "Null", lng主页ID) & ","
        strSQL = strSQL & "'" & strArr(0) & "',"
        strSQL = strSQL & "'" & strArr(1) & "',"
        strSQL = strSQL & "'" & strArr(2) & "',"
        strSQL = strSQL & "" & Val(strArr(3)) & ","
        strSQL = strSQL & "'" & strArr(4) & "',"
        strSQL = strSQL & "'" & strArr(5) & "',"
        strSQL = strSQL & "'" & strArr(6) & "',"
        strSQL = strSQL & "'" & strArr(7) & "',"  '医疗补助类别
        strSQL = strSQL & "'" & strArr(8) & "',"
        
        '10列为XML值,需分解值
        strXMLText = strArr(9)
        str基本信息 = strXMLText
        'GKC010  string  800     基本医疗子段信息，下面缩排的元素是它的子元素
        'SubRkn 下面缩排的元素是它的子元素
        '    AKA160  number  14  2   子段下限金额
        '    YKA106  number  14  2   自付金额
        '    YKA 107 number  14  2   支付金额
        '    YKA 063 number  14  2   公务员补助金额
        '    YKA057  number  14  2   先行自付部分
        
        strSQL = strSQL & "'" & strArr(9) & "',"
        strSQL = strSQL & 0 & ","
        strSQL = strSQL & 0 & ","
        strSQL = strSQL & 0 & ","
        strSQL = strSQL & 0 & ","
        strSQL = strSQL & 0 & ","
        
        '累计缴费月数
        strSQL = strSQL & "" & Val(strArr(10)) & ","
        strSQL = strSQL & "" & Val(strArr(11)) & ","
        strSQL = strSQL & "'" & strArr(12) & "',"
        
        strSQL = strSQL & "" & Val(strArr(13)) & ","
        '分段标准
        strSQL = strSQL & "'" & strArr(22) & "',"
        
        strSQL = strSQL & "" & Val(strArr(23)) & ","
        
        strSQL = strSQL & "" & Val(strArr(24)) & ","
        strSQL = strSQL & "" & Val(strArr(25)) & ","
        strSQL = strSQL & "" & Val(strArr(26)) & ","
        strSQL = strSQL & "" & Val(strArr(27)) & ","
        strSQL = strSQL & "" & Val(strArr(28)) & ","
        strSQL = strSQL & "" & Val(strArr(29)) & ")"
        
        If g病人身份_重庆渝北.虚拟结算 Then
            '虚拟结算用不着保存相关数据
        Else
            '存入数据库中
            gcnOracle_CQYB.Execute strSQL, , adCmdStoredProc
            If insertInto子项(lngID, str基本信息) = False Then
                DebugTool "插入子项目错误!"
            End If
        End If
        
        'XML费用结果写入
        If blnFirst Then
            AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(g病人身份_重庆渝北.社保经办构构代码, 1, 4)
            AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "12"
            AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    
            'BaseInfo                多个待遇享受段共有的相同的基本信息部分，下面缩排的元素是它的子元素
            Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "BaseInfo", "")
            '    akc190  string  20      就诊编号
            AppendXMLNode objXMLItem, "akc190", strArr(0)
            '    yka103  string  20      结算编号
            AppendXMLNode objXMLItem, "yka103", strArr(1)
            '    yka198  string  20      退单对应结算编号
            AppendXMLNode objXMLItem, "yka198", strArr(2)
            '    ykc114  number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
            AppendXMLNode objXMLItem, "ykc114", strArr(3)
            '    yab003  string  4       社保经办机构代码
            AppendXMLNode objXMLItem, "yab003", strArr(4)
            strValues = ""
            
            strValues = strValues & strArr(0) & vbTab
            strValues = strValues & strArr(1) & vbTab
            strValues = strValues & strArr(2) & vbTab
            strValues = strValues & Val(strArr(3)) & vbTab
            str审批记录号 = strArr(3)
            strValues = strValues & strArr(4) & vbTab
            strValues = strValues & strArr(12) & vbTab
            blnFirst = False
        End If
        
        '需确定相关的字串
        'ReckonInfo              多个待遇享受段的结算分段信息，下面缩排的元素是它的子元素
        Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "ReckonInfo", "")
        
        'akc190  string  20      就诊编号
         AppendXMLNode objXMLItem, "akc190", strArr(0)
        'yka103  string  20      结算编号
         AppendXMLNode objXMLItem, "yka103", strArr(1)
        'yka198  string  20      退单对应结算编号
         AppendXMLNode objXMLItem, "yka198", strArr(2)
        'ykc114  number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
         AppendXMLNode objXMLItem, "ykc114", strArr(3)
        'yab003  string  4       社保经办机构代码
         AppendXMLNode objXMLItem, "yab003", strArr(4)
        'aka213  string  2       分段标准，03 起付线， 05 基本医疗 ，06 大额医疗，07 超限
         AppendXMLNode objXMLItem, "aka213", strArr(22)
        'yka056  number  14  2   全自费金额
         AppendXMLNode objXMLItem, "yka056", strArr(23)
        'yka057  number  14  2   挂钩自费金额
         AppendXMLNode objXMLItem, "yka057", strArr(24)
        'yka111  number  14  2   符合范围金额
         AppendXMLNode objXMLItem, "yka111", strArr(25)
        'yka106  number  14  2   自付金额
         AppendXMLNode objXMLItem, "yka106", strArr(26)
        'yka107  number  14  2   支付金额
         AppendXMLNode objXMLItem, "yka107", strArr(27)
        'yka063  number  14  2   公务员统筹支付金额
         AppendXMLNode objXMLItem, "yka063", strArr(28)
        'yka221  number  14  2   享受医疗补助个人自付累计金额
         AppendXMLNode objXMLItem, "yka221", strArr(29)
        'Akc315  String  3       医疗行政职务
         AppendXMLNode objXMLItem, "Akc315", strArr(12)
         
        
        '根据分段标准,计算相的值
        '0-全自费:decode(分段标准,'02',yka056+yka057,0)
        '1-先自付部份:decode(分段标准,'05',yka057,'06',yka057,0)
        '2-符合金额:decode(分段标准,'03',yka11,'04',yka11,'05',yka11,'06',yka11,'07',yka11,'10',yka11,0)
        '门诊
        '3-费用总额:decode(分段标准,'02','yka056,'02','yka057+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka063+yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
        
        '住院
        '3-费用总额:decode(分段标准,'02','yka056,'02','yka057+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
        
        '4-进入起付线:decode(分段标准,'03',yka106+yka057,0)
        '5-基本医疗自付:decode(分段标准,'05',yka106,0)
        
        '6-基本医疗统筹支付:decode(分段标准,'05',yka107,0)
        '7-大额自付:decode(分段标准,'06','yka106,0)
        '8-大额支付:decode(分段标准,'06','yka107
        '门诊
        '   9-公务员补助:decode(分段标准,'05',yka063,'03',yka063,'07',yka063,0)
        
        '住院
        '   9-公务员补助:decode(分段标准,'07',yka063,0)
        '10-超限自付:decode(分段标准,'08',yka106+yka057,0)
        '11-单病种超限decode(分段标准,'11','yka056,0)
        '求和
        
        dblTmp(0) = Decode(strArr(22), "02", Val(strArr(23)) + Val(strArr(24)), 0)
        dblTmp(1) = Decode(strArr(22), "05", Val(strArr(24)), "06", Val(strArr(24)), 0)
        dblTmp(2) = Decode(strArr(22), "03", Val(strArr(25)), "03", Val(strArr(25)), "04", Val(strArr(25)), "05", Val(strArr(25)), "06", Val(strArr(25)), "07", Val(strArr(25)), "10", Val(strArr(25)), 0)
        
        If g病人身份_重庆渝北.结算标志 = 1 Then
            dblTmp(3) = Decode(strArr(22), "02", Val(strArr(23)) + Val(strArr(24)) + Val(strArr(28)), "03", Val(strArr(26)) + Val(strArr(24)), "07", Val(strArr(28)), "04", Val(strArr(25)), "05", Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "06", Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "08", Val(strArr(25)) + Val(strArr(24)), "10", Val(strArr(25)), "11", Val(strArr(23)), 0)
        Else
            'dblTmp(3) = Decode(strArr(22), "02", Val(strArr(23)) + Val(strArr(24)) + Val(strArr(28)), "03", Val(strArr(26)) + Val(strArr(24)), "07", Val(strArr(28)), "04", Val(strArr(25)), "05", Val(strArr(28)) + Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "06", Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "08", Val(strArr(25)) + Val(strArr(24)), "10", Val(strArr(25)), 0)
            '临时加入“yka063(03段)
            dblTmp(3) = Decode(strArr(22), "02", Val(strArr(23)) + Val(strArr(24)) + Val(strArr(28)), "03", Val(strArr(26)) + Val(strArr(24)) + Val(strArr(28)), "07", Val(strArr(28)), "04", Val(strArr(25)), "05", Val(strArr(28)) + Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "06", Val(strArr(26)) + Val(strArr(27)) + Val(strArr(24)), "08", Val(strArr(25)) + Val(strArr(24)), "10", Val(strArr(25)), "11", Val(strArr(23)), 0)
        End If
        
        dblTmp(4) = Decode(strArr(22), "03", Val(strArr(24)) + Val(strArr(26)), 0)
        dblTmp(5) = Decode(strArr(22), "05", Val(strArr(26)), 0)
        
        dblTmp(6) = Decode(strArr(22), "05", Val(strArr(27)), 0)
        
        dblTmp(7) = Decode(strArr(22), "06", Val(strArr(26)), 0)
        
        dblTmp(8) = Decode(strArr(22), "06", Val(strArr(27)), 0)
        
        If g病人身份_重庆渝北.结算标志 = 1 Then
            dblTmp(9) = Decode(strArr(22), "07", Val(strArr(28)), 0)
        Else
            'dblTmp(9) = Decode(strArr(22), "05", Val(strArr(28)), "07", Val(strArr(28)), 0)
            '增加了03段的yka063
            dblTmp(9) = Decode(strArr(22), "05", Val(strArr(28)), "03", Val(strArr(28)), "07", Val(strArr(28)), 0)
        End If
        dblTmp(10) = Decode(strArr(22), "08", Val(strArr(26)) + Val(strArr(24)), 0)
        dblTmp(11) = Decode(strArr(22), "11", Val(strArr(23)), 0)
        '结算结果多出一个费用分段标准AKA213='11',即单病种超限部分。该分段的全自费金额YKA056，即为单病种超限部分。这部分费用病人不用支付，医保中心也不予支付，由医院承担。故在结算结果处应显示这笔费用，同时也应加在本次结算的总费用上；


        '分别求和
        If strArr(1) = strArr(2) Then
            For i = 0 To 11
                dblSumSubMony(i) = dblSumSubMony(i) + dblTmp(i)
            Next
        Else
            For i = 0 To 11
                dblSumSubmony1(i) = dblSumSubmony1(i) + dblTmp(i)
            Next
            
            strvalues1 = strvalues1 & strArr(0) & vbTab
            strvalues1 = strvalues1 & strArr(1) & vbTab
            strvalues1 = strvalues1 & strArr(2) & vbTab
            strvalues1 = strvalues1 & Val(str审批记录号) & vbTab
            strvalues1 = strvalues1 & strArr(4) & vbTab
            strvalues1 = strvalues1 & strArr(12) & vbTab
        End If
        
        '求总和
        For i = 0 To 11
            dblSumMony(i) = dblSumMony(i) + dblTmp(i)
        Next
    Loop
    
    For i = 0 To 11
        dblSumMony(i) = Round(dblSumMony(i), 2)
    Next
    For i = 0 To 11
        dblSumSubmony1(i) = Round(dblSumSubmony1(i), 2)
    Next
    
    objText.Close
    
    '由于无虚拟结算,所以需将值写入
    If Get结算方式(dblSumMony) = False Then
        DebugTool "显示结算方式失败!"
        Exit Function
    End If
    DebugTool "显示结算方式成功!"
    
'    If Format(g病人身份_重庆渝北.费用总额, "###0.00;-###0.00;0;0") <> Format(dblSumMony(3), "###0.00;-###0.00;0;0") Then
'        Dim blnYes As Boolean
'        '费用总额与医保中心返回总额不致,不能进行结算
'        ShowMsgbox "本次结算总额(" & g病人身份_重庆渝北.费用总额 & ") 与" & vbCrLf & _
'                    "   中心返回的总额(" & dblSumMony(3) & ")不致产能结算?"
'        Exit Function
'    End If
    
    '写入费用结算结果
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    strXMLtext1 = strXMLText
        
        
    '0-全自费:decode(分段标准,'02',yka056+yka057,0)
    '1-先自付部份:decode(分段标准,'05',yka057,'06',yka057,0)
    '2-符合金额:decode(分段标准,'03',yka11,'04',yka11,'05',yka11,'06',yka11,'07',yka11,'10',yka11,0)
    '3-费用总额:decode(分段标准,'02','yka056,'02','yka053+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
    '4-进入起付线:decode(分段标准,'03',yka106+yka057,0)
    '5-基本医疗自付:decode(分段标准,'05',yka106,0)
    '6-基本医疗统筹支付:decode(分段标准,'05',yka107,0)
    '7-大额自付:decode(分段标准,'06','yka106,0)
    '8-大额支付:decode(分段标准,'06','yka107
    '门诊
        '   9-公务员补助:decode(分段标准,'05',yka063,'03',yka063,'07',yka063,0)
    
    '住院
    '   9-公务员补助:decode(分段标准,'07',yka063,0)
    '10-超限自付:decode(分段标准,'08',yka106+yka057,0)
    Dim dbl个人帐户 As Double
    Dim dbl现金   As Double
    
    
    
    i = 0
ECal2:
    
    DebugTool "开始写入费用结算基本信息!"
    
    '写入费用结算基本信息
    strArr = Split(strValues, vbTab)
    
    If g病人身份_重庆渝北.结算标志 = 1 Then
        dbl现金 = IIf(i = 99, dblSumSubmony1(0), dblSumSubMony(0))
        If 医保病人已经出院(lng病人ID) Then
            dbl个人帐户 = IIf(i = 99, dblSumSubmony1(1) + dblSumSubmony1(5) + dblSumSubmony1(7) + dblSumSubmony1(10) + dblSumSubmony1(4), dblSumSubMony(1) + dblSumSubMony(5) + dblSumSubMony(7) + dblSumSubMony(10) + dblSumSubMony(4))
        Else
            dbl现金 = dbl现金 + IIf(i = 99, dblSumSubmony1(1) + dblSumSubmony1(5) + dblSumSubmony1(7) + dblSumSubmony1(10) + dblSumSubmony1(4), dblSumSubMony(1) + dblSumSubMony(5) + dblSumSubMony(7) + dblSumSubMony(10) + dblSumSubMony(4))
            dbl个人帐户 = 0
        End If
        
        If g病人身份_重庆渝北.帐户余额 <= dbl个人帐户 Then
            dbl现金 = dbl现金 + dbl个人帐户 - g病人身份_重庆渝北.帐户余额
            dbl个人帐户 = g病人身份_重庆渝北.帐户余额
        End If
    Else
        dbl个人帐户 = dblSumMony(1) + dblSumMony(5) + dblSumMony(7) + dblSumMony(10) + dblSumMony(4)
        
        dbl现金 = dblSumMony(0)
        
        If g病人身份_重庆渝北.帐户余额 <= dbl个人帐户 Then
            dbl现金 = dbl现金 + dbl个人帐户 - g病人身份_重庆渝北.帐户余额
            dbl个人帐户 = g病人身份_重庆渝北.帐户余额
        End If
    End If
    dbl现金 = Round(dbl现金, 2)
    dbl个人帐户 = Round(dbl个人帐户, 2)
    
    
    '过程参数:
    '    病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 个人编号, 单位编号, 姓名, 性别, 出生日期, 实足年龄,
    '    累计缴费月数, 医疗人员类别, 医疗机构编码, 分支机构编码, 医疗机构类别, 特种病标志, 支付类别, 病种编码, 本次起付线,
    '    医疗费总额, 全自费总额, 挂钩自费总额, 符合范围总额, 个人帐户支付总额, 个人现金支付总额, 经办时间, 经办机构代码,
    '    医疗照顾类别 , 医疗补助类别, 就诊结算方式, 发票号, 备注, 分段计算情况, 医疗行政类别
    
    strSQL = "ZL_费用基本信息_INSERT(" & lng病人ID & ","
    strSQL = strSQL & IIf(lng主页ID = 0, "NULL", lng主页ID) & ","
    strSQL = strSQL & "'" & strArr(0) & "',"
    strSQL = strSQL & "'" & strArr(1) & "',"
    strSQL = strSQL & "'" & strArr(2) & "',"
    strSQL = strSQL & "" & Val(strArr(3)) & ","
    strSQL = strSQL & "" & g病人身份_重庆渝北.个人编号 & ","
    strSQL = strSQL & "" & g病人身份_重庆渝北.单位编码 & ","
    strSQL = strSQL & "'" & g病人身份_重庆渝北.姓名 & "',"
    strSQL = strSQL & "'" & g病人身份_重庆渝北.性别 & "',"
    
    If g病人身份_重庆渝北.出生日期 = "" Then
        strSQL = strSQL & "NULL,"
    Else
        strSQL = strSQL & "to_date('" & g病人身份_重庆渝北.出生日期 & "','yyyy-mm-dd'),"
    End If
    
    strSQL = strSQL & "" & g病人身份_重庆渝北.年龄 & ","
    strSQL = strSQL & "" & g病人身份_重庆渝北.累计缴费月数 & ","
    strSQL = strSQL & "'" & g病人身份_重庆渝北.医疗人员类别 & "',"
    strSQL = strSQL & "'" & InitInfor_重庆渝北.医院编码 & "',"
    strSQL = strSQL & "'" & "01" & "',"
    strSQL = strSQL & "'" & "" & "',"       '医疗机构类别
    strSQL = strSQL & "'" & IIf(g病人身份_重庆渝北.病种ID <> 0, "1", "0") & "',"
        
    strSQL = strSQL & "'" & g病人身份_重庆渝北.支付类别 & "',"
    strSQL = strSQL & "'" & g病人身份_重庆渝北.病种编码 & "',"
    
    strSQL = strSQL & "" & 0 & ","      '本次起付线
    If g病人身份_重庆渝北.结算标志 = 1 Then
        strSQL = strSQL & "" & IIf(i = 99, dblSumSubmony1(3), dblSumSubMony(3)) & ","
        strSQL = strSQL & "" & IIf(i = 99, dblSumSubmony1(0), dblSumSubMony(0)) & ","
        strSQL = strSQL & "" & IIf(i = 99, dblSumSubmony1(1), dblSumSubMony(1)) & ","
        strSQL = strSQL & "" & IIf(i = 99, dblSumSubmony1(2), dblSumSubMony(2)) & ","
    Else
        strSQL = strSQL & "" & dblSumMony(3) & ","
        strSQL = strSQL & "" & dblSumMony(0) & ","
        strSQL = strSQL & "" & dblSumMony(1) & ","
        strSQL = strSQL & "" & dblSumMony(2) & ","
    End If
    strSQL = strSQL & "" & Format(dbl个人帐户, "####0.00;-####0.00") & ","
    strSQL = strSQL & "" & Format(dbl现金, "####0.00;-####0.00") & ","
    strSQL = strSQL & "to_date('" & str经办时间 & "','yyyy-mm-dd HH24:mi:ss'),"
    strSQL = strSQL & "'" & strArr(4) & "',"
    strSQL = strSQL & "'" & g病人身份_重庆渝北.医疗照顾类别 & "',"
    strSQL = strSQL & "'" & g病人身份_重庆渝北.医疗补助类别 & "',"
    strSQL = strSQL & "'" & g病人身份_重庆渝北.就诊结算方式 & "',"
    strSQL = strSQL & "'" & g病人身份_重庆渝北.发票号 & "',"
    strSQL = strSQL & "'" & "" & "',"
    strSQL = strSQL & "'" & str基本信息 & "',"
    strSQL = strSQL & "'" & strArr(5) & "')"
            
    
    DebugTool "准备执行写入费用结算基本信息!SQL=:" & vbCrLf & strSQL
    If g病人身份_重庆渝北.虚拟结算 Then
        '虚拟结算不保存数据
    Else
        '保存数据
        gcnOracle_CQYB.Execute strSQL, , adCmdStoredProc
    End If
    DebugTool "写入费用结算基本信息成功"
    
    Call intXML
    
    'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(g病人身份_重庆渝北.社保经办构构代码, 1, 4)
    'SvrcID  string  2       远程数据服务标识，定值10, 标识大小写敏感，足两位长
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "10"
    'CtrInf  string  20      控制信息，预留, 标识大小写敏感
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    
    'akc190  string  20      就诊编号
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", strArr(0)
    'yka103  string  20      结算编号
    AppendXMLNode gobjXMLInPut.documentElement, "yka103", strArr(1)
    'yka198  string  20      退单对应结算编号
    AppendXMLNode gobjXMLInPut.documentElement, "yka198", strArr(2)
    
    'ykc114  number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
    AppendXMLNode gobjXMLInPut.documentElement, "ykc114", strArr(3)
    'aac001  number  15  0   个人编号
    AppendXMLNode gobjXMLInPut.documentElement, "aac001", g病人身份_重庆渝北.个人编号
    'aab001  number  15  0   单位编号
    AppendXMLNode gobjXMLInPut.documentElement, "aab001", g病人身份_重庆渝北.单位编码
    'aac003  string  20      姓名
    AppendXMLNode gobjXMLInPut.documentElement, "aac003", g病人身份_重庆渝北.姓名
    'aac004  string  1       性别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "aac004", g病人身份_重庆渝北.性别
    
    'aac006  date    日      出生日期
    AppendXMLNode gobjXMLInPut.documentElement, "aac006", g病人身份_重庆渝北.出生日期
    'akc023  number  3       实足年龄
    AppendXMLNode gobjXMLInPut.documentElement, "akc023", g病人身份_重庆渝北.年龄
    'ykc021  number  3       累计缴费月数
    AppendXMLNode gobjXMLInPut.documentElement, "ykc021", g病人身份_重庆渝北.累计缴费月数
    'akc021  string  6       医疗人员类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "akc021", g病人身份_重庆渝北.医疗人员类别
    'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", InitInfor_重庆渝北.医院编码
    'ykb006  string  3       定点医疗机构分支机构编号
    AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"            '分支机构编码
    'akb023  string  6       医疗机构类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "akb023", InitInfor_重庆渝北.机构类别
    
    'aka123  string  1       特种病标志，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "aka123", IIf(g病人身份_重庆渝北.病种ID <> 0, "1", "0")         '特种病标志
    'aka130  string  6       支付类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", g病人身份_重庆渝北.支付类别
    'yka026  string  20      病种编码
    AppendXMLNode gobjXMLInPut.documentElement, "yka026", g病人身份_重庆渝北.病种编码
    'yka115  number  14  2   本次起付线
    AppendXMLNode gobjXMLInPut.documentElement, "yka115", g病人身份_重庆渝北.本次起付线            '本次起付线
    
    If g病人身份_重庆渝北.结算标志 = 1 Then
        'yka055  number  14  2   医疗费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka055", IIf(i = 99, dblSumSubmony1(3), dblSumSubMony(3))            '
        'yka056  number  14  2   全自费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka056", IIf(i = 99, dblSumSubmony1(0), dblSumSubMony(0))             '
        'yka057  number  14  2   挂钩自费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka057", IIf(i = 99, dblSumSubmony1(1), dblSumSubMony(1))              '
        'yka111  number  14  2   符合范围总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka111", IIf(i = 99, dblSumSubmony1(2), dblSumSubMony(2))                 '
    Else
        'yka055  number  14  2   医疗费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka055", dblSumMony(3)                '
        'yka056  number  14  2   全自费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka056", dblSumMony(0)              '
        'yka057  number  14  2   挂钩自费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka057", dblSumMony(1)               '
        'yka111  number  14  2   符合范围总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka111", dblSumMony(2)                '
    End If
    
    'yka112  number  14  2   个人账户支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "yka112", dbl个人帐户                 '
    'yka113  number  14  2   个人现金支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "yka113", dbl现金                  '
    'aae036  date        秒  经办时间
    '经办时间
    AppendXMLNode gobjXMLInPut.documentElement, "aae036", str经办时间                  '
    'yab003  string  4       社保经办机构代码
    AppendXMLNode gobjXMLInPut.documentElement, "yab003", strArr(4)                  '
    'ykc120  string  6       医疗照顾类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "ykc120", g病人身份_重庆渝北.医疗照顾类别                   '
    'ykc121  string  6       享受医疗补助类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "ykc121", g病人身份_重庆渝北.医疗补助类别                    '
    'yka222  string  6       就诊结算方式
    AppendXMLNode gobjXMLInPut.documentElement, "yka222", g病人身份_重庆渝北.就诊结算方式                    '
    'yka110  string  20      发票号
    AppendXMLNode gobjXMLInPut.documentElement, "yka110", g病人身份_重庆渝北.发票号                                '
    'aae013  string  100     备注
    AppendXMLNode gobjXMLInPut.documentElement, "aae013", ""                              '
    
    'gkc010  string  800     分段计算情况(住院用)
    AppendXMLNode gobjXMLInPut.documentElement, "gkc010", "||GKC010_LXH||"                              '
    'akc315  string  3       医疗待遇行政类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "akc315", strArr(5)                              '
        
    '写入基本信息
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    strXMLText = Replace(strXMLText, "||GKC010_LXH||", str基本信息)
    DebugTool "产生XML串!XML:=:" & vbCrLf & strXMLText
    
    If g病人身份_重庆渝北.虚拟结算 Then
    Else
        DebugTool "开始向医保中心提交XML串" & vbCrLf & strXMLText
    
        If 业务请求_重庆渝北(结算基本信息写入, strXMLText, strOutput) = False Then
            DebugTool "向医保中心提交XML串失败"
            If g病人身份_重庆渝北.结算标志 = 1 And i = 99 Then
                '因为存在被冲销的基本信息,所以需冲销已经上传的明细记录
            Else
                '如果是基本信息写入失败,则直接退出即可
            End If
            Exit Function
        End If
        
        If g病人身份_重庆渝北.结算标志 = 1 And i <> 99 And Trim(strvalues1) <> "" Then
            '需进行第二次写入基本信息
            i = 99
            strValues = strvalues1
            GoTo ECal2:
        End If
        strOutput = ""
        '写结果
        DebugTool "开始向医保中心提交结算结果XML串" & vbCrLf & strXMLtext1
        If 业务请求_重庆渝北(结算结果写入, strXMLtext1, strOutput) = False Then
            '肯定传了明细基本信息,所以需传递相反数,冲销基本信息和结果
            Call 费用基本信息冲销(g病人身份_重庆渝北.结算编号)
            DebugTool "开始向医保中心提交结算结果失败"
            Exit Function
        End If
    
    End If
    
    
    '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
  '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(公务员补助),帐户累计支出_IN(大额支付),累计进入统筹_IN(基本医疗自付),累计统筹报销_IN,住院次数_IN,起付线(进入起付线),封顶线_IN(支付类别+10000),实际起付线_IN(单病种超限),
    '   发生费用金额_IN(发生费用),全自付金额_IN(全自付),首先自付金额_IN(首先自付),
    '   进入统筹金额_IN(符合金额),统筹报销金额_IN(基本医疗统筹支付),大病自付金额_IN(大额自付),超限自付金额_IN(超限自付),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算编号),主页ID_IN,中途结帐_IN,备注_IN(就诊编号)
     
     '0-全自费:decode(分段标准,'02',yka056+yka057,0)
    '1-先自付部份:decode(分段标准,'05',yka057,'06',yka057,0)
    '2-符合金额:decode(分段标准,'03',yka11,'04',yka11,'05',yka11,'06',yka11,'07',yka11,'10',yka11,0)
    '3-费用总额:decode(分段标准,'02','yka056,'02','yka053+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
    '4-进入起付线:decode(分段标准,'03',yka106+yka057,0)
    '5-基本医疗自付:decode(分段标准,'05',yka106,0)
    '6-基本医疗统筹支付:decode(分段标准,'05',yka107,0)
    '7-大额自付:decode(分段标准,'06','yka106,0)
    '8-大额支付:decode(分段标准,'06','yka107
    '9-公务员补助:decode(分段标准,'07',yka063,0)
    '10-超限自付:decode(分段标准,'08',yka106+yka057,0)
    
    Dim blnUpdate As Boolean
    DebugTool "开始保存结算记录"
    If g病人身份_重庆渝北.虚拟结算 Then
        '虚拟结算不保存数据
    Else
        Err = 0
        On Error Resume Next
        If g病人身份_重庆渝北.结算标志 = 0 Then
            #If gverControl < 2 Then
                blnUpdate = False
            #Else
                blnUpdate = True
            #End If
        Else
            If g病人身份_重庆渝北.结算标志 = 1 Then
                blnUpdate = m虚拟结算信息.bln验证
            Else
                blnUpdate = False
            End If
        End If
        
        gstrSQL = "zl_保险结算记录_insert(" & IIf(g病人身份_重庆渝北.结算标志 = 1, 2, 1) & "," & g病人身份_重庆渝北.结帐ID & "," & TYPE_重庆渝北 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          dblSumMony(9) & "," & dblSumMony(8) & "," & dblSumMony(5) & ",NULL,NULL," & dblSumMony(4) & "," & "1" & g病人身份_重庆渝北.支付类别 & "," & dblSumMony(11) & "," & _
           dblSumMony(3) & "," & dblSumMony(0) & "," & dblSumMony(1) & "," & _
            "" & dblSumMony(2) & "," & dblSumMony(6) & "," & dblSumMony(7) & "," & dblSumMony(10) & "," & dbl个人帐户 & ",'" & _
           g病人身份_重庆渝北.结算编号 & "'," & IIf(lng主页ID = 0, "NULL", lng主页ID) & "," & IIf(g病人身份_重庆渝北.中途结帐 = 1, "1", "NULL") & ",'" & _
           g病人身份_重庆渝北.就诊编号 & "'" & IIf(blnUpdate, ",1", "") & ")"
           
           
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
        If Err <> 0 Then
            DebugTool "更新保险结算记录时出错!" & vbCrLf & " 错误号:" & Err.Number & " 错误描述:" & Err.Description
            Err.Clear
            '肯定传了明细基本信息和费用结果的,所以需传递相反数,冲销基本信息和结果
            Call 费用基本信息冲销(g病人身份_重庆渝北.结算编号)
            Call 费用结算结果冲销(g病人身份_重庆渝北.结算编号)
            Exit Function
        End If
        '冲个人帐户
        If g病人身份_重庆渝北.结算标志 = 1 Then
            If 医保病人已经出院(g病人身份_重庆渝北.lng病人ID) Then
                '扣减个人帐户
                If IC卡帐户支付_重庆渝北(dbl个人帐户, str经办时间, g病人身份_重庆渝北.结算编号) = False Then
                
                    '肯定传了明细基本信息和费用结果的,所以需传递相反数,冲销基本信息和结果
                    Call 费用基本信息冲销(g病人身份_重庆渝北.结算编号)
                    Call 费用结算结果冲销(g病人身份_重庆渝北.结算编号)
                    Exit Function
                End If
            End If
        Else
            '扣减个人帐户
            If IC卡帐户支付_重庆渝北(dbl个人帐户, str经办时间, g病人身份_重庆渝北.结算编号) = False Then
                '肯定传了明细基本信息和费用结果的,所以需传递相反数,冲销基本信息和结果
                Call 费用基本信息冲销(g病人身份_重庆渝北.结算编号)
                Call 费用结算结果冲销(g病人身份_重庆渝北.结算编号)
                Exit Function
            End If
        End If
    
    End If
    DebugTool "保存结算记录成功"
    
    费用结果分解 = True
    
    Exit Function
errHand:
   DebugTool "费用结算结果出错(费用结果分解)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
   
    objText.Close
End Function
Private Function Get结算方式(ByVal strDblCur As Variant) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:根据传入值确定相应的结算方式
    '--入参数:
    '--出参数:str结算方式
    '--返  回:成功返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String
    Dim dbl个人帐户 As Double
    Dim dbl现金 As Double
    Dim dbl单病种超限 As Double
    Dim obj结算信息 As 结算信息
    'strDblCur根据分段标准,计算相的值
        '0-全自费:decode(分段标准,'02',yka056+yka057,0)
        '1-先自付部份:decode(分段标准,'05',yka057,'06',yka057,0)
        '2-符合金额:decode(分段标准,'03',yka11,'04',yka11,'05',yka11,'06',yka11,'07',yka11,'10',yka11,0)
        '3-费用总额:decode(分段标准,'02','yka056,'02','yka057+yka063,'03',yka106+yka057,'07',yka063,'04',yka111,'05',yka106+yka107+yka057,'06','yka106+yka107+yka057,'08',yka111+yka057,'10',yka111,0)
        
        '4-进入起付线:decode(分段标准,'03',yka106+yka057,0)
        '5-基本医疗自付:decode(分段标准,'05',yka106,0)
        
        '6-基本医疗统筹支付:decode(分段标准,'05',yka107,0)
        '7-大额自付:decode(分段标准,'06','yka106,0)
        '8-大额支付:decode(分段标准,'06','yka107
        '9-公务员补助:decode(分段标准,'07',yka063,0)
        '10-超限自付:decode(分段标准,'08',yka106+yka057,0)
        '11-单病种超限decode(分段标准,'11','yka056,0)
    '医保基金=基本医疗统筹支付+
    '个人帐户支付=先自付部份+基本医疗自付+超限自付+先自付部份
    '公务补助=公务员补助
    
    Err = 0
    On Error GoTo errHand:
    DebugTool "进入(" & "Get结算方式" & ")"
    
    dbl现金 = strDblCur(0)
    
    If g病人身份_重庆渝北.结算标志 = 1 Then
        If 医保病人已经出院(g病人身份_重庆渝北.lng病人ID) Then
            '出院结算需算个人帐户
            dbl个人帐户 = strDblCur(1) + strDblCur(5) + strDblCur(7) + strDblCur(10) + strDblCur(4)
        Else
            '中途结算无个人帐户支付
            dbl个人帐户 = 0
        End If
    Else
            dbl个人帐户 = strDblCur(1) + strDblCur(5) + strDblCur(7) + strDblCur(10) + strDblCur(4)
    End If
    
    
    If g病人身份_重庆渝北.帐户余额 <= dbl个人帐户 Then
        dbl现金 = dbl现金 + dbl个人帐户 - g病人身份_重庆渝北.帐户余额
        dbl个人帐户 = g病人身份_重庆渝北.帐户余额
    End If
    
    With obj结算信息
        .dbl大额补助 = Round(strDblCur(8), 2)
        .dbl单病种超限 = Round(strDblCur(11), 2)
        .dbl个人帐户 = Round(dbl个人帐户, 2)
        .dbl公务员补助 = Round(strDblCur(9), 2)
        .dbl医保基金 = Round(strDblCur(6), 2)
        .dbl费用总额 = Round(strDblCur(3), 2)
        .bln验证 = False
    End With
    str结算方式 = "||医保基金|" & obj结算信息.dbl医保基金
    str结算方式 = str结算方式 & "||大额补助|" & obj结算信息.dbl大额补助
    str结算方式 = str结算方式 & "||公务员补助|" & obj结算信息.dbl公务员补助
    str结算方式 = str结算方式 & "||个人帐户|" & obj结算信息.dbl个人帐户
    str结算方式 = str结算方式 & "||单病种超限|" & obj结算信息.dbl单病种超限
    
    If Round(g病人身份_重庆渝北.费用总额, 2) <> obj结算信息.dbl费用总额 Then
        '费用总额与医保中心返回总额不致,不能进行结算
        ShowMsgbox "本次结算总额(" & g病人身份_重庆渝北.费用总额 & ") 与" & vbCrLf & _
                    "   中心返回的总额(" & obj结算信息.dbl费用总额 & ")不致产能结算?"
        Exit Function
    End If
    
   Dim blnUpdate As Boolean
   
    '如果存在
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        g病人身份_重庆渝北.结算信息 = str结算方式
        If g病人身份_重庆渝北.虚拟结算 Then
            '虚拟结算不保存相关数据
             m虚拟结算信息 = obj结算信息
        Else
            ' 0-门诊,1-住院,2-挂号,3-住院处方登记
            If g病人身份_重庆渝北.结算标志 = 0 Or g病人身份_重庆渝北.结算标志 = 2 Then
                If g病人身份_重庆渝北.结算标志 = 2 Then
                    blnUpdate = True
                Else
                    #If gverControl < 2 Then
                        blnUpdate = True
                    #Else
                        blnUpdate = False
                    #End If
                End If
                
                If blnUpdate Then
                    gstrSQL = "zl_病人结算记录_Update(" & g病人身份_重庆渝北.结帐ID & ",'" & str结算方式 & "',0)"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
                Else
                    gstrSQL = "zl_医保核对表_Insert(" & g病人身份_重庆渝北.结帐ID & ",'" & str结算方式 & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "插入医保核对表")
                End If

            Else
                #If gverControl < 2 Then
                    blnUpdate = True
                #Else
                    blnUpdate = False
                #End If
                
                If m虚拟结算信息.dbl大额补助 <> obj结算信息.dbl大额补助 Or _
                    m虚拟结算信息.dbl单病种超限 <> obj结算信息.dbl单病种超限 Or _
                    m虚拟结算信息.dbl费用总额 <> obj结算信息.dbl费用总额 Or _
                    m虚拟结算信息.dbl个人帐户 <> obj结算信息.dbl个人帐户 Or _
                    m虚拟结算信息.dbl公务员补助 <> obj结算信息.dbl公务员补助 Or _
                    m虚拟结算信息.dbl医保基金 <> obj结算信息.dbl医保基金 Then
                    If blnUpdate Then
                        gstrSQL = "zl_病人结算记录_Update(" & g病人身份_重庆渝北.结帐ID & ",'" & str结算方式 & "',1)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
                    Else
                        m虚拟结算信息.bln验证 = True
                    End If
                End If
                
                If blnUpdate = False Then
                    gstrSQL = "zl_医保核对表_Insert(" & g病人身份_重庆渝北.结帐ID & ",'" & str结算方式 & "'," & IIf(m虚拟结算信息.bln验证, 1, 0) & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "插入医保核对表")
                End If
            End If
        End If
        If g病人身份_重庆渝北.虚拟结算 Then
            g病人身份_重庆渝北.结算信息 = Replace(g病人身份_重庆渝北.结算信息, "||", "[")
            g病人身份_重庆渝北.结算信息 = Replace(g病人身份_重庆渝北.结算信息, "|", ";")
            g病人身份_重庆渝北.结算信息 = Replace(g病人身份_重庆渝北.结算信息, "[", ";0|")
            g病人身份_重庆渝北.结算信息 = g病人身份_重庆渝北.结算信息 & ";0"
        End If
    End If
    
    '显示结算信息
    If g病人身份_重庆渝北.虚拟结算 Or g病人身份_重庆渝北.结算标志 = 1 Then
    Else
        #If gverControl < 2 Then
            If frm结算信息.ShowME(g病人身份_重庆渝北.结帐ID, True) = False Then
                Get结算方式 = False
                Exit Function
            End If
        #End If
    End If
    Get结算方式 = True
    Exit Function
errHand:
  DebugTool "保存病结算记录出错(Get结算方式)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
  
End Function
'20041012:刘兴宏:因为银海产生的文本文件的格式不对(但费用结果又是正确的.）造成读取记录为所有的，而不是行.所以只有采取按回车来单独处理这种文件
'Private Function Save费用明细结算分割(ByVal strFile As String) As Boolean
'    '-----------------------------------------------------------------------------------------------------------
'    '--功  能:保存费用结算后产生的明细
'    '--入参数:
'    '--出参数:
'    '--返  回:
'    '-----------------------------------------------------------------------------------------------------------
'    DebugTool "进入:Save费用明细结算分割"
'    Dim strSql As String
'    Dim objFile As New FileSystemObject
'    Dim objText As TextStream
'    Dim strText As String
'    Dim strTemp  As String
'    Dim strArr
'
'    Dim strXMLText As String
'
'    If g病人身份_重庆渝北.结算标志 <> 1 Then
'        '门诊部份,由于没有输出文本,所以无保存相关明细信息
'        Save费用明细结算分割 = True
'        Exit Function
'    End If
'
'    Err = 0
'    On Error GoTo ErrHand:
'
'    Set objText = objFile.OpenTextFile(strFile)
'    '明细过程参数(现暂无用):
'    '   记帐流水号,就诊编号, 项目编码, 商品名编码, 审批商品编码, 结算编号, 经办机构代码, 项目结算方式, 费用总额, 帐户支付额, 分段标准,
'    '   全自费金额, 挂钩自费金额, 符合范围金额, 自付金额, 支付金额, 公务员统筹支付, 补助自付累计, 自付比例
'
'
'    Do While Not objText.AtEndOfStream
'        strTemp = Trim(objText.ReadLine)
'        strArr = Split(strTemp, vbTab)
'            '文本格式
'                '            AKC190  string  20      就诊编号
'                '            YKA104  number  15  0   退单对应记账流水号
'                '            YKA002  string  20      医保项目编码
'                '            YKA231  string  20      医保项目商品名编码
'                '            YKA247  string  20      特殊审批医保项目商品名编码
'                '            YKA096  number  20      自付比例
'                '            YKA272  string  4       目录分类
'                '            AKC225  string  6       实际价格
'                '            AKC226  number  14  2   数量
'                '            YKA055  number  14  2   费用总额
'                '            YKA056  number  14  2   自付金额
'                '            YKA057  number  14  2   挂钩自付金额
'                '            YKA111  number  14  2   符合范围部分金额
'                '            YKA103  number  14  2   退单对应结算编号
'            '过程参数:
'            '        就诊编号_IN 医保费用分类信息.就诊编号%type,
'            '        退单流水号_IN   医保费用分类信息.退单流水号%type,
'            '        项目编码_IN 医保费用分类信息.项目编码%type,
'            '        商品名编码_IN   医保费用分类信息.商品名编码%type,
'            '        特殊商品名编码_IN   医保费用分类信息.特殊商品名编码%type,
'            '        自付比例_IN 医保费用分类信息.自付比例%type,
'            '        目录分类_IN 医保费用分类信息.目录分类%type,
'            '        实际价格_IN 医保费用分类信息.目录分类%type,
'            '        数量_IN     医保费用分类信息.数量%type,
'            '        费用总额_IN 医保费用分类信息.费用总额%type,
'            '        自付金额_IN 医保费用分类信息.自付金额%type,
'            '        挂钩自付金额_IN 医保费用分类信息.挂钩自付金额%type,
'            '        符合范围金额_IN 医保费用分类信息.符合范围金额%type,
'            '        退单结算编号_IN 医保费用分类信息.退单结算编号%type
'
'            strSql = "ZL_医保费用分类信息_INSERT("
'            strSql = strSql & "'" & strArr(0) & "',"
'            strSql = strSql & "" & Val(strArr(1)) & ","
'            strSql = strSql & "'" & strArr(2) & "',"
'            strSql = strSql & "'" & strArr(3) & "',"
'            strSql = strSql & "'" & strArr(4) & "',"
'            strSql = strSql & "" & Val(strArr(5)) & ","
'            strSql = strSql & "'" & strArr(6) & "',"
'            strSql = strSql & "" & Val(strArr(7)) & ","
'            strSql = strSql & "" & Val(strArr(8)) & ","
'            strSql = strSql & "" & Val(strArr(9)) & ","
'            strSql = strSql & "" & Val(strArr(10)) & ","
'            strSql = strSql & "" & Val(strArr(11)) & ","
'            strSql = strSql & "" & Val(strArr(12)) & ","
'            strSql = strSql & "" & Val(strArr(13)) & ")"
'
'
'            '只有住院才有
'            '20040720取消
'
'
'            '       StrSQL = "ZL_医保明细费用_UPDATE("
'
'            '            '记帐流水号
'            '            StrSQL = StrSQL & Val(strArr(1)) & ","
'            '            StrSQL = StrSQL & "'" & strArr(0) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(2) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(3) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(4) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(5) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(6) & "',"
'            '            StrSQL = StrSQL & "'" & strArr(7) & "',"
'            '
'            '            StrSQL = StrSQL & "" & Val(strArr(8)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(9)) & ","
'            '
'            '            '分段标准
'            '            StrSQL = StrSQL & "" & Val(strArr(18)) & ","
'            '
'            '            StrSQL = StrSQL & "" & Val(strArr(19)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(20)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(21)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(22)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(23)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(24)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(25)) & ","
'            '            StrSQL = StrSQL & "" & Val(strArr(26)) & ")"
'            '修改费用分类信息
'            gcnOracle_CQYB.Execute strSql, , adCmdStoredProc
'    Loop
'    objText.Close
'    Save费用明细结算分割 = True
'    Exit Function
'ErrHand:
'
'    DebugTool "明细分割保存(Save费用明细结算分割)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
'   objText.Close
'End Function
Private Function Save费用明细结算分割(ByVal strFile As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存费用结算后产生的明细
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    '20041012:刘兴宏:因为银海产生的文本文件的格式不对(但费用结果又是正确的.）造成读取记录为所有的，而不是行.所以只有采取按回车来单独处理这种文件
    DebugTool "进入:Save费用明细结算分割"
    Dim strSQL As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strText As String
    Dim strTemp  As String
    Dim strArr
    Dim strArr1
    Dim i As Long
    
    Dim strXMLText As String
    
    If g病人身份_重庆渝北.结算标志 <> 1 Then
        '门诊部份,由于没有输出文本,所以无保存相关明细信息
        Save费用明细结算分割 = True
        Exit Function
    End If
    If g病人身份_重庆渝北.虚拟结算 Then
            '虚拟结算用不着保存相关数据
            Save费用明细结算分割 = True
            Exit Function
    End If
    Err = 0
    On Error GoTo errHand:
    
    Set objText = objFile.OpenTextFile(strFile)
    '明细过程参数(现暂无用):
    '   记帐流水号,就诊编号,病人ID, 项目编码, 商品名编码, 审批商品编码, 结算编号, 经办机构代码, 项目结算方式, 费用总额, 帐户支付额, 分段标准,
    '   全自费金额, 挂钩自费金额, 符合范围金额, 自付金额, 支付金额, 公务员统筹支付, 补助自付累计, 自付比例
    strText = Trim(objText.ReadAll)
    strArr1 = Split(strText, vbCr)
    
    For i = 0 To UBound(strArr1)
        If Trim(strArr1(i)) <> "" And Len(strArr1(i)) > 2 Then
            strArr = Split(strArr1(i), vbTab)
                '文本格式
                    '            AKC190  string  20      就诊编号
                    '            YKA104  number  15  0   退单对应记账流水号
                    '            YKA002  string  20      医保项目编码
                    '            YKA231  string  20      医保项目商品名编码
                    '            YKA247  string  20      特殊审批医保项目商品名编码
                    '            YKA096  number  20      自付比例
                    '            YKA272  string  4       目录分类
                    '            AKC225  string  6       实际价格
                    '            AKC226  number  14  2   数量
                    '            YKA055  number  14  2   费用总额
                    '            YKA056  number  14  2   自付金额
                    '            YKA057  number  14  2   挂钩自付金额
                    '            YKA111  number  14  2   符合范围部分金额
                    '            YKA103  number  14  2   退单对应结算编号
                '过程参数:
                '        病人ID,
                '        主页ID,
                '        就诊编号_IN 医保费用分类信息.就诊编号%type,
                '        退单流水号_IN   医保费用分类信息.退单流水号%type,
                '        项目编码_IN 医保费用分类信息.项目编码%type,
                '        商品名编码_IN   医保费用分类信息.商品名编码%type,
                '        特殊商品名编码_IN   医保费用分类信息.特殊商品名编码%type,
                '        自付比例_IN 医保费用分类信息.自付比例%type,
                '        目录分类_IN 医保费用分类信息.目录分类%type,
                '        实际价格_IN 医保费用分类信息.目录分类%type,
                '        数量_IN     医保费用分类信息.数量%type,
                '        费用总额_IN 医保费用分类信息.费用总额%type,
                '        自付金额_IN 医保费用分类信息.自付金额%type,
                '        挂钩自付金额_IN 医保费用分类信息.挂钩自付金额%type,
                '        符合范围金额_IN 医保费用分类信息.符合范围金额%type,
                '        退单结算编号_IN 医保费用分类信息.退单结算编号%type
                        
                strSQL = "ZL_医保费用分类信息_INSERT("
                strSQL = strSQL & "" & lng病人ID & ","
                strSQL = strSQL & "" & lng主页ID & ","
                strSQL = strSQL & "'" & strArr(0) & "',"
                strSQL = strSQL & "" & Val(strArr(1)) & ","
                strSQL = strSQL & "'" & strArr(2) & "',"
                strSQL = strSQL & "'" & strArr(3) & "',"
                strSQL = strSQL & "'" & strArr(4) & "',"
                strSQL = strSQL & "" & Val(strArr(5)) & ","
                strSQL = strSQL & "'" & strArr(6) & "',"
                strSQL = strSQL & "" & Val(strArr(7)) & ","
                strSQL = strSQL & "" & Val(strArr(8)) & ","
                strSQL = strSQL & "" & Val(strArr(9)) & ","
                strSQL = strSQL & "" & Val(strArr(10)) & ","
                strSQL = strSQL & "" & Val(strArr(11)) & ","
                strSQL = strSQL & "" & Val(strArr(12)) & ","
                strSQL = strSQL & "" & Val(strArr(13)) & ")"
                
                
                '只有住院才有
                '20040720取消
                
                
                '       StrSQL = "ZL_医保明细费用_UPDATE("
                             
                '            '记帐流水号
                '            StrSQL = StrSQL & Val(strArr(1)) & ","
                '            StrSQL = StrSQL & "'" & strArr(0) & "',"
                '            StrSQL = StrSQL & "'" & strArr(2) & "',"
                '            StrSQL = StrSQL & "'" & strArr(3) & "',"
                '            StrSQL = StrSQL & "'" & strArr(4) & "',"
                '            StrSQL = StrSQL & "'" & strArr(5) & "',"
                '            StrSQL = StrSQL & "'" & strArr(6) & "',"
                '            StrSQL = StrSQL & "'" & strArr(7) & "',"
                '
                '            StrSQL = StrSQL & "" & Val(strArr(8)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(9)) & ","
                '
                '            '分段标准
                '            StrSQL = StrSQL & "" & Val(strArr(18)) & ","
                '
                '            StrSQL = StrSQL & "" & Val(strArr(19)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(20)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(21)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(22)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(23)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(24)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(25)) & ","
                '            StrSQL = StrSQL & "" & Val(strArr(26)) & ")"
                '修改费用分类信息
                gcnOracle_CQYB.Execute strSQL, , adCmdStoredProc
            End If
    Next
    objText.Close
    Save费用明细结算分割 = True
    Exit Function
errHand:
     DebugTool "明细分割保存(Save费用明细结算分割)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
   objText.Close
End Function

Private Function 病人费用结算(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:费用结算
    '--入参数:
    '--出参数:
    '--返  回:结算成功返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInFile1 As String          '本次结算待遇审批信息文件存放地址及文件名
    Dim strInFile2 As String          '本次结算明细信息存放地址及文件名
    Dim strInFile3 As String          '历次结算信息存放地址及文件名（可以为空）
    Dim strOutFile1 As String         '导出明细分割后文本文件存放地址及文件名，以制表符分割
    Dim strOutXMLFile1 As String      '导出明细分割后文本文件存放地址及文件名，以XML形式
    Dim stroutFile2 As String         '导出结算结果信息的文本文件存放地址及文件名，制表符分割
    Dim stroutXMLFile2 As String      '导出结算结果信息的文本文件存放地址及文件名，以XML形式
    Dim strErrMsg As String
    Dim lngAppCode As Long
    Dim gobjFile As New FileSystemObject
    
    Dim blnReturn As Boolean
    
    strInFile1 = gstrAppPath & "\待遇审批信息.txt"
    strInFile2 = gstrAppPath & "\结算明细信息.txt"
    strInFile3 = gstrAppPath & "\历次结算信息.txt"
    
    strOutFile1 = gstrAppPath & "\导出明细分割.txt"
    strOutXMLFile1 = gstrAppPath & "\导出明细分割XML.txt"
    
    stroutFile2 = gstrAppPath & "\导出结算结果信息.txt"
    stroutXMLFile2 = gstrAppPath & "\导出结算结果信息XML.txt"
    病人费用结算 = False
    
    DebugTool "病人费用结算进入"
    Err = 0
    On Error GoTo errHand:
    病人费用结算 = False
    If InitInfor_重庆渝北.模拟数据 Then
        Read模拟数据 费用结算, "", ""
    Else
        Debug.Print Time
        blnReturn = gobj费用结算.chargereckoning(strInFile1, strInFile2, strInFile3, g病人身份_重庆渝北.社保经办构构代码, g病人身份_重庆渝北.就诊编号, g病人身份_重庆渝北.结算编号, strOutFile1, strOutXMLFile1, stroutFile2, stroutXMLFile2, lngAppCode, strErrMsg)
        Debug.Print Time
        If blnReturn = False Then
            ShowMsgbox "错误号:" & lngAppCode & vbCrLf & "错误信息:" & strErrMsg
            GoTo DelFile:
            Exit Function
        End If
    End If
    '明细分解
    If Save费用明细结算分割(strOutFile1, lng病人ID, lng主页ID) = False Then
        GoTo DelFile:
        Exit Function
    End If
    
    '费用结果分解
    If 费用结果分解(stroutFile2, lng病人ID, lng主页ID) = False Then
        GoTo DelFile:
        Exit Function
    End If
    
    病人费用结算 = True
    GoTo DelFile:
    
    Exit Function
errHand:
    
    DebugTool "结算错误:" & Err.Number & "   信息:" & Err.Description
DelFile:
    '清除零时文件.
    Err = 0
    On Error Resume Next
    If gobjFile.FileExists(strOutFile1) = True Then
        gobjFile.DeleteFile strOutFile1, True
    End If
    If gobjFile.FileExists(stroutFile2) = True Then
        gobjFile.DeleteFile stroutFile2, True
    End If
    If gobjFile.FileExists(strOutXMLFile1) = True Then
        gobjFile.DeleteFile strOutXMLFile1, True
    End If
    If gobjFile.FileExists(stroutXMLFile2) = True Then
        gobjFile.DeleteFile stroutXMLFile2, True
    End If
End Function

Private Function Get结算编号() As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取结算编号
    '--入参数:
    '--出参数:
    '--返  回:新的结算号码
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String, strErrMsg As String, intAppCode As Integer
    Dim blnReturn As Boolean
    
    If Not g病人身份_重庆渝北.冲销 Then
        Get结算编号 = ""
        Err = 0
        On Error GoTo errHand:
        If InitInfor_重庆渝北.模拟数据 Then
            gstrSQL = "Select 医保病种目录_ID.nextval as 序号 from dual"
            OpenRecordset_ZLYB rsTemp, "获取结算编号"
            Get结算编号 = Nvl(rsTemp!序号)
            Exit Function
        End If
        Call intXML
         AppendXMLNode gobjXMLInPut.documentElement, "YAB003", g病人身份_重庆渝北.社保经办构构代码
        'SvrcID  string  2       远程数据服务标识，定值102, 标识大小写敏感，足两位长
        AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "102"
        
        Get结算编号 = ""
        StrInput = 取掉XML的前导标识(gobjXMLInPut.xml)
        
        Err = 0
        On Error GoTo errHand:
        blnReturn = gobjYingHaiDll.dll_main_in(StrInput, strOutput, intAppCode, strErrMsg)
        
        If blnReturn = False Then
          '发生错误,提示出信息
            ShowMsgbox strErrMsg
            Get结算编号 = ""
            Exit Function
        End If
        Get结算编号 = strOutput
        Exit Function
    End If
    
    gstrSQL = "Select 支付顺序号,备注 摘要 From 保险结算记录 where 记录ID= [1] And 险类=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取结算编号", g病人身份_重庆渝北.结帐ID, TYPE_重庆渝北)
    
    If rsTemp.EOF Then
        Get结算编号 = ""
    Else
        Get结算编号 = Nvl(rsTemp!支付顺序号)
        g病人身份_重庆渝北.就诊编号 = Substr(Nvl(rsTemp!摘要), 1, 20)
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function Get就诊编号_重庆渝北() As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取结算编号
    '--入参数:
    '--出参数:
    '--返  回:新的结算号码
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String
    Dim strOutput As String, strErrMsg As String, intAppCode As Integer
    Dim blnReturn As Boolean
    
    
    If InitInfor_重庆渝北.模拟数据 Then
        gstrSQL = "Select 医保病种目录_ID.nextval as 序号 from dual"
        OpenRecordset_ZLYB rsTemp, "获取结算编号"
        Get就诊编号_重庆渝北 = Nvl(rsTemp!序号)
        Exit Function
    End If
    
     Call intXML
     AppendXMLNode gobjXMLInPut.documentElement, "YAB003", g病人身份_重庆渝北.社保经办构构代码
    'SvrcID  string  2       远程数据服务标识，定值08, 标识大小写敏感，足两位长
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "101"
    'CtrInf  string  20      控制信息，预留, 标识大小写敏感

    Get就诊编号_重庆渝北 = ""
    StrInput = 取掉XML的前导标识(gobjXMLInPut.xml)
    
    Err = 0
    On Error GoTo errHand:
    blnReturn = gobjYingHaiDll.dll_main_in(StrInput, strOutput, intAppCode, strErrMsg)
    
    If blnReturn = False Then
      '发生错误,提示出信息
        ShowMsgbox strErrMsg
        Get就诊编号_重庆渝北 = ""
        Exit Function
    End If
    Get就诊编号_重庆渝北 = strOutput
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function 门诊虚拟结算_重庆渝北(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    
    '目前不支付门诊虚拟结算
    
    str结算方式 = ""
    门诊虚拟结算_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Get明细记录(ByRef lng结帐ID As Long, Optional strNO As String, Optional lng记录性质 As Long, Optional lng记录状态 As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取本次结帐的明细记录
    '--入参数:lng结帐ID-本次结帐的ID记录
    '         strno-本次处方的单据号,lng记录性质=记录性质,lng记录状态
    '--出参数:
    '--返  回:SQL语句
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strTable As String
    Dim rsTemp As New ADODB.Recordset
    Dim strFields As String
    
    strTable = IIf(g病人身份_重庆渝北.结算标志 = 0 Or g病人身份_重庆渝北.结算标志 = 2, "门诊费用记录", "住院费用记录")
    If lng结帐ID = 0 And g病人身份_重庆渝北.结算标志 <> 1 Then
            '--需确定审批标志
            strSQL = " " & _
                "  Select Rownum 标识号,A.ID,A.病人ID,a.主页id,A.收费细目id,收入项目id,A.NO,A.序号 ,A.记录性质,A.记录状态,A.登记时间 as 经办时间 ,c.名称 as 开单部门,a.开单人 as 开单医生,nvl(a.是否上传,0) 是否上传, " & _
                "      A.数次*A.付数 as 数量,A.计算单位,Round(A.实收金额/(A.数次*A.付数),4) as 实际价格,A.实收金额 as 实收金额, " & _
                "      A.收费类别,D.项目编码,D.项目名称,l.人员身份 as 经办机构代码,l.结算编号 结算号,l.支付类别," & _
                "      l.就诊编号 ,'' as 审批标志,L.险类,L.中心,L.卡号,L.医保号,L.人员身份,L.病种ID,L.就诊时间 ,J.名称 as 商品名" & _
                "  From (Select * From " & strTable & " Where 记录状态<>0 and nvl(实收金额,0)<>0 and NO='" & strNO & "' and 记录性质=" & lng记录性质 & " and 记录状态=" & lng记录状态 & " and  Nvl(附加标志,0)<>9 ) A,部门表 C," & _
                "       保险支付项目 D,保险帐户 L,收费细目 J " & _
                "  Where A.开单部门id=C.id(+) and  A.病人id=L.病人id  and a.收费细目id=J.id and L.险类=" & TYPE_重庆渝北 & "  And A.收费细目ID=D.收费细目ID And D.险类= " & TYPE_重庆渝北 & _
                "  Order by a.NO,A.记录性质,A.记录状态,a.序号"
    Else
        If g病人身份_重庆渝北.结算标志 = 1 And lng结帐ID = 0 Then
            '住院需传递所明细记录,需根据结算编号及就诊号来确定.
            strSQL = "Select Rownum 标识号, " & _
                     "          A.ID,A.病人ID,a.主页id,A.收费细目id ,收入项目id,A.NO,A.序号,A.记录性质,A.记录状态,A.登记时间 as 经办时间,c.名称 as 开单部门,a.开单人 as 开单医生, " & _
                     "          nvl(a.是否上传, 0) 是否上传,A.数次 * A.付数 as 数量,A.计算单位,Round(A.实收金额 / (A.数次 * A.付数), 4) as 实际价格, " & _
                     "          A.实收金额 as 实收金额,A.收费类别 ,b.记帐流水号,b.退单流水号,D.项目编码,D.项目名称, " & _
                     "          L.人员身份 as 经办机构代码,l.结算编号  as 结算号,L.支付类别 ,l.就诊编号,b.审批标志, " & _
                     "          L.险类,L.中心,L.卡号,L.医保号,L.人员身份,L.病种ID,L.就诊时间 ,J.名称 as 商品名" & _
                     "   From  " & strTable & " a , " & _
                     "          医保明细费用 b,部门表 C,保险支付项目 D,保险帐户 L,收费细目 J  " & _
                     "   Where a.记录状态<>0 and nvl(a.附加标志,0)<>9 and nvl(a.实收金额,0)<>0  and A.开单部门id = C.id(+) and a.id=b.费用id and b.就诊编号='" & g病人身份_重庆渝北.就诊编号 & "' " & _
                                IIf(g病人身份_重庆渝北.lng病人ID = 0, "", " And A.病人id =" & g病人身份_重庆渝北.lng病人ID) & _
                                IIf(g病人身份_重庆渝北.lng主页ID = 0, "", " And A.主页id =" & g病人身份_重庆渝北.lng主页ID) & _
                     "          and A.病人id = L.病人id and L.险类 = " & TYPE_重庆渝北 & "  And " & _
                     "          A.收费细目ID = D.收费细目ID and a.收费细目id=J.id And D.险类 =  " & TYPE_重庆渝北 & _
                     "    Order by a.NO,A.记录性质,A.记录状态,a.序号"
        Else
            '--需确定审批标志
            strSQL = " " & _
                "  Select Rownum 标识号,A.ID,A.病人ID,a.主页id,A.收费细目id,收入项目id,A.NO,A.序号 ,A.记录性质,A.记录状态,A.登记时间 as 经办时间 ,c.名称 as 开单部门,a.开单人 as 开单医生,nvl(a.是否上传,0) 是否上传, " & _
                "      A.数次*A.付数 as 数量,A.计算单位,Round(A.结帐金额/(A.数次*A.付数),4) as 实际价格,A.结帐金额 as 实收金额, " & _
                "      A.收费类别,D.项目编码,D.项目名称,人员身份 as 经办机构代码,'" & g病人身份_重庆渝北.结算编号 & "' as 结算号,L.支付类别," & _
                "      '" & g病人身份_重庆渝北.就诊编号 & "' as 就诊编号 ,'' as 审批标志,L.险类,L.中心,L.卡号,L.医保号,L.人员身份,L.病种ID,L.就诊时间 ,J.名称 as 商品名" & _
                "  From (Select * From " & strTable & " Where 记录状态<>0 and nvl(实收金额,0)<>0 and 结帐ID=" & lng结帐ID & " and  Nvl(附加标志,0)<>9 ) A,部门表 C," & _
                "       保险支付项目 D,保险帐户 L,收费细目 J " & _
                "  Where A.开单部门id=C.id(+) and  A.病人id=L.病人id and a.收费细目id=J.id and L.险类=" & TYPE_重庆渝北 & "  And A.收费细目ID=D.收费细目ID And D.险类= " & TYPE_重庆渝北 & _
                "   Order by a.NO,A.记录性质,A.记录状态,a.序号"
        End If
    End If

    Get明细记录 = strSQL
End Function

Private Function 处方明细上传(ByVal rs明细 As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:处方明细上传
    '--入参数:明细记录的字段:'ID,病人ID,收费细目ID,NO,记帐流水号,退单流水号,记录性质,记录状态,经办时间,开单部门,开单医生,,数量,计算单位,实际价格,实收金额,收费类别,项目编码,项目名称,经办机构代码,结算号,就诊编号,审批标志,险类,中心,卡号,医保号,人员身份,病种ID,就诊时间
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
   Dim rsTemp As New ADODB.Recordset
   Dim rs项目 As New ADODB.Recordset
   Dim strXMLText As String
   Dim blnTrue As Boolean
   Dim strOutput As String
   
    Err = 0
    On Error GoTo errHand:
    DebugTool "进入(" & "处方明细上传" & ")"
    If rs明细 Is Nothing Then Exit Function
    If rs明细.RecordCount = 0 Then Exit Function
    
     ''ID,病人ID,收费细目ID,NO,记帐流水号,退单流水号,记录性质,记录状态,经办时间,数量,计算单位,实际价格,实收金额,收费类别,项目编码,项目名称,经办机构代码,结算号,就诊编号,审批标志,险类,中心,卡号,医保号,人员身份,病种ID,就诊时间
    With rs明细
        .Filter = 0
        .Filter = "是否上传=0"
        If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
        blnTrue = True
        Do While Not .EOF
                 If Nvl(!项目编码) = "" Then
                     ShowMsgbox "存在未构对医保项目,请在保险项目中设置相应的对应关系!"
                     Exit Function
                 End If
                Call intXML
                Set rsTemp = Get医保明细费用(!ID)
                If g病人身份_重庆渝北.结算标志 = 2 And Nvl(!收入项目id, 0) = InitInfor_重庆渝北.收入项目id Then
                    Set rs项目 = Get保险项目(Nvl(rsTemp!项目编码))
                Else
                    Set rs项目 = Get保险项目(Nvl(!项目编码))
                    
                End If
                If rs项目.RecordCount = 0 Then Exit Function
                
                If g病人身份_重庆渝北.结算标志 = 3 Then
                    g病人身份_重庆渝北.支付类别 = Nvl(!支付类别)
                    g病人身份_重庆渝北.社保经办构构代码 = Nvl(!人员身份)
                End If
                'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
                AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(g病人身份_重庆渝北.社保经办构构代码, 1, 4)
                'SvrcID  string  2       远程数据服务标识，定值09, 标识大小写敏感，足两位长
                AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "09"
                'CtrInf  string  20      控制信息，预留, 标识大小写敏感
                AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
                'akc190  string  20      就诊编号
                AppendXMLNode gobjXMLInPut.documentElement, "akc190", Nvl(!就诊编号)
                'yka105  number  15  0   记账流水号
                AppendXMLNode gobjXMLInPut.documentElement, "yka105", Nvl(rsTemp!记帐流水号, 0)
                'yka002  string  20      医保项目编码
                AppendXMLNode gobjXMLInPut.documentElement, "yka002", Nvl(!项目编码)
                
                'yka103  string  20      结算编号
                AppendXMLNode gobjXMLInPut.documentElement, "yka103", Nvl(rsTemp!结算编号)
                'yka104  number  15      退单对应记账流水号
                AppendXMLNode gobjXMLInPut.documentElement, "yka104", Nvl(rsTemp!退单流水号)
                'aka130  string  6       支付类别，见代码表
                AppendXMLNode gobjXMLInPut.documentElement, "aka130", g病人身份_重庆渝北.支付类别
                'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
                AppendXMLNode gobjXMLInPut.documentElement, "akb020", InitInfor_重庆渝北.医院编码
                'ykb006  string  3       定点医疗机构分支机构编号
                AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"                '/***/需确定的问题
                'aac001  number  15  0   个人编号
                AppendXMLNode gobjXMLInPut.documentElement, "aac001", Nvl(!医保号)
                'akc226  number  14  4   数量
                AppendXMLNode gobjXMLInPut.documentElement, "akc226", Nvl(!数量, 0)
                'akc225  number  14  4   实际价格
                AppendXMLNode gobjXMLInPut.documentElement, "akc225", Nvl(!实际价格, 0)
                'yka055  number  14  2   医疗费总额
                AppendXMLNode gobjXMLInPut.documentElement, "yka055", Nvl(!实收金额, 0)
                'yka096  number  14  4   自付比例
                AppendXMLNode gobjXMLInPut.documentElement, "yka096", Nvl(rsTemp!自付比例, 0)
                'yka056  number  14  2   全自费金额
                AppendXMLNode gobjXMLInPut.documentElement, "yka056", Nvl(rsTemp!全自费金额, 0)
                'yka057  number  14  2   挂钩自费金额
                AppendXMLNode gobjXMLInPut.documentElement, "yka057", Nvl(rsTemp!挂钩自费金额, 0)
                'yka111  number  14  2   符合范围金额
                AppendXMLNode gobjXMLInPut.documentElement, "yka111", Nvl(rsTemp!符合范围金额, 0)
                'yka012  string  6       服务项目结算方式，见代码表
                AppendXMLNode gobjXMLInPut.documentElement, "yka012", "0"
                'yka098  string  50      开单科室名称
                AppendXMLNode gobjXMLInPut.documentElement, "yka098", Nvl(!开单部门)
                'yka099  string  20      开单医生
                AppendXMLNode gobjXMLInPut.documentElement, "yka099", Nvl(!开单医生)
                'yka101  string  50      受单科室名称
                AppendXMLNode gobjXMLInPut.documentElement, "yka101", Nvl(!开单部门)
                'yka102  string  20      受单医生
                AppendXMLNode gobjXMLInPut.documentElement, "yka102", Nvl(!开单医生)
                'aae036  date        秒  经办时间
                AppendXMLNode gobjXMLInPut.documentElement, "aae036", Format(!经办时间, "yyyy-mm-dd HH:MM:SS")
                'ykc166  date        秒  明细发生时间
                AppendXMLNode gobjXMLInPut.documentElement, "ykc166", Format(!经办时间, "yyyy-mm-dd HH:MM:SS")
                'yab003  string  4       社保经办机构代码
                AppendXMLNode gobjXMLInPut.documentElement, "yab003", g病人身份_重庆渝北.社保经办构构代码
                'yka231  string  20      商品名代码
                AppendXMLNode gobjXMLInPut.documentElement, "yka231", Nvl(rs项目!商品代码)
                'yka247  String  20      自费药品对应商品名代码
                AppendXMLNode gobjXMLInPut.documentElement, "yka247", IIf(rs项目!新增标志 = 1, Nvl(rs项目!标准编号), Nvl(rs项目!商品代码))
                'yka232  string  100     商品名
                AppendXMLNode gobjXMLInPut.documentElement, "yka232", Nvl(rs项目!商品名)
                'ykc130  string  6       购药类别
                AppendXMLNode gobjXMLInPut.documentElement, "ykc130", "0" '/**/需确定的问题
                'yka249  string  20      审批医生名称
                AppendXMLNode gobjXMLInPut.documentElement, "yka249", Nvl(rsTemp!审批人)
                'yka250  string  20      审批医生职称
                AppendXMLNode gobjXMLInPut.documentElement, "yka250", Nvl(rsTemp!审批人职称)
                'aae013  string  100     备注
                AppendXMLNode gobjXMLInPut.documentElement, "aae013", ""        '/**
                'gkc013  string  6       项目审批标志
                AppendXMLNode gobjXMLInPut.documentElement, "yka250", Nvl(rsTemp!审批标志, 0)
                'gkc014  string  50      剂型
                AppendXMLNode gobjXMLInPut.documentElement, "gkc014", Nvl(rs项目!剂型, 0)
                'yka272  String  6       目录分类
                AppendXMLNode gobjXMLInPut.documentElement, "yka272", Nvl(rs项目!目录分类, 0)
                
                strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
                
                WriteDebugInfor_重庆渝北 strXMLText
                
                '业务请求明细提交
                If 业务请求_重庆渝北(处方明细写入, strXMLText, strOutput) = False Then
                    blnTrue = False
                Else
                    '更新上传标志
                    '为病人费用记录打上标记，以便随时上传
                    'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                    gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                    zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
                End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    处方明细上传 = blnTrue
    Exit Function
errHand:
  DebugTool "处方明细上传出错(处方明细上传)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
End Function
Private Function Get发票号码(ByVal lng结帐ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 病人ID,实际票号 From 病人结帐记录 Where ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取发票号", lng结帐ID)
    If rsTemp.EOF Then
    Get发票号码 = ""
    Else
        Get发票号码 = Nvl(rsTemp!实际票号)
    End If

End Function


Public Function 门诊结算_重庆渝北(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, _
    Optional ByRef strAdvance As String = "") As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    
    Dim lng病人ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim str开始时间 As String
    Dim str结束时间 As String
    
    门诊结算_重庆渝北 = False
    
    WriteDebugDate_重庆渝北 "================================================================================================================================================================================================================================================"
    WriteDebugDate_重庆渝北 "===功    能:门诊结算"
    WriteDebugDate_重庆渝北 "===开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    WriteDebugDate_重庆渝北 "================================================================================================================================================================================================================================================"
    
    g病人身份_重庆渝北.结算标志 = 0
    g病人身份_重庆渝北.冲销 = False
    
    g病人身份_重庆渝北.结算编号 = Get结算编号
    g病人身份_重庆渝北.结帐ID = lng结帐ID
    g病人身份_重庆渝北.发票号 = Get发票号码(lng结帐ID)
    g病人身份_重庆渝北.虚拟结算 = False

  
    gstrSQL = "Select 病人id, 登记时间 From 门诊费用记录 where rownum<=1 and 结帐id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取登记时间"
    
    If g病人身份_重庆渝北.就诊时间 > Format(rsTemp!登记时间, "yyyy-MM-dd HH:mm:ss") Then
        g病人身份_重庆渝北.就诊时间 = Format(rsTemp!登记时间, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '保存当前状态的结算编号
    gstrSQL = "zl_保险帐户_更新信息(" & Nvl(rsTemp!病人ID, 0) & "," & TYPE_重庆渝北 & ",'结算编号','''" & g病人身份_重庆渝北.结算编号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算编号")
    
    gcnOracle_CQYB.BeginTrans
    
    
    If 病人结算(lng结帐ID) = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    
    gcnOracle_CQYB.CommitTrans
    #If gverControl < 2 Then
       strAdvance = ""
    #Else
        strAdvance = g病人身份_重庆渝北.结算信息
    #End If
    门诊结算_重庆渝北 = True
    Exit Function
errHand:
    gcnOracle_CQYB.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    DebugTool "门诊结算(门诊结算_重庆渝北)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
    Exit Function
End Function
Private Function Get冲销ID() As Long
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取当前冲销记录的ID值
    '--入参数:
    '--出参数:
    '--返  回:冲销ID
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    '取冲销记录的结帐ID
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", g病人身份_重庆渝北.结帐ID)
    If rsTemp.EOF Then
        Get冲销ID = 0
    Else
        Get冲销ID = Nvl(rsTemp!结帐ID, 0)
    End If

End Function

Public Function 门诊结算冲销_重庆渝北(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    

    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    
    Err = 0
    On Error GoTo errHand
    
    门诊结算冲销_重庆渝北 = False
    WriteDebugDate_重庆渝北 "================================================================================================================================================================================================================================================"
    WriteDebugDate_重庆渝北 "===功    能:门诊结算冲销`"
    WriteDebugDate_重庆渝北 "===开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    WriteDebugDate_重庆渝北 "================================================================================================================================================================================================================================================"
    
    '获取基本信息
    Call Get病人信息(lng病人ID)
    
    g病人身份_重庆渝北.结帐ID = lng结帐ID
    g病人身份_重庆渝北.结算标志 = 0
    g病人身份_重庆渝北.冲销 = False
    g病人身份_重庆渝北.结算编号 = Get结算编号
    g病人身份_重庆渝北.冲销 = True
    g病人身份_重庆渝北.冲销ID = Get冲销ID
    g病人身份_重庆渝北.虚拟结算 = False
    g病人身份_重庆渝北.lng病人ID = lng病人ID
    
    '保存当前状态的结算编号
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆渝北 & ",'结算编号','''" & lng病人ID & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算编号")
    
    gcnOracle_CQYB.BeginTrans
    门诊结算冲销_重庆渝北 = 病人结算冲销(lng结帐ID)
    If 门诊结算冲销_重庆渝北 = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    gcnOracle_CQYB.CommitTrans
    Exit Function
errHand:
    gcnOracle_CQYB.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Public Function 医保设置_重庆渝北() As Boolean
    医保设置_重庆渝北 = frmSet重庆渝北.参数设置
End Function

Public Function 入院登记_重庆渝北(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    
    Dim rsTemp As New ADODB.Recordset
    Dim blnYes As Boolean
    On Error GoTo errHand
    
    
    If 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "病人存在未结算费用,是否继续?", True, blnYes
        If blnYes = False Then
            Exit Function
        End If
    End If
    
        
    g病人身份_重庆渝北.冲销 = False
    g病人身份_重庆渝北.虚拟结算 = False
    g病人身份_重庆渝北.结算标志 = 2
    ''获取就诊号
    'g病人身份_重庆渝北.就诊编号 = Get就诊编号_重庆渝北
    g病人身份_重庆渝北.结算编号 = Get结算编号
    
    
    '更新就诊编号
'    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆渝北 & ",'就诊编号','''" & g病人身份_重庆渝北.就诊编号 & "''')"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊编号")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆渝北 & ",'结算编号','''" & g病人身份_重庆渝北.结算编号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算编号")
    
    '获取相关病人信息
    '(SELECT A.病人id,A.主页id,max(decode(A.序号,1,B.编码,'')) 手术编码,max(decode(A.序号,1,A.描述信息,'')) AS 手术名称,max(decode(A.序号,1,a.切口,'')) 切口,max(decode(A.序号,1,A.愈合,'')) 愈合    from 手术情况 A ,疾病编码目录 B     where A.手术ID=B.ID and A.序号=1 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & "    Group BY  A.病人id,A.主页id ) E,
    gstrSQL = "Select C.住院号,C.当前病区id,C.当前床号,c.就诊卡号 as 病历号,c.住院号,to_char(A.确诊日期,'yyyy-MM-dd hh24:mi:ss') as 确诊日期,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 入院经办时间," & _
        " to_char(A.登记时间,'yyyy-MM-dd') 入院日期  ,to_char(A.登记时间,'yyyy-MM-dd') 入院时间,D.入院诊断,D.入院诊断1,D.入院诊断2,D.入院诊断3,'' 手术编码,'' 手术名称,'' 切口,'' 愈合,'' 主治医师,'' 主任医师" & _
        " From 病案主页 A,部门表 B,病人信息 C, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,a.描述信息,'')) AS 入院诊断, max(DECODE(a.诊断次序,2,描述信息,'')) AS 入院诊断1,max(DECODE(a.诊断次序,3,a.描述信息,'')) AS 入院诊断2, max(DECODE(a.诊断次序,4,a.描述信息,'')) AS 入院诊断3 From 诊断情况 A  Where  a.诊断类型 =1 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by  病人id,主页id)   D" & _
        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
        "       and A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
        ""
        'and A.主页id=F.主页id(+) and a.病人id=F.病人id(+)
        '(SELECT 病人id,主页id,max(decode(信息名,'主治医师',信息值,'')) 主治医师,max(decode(信息名,'主任医师',信息值,'')) 主任医师 from 病案主页从表 where 主页id=" & lng主页ID & " and 病人id=" & lng病人ID & "    Group BY  病人id,主页id ) F
        'and A.主页id=E.主页id(+) and a.病人id=E.病人id(+)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院信息", lng病人ID, lng主页ID)
    
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    '需进行审批.
 
    If 资格审核待遇核定(lng病人ID, Format(rsTemp!入院时间, "yyyy-MM-dd HH:mm:ss"), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) = False Then
        Exit Function
    End If
    
    '第二步:写入资格审批待遇,并产生本产待遇文件
    If Save审批信息(lng病人ID, False) = False Then
        Exit Function
    End If
    
    Call intXML
    'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", g病人身份_重庆渝北.社保经办构构代码
    'SvrcID  string  2       远程数据服务标识，定值08, 标识大小写敏感，足两位长
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "08"
    'CtrInf  string  20      控制信息，预留, 标识大小写敏感
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'aac001  number  15  0   个人编号
    AppendXMLNode gobjXMLInPut.documentElement, "aac001", g病人身份_重庆渝北.个人编号
    'akc021  string  6       医疗人员类别
    AppendXMLNode gobjXMLInPut.documentElement, "akc021", g病人身份_重庆渝北.医疗人员类别
    'akc190  string  20      就诊编号
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g病人身份_重庆渝北.就诊编号
    'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", InitInfor_重庆渝北.医院编码
    'ykb006  string  3       定点医疗机构分支机构编号
    AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"
    'aka130  string  6       支付类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", g病人身份_重庆渝北.支付类别
    'akc192  date    日      入院日期
    AppendXMLNode gobjXMLInPut.documentElement, "akc192", Nvl(rsTemp!入院日期)
    
    'akc193  string  100     入院诊断
    AppendXMLNode gobjXMLInPut.documentElement, "akc193", g病人身份_重庆渝北.病情名称
    
    'ykc011  string  50      入院科室
    AppendXMLNode gobjXMLInPut.documentElement, "ykc011", Nvl(rsTemp!入院科室)
    'ykc013  string  20      入院经办人
    AppendXMLNode gobjXMLInPut.documentElement, "ykc013", Nvl(rsTemp!经办人)
    'ykc014  date        秒  入院经办时间
    AppendXMLNode gobjXMLInPut.documentElement, "ykc014", Nvl(rsTemp!入院经办时间)
    'akc195  string  6       出院原因，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "akc195", ""
    'akc194  date    日      出院日期
    AppendXMLNode gobjXMLInPut.documentElement, "akc194", ""
    'akc196  string  100     出院诊断
    AppendXMLNode gobjXMLInPut.documentElement, "akc196", ""
    'ykc015  string  50      出院科室
    AppendXMLNode gobjXMLInPut.documentElement, "ykc015", ""
    'ykc016  string  20      出院经办人
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", ""
    'ykc017  date        秒  出院经办时间
    AppendXMLNode gobjXMLInPut.documentElement, "ykc017", ""
    'ykc023  string  6       住院状态
    '0-在院,1-出院 2-转院
    AppendXMLNode gobjXMLInPut.documentElement, "ykc023", "0"
    'ykc009  string  20      病历号
    AppendXMLNode gobjXMLInPut.documentElement, "ykc009", Nvl(rsTemp!病历号)
    'ykc010  string  20      住院号
    AppendXMLNode gobjXMLInPut.documentElement, "ykc010", Nvl(rsTemp!住院号)
    'ykc149  string  100     入院诊断1(小苏要求,对应:病种编码)
    AppendXMLNode gobjXMLInPut.documentElement, "ykc149", g病人身份_重庆渝北.病情编码
    'ykc150  string  100     入院诊断2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc150", g病人身份_重庆渝北.病情名称1
    'ykc151  string  100     入院诊断3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc151", g病人身份_重庆渝北.病情名称2
    'ykc012  string  12      入院床位
    AppendXMLNode gobjXMLInPut.documentElement, "ykc012", Nvl(rsTemp!当前床号)
    'ykc152  string  100     出院诊断1
    AppendXMLNode gobjXMLInPut.documentElement, "ykc152", ""
    'ykc153  string  100     出院诊断2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc153", ""
    'ykc154  string  100     出院诊断3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc154", ""
    'ykc016  string  12      出院床位
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", ""
    'ykc155  string  20      手术编码
    AppendXMLNode gobjXMLInPut.documentElement, "ykc155", Nvl(rsTemp!手术编码)
    
    'ykc156  string  100     手术名称
    AppendXMLNode gobjXMLInPut.documentElement, "ykc156", Nvl(rsTemp!手术名称)
    'ykc157  date        秒  确诊时间
    AppendXMLNode gobjXMLInPut.documentElement, "ykc157", Nvl(rsTemp!确诊日期)
    'ykc158  string  4       手术切口分类
    AppendXMLNode gobjXMLInPut.documentElement, "ykc158", Nvl(rsTemp!切口)
    'ykc159  string  4       手术切口愈合级别
    AppendXMLNode gobjXMLInPut.documentElement, "ykc159", Nvl(rsTemp!愈合)
    'ykc160  string  20      住院医师姓名
    AppendXMLNode gobjXMLInPut.documentElement, "ykc160", Nvl(rsTemp!住院医师)
    'ykc161  string  20      主治医师姓名
    AppendXMLNode gobjXMLInPut.documentElement, "ykc161", Nvl(rsTemp!主治医师)
    'ykc162  string  20      主任医师姓名
    AppendXMLNode gobjXMLInPut.documentElement, "ykc162", Nvl(rsTemp!主任医师)
    'aae013  string  100     备注
    AppendXMLNode gobjXMLInPut.documentElement, "aae013", ""
    
    Dim strXMLText As String
    Dim strOutput As String
    strXMLText = gobjXMLInPut.xml
    strXMLText = 取掉XML的前导标识(strXMLText)
        
        
        
    If 业务请求_重庆渝北(就诊信息写入, strXMLText, strOutput, "") = False Then
        入院登记_重庆渝北 = False
        Exit Function
    End If
    '保存病情
    gcnOracle_CQYB.BeginTrans
    Call Save病情信息(lng病人ID, lng主页ID, 1)
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆渝北 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    gcnOracle_CQYB.CommitTrans
    
    入院登记_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_重庆渝北(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    ShowMsgbox "本医保不支持入院撤销!"
    Exit Function
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重庆渝北 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    入院登记撤销_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_重庆渝北(lng病人ID As Long, lng主页ID As Long) As Boolean
    
    Dim str出院原因 As String
    Dim intMouse  As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim rs病情 As New ADODB.Recordset
    Dim strTemp As String
    On Error GoTo errHand
    
    '1       出院原因 治愈
    '2       出院原因 好转
    '3       出院原因 未愈
    '4       出院原因 死亡
    '5       出院原因 转院
    '9       出院原因 其它
   '获取基本信息
   
    Call Get病人信息(lng病人ID)

    intMouse = Screen.MousePointer
    Screen.MousePointer = 1

    If frm补录病情_重庆渝北.ShowSelect(TYPE_重庆渝北, lng病人ID, lng主页ID) = False Then
        Screen.MousePointer = intMouse
        Exit Function
    End If
    Screen.MousePointer = intMouse

    
    gstrSQL = "Select 性质,序号,病情ID,病情编码,病情 from 病人诊断情况_91 where  病人id=" & lng病人ID & IIf(lng主页ID = 0, " and 主页id is null ", " and 主页id=" & lng主页ID) & " and 性质 IN (1,2)"
    Call OpenRecordset_OtherBase(rs病情, "获取诊断情况", gstrSQL, gcnOracle_CQYB)
    

    '获取原就诊流水号
    
    gstrSQL = "Select 就诊编号,结算编号 From 保险帐户 Where 险类=[1] And 病人ID=" & lng病人ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取就诊编号和结算编号", TYPE_重庆渝北)
    
    g病人身份_重庆渝北.就诊编号 = Nvl(rsTemp!就诊编号)
    g病人身份_重庆渝北.结算编号 = Nvl(rsTemp!结算编号)
    
    '获取相关病人信息
    gstrSQL = Get入出SQL(lng病人ID, lng主页ID)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院信息")
    
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    rs病情.Filter = "性质=1 and 序号=1"
    If rs病情.EOF Then
        ShowMsgbox "无入院病种,不能继续!"
        Exit Function
    End If


'    If 资格审核待遇核定(lng病人ID, Format(rsTemp!入院时间, "yyyy-MM-dd HH:mm:ss"), Format(rsTemp!出院时间, "yyyy-MM-dd HH:mm:ss")) = False Then
'        Exit Function
'    End If
'
'    '第二步:写入资格审批待遇,并产生本产待遇文件
'    If Save审批信息(lng病人ID, False) = False Then
'        Exit Function
'    End If

    Call intXML
    'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", g病人身份_重庆渝北.社保经办构构代码
    'SvrcID  string  2       远程数据服务标识，定值08, 标识大小写敏感，足两位长
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "08"
    'CtrInf  string  20      控制信息，预留, 标识大小写敏感
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'aac001  number  15  0   个人编号
    AppendXMLNode gobjXMLInPut.documentElement, "aac001", g病人身份_重庆渝北.个人编号
    'akc021  string  6       医疗人员类别
    AppendXMLNode gobjXMLInPut.documentElement, "akc021", g病人身份_重庆渝北.医疗人员类别
    'akc190  string  20      就诊编号
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g病人身份_重庆渝北.就诊编号
    'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", InitInfor_重庆渝北.医院编码
    'ykb006  string  3       定点医疗机构分支机构编号
    AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"
    'aka130  string  6       支付类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", g病人身份_重庆渝北.支付类别
    'akc192  date    日      入院日期
    AppendXMLNode gobjXMLInPut.documentElement, "akc192", Nvl(rsTemp!入院日期)
    
    'akc193  string  100     入院诊断
    AppendXMLNode gobjXMLInPut.documentElement, "akc193", Nvl(rs病情!病情)
    'ykc011  string  50      入院科室
    AppendXMLNode gobjXMLInPut.documentElement, "ykc011", Nvl(rsTemp!入院科室)
    'ykc013  string  20      入院经办人
    AppendXMLNode gobjXMLInPut.documentElement, "ykc013", Nvl(rsTemp!经办人)
    'ykc014  date        秒  入院经办时间
    AppendXMLNode gobjXMLInPut.documentElement, "ykc014", Nvl(rsTemp!入院经办时间)
    'akc195  string  6       出院原因，见代码表
    str出院原因 = IIf(IsNull(rsTemp!出院方式), "", rsTemp!出院方式)
    '1       出院原因 治愈
    '2       出院原因 好转
    '3       出院原因 未愈
    '4       出院原因 死亡
    '5       出院原因 转院
    '9       出院原因 其它
    '治愈、好转、未愈、死亡、其他
      Select Case str出院原因
      Case "治愈"
          str出院原因 = 1
      Case "好转"
          str出院原因 = 2
      Case "未愈"
          str出院原因 = 3
      Case "死亡"
          str出院原因 = 4
      Case "转院"
          str出院原因 = 5
      Case Else
          str出院原因 = 9
      End Select
      
    AppendXMLNode gobjXMLInPut.documentElement, "akc195", str出院原因
    'akc194  date    日      出院日期
    AppendXMLNode gobjXMLInPut.documentElement, "akc194", Nvl(rsTemp!出院日期)
    'akc196  string  100     出院诊断
    AppendXMLNode gobjXMLInPut.documentElement, "akc196", Nvl(rsTemp!出院诊断)
    'ykc015  string  50      出院科室
    AppendXMLNode gobjXMLInPut.documentElement, "ykc015", Nvl(rsTemp!出院科室)
    'ykc016  string  20      出院经办人
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", Nvl(rsTemp!操作员)
    'ykc017  date        秒  出院经办时间
    AppendXMLNode gobjXMLInPut.documentElement, "ykc017", Format(rsTemp!终止时间, "yyyy-MM-dd HH:mm:ss")
    'ykc023  string  6       住院状态
    '0-在院,1-出院 2-转院
    AppendXMLNode gobjXMLInPut.documentElement, "ykc023", IIf(str出院原因 = "5", "2", "1")
    'ykc009  string  20      病历号
    AppendXMLNode gobjXMLInPut.documentElement, "ykc009", Nvl(rsTemp!病历号)
    'ykc010  string  20      住院号
    AppendXMLNode gobjXMLInPut.documentElement, "ykc010", Nvl(rsTemp!住院号)
    
    'ykc149  string  100     入院诊断1(要求传ICD-10编码)
    AppendXMLNode gobjXMLInPut.documentElement, "ykc149", Nvl(rs病情!病情编码)
    
    rs病情.Filter = "性质=1 and 序号=2"
    If rs病情.EOF Then
        strTemp = ""
    Else
        strTemp = Nvl(rs病情!病情)
    End If
    'ykc150  string  100     入院诊断2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc150", strTemp
    rs病情.Filter = "性质=1 and 序号=3"
    If rs病情.EOF Then
        strTemp = ""
    Else
        strTemp = Nvl(rs病情!病情)
    End If
    
    'ykc151  string  100     入院诊断3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc151", strTemp
    'ykc012  string  12      入院床位
    AppendXMLNode gobjXMLInPut.documentElement, "ykc012", Nvl(rsTemp!入院病床)
    rs病情.Filter = "性质=2 and 序号=1"
    If rs病情.EOF Then
        strTemp = ""
    Else
        strTemp = Nvl(rs病情!病情)
    End If
    'ykc152  string  100     出院诊断1
    AppendXMLNode gobjXMLInPut.documentElement, "ykc152", strTemp
    rs病情.Filter = "性质=2 and 序号=2"
    If rs病情.EOF Then
        strTemp = ""
    Else
        strTemp = Nvl(rs病情!病情)
    End If
    
    'ykc153  string  100     出院诊断2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc153", strTemp
    
    rs病情.Filter = "性质=2 and 序号=3"
    If rs病情.EOF Then
        strTemp = ""
    Else
        strTemp = Nvl(rs病情!病情)
    End If
    
    'ykc154  string  100     出院诊断3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc154", strTemp
    'ykc016  string  12      出院床位
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", Nvl(rsTemp!出院病床)
    'ykc155  string  20      手术编码
    AppendXMLNode gobjXMLInPut.documentElement, "ykc155", Nvl(rsTemp!手术编码)
    
    'ykc156  string  100     手术名称
    AppendXMLNode gobjXMLInPut.documentElement, "ykc156", Nvl(rsTemp!手术名称)
    'ykc157  date        秒  确诊时间
    AppendXMLNode gobjXMLInPut.documentElement, "ykc157", IIf(Nvl(rsTemp!确诊日期) = "", Format(rsTemp!入院时间, "yyyy-MM-dd HH:mm:ss"), Nvl(rsTemp!确诊日期))
    'ykc158  string  4       手术切口分类
    AppendXMLNode gobjXMLInPut.documentElement, "ykc158", Nvl(rsTemp!切口)
    'ykc159  string  4       手术切口愈合级别
    AppendXMLNode gobjXMLInPut.documentElement, "ykc159", Nvl(rsTemp!愈合)
    'ykc160  string  20      住院医师姓名
    AppendXMLNode gobjXMLInPut.documentElement, "ykc160", Nvl(rsTemp!住院医师)
    'ykc161  string  20      主治医师姓名
    AppendXMLNode gobjXMLInPut.documentElement, "ykc161", Nvl(rsTemp!主治医师)
    'ykc162  string  20      主任医师姓名
    AppendXMLNode gobjXMLInPut.documentElement, "ykc162", Nvl(rsTemp!主任医师)
    'aae013  string  100     备注
    AppendXMLNode gobjXMLInPut.documentElement, "aae013", ""
    
    Dim strXMLText As String
    Dim strOutput As String
    strXMLText = gobjXMLInPut.xml
    strXMLText = 取掉XML的前导标识(strXMLText)
    
    If 业务请求_重庆渝北(就诊信息写入, strXMLText, strOutput, "") = False Then
        出院登记_重庆渝北 = False
          Exit Function
    End If
     
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        '如果不存在未结费用,需删除审批信息
        '需审批记录作废
        Call 审批记录作废_重庆渝北
    End If
     
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重庆渝北 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    
    出院登记_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Get入出SQL(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    Dim strSQL As String
    
    strSQL = "Select C.住院号,C.当前病区id,A.入院病床 ,c.就诊卡号 as 病历号,c.住院号,to_char(A.确诊日期,'yyyy-MM-dd hh24:mi:ss') as 确诊日期,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 入院经办时间," & _
        " to_char(A.入院日期,'yyyy-MM-dd') 入院日期,to_char(A.入院日期,'yyyy-MM-dd') 入院时间,J.终止时间,J.操作员,D.入院诊断,D.入院诊断1,D.入院诊断2,D.入院诊断3,A.出院方式,to_Char(a.出院日期,'yyyy-MM-DD') as 出院日期,a.出院日期 as 出院时间,a.出院病床,H.名称 as 出院科室,'' 手术编码,'' 手术名称,'' 切口,'' 愈合,'' 主治医师,'' 主任医师,G.出院诊断,G.出院诊断1,g.出院诊断2,g.出院诊断3" & _
        " From 病案主页 A,部门表 B,病人信息 C,部门表 H, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,a.描述信息,'')) AS 入院诊断, max(DECODE(a.诊断次序,2,a.描述信息,'')) AS 入院诊断1,max(DECODE(a.诊断次序,3,a.描述信息,'')) AS 入院诊断2, max(DECODE(a.诊断次序,4,a.描述信息,'')) AS 入院诊断3 From 诊断情况 A  Where   a.诊断类型 =1  and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by 病人id,主页id)   D," & _
        "       (Select 病人id,主页id,Max(终止时间) as 终止时间,max(操作员姓名) 操作员 From 病人变动记录 where  终止原因=1 and 病人id=" & lng病人ID & " and 主页id=" & lng主页ID & " Group by 病人id,主页id) J," & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,a.描述信息,'')) AS 出院诊断, max(DECODE(a.诊断次序,2,a.描述信息,'')) AS 出院诊断1,max(DECODE(a.诊断次序,3,a.描述信息,'')) AS 出院诊断2, max(DECODE(a.诊断次序,4,a.描述信息,'')) AS 出院诊断3 From 诊断情况 A   Where  a.诊断类型 = 3 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by 病人id,主页id)   G" & _
        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
        "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID and A.出院科室ID=H.id(+) " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
        "       and A.主页id=J.主页id(+) and a.病人id=J.病人id(+)" & _
        "       and A.主页id=G.主页id(+) and a.病人id=G.病人id(+) " & _
        ""
    Get入出SQL = strSQL
End Function

Public Function 出院登记撤销_重庆渝北(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rs病情 As New ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo errHand
     
     '获取相关病人信息
      '获取基本信息
    
    '由于在结帐后要冲销IC卡,所以不能对已经结完了的病人进行取消记帐
    
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "不能对不存在未结费用的病人进行撤销出院,需重新办理入院."
        Exit Function
    End If
    
    Call Get病人信息(lng病人ID)
 
    gstrSQL = "Select 性质,序号,病情ID,病情编码,病情 from 病人诊断情况_91 where   病人id=" & lng病人ID & IIf(lng主页ID = 0, " and 主页id is null ", " and 主页id=" & lng主页ID) & " and 性质 IN (1,2)"
    Call OpenRecordset_OtherBase(rs病情, "获取诊断情况", gstrSQL, gcnOracle_CQYB)
    
    rs病情.Filter = "性质=1 and 序号=1"
    If rs病情.EOF Then
        ShowMsgbox "无入院病种,不能继续!"
        Exit Function
    End If

    
    gstrSQL = Get入出SQL(lng病人ID, lng主页ID)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取入院信息")
    
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    Call intXML
    'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", g病人身份_重庆渝北.社保经办构构代码
    'SvrcID  string  2       远程数据服务标识，定值08, 标识大小写敏感，足两位长
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "08"
    'CtrInf  string  20      控制信息，预留, 标识大小写敏感
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'aac001  number  15  0   个人编号
    AppendXMLNode gobjXMLInPut.documentElement, "aac001", g病人身份_重庆渝北.个人编号
    'akc021  string  6       医疗人员类别
    AppendXMLNode gobjXMLInPut.documentElement, "akc021", g病人身份_重庆渝北.医疗人员类别
    'akc190  string  20      就诊编号
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g病人身份_重庆渝北.就诊编号
    'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", InitInfor_重庆渝北.医院编码
    'ykb006  string  3       定点医疗机构分支机构编号
    AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"
    'aka130  string  6       支付类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", g病人身份_重庆渝北.支付类别
    'akc192  date    日      入院日期
    AppendXMLNode gobjXMLInPut.documentElement, "akc192", Nvl(rsTemp!入院日期)
    
    'akc193  string  100     入院诊断
    AppendXMLNode gobjXMLInPut.documentElement, "akc193", Nvl(rs病情!病情)
    'ykc011  string  50      入院科室
    AppendXMLNode gobjXMLInPut.documentElement, "ykc011", Nvl(rsTemp!入院科室)
    'ykc013  string  20      入院经办人
    AppendXMLNode gobjXMLInPut.documentElement, "ykc013", Nvl(rsTemp!经办人)
    'ykc014  date        秒  入院经办时间
    AppendXMLNode gobjXMLInPut.documentElement, "ykc014", Nvl(rsTemp!入院经办时间)
    'akc195  string  6       出院原因，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "akc195", ""
    'akc194  date    日      出院日期
    AppendXMLNode gobjXMLInPut.documentElement, "akc194", ""
    'akc196  string  100     出院诊断
    AppendXMLNode gobjXMLInPut.documentElement, "akc196", ""
    'ykc015  string  50      出院科室
    AppendXMLNode gobjXMLInPut.documentElement, "ykc015", ""
    'ykc016  string  20      出院经办人
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", ""
    'ykc017  date        秒  出院经办时间
    AppendXMLNode gobjXMLInPut.documentElement, "ykc017", ""
    'ykc023  string  6       住院状态
    
    '0-在院,1-出院 2-转院
    AppendXMLNode gobjXMLInPut.documentElement, "ykc023", "0"
    'ykc009  string  20      病历号
    AppendXMLNode gobjXMLInPut.documentElement, "ykc009", Nvl(rsTemp!病历号)
    'ykc010  string  20      住院号
    AppendXMLNode gobjXMLInPut.documentElement, "ykc010", Nvl(rsTemp!住院号)
    
    'ykc149  string  100     入院诊断1
    AppendXMLNode gobjXMLInPut.documentElement, "ykc149", Nvl(rs病情!病情编码)
    rs病情.Filter = "性质=1 and 序号=2"
    If rs病情.EOF Then
        strTemp = ""
    Else
        strTemp = Nvl(rs病情!病情)
    End If
    
    'ykc150  string  100     入院诊断2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc150", strTemp
    rs病情.Filter = "性质=1 and 序号=3"
    If rs病情.EOF Then
        strTemp = ""
    Else
        strTemp = Nvl(rs病情!病情)
    End If
    
    'ykc151  string  100     入院诊断3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc151", strTemp
    'ykc012  string  12      入院床位
    AppendXMLNode gobjXMLInPut.documentElement, "ykc012", Nvl(rsTemp!入院病床)
    
    rs病情.Filter = "性质=2 and 序号=1"
    If rs病情.EOF Then
        strTemp = ""
    Else
        strTemp = Nvl(rs病情!病情)
    End If
    
    'ykc152  string  100     出院诊断1
    AppendXMLNode gobjXMLInPut.documentElement, "ykc152", strTemp
    rs病情.Filter = "性质=2 and 序号=2"
    If rs病情.EOF Then
        strTemp = ""
    Else
        strTemp = Nvl(rs病情!病情)
    End If
    
    'ykc153  string  100     出院诊断2
    AppendXMLNode gobjXMLInPut.documentElement, "ykc153", strTemp
    rs病情.Filter = "性质=2 and 序号=2"
    If rs病情.EOF Then
        strTemp = ""
    Else
        strTemp = Nvl(rs病情!病情)
    End If
    'ykc154  string  100     出院诊断3
    AppendXMLNode gobjXMLInPut.documentElement, "ykc154", strTemp
    'ykc016  string  12      出院床位
    AppendXMLNode gobjXMLInPut.documentElement, "ykc016", ""
    'ykc155  string  20      手术编码
    AppendXMLNode gobjXMLInPut.documentElement, "ykc155", Nvl(rsTemp!手术编码)
    
    'ykc156  string  100     手术名称
    AppendXMLNode gobjXMLInPut.documentElement, "ykc156", Nvl(rsTemp!手术名称)
    'ykc157  date        秒  确诊时间
    AppendXMLNode gobjXMLInPut.documentElement, "ykc157", Nvl(rsTemp!确诊日期)
    'ykc158  string  4       手术切口分类
    AppendXMLNode gobjXMLInPut.documentElement, "ykc158", Nvl(rsTemp!切口)
    'ykc159  string  4       手术切口愈合级别
    AppendXMLNode gobjXMLInPut.documentElement, "ykc159", Nvl(rsTemp!愈合)
    'ykc160  string  20      住院医师姓名
    AppendXMLNode gobjXMLInPut.documentElement, "ykc160", Nvl(rsTemp!住院医师)
    'ykc161  string  20      主治医师姓名
    AppendXMLNode gobjXMLInPut.documentElement, "ykc161", Nvl(rsTemp!主治医师)
    'ykc162  string  20      主任医师姓名
    AppendXMLNode gobjXMLInPut.documentElement, "ykc162", Nvl(rsTemp!主任医师)
    'aae013  string  100     备注
    AppendXMLNode gobjXMLInPut.documentElement, "aae013", ""
    
    Dim strXMLText As String
    Dim strOutput As String
    strXMLText = gobjXMLInPut.xml
    strXMLText = 取掉XML的前导标识(strXMLText)
    
    If 业务请求_重庆渝北(就诊信息写入, strXMLText, strOutput, "") = False Then
        出院登记撤销_重庆渝北 = False
        Exit Function
    End If
    
'    If Not 存在未结费用(lng病人ID, lng主页ID) Then
'
'        If 资格审核待遇核定(lng病人ID, Format(rsTemp!入院时间, "yyyy-MM-dd HH:mm:ss"), Format(zlDataBase.Currentdate, "yyyy-MM-dd HH:mm:ss")) = False Then
'            Exit Function
'        End If
'        '第二步:写入资格审批待遇,并产生本产待遇文件
'        If Save审批信息(lng病人ID, False) = False Then
'            Exit Function
'        End If
'    End If
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆渝北 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销出院登记")
    
    出院登记撤销_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_重庆渝北(ByVal lng病人ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    
    '读卡失败则退出
    gstrSQL = "Select Nvl(帐户余额,0) 帐户余额,退休证号 From 保险帐户 Where 险类=[1] And 病人id=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取帐户余额", TYPE_重庆渝北, lng病人ID)
    
    With g病人身份_大连
        个人余额_重庆渝北 = Nvl(rsTemp!帐户余额, 0)
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function 住院虚拟结算_重庆渝北(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String

    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim str开始时间 As String
    Dim str结束时间 As String
    Dim intMouse As Integer
    Dim lng主页ID As Long
    
    住院虚拟结算_重庆渝北 = ""
    
    DebugTool "进入住院虚拟结算"
    
    Call Get病人信息(lng病人ID)
    DebugTool "已经获取病人信息,并进入身份验证."
    WriteDebugDate_重庆渝北 "================================================================================================================================================================================================================================================"
    WriteDebugDate_重庆渝北 "===功    能:住院虚拟结算"
    WriteDebugDate_重庆渝北 "===开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    WriteDebugDate_重庆渝北 "================================================================================================================================================================================================================================================"
    
    If bln结帐处 Then
        '需重新进行验卡
        intMouse = Screen.MousePointer
        Screen.MousePointer = 1
        If Trim(frmIdentify重庆渝北.GetPatient(4, 0)) = "" Then
            Exit Function
        End If
        Screen.MousePointer = intMouse
    Else
        '其他地方
    End If
    
    
    g病人身份_重庆渝北.结算标志 = 1
    g病人身份_重庆渝北.冲销 = False
    
    g病人身份_重庆渝北.费用总额 = 0
    
    '求本次总额
    DebugTool "取本次费用总额"
    With rsExse
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            g病人身份_重庆渝北.费用总额 = g病人身份_重庆渝北.费用总额 + Nvl(rsExse!金额, 0)
            If lng主页ID < Nvl(rsExse!主页ID, 0) Then
                lng主页ID = Nvl(rsExse!主页ID, 0)
            End If
            
            .MoveNext
        Loop
    End With
    
    '取就诊时间
    DebugTool "取就诊信息!"
    gstrSQL = "Select 就诊时间,就诊编号,结算编号 From 保险帐户 where 险类=" & TYPE_重庆渝北 & " and 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取就诊时间"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "在保险帐户中不存在此医保病人!"
        Exit Function
    End If
    
    g病人身份_重庆渝北.就诊时间 = Format(rsTemp!就诊时间, "yyyy-MM-dd HH:mm:ss")
    g病人身份_重庆渝北.就诊编号 = Nvl(rsTemp!就诊编号)
    g病人身份_重庆渝北.结算编号 = Nvl(rsTemp!结算编号)
    g病人身份_重庆渝北.lng病人ID = lng病人ID
    
    '需读出在费用记录中的自动记帐部份,更新在表中
    gstrSQL = "" & _
        "   Select ID,NO,记录状态,记录性质,序号 From 住院费用记录 " & _
        "   Where 病人id=" & lng病人ID & " and 主页id=" & lng主页ID & " and id not in(Select a.id From 住院费用记录 a,医保明细费用 b Where a.id=b.费用id And a.病人id=" & lng病人ID & " and a.主页id=" & lng主页ID & ") " & _
        "   Order by 记录性质,NO,记录状态,序号"
    
    DebugTool "插入自动记帐部分的明细记录(向中间库插入)!"
    
    zlDatabase.OpenRecordset rs明细, gstrSQL, "获取自动记帐明细记录"
    With rs明细
        .Filter = "记录状态<>2"
        g病人身份_重庆渝北.冲销 = False
        Do While Not .EOF
             IsertInto医保明细 !ID, Nvl(!NO), Nvl(!序号, 0), Nvl(!记录性质, 0), ""
            .MoveNext
        Loop
        .Filter = "记录状态=2"
        If .RecordCount <> 0 Then .MoveFirst
        g病人身份_重庆渝北.冲销 = True
        
        '需插入退单的流水号
        Do While Not .EOF
            IsertInto医保明细 !ID, Nvl(!NO), Nvl(!序号, 0), Nvl(!记录性质, 0), ""
            .MoveNext
        Loop
    End With
        
    '更新所有医保明细费用记录的就诊编号及结算编号
    Err = 0
    On Error Resume Next
    
    DebugTool "更新本次结算的结算编号(向中间库医保明细更新)!"
    gcnOracle_CQYB.Execute "UPdate 医保明细费用 set 结算编号='" & g病人身份_重庆渝北.结算编号 & "' where 结算编号 is null and 就诊编号='" & g病人身份_重庆渝北.就诊编号 & "'"
    
    If Err <> 0 Then
        ShowMsgbox "在更新医保费用时出错!"
        Exit Function
    End If
    
    g病人身份_重庆渝北.冲销 = False
    g病人身份_重庆渝北.虚拟结算 = True
    g病人身份_重庆渝北.lng主页ID = lng主页ID
    Call 补传住院明细记录(lng病人ID, lng主页ID)
    
    gcnOracle_CQYB.BeginTrans
    
    DebugTool "进入病人结算!"
    If 病人结算(0) = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    DebugTool "结算完成!"
    住院虚拟结算_重庆渝北 = g病人身份_重庆渝北.结算信息
    gcnOracle_CQYB.CommitTrans
    Exit Function
errHand:
    DebugTool "虚拟结算时发生错误" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & "错误信息:" & Err.Description
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_重庆渝北(lng结帐ID As Long, ByVal lng病人ID As Long, Optional ByRef strAdvance As String = "") As Boolean

    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim str病种 As String
    Dim str单位编码 As String
    WriteDebugDate_重庆渝北 "================================================================================================================================================================================================================================================"
    WriteDebugDate_重庆渝北 "===功    能:住院结算"
    WriteDebugDate_重庆渝北 "===开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    WriteDebugDate_重庆渝北 "================================================================================================================================================================================================================================================"
    
    住院结算_重庆渝北 = False
    
    str病种 = g病人身份_重庆渝北.病种编码
    str单位编码 = g病人身份_重庆渝北.单位编码
    Call Get病人信息(lng病人ID)
    g病人身份_重庆渝北.病种编码 = str病种
     
    g病人身份_重庆渝北.单位编码 = str单位编码
    
    g病人身份_重庆渝北.结帐ID = lng结帐ID
    g病人身份_重庆渝北.结算标志 = 1
    g病人身份_重庆渝北.冲销 = False
    g病人身份_重庆渝北.发票号 = Get发票号码(lng结帐ID)
    g病人身份_重庆渝北.lng病人ID = lng病人ID
    
    '求本次结算的费用总额
    gstrSQL = "Select Sum(nvl(结帐金额,0)) as 总费用, max(主页id) as 主页id From 住院费用记录 where 结帐id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取总费用"
    g病人身份_重庆渝北.费用总额 = Nvl(rsTemp!总费用, 0)
    g病人身份_重庆渝北.lng主页ID = Nvl(rsTemp!主页ID, 0)
    
    '取就诊时间
    gstrSQL = "Select 就诊时间,就诊编号,结算编号 From 保险帐户 where 险类=" & TYPE_重庆渝北 & " and 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取就诊时间"
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "在保险帐户中不存在此医保病人!"
        Exit Function
    End If
    
    g病人身份_重庆渝北.就诊时间 = Format(rsTemp!就诊时间, "yyyy-MM-dd HH:mm:ss")
    g病人身份_重庆渝北.就诊编号 = Nvl(rsTemp!就诊编号)
    g病人身份_重庆渝北.结算编号 = Nvl(rsTemp!结算编号)
    g病人身份_重庆渝北.虚拟结算 = False
    
       
    Err = 0
    On Error GoTo errHand
    
    gcnOracle_CQYB.BeginTrans
    
    If 病人结算(lng结帐ID) = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    '结算完成后，将当前的结算编号置为空.
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆渝北 & ",'结算编号','''" & "" & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算编号")
    gcnOracle_CQYB.CommitTrans
    If m虚拟结算信息.bln验证 Then
        strAdvance = g病人身份_重庆渝北.结算信息
    End If
    住院结算_重庆渝北 = True
    Exit Function
errHand:
    DebugTool "住院结算(住院结算_重庆渝北)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
    gcnOracle_CQYB.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_重庆渝北(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    MsgBox "医保不支持结帐作废，请直接作废记帐单据后，再结帐！", vbInformation, gstrSysName
    住院结算冲销_重庆渝北 = False
End Function
Public Function 处方登记_重庆渝北(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng病人ID As Long
    Dim blnUpload As Boolean
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errHand:
    
    
    g病人身份_重庆渝北.结算标志 = 3
    g病人身份_重庆渝北.冲销 = lng记录状态 <> 1
    
    
    
    '更新分配结算号
    gstrSQL = " Select distinct a.病人id,b.就诊编号,b.结算编号  From 住院费用记录 a,保险帐户 b  " & _
        " where a.no=[1] and a.记录状态=[2] and a.记录性质=[3] and a.病人id=b.病人id  and b.险类=[4]"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人相关信息", str单据号, lng记录状态, lng记录性质, TYPE_重庆渝北)
    '主要考虑记帐表的这种表况
    
    Do While Not rsTemp.EOF
        If IsNull(rsTemp!结算编号) Then
            g病人身份_重庆渝北.结算编号 = Get结算编号
            gstrSQL = "zl_保险帐户_更新信息(" & Nvl(rsTemp!病人ID, 0) & "," & TYPE_重庆渝北 & ",'结算编号','''" & g病人身份_重庆渝北.结算编号 & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算编号")
        End If
        rsTemp.MoveNext
    Loop
    
    
    '第一步: 读取费用明细记录
    gstrSQL = Get明细记录(0, str单据号, lng记录性质, lng记录状态)
    
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取处方明细")
    If rs明细.RecordCount = 0 Then
        ShowMsgbox "没有明细记录，可能相关项目未进行相应的对码"
        Exit Function
    End If
    gcnOracle_CQYB.BeginTrans
    
    If Save医保明细数据(rs明细) = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    
    '第二步:处方明细上传
    If 处方明细上传(rs明细) = False Then
        ShowMsgbox "在进行处方明细上传时存在一条以上的明细上传失败,请以后注意补传!"
    End If
    gcnOracle_CQYB.CommitTrans
    处方登记_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    gcnOracle_CQYB.RollbackTrans
End Function

'---------------------------------------------------------------------------------------------------------------------------
'声明:
'
'---------------------------------------------------------------------------------------------------------------------------
Public Function 业务请求_重庆渝北(ByVal int业务类型 As 业务类型_重庆渝北, ByVal strInputString As String, ByRef strOutPutstring As String, Optional ByRef strErrMsg As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串
    '         strOutPutString-输出串
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String
    Dim strOutput As String
    Dim AppStruct As Struct
    Dim blnReturn As Boolean    '返回错误
    Dim intAppCode As Integer
    
    
    
    StrInput = strInputString
    DebugTool "进入:" & int业务类型 & ":" & strInputString
    strOutput = ""
    If InitInfor_重庆渝北.模拟数据 Then
        '读取模拟数据
        Read模拟数据 int业务类型, strInputString, strOutPutstring
         业务请求_重庆渝北 = True
        Exit Function
    End If
  
    AppStruct.strErrMsg = Space(4500)
    strOutput = ""
    
    '业务请求
    'blnReturn = DataUpload(strInPut, strOutput, AppStruct)
    '按新接口定义
    
     
    Err = 0
    On Error GoTo errHand:
    
    blnReturn = gobjYingHaiDll.dll_main_in(StrInput, strOutput, intAppCode, strErrMsg)
    
    DebugTool "业务请求:" & int业务类型 & " 参:" & strOutput & " Err:" & strErrMsg
    
    If blnReturn = False Then
      '发生错误,提示出信息
        ShowMsgbox strErrMsg
        业务请求_重庆渝北 = False
        Exit Function
    End If
    strOutPutstring = strOutput
    业务请求_重庆渝北 = True
    Exit Function
    
    
    strErrMsg = ""
    If AppStruct.lngAppCode = 1 Then
        业务请求_重庆渝北 = True
    ElseIf AppStruct.lngAppCode < 0 Then
        '发生错误,提示出信息
        ShowMsgbox AppStruct.strErrMsg
        strErrMsg = AppStruct.strErrMsg
        业务请求_重庆渝北 = False
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Read模拟数据(ByVal int业务类型 As 业务类型_重庆渝北, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--功  能:通过该功能读取模拟数据,以便测试
    '--入参数:
    '--出参数:
    '--返  回:字串
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    
    strFile = App.Path & "\模拟提交串.txt"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
        
    Select Case int业务类型
        Case 获取系统时间
            strTemp = "获取系统时间"
            '以本地时间为准
        Case 身份鉴别
            strTemp = "身份鉴别"
        Case 修改密码
            strTemp = "修改密码"
        Case IC卡帐户支付
            strTemp = "IC卡帐户支付"
        Case 资格审批待遇核定
            strTemp = "资格审批待遇核定"
        Case 就诊信息写入
            strTemp = "就诊信息写入"
        Case 处方明细写入
            strTemp = "处方明细写入"
        Case 结算基本信息写入
            strTemp = "结算基本信息写入"
        Case 结算结果写入
            strTemp = "结算结果写入"
        Case 核对帐户支付信息
            strTemp = "核对帐户支付信息"
        Case 核对就诊信息
            strTemp = "核对就诊信息"
        Case 核对处方明细信息
            strTemp = "核对处方明细信息"
        Case 核对费用结算结果
            strTemp = "核对费用结算结果"
        Case 核对费用结算基本信息
            strTemp = "核对费用结算基本信息"
        Case 导出服务项目目录
            strTemp = "导出服务项目目录"
        Case 导出ICD_10信息
            strTemp = "导出ICD_10信息"
        Case 导出病种目录
            strTemp = "导出病种目录"
        Case 导出病种就诊结算信息
            strTemp = "导出病种就诊结算信息"
        Case 导出医保定点信息
            strTemp = "导出医保定点信息"
        Case 获取客户机标识号
            strTemp = "获取客户机标识号"
        Case 解析卡内数据
            strTemp = "解析卡内数据"
        Case 获取就诊编号
            strTemp = "获取就诊编号"
        Case 获取记帐流水号
            strTemp = "获取记帐流水号"
        Case 获取结算编号
            strTemp = "获取结算编号"
        Case 费用结算
            strTemp = "费用结算"
        Case 作废审批记录
            strTemp = "作废审批记录"
    End Select
    
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    objText.WriteLine "[" & strTemp & "]"
    objText.WriteLine strInputString
    objText.Close
    If int业务类型 = 解析卡内数据 Then
        strFile = App.Path & "\解析卡.txt"
    Else
        strFile = App.Path & "\医保模拟数据.txt"
    End If
    
    Dim blnStart As Boolean
    Dim strArr
    Err = 0
    On Error GoTo errHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                If int业务类型 = 解析卡内数据 Then
                    strArr = Split(strText, vbTab)
                    If Val(strArr(0)) = 1 Then
                        With g病人身份_重庆渝北
                             .个人编号 = strArr(1)
                             .卡号 = strArr(2)
                         End With
                        str = strArr(1) & vbTab & strArr(2)
                        Exit Do
                    End If
                Else
                    If blnStart Then
                        If strText = "" Then
                            strText = "" & vbTab
                        End If
                        strArr = Split(strText, vbTab)
                        If strArr(0) = strInputString Then
                            str = strArr(1)
                            Exit Do
                        End If
                   Else
                        If "<" & strTemp & ">" = strText Then
                            blnStart = True
                        End If
                   End If
                    If "</" & strTemp & ">" = strText Then
                        Exit Do
                    End If
                End If
            Loop
            objText.Close
            strOutPutstring = str
    
    End If
    
    Exit Function
errHand:
    DebugTool Err.Description
   
End Function
Private Function GetXML串(ByVal strInputXMLString As String, Optional blnLoadRoot As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:将XML串加入对角gobjOutput
    '--入参数:blnLoadRoot-是否自动加入Root接点
    '--出参数:
    '--返  回:加载成功,返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    
    If blnLoadRoot Then
        strXMLText = "<" & gstrXMLRootPart & ">" & strInputXMLString & "</" & gstrXMLRootPart & ">"
    Else
        strXMLText = strInputXMLString
    End If
    
    GetXML串 = gobjXMLOutput.loadXML(strXMLText)
End Function
Private Function intXML() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:初始XML
    '--入参数:
    '--出参数:
    '--返  回:初始成功返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim nodData As MSXML2.IXMLDOMElement
    
       
    On Error Resume Next
    Set gobjXMLInPut = New MSXML2.DOMDocument
    Set gobjXMLOutput = New MSXML2.DOMDocument
    If Err <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set nodData = gobjXMLInPut.createElement(gstrXMLRootPart)
    Set gobjXMLInPut.documentElement = nodData
    intXML = True
End Function
Private Function AppendXMLNode(nodParent As MSXML2.IXMLDOMElement, ByVal Name As String, ByVal Value As String) As MSXML2.IXMLDOMElement
    '功能：在指定XML元素下增加子元素
    Set AppendXMLNode = gobjXMLInPut.createElement(Name)
    AppendXMLNode.Text = Value
    nodParent.appendChild AppendXMLNode
End Function
Public Function GetXMLOutput(ByVal Name As String, Optional blnName As Boolean = True, Optional lngRow As Long = 0) As String

    '功能：得到指定元素的值
    '参数:blnName-根据名称来取值
    Dim xmlElement As MSXML2.IXMLDOMElement
    If blnName Then
        Set xmlElement = gobjXMLOutput.getElementsByTagName(Name).Item(lngRow)
    Else
        Set xmlElement = gobjXMLOutput.documentElement.selectSingleNode(Name)
    End If
    If Not xmlElement Is Nothing Then
        '找到指定子元素
        GetXMLOutput = xmlElement.Text
    End If
End Function
Public Function 修改密码_重庆渝北(ByVal strOldPassWord As String, ByVal strNewPassWord As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:远程修改医保人员的帐户密码
    '--入参数:
    '--出参数:
    '--返  回:成功true,否则False
    '-----------------------------------------------------------------------------------------------------------
    Dim strOutput As String
    Dim strXMLText As String
    Dim blnReturn As Boolean
    Err = 0
    On Error GoTo errHand:
    
    修改密码_重庆渝北 = False
    If intXML = False Then Exit Function
        
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(g病人身份_重庆渝北.社保经办构构代码, 1, 4)
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "04"
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    AppendXMLNode gobjXMLInPut.documentElement, "code", Substr(g病人身份_重庆渝北.卡号, 1, 20)
    AppendXMLNode gobjXMLInPut.documentElement, "ykc005", Substr(strOldPassWord, 1, 6)
    AppendXMLNode gobjXMLInPut.documentElement, "New_ykc005", Substr(strNewPassWord, 1, 6)
    
    strXMLText = gobjXMLInPut.documentElement.xml
    '取掉前导XML串
    strXMLText = Mid(strXMLText, Len(gstrXMLRootPart) + 3, Len(strXMLText) - 3)
        
    '业务请求
    
    blnReturn = 业务请求_重庆渝北(修改密码, strXMLText, strOutput)
    If blnReturn = False Then
        Exit Function
    End If
    
    '输出串
    修改密码_重庆渝北 = True
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    修改密码_重庆渝北 = False
End Function
Public Function 解析卡_重庆渝北() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:解析卡内数据
    '--入参数:strCardData-卡内数据
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strNO As String
    Dim strCardNO As String
    Dim strExInfor As String
   ' Dim objTest As New Mwic_32.clsMwic_32
    
    Err = 0
    On Error GoTo errHand:
    If InitInfor_重庆渝北.解析卡内数据 Then
        Read模拟数据 解析卡内数据, "", strExInfor
        
    Else
        strNO = Space(10)
        strCardNO = Space(12)
        strExInfor = Space(4)
        '--Bug需要屏蔽
        Call srd_4428_info(strNO, strCardNO, strExInfor)
        With g病人身份_重庆渝北
            .个人编号 = strNO
            .卡号 = strCardNO
        End With
    End If
    解析卡_重庆渝北 = True
    Exit Function
errHand:
    解析卡_重庆渝北 = False
    ShowMsgbox "IC卡错误,不能识别!"
End Function
Public Function 核对病人就诊信息_重庆渝北(ByVal lng病人ID As Long) As String
  '-----------------------------------------------------------------------------------------------------------
    '--功  能:核对病人就诊信息
    '--入参数:
    '--出参数:
    '--返  回:返回核对范围内的记录数
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    Dim strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    Dim strTemp As String
    
    Call Get病人信息(lng病人ID)
    
    Call intXML
    
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "14"
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_重庆渝北.医院编码, 1, 8)
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g病人身份_重庆渝北.就诊编号
    
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    
    If 业务请求_重庆渝北(核对就诊信息, strXMLText, strOutput, "") = False Then
        ShowMsgbox "读取核对就诊信息时出院"
        Exit Function
    End If
    
    If GetXML串(strOutput) = False Then
        ShowMsgbox "核对就诊信息中返回串不是一个有效的XML串！"
        Exit Function
    End If
    lngCount = Val(GetXMLOutput("RecordCount"))
    
    
    '查找入出记录数
    gstrSQL = "Select count(distinct a.病人id||' '||a.主页id)  as 总数 From 病案主页 a,保险帐户 b where a.病人id=b.病人id "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医院医保病人的入出信息")
    
    strTemp = "中心记录数为:" & lngCount & "|" & Nvl(rsTemp!总数, 0)
    frmShowMsg.ShowInFor strTemp
End Function

Public Function 核对病人帐户支付信息_重庆渝北(ByVal lng病人ID As Long) As String
  '-----------------------------------------------------------------------------------------------------------
    '--功  能:核对病人就诊信息
    '--入参数:
    '--出参数:
    '--返  回:返回核对范围内的记录数
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    Dim strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    Dim dbl总额 As Double
    Dim strTemp As String
    
    Get病人信息 lng病人ID
    
    Call intXML
    'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    'SvrcID  string  2       远程数据服务标识，定值13, 标识大小写敏感，足两位长
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "13"
    'CtrInf  string  20      控制信息，预留, 标识大小写敏感
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'yab003  string  4       社保经办机构代码
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_重庆渝北.医院编码, 1, 8)
    'akc190  string  20      需要核对的账户支付信息的就诊编号
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g病人身份_重庆渝北.就诊编号
    
    
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    
    If 业务请求_重庆渝北(核对帐户支付信息, strXMLText, strOutput, "") = False Then
        ShowMsgbox "核对帐户支付信息时,业务请求失败！"
        Exit Function
    End If
    
    If GetXML串(strOutput) = False Then
        ShowMsgbox "核对帐户支付信息中返回串不是一个有效的XML串！"
        Exit Function
    End If
    'RecordCount number  15      在核对范围内的所有信息的记录数量
    lngCount = Val(GetXMLOutput("RecordCount"))
    'DefrayAmount    string  14  2   在核对范围内的所有记录的帐户（卡）支付总额累加值
    dbl总额 = Val(GetXMLOutput("DefrayAmount"))
    
    '查找入出记录数
    gstrSQL = "Select count(记录ID) as 记录数,sum(个人帐户支付) as 总额  From 保险结算记录 where nvl(个人帐户支付,0)<>0 and 备注=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医院核对帐户支付信息", g病人身份_重庆渝北.就诊编号)
    
    strTemp = "记录数为:" & lngCount & "|" & Nvl(rsTemp!记录数, 0) & "||支付总额为:" & Format(dbl总额, "####0.00;-####0.00; ;") & "元|" & Format(Nvl(rsTemp!总额, 0), "####0.00;-####0.00; ;") & "元"
    frmShowMsg.ShowInFor strTemp
End Function


Public Function 核对病人费用结算基本信息_重庆渝北(ByVal lng病人ID As Long) As String
  '-----------------------------------------------------------------------------------------------------------
    '--功  能:核对费用结算基本信息
    '--入参数:
    '--出参数:
    '--返  回:返回核对范围内的记录数
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    Dim strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCount As Long
    Dim dbl总额 As Double
    Get病人信息 lng病人ID
    Call intXML
    'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    'SvrcID  string  2       远程数据服务标识，定值15, 标识大小写敏感，足两位长
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "13"
    'CtrInf  string  20      控制信息，预留, 标识大小写敏感
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'yab003  string  4       社保经办机构代码
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_重庆渝北.医院编码, 1, 8)
    'akc190  string  20      需要核对的账户支付信息的就诊编号
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g病人身份_重庆渝北.就诊编号
    
    
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    
    If 业务请求_重庆渝北(核对费用结算基本信息, strXMLText, strOutput, "") = False Then
        ShowMsgbox "核对费用结算基本信息时,业务请求失败！"
        Exit Function
    End If
    
    If GetXML串(strOutput) = False Then
        ShowMsgbox "核对费用结算基本信息中返回串不是一个有效的XML串！"
        Exit Function
    End If
    gstrSQL = "Select count(distinct 结算编号 ) as 记录数,0 as 费用总额,sum(全自费金额) as 全自费,sum(挂构自费) as 挂构自费,sum(符合金额) as 符合金额,sum(本次自付) as 个人帐户,0 as 个人现金    From 费用结算结果 where 就诊编号='" & g病人身份_重庆渝北.就诊编号 & "'"
    OpenRecordset_ZLYB rsTemp, "获取费用结算基本信息"
    
    'RecordCount number  15      在核对范围内的所有信息的记录数量
    strTemp = "记录数:" & Val(GetXMLOutput("RecordCount")) & "|" & Nvl(rsTemp!记录数, 0)
    'yka055  number  14  2   在核对范围内的所有记录的医疗费总额累加值
    strTemp = strTemp & "||" & "医疗费用总额:" & Format(Val(GetXMLOutput("yka055")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!费用总额, 0)
    'yka056  number  14  2   在核对范围内的所有记录的全自费总额累加值
    strTemp = strTemp & "||" & "全自费  总额:" & Format(Val(GetXMLOutput("yka056")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!全自费, 0)
    'yka057  number  14  2   在核对范围内的所有记录的挂钩自费总额累加值
    strTemp = strTemp & "||" & "挂钩自费总额:" & Format(Val(GetXMLOutput("yka057")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!挂构自费, 0)
    'yka111  number  14  2   在核对范围内的所有记录的符合范围总额累加值
    strTemp = strTemp & "||" & "符合金额总额:" & Format(Val(GetXMLOutput("yka111")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!符合金额, 0)
    'yka112  number  14  2   在核对范围内的所有记录的个人账户支付总额累加值
    strTemp = strTemp & "||" & "个人帐户总额:" & Format(Val(GetXMLOutput("yka112")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!个人帐户, 0)
    'yka113  number  14  2   在核对范围内的所有记录的个人现金支付总额累加值
    strTemp = strTemp & "||" & "个人现金总额:" & Format(Val(GetXMLOutput("yka113")), "####0.00;-####0.00; 0;") & "|" & Nvl(rsTemp!个人现金, 0)
    '查找入出记录数
    frmShowMsg.ShowInFor strTemp
End Function


Public Function 核对处方明细信息_重庆渝北(ByVal lng病人ID As Long) As String
  '-----------------------------------------------------------------------------------------------------------
    '--功  能:核对处方明细信息
    '--入参数:
    '--出参数:
    '--返  回:返回核对范围内的记录数
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    Dim strOutput As String
    Dim strTable As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCount As Long
    Dim dbl总额 As Double
    
    Call Get病人信息(lng病人ID)
    strTable = IIf(g病人身份_重庆渝北.支付类别 Like "030*", "住院费用明细", "门诊费用明细")
    
    Call intXML
    'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    'SvrcID  string  2       远程数据服务标识，定值15, 标识大小写敏感，足两位长
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "16"
    'CtrInf  string  20      控制信息，预留, 标识大小写敏感
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'yab003  string  4       社保经办机构代码
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_重庆渝北.医院编码, 1, 8)
    'akc190  string  20      需要核对的账户支付信息的就诊编号
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g病人身份_重庆渝北.就诊编号
    
    
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    
    If 业务请求_重庆渝北(核对处方明细信息, strXMLText, strOutput, "") = False Then
        ShowMsgbox "核对处方明细信息时,业务请求失败！"
        Exit Function
    End If
    
    If GetXML串(strOutput) = False Then
        ShowMsgbox "核对处方明细信息中返回串不是一个有效的XML串！"
        Exit Function
    End If
    gstrSQL = " Select count(ID) as 记录数,Sum(付数*数次) as 数量,sum(Round(A.实收金额/(A.数次*A.付数),2)) as 价格,sum(实收金额) as 费用总额  " & _
              " From " & strTable & " A  " & _
              " Where ID in (Select 费用ID From 医保明细费用 where  就诊编号=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取处方明细信息", g病人身份_重庆渝北.就诊编号)
    
    'RecordCount number  15      在核对范围内的所有信息的记录数量
    strTemp = "明细记录数:" & Val(GetXMLOutput("RecordCount")) & "|" & Nvl(rsTemp!记录数, 0)
    'akc226  number  14  2   在核对范围内的所有记录的数量累加值
    strTemp = strTemp & "||" & "明细总数量:" & Format(Val(GetXMLOutput("akc226")), "####0.0000;-####0.00; 0;") & "|" & Nvl(rsTemp!数量, 0)
    'akc225  number  14  2   在核对范围内的所有记录的实际价格总额累加值
    strTemp = strTemp & "||" & "实际价格总额:" & Format(Val(GetXMLOutput("akc225")), "####0.00;-####0.000;0 ;") & "|" & Nvl(rsTemp!价格, 0)
    'yka055  number  14  2   在核对范围内的所有记录的医疗费总额累加值
    strTemp = strTemp & "||" & "医疗费总额:" & Format(Val(GetXMLOutput("yka055")), "####0.00;-####0.000;0 ;") & "|" & Nvl(rsTemp!费用总额, 0)
    
    '查找入出记录数
    frmShowMsg.ShowInFor strTemp
End Function


Public Function 核对费用结算结果_重庆渝北(ByVal lng病人ID As Long) As String
  '-----------------------------------------------------------------------------------------------------------
    '--功  能:核对处方明细信息
    '--入参数:
    '--出参数:
    '--返  回:返回核对范围内的记录数
    '-----------------------------------------------------------------------------------------------------------
    Dim strXMLText As String
    Dim strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCount As Long
    Dim dbl总额 As Double
    
    Call Get病人信息(lng病人ID)
    
    Call intXML
    'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    'SvrcID  string  2       远程数据服务标识，定值15, 标识大小写敏感，足两位长
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "18"
    'CtrInf  string  20      控制信息，预留, 标识大小写敏感
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'yab003  string  4       社保经办机构代码
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_重庆渝北.医院编码, 1, 8)
    'akc190  string  20      需要核对的账户支付信息的就诊编号
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", g病人身份_重庆渝北.就诊编号
    
    
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    
    If 业务请求_重庆渝北(核对费用结算结果, strXMLText, strOutput, "") = False Then
        ShowMsgbox "核对费用结算结果时,业务请求失败！"
        Exit Function
    End If
    
    If GetXML串(strOutput) = False Then
        ShowMsgbox "核对费用结算结果中返回串不是一个有效的XML串！"
        Exit Function
    End If
    
    gstrSQL = "Select count(distinct 结算编号 ) as 记录数,0 as 费用总额,sum(全自费金额) as 全自费,sum(挂构自费) as 挂构自费,sum(符合金额) as 符合金额,sum(本次自付) as 个人帐户," & _
    "       sum(本次支付金额) as 支付金额,sum(公务员统筹支付) as 公务员统筹支付,sum(补助自付累计) as 补助自付累计  " & _
    "   From 费用结算结果 where 就诊编号='" & g病人身份_重庆渝北.就诊编号 & "'"
    
    OpenRecordset_ZLYB rsTemp, "获取费用结算基本信息"
    
    'RecordCount number  15      在核对范围内的所有信息的记录数量
    strTemp = "记录数:" & Val(GetXMLOutput("RecordCount")) & "|" & Nvl(rsTemp!记录数, 0)
    'aka213  string  6       分段标准
    strTemp = strTemp & "||" & "分段标准:" & Val(GetXMLOutput("aka213")) & "|" & "0"
    'yka056  number  14  2   在核对范围内的所有记录的全自费金额累加值
    strTemp = strTemp & "||" & "全自费  总额:" & Format(Val(GetXMLOutput("yka056")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!全自费, 0)
    'yka057  number  14  2   在核对范围内的所有记录的挂钩自费金额累加值
    strTemp = strTemp & "||" & "挂钩自费总额:" & Format(Val(GetXMLOutput("yka057")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!挂构自费, 0)
    'yka111  number  14  2   在核对范围内的所有记录的符合范围金额累加值
    strTemp = strTemp & "||" & "符合金额总额:" & Format(Val(GetXMLOutput("yka111")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!符合金额, 0)
    
    'yka106  number  14  2   在核对范围内的所有记录的自付金额累加值
    strTemp = strTemp & "||" & "    自付总额:" & Format(Val(GetXMLOutput("yka106")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!个人帐户, 0)
    'yka107  number  14  2   在核对范围内的所有记录的支付金额累加值
    strTemp = strTemp & "||" & "    支付金额:" & Format(Val(GetXMLOutput("yka107")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!支付金额, 0)
    'yka063  number  14  2   在核对范围内的所有记录的公务员统筹支付金额累加值
    strTemp = strTemp & "||" & "公务员统筹支付:" & Format(Val(GetXMLOutput("yka063")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!公务员统筹支付, 0)
    'yka221  number  14  2   在核对范围内的所有记录的享受医疗补助个人自付累计金额
    strTemp = strTemp & "||" & "补助自付累计:" & Format(Val(GetXMLOutput("yka063")), "####0.00;-####0.00; ;") & "|" & Nvl(rsTemp!补助自付累计, 0)
    
    frmShowMsg.ShowInFor strTemp
End Function
Private Function IC卡帐户支付_重庆渝北(ByVal dbl本年支付 As Double, ByVal str经办时间 As String, ByVal str退单结算号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:进行远程IC卡支付
    '--入参数:
    '--出参数:
    '--返  回:支付成功,返回true,返回False
    '-----------------------------------------------------------------------------------------------------------
    IC卡帐户支付_重庆渝北 = False
    
    Err = 0
    On Error GoTo errHand:
    
    Call intXML
    
    'YAB003  String  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    If g病人身份_重庆渝北.冲销 = True Then
        'SvrcID  String  2       远程数据服务标识，定值05, 标识大小写敏感，足两位长
        AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "19"
    Else
        'SvrcID  String  2       远程数据服务标识，定值05, 标识大小写敏感，足两位长
        AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "05"
    End If
    'CtrInf  String  20      控制信息，预留, 标识大小写敏感
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'code    String  20      参保人员的医保卡号
    AppendXMLNode gobjXMLInPut.documentElement, "code", Substr(g病人身份_重庆渝北.卡号, 1, 20)
    'ykc005  String  6       就诊参保人员输入的医保认证密码，足六位长的数字字符
    AppendXMLNode gobjXMLInPut.documentElement, "ykc005", Substr(g病人身份_重庆渝北.密码, 1, 6)
    'akc190  String  20      就诊编号
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", Substr(g病人身份_重庆渝北.就诊编号, 1, 20)
    'aka130  String  6       支付类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", Substr(g病人身份_重庆渝北.支付类别, 1, 6)
    'akb020  String  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_重庆渝北.医院编码, 1, 8)
    'ykb006  String  3       定点医疗机构分支机构编号
    AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"
    


    If g病人身份_重庆渝北.冲销 = True Then
        AppendXMLNode gobjXMLInPut.documentElement, "DefrayAmount", dbl本年支付
    End If
    
    'PastBaseDefray  Number  14  2   基本医疗历年支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "PastBaseDefray", 0
    'LastBaseDefray  Number  14  2   基本医疗上年支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "LastBaseDefray", 0
    'ThisBaseDefray  Number  14  2   基本医疗本年支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "ThisBaseDefray", dbl本年支付
    'NotPastBaseDefray   Number  14  2   基本医疗本年划入非本年账户历年支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "NotPastBaseDefray", 0
    'NotLastBaseDefray   Number  14  2   基本医疗本年划入非本年账户上年支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "NotLastBaseDefray", 0
    'NotThisBaseDefray   Number  14  2   基本医疗本年划入非本年账户本年支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "NotThisBaseDefray", 0
    'PastOfficialDefray  Number  14  2   公务员历年支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "PastOfficialDefray", 0
    'LastOfficialDefray  Number  14  2   公务员上年支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "LastOfficialDefray", 0
    'ThisOfficialDefray  Number  14  2   公务员本年支付总额
    AppendXMLNode gobjXMLInPut.documentElement, "ThisOfficialDefray", 0
    'aae036  Date        秒  账户支付的经办时间
    AppendXMLNode gobjXMLInPut.documentElement, "aae036", str经办时间
    'Yka198  String  20      退单对应结算编号（此处为结算编号）
    
    AppendXMLNode gobjXMLInPut.documentElement, "Yka198", str退单结算号
    
    Dim strXMLText As String, strOutput As String
    
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    
    WriteDebugInfor_重庆渝北 strXMLText
    If 业务请求_重庆渝北(IC卡帐户支付, strXMLText, strOutput, "") = False Then
        ShowMsgbox "IC卡帐户支付时,业务请求失败！"
        Exit Function
    End If
    IC卡帐户支付_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function 审批记录作废_重庆渝北() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:审批记录作废
    '--入参数:
    '--出参数:
    '--返  回:作废成功,返回true,返回False
    '-----------------------------------------------------------------------------------------------------------
    审批记录作废_重庆渝北 = False
    
    Err = 0
    On Error GoTo errHand:
    
    If g病人身份_重庆渝北.虚拟结算 Then
        审批记录作废_重庆渝北 = True
        Exit Function
    End If
        
    Call intXML
    
    'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
    AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Substr(InitInfor_重庆渝北.经办机构代码, 1, 4)
    'SvrcID  string  2       远程数据服务标识，定值26, 标识大小写敏感，足两位长
    AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "26"
    'CtrInf  string  20      控制信息，预留, 标识大小写敏感
    AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
    'aac001  number  15  0   个人编号
    AppendXMLNode gobjXMLInPut.documentElement, "aac001", g病人身份_重庆渝北.个人编号
    
    'ykc005  string  6       就诊参保人员输入的医保认证密码，足六位长的数字字符
    AppendXMLNode gobjXMLInPut.documentElement, "ykc005", Substr(g病人身份_重庆渝北.密码, 1, 6)
    'akc190  string  20      就诊编号
    AppendXMLNode gobjXMLInPut.documentElement, "akc190", Substr(g病人身份_重庆渝北.就诊编号, 1, 20)
    'aka130  string  6       支付类别，见代码表
    AppendXMLNode gobjXMLInPut.documentElement, "aka130", Substr(g病人身份_重庆渝北.支付类别, 1, 6)
    'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
    AppendXMLNode gobjXMLInPut.documentElement, "akb020", Substr(InitInfor_重庆渝北.医院编码, 1, 8)
    
    Dim strXMLText As String, strOutput As String
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    
    If 业务请求_重庆渝北(作废审批记录, strXMLText, strOutput, "") = False Then
        ShowMsgbox "审批记录作废时,业务请求失败！"
        Exit Function
    End If
    审批记录作废_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Get代码数据_重庆渝北(ByVal intCode As CodeType, ByVal strCode As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:根据名称,获取固定值
    '--入参数:
    '--出参数:
    '--返  回:名称
    '-----------------------------------------------------------------------------------------------------------
    Dim STRNAME As String
    Select Case intCode
    Case 医疗人员类别
        STRNAME = Switch(strCode = "41", "下岗职工", strCode = "21", "退休", strCode = "22", "退休异地安置", strCode = "34", "离休异地安置", strCode = "12", "在职长期驻外", strCode = "11", "在职", strCode = "31", "离休", strCode = "33", "二等乙级伤残军人", strCode = "32", "老红军", strCode = "51", "建国前老工人", True, "其他人员")
    Case 医疗补助类别
        STRNAME = Switch(strCode = "1", "享受医疗补助", True, "不享受医疗补助")
    Case 医疗照顾类别
        STRNAME = Switch(strCode = "0", "非医疗照顾人员", strCode = "1", "医疗照顾人员甲类", True, "医疗照顾人员乙类")
    End Select
    Get代码数据_重庆渝北 = STRNAME
End Function
Private Function Get就诊时间(ByVal lng病人ID As Long) As String
    '功能：获取就诊时间
    '参数：
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 就诊时间 From 保险帐户 where 险类=" & TYPE_重庆渝北 & " and 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取就诊时间"
    
    If rsTemp.RecordCount = 0 Then
        Get就诊时间 = ""
        Exit Function
    End If
    Get就诊时间 = Format(rsTemp!就诊时间, "yyyy-mm-dd")
End Function


Public Function 挂号结算_重庆渝北(ByVal lng结帐ID As Long) As Boolean
  '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    
    Dim str时间 As String
    Dim rsTemp As New ADODB.Recordset
    
    挂号结算_重庆渝北 = False
    
    g病人身份_重庆渝北.结算标志 = 2
    g病人身份_重庆渝北.冲销 = False
    g病人身份_重庆渝北.结算编号 = Get结算编号
    g病人身份_重庆渝北.结帐ID = lng结帐ID
    g病人身份_重庆渝北.发票号 = Get发票号码(lng结帐ID)
    g病人身份_重庆渝北.虚拟结算 = False

    
    
    gstrSQL = "Select 病人id,登记时间 From 门诊费用记录 where rownum<=1 and 结帐id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取登记时间"
    
    If g病人身份_重庆渝北.就诊时间 > Format(rsTemp!登记时间, "yyyy-MM-dd HH:mm:ss") Then
        g病人身份_重庆渝北.就诊时间 = Format(rsTemp!登记时间, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '保存当前状态的结算编号
    gstrSQL = "zl_保险帐户_更新信息(" & Nvl(rsTemp!病人ID, 0) & "," & TYPE_重庆渝北 & ",'结算编号','''" & g病人身份_重庆渝北.结算编号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算编号")
        
    Err = 0
    On Error GoTo errHand
    gcnOracle_CQYB.BeginTrans
    
    挂号结算_重庆渝北 = 病人结算(lng结帐ID)
    
    If 挂号结算_重庆渝北 = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
        gcnOracle_CQYB.CommitTrans
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
  gcnOracle_CQYB.RollbackTrans
End Function


Public Function 挂号冲销_重庆渝北(ByVal lng结帐ID As Long) As Boolean

    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errHand
    
    挂号冲销_重庆渝北 = False
    gstrSQL = "Select 病人ID From 门诊费用记录  where 结帐id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人id"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "没有被冲销的挂号项目,不能冲销"
        
        Exit Function
    End If
    '需确定该时间里是否有费用结算,否则不能进行冲销
    
    
    g病人身份_重庆渝北.lng病人ID = Nvl(rsTemp!病人ID, 0)
    '获取基本信息
    Call Get病人信息(g病人身份_重庆渝北.lng病人ID)
    
    g病人身份_重庆渝北.结帐ID = lng结帐ID
    g病人身份_重庆渝北.结算标志 = 2
    g病人身份_重庆渝北.冲销 = False
    g病人身份_重庆渝北.结算编号 = Get结算编号
    g病人身份_重庆渝北.冲销 = True
    g病人身份_重庆渝北.冲销ID = Get冲销ID
    g病人身份_重庆渝北.虚拟结算 = False

    '保存当前状态的结算编号
    gstrSQL = "zl_保险帐户_更新信息(" & g病人身份_重庆渝北.lng病人ID & "," & TYPE_重庆渝北 & ",'结算编号','''" & g病人身份_重庆渝北.结算编号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算编号")
    
    gcnOracle_CQYB.BeginTrans
    
    挂号冲销_重庆渝北 = 病人结算冲销(lng结帐ID)
    If 挂号冲销_重庆渝北 = False Then
        gcnOracle_CQYB.RollbackTrans
        Exit Function
    End If
    gcnOracle_CQYB.CommitTrans
    挂号冲销_重庆渝北 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle_CQYB.RollbackTrans
End Function

Private Function 病人结算冲销(ByVal lng结帐ID As Long) As Boolean
    '冲销病人结算
    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim str结算号 As String
    On Error GoTo errHand
    
    病人结算冲销 = False
    
    Err = 0
    On Error GoTo errHand
        
    
    '第一步:由于是冲销,所以不需进行待遇核定,但需获取原单据的结算编号
    gstrSQL = "select  支付顺序号,备注 from 保险结算记录 where 险类=" & TYPE_重庆渝北 & " and 记录id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取结算信息"
    
    str结算号 = Nvl(rsTemp!支付顺序号)
    g病人身份_重庆渝北.就诊编号 = Nvl(rsTemp!备注)
    
    '第二步:写入本次需明细结算的文本
    
    '   读取费用明细记录
    gstrSQL = Get明细记录(g病人身份_重庆渝北.冲销ID)
    
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取处方明细")
    
    If rs明细.RecordCount = 0 Then
        Exit Function
    End If
    
    If Save医保明细数据(rs明细) = False Then Exit Function
    
    
    '第三步:将历将的费用结算结果及基本信息上传(以负方式上传)
     If 费用结算结果冲销上传(str结算号) = False Then Exit Function
     
    '第四步:处方明细上传
    If 处方明细上传(rs明细) = False Then
        ShowMsgbox "在进行处方明细上传时存在一条以上的明细上传失败,请以后注意补传!"
    End If
    病人结算冲销 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 费用结算结果冲销(ByVal str结算编号 As String) As Boolean
    
    Dim rsTemp As New ADODB.Recordset
    Dim strOutput As String
    Dim strXMLText As String
    Dim blnFirst As Boolean
    Dim strSQL  As String
    Dim objXMLItem As MSXML2.IXMLDOMElement

    费用结算结果冲销 = False
    
    '冲销费用结算结果,写相反数
    gstrSQL = "" & _
                "   Select id, 病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 医疗人员类别, 医疗照顾类别, 医疗补助类别," & _
                "       年度, 基本医疗信息, -1*下限金额 下限金额 , -1*自付金额 自付金额, 经办机构代码,支付金额 支付金额,-1* 公务员补助 公务员补助," & _
                "       -1*先行自付金额 先行自付金额, 累计缴费月数, 实足年龄," & _
                "       医疗行政类别, -1*帐户支付 帐户支付, 分段标准, -1*全自费金额 全自费金额, -1*挂构自费 挂构自费, -1*符合金额 符合金额, -1*本次自付 本次自付," & _
                "       -1*本次支付金额 本次支付金额, -1*公务员统筹支付 公务员统筹支付,-1*补助自付累计  补助自付累计 " & _
                "   From 费用结算结果 " & _
                "   where 结算编号='" & str结算编号 & "' and 就诊编号='" & g病人身份_重庆渝北.就诊编号 & "'"
                    
    Err = 0
    On Error GoTo errHand:
    
    OpenRecordset_ZLYB rsTemp, "冲销结算结果"
    
    If rsTemp.EOF Then
        ShowMsgbox "结算编号为:" & str结算编号 & " 的结算号不存在,冲销失败!"
        Exit Function
    End If
 
    '存储过程参数:
    'ID,病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 经办机构代码, 医疗人员类别, 医疗照顾类别, 医疗补助类别,
    '年度,基本医疗信息, 下限金额, 自付金额, 支付金额, 公务员补助, 先行自付金额, 累计缴费月数, 实足年龄, 医疗行政类别, 帐户支付, 分段标准,
    '全自费金额, 挂构自费,符合金额_IN, 本次自付, 本次支付金额, 公务员统筹支付, 补肋自付累计
    
    Call intXML
    blnFirst = True
    Dim lngID As Long
    With rsTemp
        Do While Not .EOF
            gstrSQL = "Select 费用结算结果_ID.nextval as ID from dual"
            OpenRecordset_ZLYB rsTemp, "获取结算编号"
            lngID = Nvl(rsTemp!ID, 0)
            
            strSQL = "ZL_费用结算结果_INSERT("
            
            
            strSQL = strSQL & lngID & ","
            strSQL = strSQL & Nvl(!病人ID, 0) & ","
            
            strSQL = strSQL & Nvl(!主页ID, 0) & ","
            strSQL = strSQL & "'" & Nvl(!就诊编号) & "',"
            strSQL = strSQL & "'" & g病人身份_重庆渝北.结算编号 & "',"
            strSQL = strSQL & "'" & Nvl(!结算编号) & "',"
            strSQL = strSQL & "" & Nvl(!审批记录序号, 0) & ","
            strSQL = strSQL & "'" & Nvl(!经办机构代码) & "',"
            strSQL = strSQL & "'" & Nvl(!医疗人员类别) & "',"
            strSQL = strSQL & "'" & Nvl(!医疗照顾类别) & "',"
            strSQL = strSQL & "'" & Nvl(!医疗补助类别) & "',"  '医疗补助类别
            strSQL = strSQL & "'" & Nvl(!年度) & "',"
            strSQL = strSQL & "'" & Nvl(!基本医疗信息) & "',"
            strSQL = strSQL & "" & Nvl(!下限金额, 0) & ","
            
            
            strSQL = strSQL & "" & Nvl(!自付金额, 0) & ","
            strSQL = strSQL & "" & Nvl(!支付金额, 0) & ","
            strSQL = strSQL & "" & Nvl(!公务员补助, 0) & ","
            strSQL = strSQL & "" & Nvl(!先行自付金额, 0) & ","
            strSQL = strSQL & "" & Nvl(!累计缴费月数, 0) & ","
            strSQL = strSQL & "" & Nvl(!实足年龄, 0) & ","
            
            strSQL = strSQL & "'" & Nvl(!医疗行政类别) & "',"
            strSQL = strSQL & "" & Nvl(!帐户支付, 0) & ","
            strSQL = strSQL & "'" & Nvl(!分段标准) & "',"
            
            
            
            strSQL = strSQL & "" & Nvl(!全自费金额, 0) & ","
            strSQL = strSQL & "" & Nvl(!挂构自费, 0) & ","
            strSQL = strSQL & "" & Nvl(!符合金额, 0) & ","
            strSQL = strSQL & "" & Nvl(!本次自付, 0) & ","
            strSQL = strSQL & "" & Nvl(!本次支付金额, 0) & ","
            
            strSQL = strSQL & "" & Nvl(!公务员统筹支付, 0) & ","
            strSQL = strSQL & "" & Nvl(!补助自付累计, 0) & ")"
            
                        
            '存入数据库中
            gcnOracle_CQYB.Execute strSQL, , adCmdStoredProc
            
            If insertInto子项(lngID, Nvl(!基本医疗信息)) = False Then
                DebugTool "插入子项错误!"
            End If
            'XML费用结果写入
            If blnFirst Then
                AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Nvl(!经办机构代码)
                AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "12"
                AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
        
                'BaseInfo                多个待遇享受段共有的相同的基本信息部分，下面缩排的元素是它的子元素
                Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "BaseInfo", "")
                '    akc190  string  20      就诊编号
                AppendXMLNode objXMLItem, "akc190", Nvl(!就诊编号)
                '    yka103  string  20      结算编号
                AppendXMLNode objXMLItem, "yka103", g病人身份_重庆渝北.结算编号
                '    yka198  string  20      退单对应结算编号
                AppendXMLNode objXMLItem, "yka198", Nvl(!结算编号)
                '    ykc114  number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
                AppendXMLNode objXMLItem, "ykc114", Nvl(!审批记录序号, 0)
                '    yab003  string  4       社保经办机构代码
                AppendXMLNode objXMLItem, "yab003", Nvl(!经办机构代码)
                blnFirst = False
            End If
            
            
            '需确定相关的字串
            'ReckonInfo              多个待遇享受段的结算分段信息，下面缩排的元素是它的子元素
            Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "ReckonInfo", "")
            
            'akc190  string  20      就诊编号
             AppendXMLNode objXMLItem, "akc190", Nvl(!就诊编号)
            'yka103  string  20      结算编号
             AppendXMLNode objXMLItem, "yka103", g病人身份_重庆渝北.结算编号
             
            'yka198  string  20      退单对应结算编号
             AppendXMLNode objXMLItem, "yka198", Nvl(!结算编号)
            'ykc114  number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
             AppendXMLNode objXMLItem, "ykc114", Nvl(!审批记录序号)
            'yab003  string  4       社保经办机构代码
             AppendXMLNode objXMLItem, "yab003", Nvl(!经办机构代码)
            'aka213  string  2       分段标准，03 起付线， 05 基本医疗 ，06 大额医疗，07 超限
             AppendXMLNode objXMLItem, "aka213", Nvl(!分段标准)
            'yka056  number  14  2   全自费金额
             AppendXMLNode objXMLItem, "yka056", Nvl(!全自费金额, 0)
            'yka057  number  14  2   挂钩自费金额
             AppendXMLNode objXMLItem, "yka057", Nvl(!挂构自费, 0)
            'yka111  number  14  2   符合范围金额
             AppendXMLNode objXMLItem, "yka111", Nvl(!符合金额, 0)
            'yka106  number  14  2   自付金额
             AppendXMLNode objXMLItem, "yka106", Nvl(!本次自付, 0)
            'yka107  number  14  2   支付金额
             AppendXMLNode objXMLItem, "yka107", Nvl(!支付金额, 0)
            'yka063  number  14  2   公务员统筹支付金额
             AppendXMLNode objXMLItem, "yka063", Nvl(!公务员统筹支付, 0)
            'yka221  number  14  2   享受医疗补助个人自付累计金额
             AppendXMLNode objXMLItem, "yka221", Nvl(!补助自付累计, 0)
            'Akc315  String  3       医疗行政职务
             AppendXMLNode objXMLItem, "Akc315", Nvl(!医疗行政类别)
            .MoveNext
        Loop
    End With
      
    '写入费用结算结果
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    WriteDebugInfor_重庆渝北 strXMLText
    
    If 业务请求_重庆渝北(结算结果写入, strXMLText, strOutput) = False Then
        Exit Function
    End If
    WriteDebugInfor_重庆渝北 strOutput
    费用结算结果冲销 = True
    Exit Function
    
errHand:
    DebugTool "费用结算结果冲销失败!" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & " 错误描述: " & Err.Description
 End Function
 Private Function 费用基本信息冲销(ByVal str结算编号 As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strXMLText As String
    Dim strSQL As String
    Dim str经办时间 As String, strOutput As String
    
    
    费用基本信息冲销 = False
    '写入费用结算基本信息
    '费用基本信息冲销
    gstrSQL = " " & _
    "           Select 病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 个人编号, 单位编号, 姓名, 性别, 出生日期, 实足年龄, " & _
    "                   累计缴费月数, 医疗人员类别, 医疗机构编码, 分支机构编码, 医疗机构类别, 特种病标志, 支付类别, 病种编码, 本次起付线, " & _
    "                   -1*医疗费总额 医疗费总额, -1*全自费总额 全自费总额, -1*挂钩自费总额 挂钩自费总额, -1*符合范围总额 符合范围总额, -1*个人帐户支付总额 个人帐户支付总额," & _
    "                  -1*个人现金支付总额 个人现金支付总额, 经办时间, 经办机构代码, " & _
    "                   医疗照顾类别 , 医疗补助类别, 就诊结算方式, 发票号, 备注, 分段计算情况, 医疗行政类别 " & _
    "           From 费用基本信息 " & _
    "           where 就诊编号='" & g病人身份_重庆渝北.就诊编号 & "' and 结算编号='" & str结算编号 & "'"

    '过程参数:
    '    病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 个人编号, 单位编号, 姓名, 性别, 出生日期, 实足年龄,
    '    累计缴费月数, 医疗人员类别, 医疗机构编码, 分支机构编码, 医疗机构类别, 特种病标志, 支付类别, 病种编码, 本次起付线,
    '    医疗费总额, 全自费总额, 挂钩自费总额, 符合范围总额, 个人帐户支付总额, 个人现金支付总额, 经办时间, 经办机构代码,
    '    医疗照顾类别 , 医疗补助类别, 就诊结算方式, 发票号, 备注, 分段计算情况, 医疗行政类别
    
    OpenRecordset_ZLYB rsTemp, "获取费用结算基本信息"
    
    With rsTemp
    
        strSQL = "ZL_费用基本信息_INSERT(" & Nvl(!病人ID, 0) & ","
        strSQL = strSQL & Nvl(!主页ID, 0) & ","
        strSQL = strSQL & "'" & Nvl(!就诊编号) & "',"
        strSQL = strSQL & "'" & g病人身份_重庆渝北.结算编号 & "',"
        strSQL = strSQL & "'" & Nvl(!结算编号) & "',"
        strSQL = strSQL & "" & Nvl(!审批记录序号, 0) & ","
        strSQL = strSQL & "" & Nvl(!个人编号, 0) & ","
        strSQL = strSQL & "" & Nvl(!单位编号, 0) & ","
        strSQL = strSQL & "'" & Nvl(!姓名) & "',"
        strSQL = strSQL & "'" & Nvl(!性别) & "',"
        If IsNull(!出生日期) Then
            strSQL = strSQL & "NULL,"
        Else
            strSQL = strSQL & "to_date('" & Format(!出生日期, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
        End If
        
        strSQL = strSQL & "" & Nvl(!实足年龄, 0) & ","
        strSQL = strSQL & "" & Nvl(!累计缴费月数, 0) & ","
        strSQL = strSQL & "'" & Nvl(!医疗人员类别) & "',"  '医疗人员类别
        strSQL = strSQL & "'" & Nvl(!医疗机构编码) & "',"
        strSQL = strSQL & "'" & Nvl(!分支机构编码) & "',"
        strSQL = strSQL & "'" & Nvl(!医疗机构类别) & "',"
        strSQL = strSQL & "'" & Nvl(!特种病标志) & "',"
        strSQL = strSQL & "'" & Nvl(!支付类别) & "',"
        strSQL = strSQL & "'" & Nvl(!病种编码) & "',"
        strSQL = strSQL & "" & Nvl(!本次起付线, 0) & ","
        strSQL = strSQL & "" & Nvl(!医疗费总额, 0) & ","

        
        strSQL = strSQL & "" & Nvl(!全自费总额, 0) & ","
        strSQL = strSQL & "" & Nvl(!挂钩自费总额, 0) & ","
        strSQL = strSQL & "" & Nvl(!符合范围总额, 0) & ","
        strSQL = strSQL & "" & Nvl(!个人帐户支付总额, 0) & ","
        strSQL = strSQL & "" & Nvl(!个人现金支付总额, 0) & ","
        If IsNull(!经办时间) Then
            strSQL = strSQL & "NULL,"
        Else
            strSQL = strSQL & "to_date('" & Format(!经办时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        End If
        
        strSQL = strSQL & "'" & Nvl(!经办机构代码) & "',"
        strSQL = strSQL & "'" & Nvl(!医疗照顾类别) & "',"
        strSQL = strSQL & "'" & Nvl(!医疗补助类别) & "',"
        strSQL = strSQL & "'" & Nvl(!就诊结算方式) & "',"
        strSQL = strSQL & "'" & Nvl(!发票号) & "',"
        strSQL = strSQL & "'" & Nvl(!备注) & "',"
        strSQL = strSQL & "'" & Nvl(!分段计算情况) & "',"
        strSQL = strSQL & "'" & Nvl(!医疗行政类别) & "')"
            
        '保存数据
        gcnOracle_CQYB.Execute strSQL, , adCmdStoredProc
        
        Call intXML
    
        'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
        AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Nvl(!经办机构代码)
        'SvrcID  string  2       远程数据服务标识，定值10, 标识大小写敏感，足两位长
        
        AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "10"
        'CtrInf  string  20      控制信息，预留, 标识大小写敏感
        AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
        
        'akc190  string  20      就诊编号
        AppendXMLNode gobjXMLInPut.documentElement, "akc190", Nvl(!就诊编号)
        'yka103  string  20      结算编号
        AppendXMLNode gobjXMLInPut.documentElement, "yka103", g病人身份_重庆渝北.结算编号
        'yka198  string  20      退单对应结算编号
        AppendXMLNode gobjXMLInPut.documentElement, "yka198", Nvl(!结算编号)
        
        'ykc114  number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
        AppendXMLNode gobjXMLInPut.documentElement, "ykc114", Nvl(!审批记录序号, 0)
        'aac001  number  15  0   个人编号
        AppendXMLNode gobjXMLInPut.documentElement, "aac001", Nvl(!个人编号, 0)
        'aab001  number  15  0   单位编号
        AppendXMLNode gobjXMLInPut.documentElement, "aab001", Nvl(!单位编号, 0)
        'aac003  string  20      姓名
        AppendXMLNode gobjXMLInPut.documentElement, "aac003", Nvl(!姓名)
        'aac004  string  1       性别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "aac004", Nvl(!性别)
        
        'aac006  date    日      出生日期
        AppendXMLNode gobjXMLInPut.documentElement, "aac006", Format(!出生日期, "yyyy-mm-dd")
        'akc023  number  3       实足年龄
        AppendXMLNode gobjXMLInPut.documentElement, "akc023", Nvl(!实足年龄, 0)
        'ykc021  number  3       累计缴费月数
        AppendXMLNode gobjXMLInPut.documentElement, "ykc021", Nvl(!累计缴费月数, 0)
        'akc021  string  6       医疗人员类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "akc021", Nvl(!医疗人员类别)
        'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
        AppendXMLNode gobjXMLInPut.documentElement, "akb020", Nvl(!医疗机构编码)
        'ykb006  string  3       定点医疗机构分支机构编号
        AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"          '分支机构编码
        'akb023  string  6       医疗机构类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "akb023", InitInfor_重庆渝北.机构类别
        
        'aka123  string  1       特种病标志，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "aka123", Nvl(!特种病标志, 0)      '特种病标志
        'aka130  string  6       支付类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "aka130", Nvl(!支付类别)
        'yka026  string  20      病种编码
        AppendXMLNode gobjXMLInPut.documentElement, "yka026", Nvl(!病种编码)
        
        '    '    病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 个人编号, 单位编号, 姓名, 性别, 出生日期, 实足年龄,
        '    累计缴费月数, 医疗人员类别, 医疗机构编码, 分支机构编码, 医疗机构类别, 特种病标志, 支付类别, 病种编码, 本次起付线,
        '    医疗费总额, 全自费总额, 挂钩自费总额, 符合范围总额, 个人帐户支付总额, 个人现金支付总额, 经办时间, 经办机构代码,
        '    医疗照顾类别 , 医疗补助类别, 就诊结算方式, 发票号, 备注, 分段计算情况, 医疗行政类别
        
        'yka115  number  14  2   本次起付线
        AppendXMLNode gobjXMLInPut.documentElement, "yka115", Nvl(!本次起付线, 0)           '本次起付线
        'yka055  number  14  2   医疗费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka055", Nvl(!医疗费总额, 0)
        'yka056  number  14  2   全自费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka056", Nvl(!全自费总额, 0)              '
        'yka057  number  14  2   挂钩自费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka057", Nvl(!挂钩自费总额, 0)               '
        'yka111  number  14  2   符合范围总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka111", Nvl(!符合范围总额, 0)                '
        'yka112  number  14  2   个人账户支付总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka112", Nvl(!个人帐户支付总额, 0)                 '
        'yka113  number  14  2   个人现金支付总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka113", Nvl(!个人现金支付总额, 0)                  '
        'aae036  date        秒  经办时间
        str经办时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        '经办时间
        AppendXMLNode gobjXMLInPut.documentElement, "aae036", str经办时间                 '
        'yab003  string  4       社保经办机构代码
        AppendXMLNode gobjXMLInPut.documentElement, "yab003", Nvl(!经办机构代码)               '
        'ykc120  string  6       医疗照顾类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "ykc120", Nvl(!医疗照顾类别)                  '
        'ykc121  string  6       享受医疗补助类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "ykc121", Nvl(!医疗补助类别)
        'yka222  string  6       就诊结算方式
        AppendXMLNode gobjXMLInPut.documentElement, "yka222", Nvl(!就诊结算方式) '
        'yka110  string  20      发票号
        AppendXMLNode gobjXMLInPut.documentElement, "yka110", Nvl(!发票号)                                '
        'aae013  string  100     备注
        AppendXMLNode gobjXMLInPut.documentElement, "aae013", Nvl(!备注)                              '
        'gkc010  string  800     分段计算情况(住院用)
        AppendXMLNode gobjXMLInPut.documentElement, "gkc010", Nvl(!分段计算情况)                              '
        'akc315  string  3       医疗待遇行政类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "akc315", Nvl(!医疗行政类别)                              '
    End With
    
    '写入基本信息
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    WriteDebugInfor_重庆渝北 strXMLText
    
    strXMLText = Replace(strXMLText, "&lt;", "<")
    strXMLText = Replace(strXMLText, "&gt;", ">")
    If 业务请求_重庆渝北(结算基本信息写入, strXMLText, strOutput) = False Then
        Exit Function
    End If
    费用基本信息冲销 = True
    Exit Function
errHand:
    DebugTool "费用基本信息冲销失败!" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & " 错误描述: " & Err.Description
End Function
Private Function 费用结算结果冲销上传(ByVal str结算编号 As String) As Boolean
    '根据结算编号,冲销历次的结算信息
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnFirst As Boolean
    Dim objXMLItem As MSXML2.IXMLDOMElement
    Dim strXMLText As String
    Dim strXMLtext1 As String
    Dim strOutput As String
    Dim str经办时间 As String
    Dim rs结算 As New ADODB.Recordset
    If g病人身份_重庆渝北.结算标志 = 0 Or g病人身份_重庆渝北.结算标志 = 2 Then
        gstrSQL = "Select * From 保险结算记录 where 性质=1 and  记录id=" & g病人身份_重庆渝北.结帐ID
    Else
        gstrSQL = "Select * From 保险结算记录 where 性质=2 and  记录id=" & g病人身份_重庆渝北.结帐ID
    End If
     
     '获取原结算记录中的数据,以便冲销
     zlDatabase.OpenRecordset rs结算, gstrSQL, "获取原结算记录中的数据"
     If rs结算.RecordCount = 0 Then
        ShowMsgbox "获取原结算记录时，发现无历史结算记录,不能保继续!"
        Exit Function
     End If
    
    gstrSQL = "" & _
            "   Select id,病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 医疗人员类别, 医疗照顾类别, 医疗补助类别," & _
            "       年度, 基本医疗信息, -1*下限金额 下限金额 , -1*自付金额 自付金额, 经办机构代码,支付金额 支付金额,-1* 公务员补助 公务员补助," & _
            "       -1*先行自付金额 先行自付金额, 累计缴费月数, 实足年龄," & _
            "       医疗行政类别, -1*帐户支付 帐户支付, 分段标准, -1*全自费金额 全自费金额, -1*挂构自费 挂构自费, -1*符合金额 符合金额, -1*本次自付 本次自付," & _
            "       -1*本次支付金额 本次支付金额, -1*公务员统筹支付 公务员统筹支付,-1*补助自付累计  补助自付累计 " & _
            "   From 费用结算结果 " & _
            "   where 结算编号='" & str结算编号 & "' and 就诊编号='" & g病人身份_重庆渝北.就诊编号 & "'"
                
    Err = 0
    On Error GoTo errHand:
    OpenRecordset_ZLYB rsTemp, "冲销结算结果"
    
    费用结算结果冲销上传 = False
    If rsTemp.EOF Then
        ShowMsgbox "结算编号为:" & str结算编号 & " 的结算号不存在,冲销失败!"
        Exit Function
    End If
 
    '存储过程参数:
    '病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 经办机构代码, 医疗人员类别, 医疗照顾类别, 医疗补助类别,
    '年度,基本医疗信息, 下限金额, 自付金额, 支付金额, 公务员补助, 先行自付金额, 累计缴费月数, 实足年龄, 医疗行政类别, 帐户支付, 分段标准,
    '全自费金额, 挂构自费,符合金额_IN, 本次自付, 本次支付金额, 公务员统筹支付, 补肋自付累计
    
    Call intXML
    Dim lngID As Long
    blnFirst = True
    With rsTemp
        Do While Not .EOF
            strSQL = "ZL_费用结算结果_INSERT("
            
            gstrSQL = "Select 费用结算结果_ID.nextval as ID from dual"
            OpenRecordset_ZLYB rsTmp, "获取结算编号"
            lngID = Nvl(rsTmp!ID, 0)
                
            strSQL = strSQL & lngID & ","
            strSQL = strSQL & Nvl(!病人ID, 0) & ","
            strSQL = strSQL & Nvl(!主页ID, 0) & ","
            strSQL = strSQL & "'" & Nvl(!就诊编号) & "',"
            strSQL = strSQL & "'" & g病人身份_重庆渝北.结算编号 & "',"
            strSQL = strSQL & "'" & Nvl(!结算编号) & "',"
            strSQL = strSQL & "" & Nvl(!审批记录序号, 0) & ","
            strSQL = strSQL & "'" & Nvl(!经办机构代码) & "',"
            strSQL = strSQL & "'" & Nvl(!医疗人员类别) & "',"
            strSQL = strSQL & "'" & Nvl(!医疗照顾类别) & "',"
            strSQL = strSQL & "'" & Nvl(!医疗补助类别) & "',"  '医疗补助类别
            strSQL = strSQL & "'" & Nvl(!年度) & "',"
            strSQL = strSQL & "'" & Nvl(!基本医疗信息) & "',"
            strSQL = strSQL & "" & Nvl(!下限金额, 0) & ","
            
            
            strSQL = strSQL & "" & Nvl(!自付金额, 0) & ","
            strSQL = strSQL & "" & Nvl(!支付金额, 0) & ","
            strSQL = strSQL & "" & Nvl(!公务员补助, 0) & ","
            strSQL = strSQL & "" & Nvl(!先行自付金额, 0) & ","
            strSQL = strSQL & "" & Nvl(!累计缴费月数, 0) & ","
            strSQL = strSQL & "" & Nvl(!实足年龄, 0) & ","
            
            strSQL = strSQL & "'" & Nvl(!医疗行政类别) & "',"
            strSQL = strSQL & "" & Nvl(!帐户支付, 0) & ","
            strSQL = strSQL & "'" & Nvl(!分段标准) & "',"
            
            
            
            strSQL = strSQL & "" & Nvl(!全自费金额, 0) & ","
            strSQL = strSQL & "" & Nvl(!挂构自费, 0) & ","
            strSQL = strSQL & "" & Nvl(!符合金额, 0) & ","
            strSQL = strSQL & "" & Nvl(!本次自付, 0) & ","
            strSQL = strSQL & "" & Nvl(!本次支付金额, 0) & ","
            
            strSQL = strSQL & "" & Nvl(!公务员统筹支付, 0) & ","
            strSQL = strSQL & "" & Nvl(!补助自付累计, 0) & ")"
            
                    
            '存入数据库中
            gcnOracle_CQYB.Execute strSQL, , adCmdStoredProc
            
            If insertInto子项(lngID, Nvl(!基本医疗信息)) = False Then
                DebugTool "插入子项错误!"
            End If
            
            'XML费用结果写入
            If blnFirst Then
                AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Nvl(!经办机构代码)
                AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "12"
                AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
        
                'BaseInfo                多个待遇享受段共有的相同的基本信息部分，下面缩排的元素是它的子元素
                Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "BaseInfo", "")
                '    akc190  string  20      就诊编号
                AppendXMLNode objXMLItem, "akc190", Nvl(!就诊编号)
                '    yka103  string  20      结算编号
                AppendXMLNode objXMLItem, "yka103", g病人身份_重庆渝北.结算编号
                '    yka198  string  20      退单对应结算编号
                AppendXMLNode objXMLItem, "yka198", Nvl(!结算编号)
                '    ykc114  number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
                AppendXMLNode objXMLItem, "ykc114", Nvl(!审批记录序号, 0)
                '    yab003  string  4       社保经办机构代码
                AppendXMLNode objXMLItem, "yab003", Nvl(!经办机构代码)
                blnFirst = False
            End If
            
            '需确定相关的字串
            'ReckonInfo              多个待遇享受段的结算分段信息，下面缩排的元素是它的子元素
            Set objXMLItem = AppendXMLNode(gobjXMLInPut.documentElement, "ReckonInfo", "")
            
            'akc190  string  20      就诊编号
             AppendXMLNode objXMLItem, "akc190", Nvl(!就诊编号)
            'yka103  string  20      结算编号
             AppendXMLNode objXMLItem, "yka103", g病人身份_重庆渝北.结算编号
             
            'yka198  string  20      退单对应结算编号
             AppendXMLNode objXMLItem, "yka198", Nvl(!结算编号)
            'ykc114  number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
             AppendXMLNode objXMLItem, "ykc114", Nvl(!审批记录序号)
            'yab003  string  4       社保经办机构代码
             AppendXMLNode objXMLItem, "yab003", Nvl(!经办机构代码)
            'aka213  string  2       分段标准，03 起付线， 05 基本医疗 ，06 大额医疗，07 超限
             AppendXMLNode objXMLItem, "aka213", Nvl(!分段标准)
            'yka056  number  14  2   全自费金额
             AppendXMLNode objXMLItem, "yka056", Nvl(!全自费金额, 0)
            'yka057  number  14  2   挂钩自费金额
             AppendXMLNode objXMLItem, "yka057", Nvl(!挂构自费, 0)
            'yka111  number  14  2   符合范围金额
             AppendXMLNode objXMLItem, "yka111", Nvl(!符合金额, 0)
            'yka106  number  14  2   自付金额
             AppendXMLNode objXMLItem, "yka106", Nvl(!本次自付, 0)
            'yka107  number  14  2   支付金额
             AppendXMLNode objXMLItem, "yka107", Nvl(!本次支付金额, 0)
            'yka063  number  14  2   公务员统筹支付金额
             AppendXMLNode objXMLItem, "yka063", Nvl(!公务员统筹支付, 0)
            'yka221  number  14  2   享受医疗补助个人自付累计金额
             AppendXMLNode objXMLItem, "yka221", Nvl(!补助自付累计, 0)
            'Akc315  String  3       医疗行政职务
             AppendXMLNode objXMLItem, "Akc315", Nvl(!医疗行政类别)
            .MoveNext
        Loop
    End With
      
    

    '写入费用结算结果
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    strXMLtext1 = strXMLText
      
    '写入费用结算基本信息
    '费用基本信息冲销
    gstrSQL = " " & _
    "           Select 病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 个人编号, 单位编号, 姓名, 性别, 出生日期, 实足年龄, " & _
    "                   累计缴费月数, 医疗人员类别, 医疗机构编码, 分支机构编码, 医疗机构类别, 特种病标志, 支付类别, 病种编码, 本次起付线, " & _
    "                   -1*医疗费总额 医疗费总额, -1*全自费总额 全自费总额, -1*挂钩自费总额 挂钩自费总额, -1*符合范围总额 符合范围总额, -1*个人帐户支付总额 个人帐户支付总额," & _
    "                  -1*个人现金支付总额 个人现金支付总额, 经办时间, 经办机构代码, " & _
    "                   医疗照顾类别 , 医疗补助类别, 就诊结算方式, 发票号, 备注, 分段计算情况, 医疗行政类别 " & _
    "           From 费用基本信息 " & _
    "           where 就诊编号='" & g病人身份_重庆渝北.就诊编号 & "' and 结算编号='" & str结算编号 & "'"

    '过程参数:
    '    病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 个人编号, 单位编号, 姓名, 性别, 出生日期, 实足年龄,
    '    累计缴费月数, 医疗人员类别, 医疗机构编码, 分支机构编码, 医疗机构类别, 特种病标志, 支付类别, 病种编码, 本次起付线,
    '    医疗费总额, 全自费总额, 挂钩自费总额, 符合范围总额, 个人帐户支付总额, 个人现金支付总额, 经办时间, 经办机构代码,
    '    医疗照顾类别 , 医疗补助类别, 就诊结算方式, 发票号, 备注, 分段计算情况, 医疗行政类别
    OpenRecordset_ZLYB rsTemp, "获取费用结算基本信息"
    With rsTemp
    
        strSQL = "ZL_费用基本信息_INSERT(" & Nvl(!病人ID, 0) & ","
        strSQL = strSQL & Nvl(!主页ID, 0) & ","
        strSQL = strSQL & "'" & Nvl(!就诊编号) & "',"
        strSQL = strSQL & "'" & g病人身份_重庆渝北.结算编号 & "',"
        strSQL = strSQL & "'" & Nvl(!结算编号) & "',"
        strSQL = strSQL & "" & Nvl(!审批记录序号, 0) & ","
        strSQL = strSQL & "" & Nvl(!个人编号, 0) & ","
        strSQL = strSQL & "" & Nvl(!单位编号, 0) & ","
        strSQL = strSQL & "'" & Nvl(!姓名) & "',"
        strSQL = strSQL & "'" & Nvl(!性别) & "',"
        If IsNull(!出生日期) Then
            strSQL = strSQL & "NULL,"
        Else
            strSQL = strSQL & "to_date('" & Format(!出生日期, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
        End If
        
        strSQL = strSQL & "" & Nvl(!实足年龄, 0) & ","
        strSQL = strSQL & "" & Nvl(!累计缴费月数, 0) & ","
        strSQL = strSQL & "'" & Nvl(!医疗人员类别) & "',"  '医疗人员类别
        strSQL = strSQL & "'" & Nvl(!医疗机构编码) & "',"
        strSQL = strSQL & "'" & Nvl(!分支机构编码) & "',"
        strSQL = strSQL & "'" & Nvl(!医疗机构类别) & "',"
        strSQL = strSQL & "'" & Nvl(!特种病标志) & "',"
        strSQL = strSQL & "'" & Nvl(!支付类别) & "',"
        strSQL = strSQL & "'" & Nvl(!病种编码) & "',"
        strSQL = strSQL & "" & Nvl(!本次起付线, 0) & ","
        strSQL = strSQL & "" & Nvl(!医疗费总额, 0) & ","

        
        strSQL = strSQL & "" & Nvl(!全自费总额, 0) & ","
        strSQL = strSQL & "" & Nvl(!挂钩自费总额, 0) & ","
        strSQL = strSQL & "" & Nvl(!符合范围总额, 0) & ","
        strSQL = strSQL & "" & Nvl(!个人帐户支付总额, 0) & ","
        strSQL = strSQL & "" & Nvl(!个人现金支付总额, 0) & ","
        If IsNull(!经办时间) Then
            strSQL = strSQL & "NULL,"
        Else
            strSQL = strSQL & "to_date('" & Format(!经办时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        End If
        
        strSQL = strSQL & "'" & Nvl(!经办机构代码) & "',"
        strSQL = strSQL & "'" & Nvl(!医疗照顾类别) & "',"
        strSQL = strSQL & "'" & Nvl(!医疗补助类别) & "',"
        strSQL = strSQL & "'" & Nvl(!就诊结算方式) & "',"
        strSQL = strSQL & "'" & Nvl(!发票号) & "',"
        strSQL = strSQL & "'" & Nvl(!备注) & "',"
        strSQL = strSQL & "'" & Nvl(!分段计算情况) & "',"
        strSQL = strSQL & "'" & Nvl(!医疗行政类别) & "')"
            
        '保存数据
        gcnOracle_CQYB.Execute strSQL, , adCmdStoredProc
        
        Call intXML
        
    
        'YAB003  string  4       在定点医疗机构就诊的参保人员所在的社保经办机构代码，足四位长
        AppendXMLNode gobjXMLInPut.documentElement, "YAB003", Nvl(!经办机构代码)
        'SvrcID  string  2       远程数据服务标识，定值10, 标识大小写敏感，足两位长
        
        AppendXMLNode gobjXMLInPut.documentElement, "SvrcID", "10"
        'CtrInf  string  20      控制信息，预留, 标识大小写敏感
        AppendXMLNode gobjXMLInPut.documentElement, "CtrInf", ""
        
        'akc190  string  20      就诊编号
        AppendXMLNode gobjXMLInPut.documentElement, "akc190", Nvl(!就诊编号)
        'yka103  string  20      结算编号
        AppendXMLNode gobjXMLInPut.documentElement, "yka103", g病人身份_重庆渝北.结算编号
        'yka198  string  20      退单对应结算编号
        AppendXMLNode gobjXMLInPut.documentElement, "yka198", Nvl(!结算编号)
        
        'ykc114  number  15  0   审批记录序号，表示在同一审批编号下的多条审批信息
        AppendXMLNode gobjXMLInPut.documentElement, "ykc114", Nvl(!审批记录序号, 0)
        'aac001  number  15  0   个人编号
        AppendXMLNode gobjXMLInPut.documentElement, "aac001", Nvl(!个人编号, 0)
        'aab001  number  15  0   单位编号
        AppendXMLNode gobjXMLInPut.documentElement, "aab001", Nvl(!单位编号, 0)
        'aac003  string  20      姓名
        AppendXMLNode gobjXMLInPut.documentElement, "aac003", Nvl(!姓名)
        'aac004  string  1       性别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "aac004", Nvl(!性别)
        
        'aac006  date    日      出生日期
        AppendXMLNode gobjXMLInPut.documentElement, "aac006", Format(!出生日期, "yyyy-mm-dd")
        'akc023  number  3       实足年龄
        AppendXMLNode gobjXMLInPut.documentElement, "akc023", Nvl(!实足年龄, 0)
        'ykc021  number  3       累计缴费月数
        AppendXMLNode gobjXMLInPut.documentElement, "ykc021", Nvl(!累计缴费月数, 0)
        'akc021  string  6       医疗人员类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "akc021", Nvl(!医疗人员类别)
        'akb020  string  8       定点医疗机构在就诊参保人员所在的医保机构中的编号
        AppendXMLNode gobjXMLInPut.documentElement, "akb020", Nvl(!医疗机构编码)
        'ykb006  string  3       定点医疗机构分支机构编号
        AppendXMLNode gobjXMLInPut.documentElement, "ykb006", "01"          '分支机构编码
        'akb023  string  6       医疗机构类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "akb023", InitInfor_重庆渝北.机构类别
        
        'aka123  string  1       特种病标志，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "aka123", Nvl(!特种病标志, 0)      '特种病标志
        'aka130  string  6       支付类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "aka130", Nvl(!支付类别)
        'yka026  string  20      病种编码
        AppendXMLNode gobjXMLInPut.documentElement, "yka026", Nvl(!病种编码)
        
        '    '    病人id, 主页id, 就诊编号, 结算编号, 退单结算号, 审批记录序号, 个人编号, 单位编号, 姓名, 性别, 出生日期, 实足年龄,
        '    累计缴费月数, 医疗人员类别, 医疗机构编码, 分支机构编码, 医疗机构类别, 特种病标志, 支付类别, 病种编码, 本次起付线,
        '    医疗费总额, 全自费总额, 挂钩自费总额, 符合范围总额, 个人帐户支付总额, 个人现金支付总额, 经办时间, 经办机构代码,
        '    医疗照顾类别 , 医疗补助类别, 就诊结算方式, 发票号, 备注, 分段计算情况, 医疗行政类别
        
        'yka115  number  14  2   本次起付线
        AppendXMLNode gobjXMLInPut.documentElement, "yka115", Nvl(!本次起付线, 0)           '本次起付线
        'yka055  number  14  2   医疗费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka055", Nvl(!医疗费总额, 0)
        'yka056  number  14  2   全自费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka056", Nvl(!全自费总额, 0)              '
        'yka057  number  14  2   挂钩自费总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka057", Nvl(!挂钩自费总额, 0)               '
        'yka111  number  14  2   符合范围总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka111", Nvl(!符合范围总额, 0)                '
        'yka112  number  14  2   个人账户支付总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka112", Nvl(!个人帐户支付总额, 0)                 '
        'yka113  number  14  2   个人现金支付总额
        AppendXMLNode gobjXMLInPut.documentElement, "yka113", Nvl(!个人现金支付总额, 0)                  '
        'aae036  date        秒  经办时间
        str经办时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        '经办时间
        AppendXMLNode gobjXMLInPut.documentElement, "aae036", str经办时间                 '
        'yab003  string  4       社保经办机构代码
        AppendXMLNode gobjXMLInPut.documentElement, "yab003", Nvl(!经办机构代码)               '
        'ykc120  string  6       医疗照顾类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "ykc120", Nvl(!医疗照顾类别)                  '
        'ykc121  string  6       享受医疗补助类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "ykc121", Nvl(!医疗补助类别)
        'yka222  string  6       就诊结算方式
        AppendXMLNode gobjXMLInPut.documentElement, "yka222", Nvl(!就诊结算方式) '
        'yka110  string  20      发票号
        AppendXMLNode gobjXMLInPut.documentElement, "yka110", Nvl(!发票号)                                '
        'aae013  string  100     备注
        AppendXMLNode gobjXMLInPut.documentElement, "aae013", Nvl(!备注)                              '
        'gkc010  string  800     分段计算情况(住院用)
        AppendXMLNode gobjXMLInPut.documentElement, "gkc010", Nvl(!分段计算情况)                              '
        'akc315  string  3       医疗待遇行政类别，见代码表
        AppendXMLNode gobjXMLInPut.documentElement, "akc315", Nvl(!医疗行政类别)                              '
            
    End With
    '写入基本信息
    strXMLText = 取掉XML的前导标识(gobjXMLInPut.xml)
    strXMLText = Replace(strXMLText, "&lt;", "<")
    strXMLText = Replace(strXMLText, "&gt;", ">")
    WriteDebugInfor_重庆渝北 strXMLText
    
    If 业务请求_重庆渝北(结算基本信息写入, strXMLText, strOutput) = False Then
        Exit Function
    End If
    
    WriteDebugInfor_重庆渝北 strXMLtext1
    
    '保存费用结算结果
    If 业务请求_重庆渝北(结算结果写入, strXMLtext1, strOutput) = False Then
        Exit Function
    End If
    
    '回退个人帐户额
    If IC卡帐户支付_重庆渝北(rsTemp!个人帐户支付总额, str经办时间, Nvl(rsTemp!结算编号)) = False Then
            Exit Function
    End If
   
   '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
  '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(公务员补助),帐户累计支出_IN(大额支付),累计进入统筹_IN(基本医疗自付),累计统筹报销_IN,住院次数_IN,起付线(进入起付线),封顶线_IN(支付类别+10000),实际起付线_IN,
    '   发生费用金额_IN(发生费用),全自付金额_IN(全自付),首先自付金额_IN(首先自付),
    '   进入统筹金额_IN(符合金额),统筹报销金额_IN(基本医疗统筹支付),大病自付金额_IN(大额自付),超限自付金额_IN(超限自付),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算编号),主页ID_IN,中途结帐_IN,备注_IN(就诊编号)
    
    gstrSQL = "zl_保险结算记录_insert(" & IIf(g病人身份_重庆渝北.结算标志 = 1, 2, 1) & "," & g病人身份_重庆渝北.冲销ID & "," & TYPE_重庆渝北 & "," & g病人身份_重庆渝北.lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
      -1 * Nvl(rs结算!帐户累计增加, 0) & "," & -1 * Nvl(rs结算!帐户累计支出, 0) & "," & -1 * Nvl(rs结算!累计进入统筹, 0) & ",NULL,NULL," & -1 * Nvl(rs结算!起付线, 0) & "," & Nvl(rs结算!封顶线, 0) & ",NULL," & _
       -1 * Nvl(rs结算!发生费用金额, 0) & "," & -1 * Nvl(rs结算!全自付金额, 0) & "," & -1 * Nvl(rs结算!首先自付金额, 0) & "," & _
        "" & -1 * Nvl(rs结算!进入统筹金额, 0) & "," & -1 * Nvl(rs结算!统筹报销金额, 0) & "," & -1 * Nvl(rs结算!大病自付金额, 0) & "," & -1 * Nvl(rs结算!超限自付金额, 0) & "," & -1 * Nvl(rs结算!个人帐户支付, 0) & ",'" & _
       g病人身份_重庆渝北.结算编号 & "'," & IIf(Nvl(rs结算!主页ID, 0) = 0, "NULL", Nvl(rs结算!主页ID, 0)) & "," & IIf(g病人身份_重庆渝北.中途结帐 = 1, "1", "NULL") & ",'" & _
       g病人身份_重庆渝北.就诊编号 & "')"
       
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
    费用结算结果冲销上传 = True
    Exit Function
errHand:
    
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function 病人结算(ByVal lng结帐ID As Long) As Boolean
    '病人费用结算
    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    
    Dim strCurrDate As String
    Dim str开始时间 As String
    Dim lng病人ID  As Long
    Err = 0
    
    On Error GoTo errHand:
    
    '第一步:需确定资格审批待遇核定
    
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59"
    
    If InitInfor_重庆渝北.模拟数据 Then
        str开始时间 = "2004-07-10 21:40:29"
        strCurrDate = "2004-07-10 21:40:29"
    Else
        str开始时间 = g病人身份_重庆渝北.就诊时间
        If g病人身份_重庆渝北.结算标志 = 1 Then
            '住院的话,其开始时间应该从00:00:00秒开始算.
            str开始时间 = Format(str开始时间, "yyyy-mm-dd" & " 00:00:00")
        End If
    End If
    lng病人ID = g病人身份_重庆渝北.lng病人ID
    
    If g病人身份_重庆渝北.虚拟结算 Then
          '需审批记录作废
        Call 审批记录作废_重庆渝北
    End If
    
    WriteDebugDate_重庆渝北 "获取资格审批信息开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If g病人身份_重庆渝北.结算标志 = 1 And g病人身份_重庆渝北.虚拟结算 = False Then
        '住院结算不能进行资格审核核写,在虚拟结算即可完成.
    Else
        If 资格审核待遇核定(lng病人ID, str开始时间, strCurrDate) = False Then
            '有可能该审批记录已经存在,所以作废一下,再进行核定.
            Call 审批记录作废_重庆渝北
            If 资格审核待遇核定(lng病人ID, str开始时间, strCurrDate) = False Then
                Exit Function
            End If
        End If
        
        WriteDebugDate_重庆渝北 "保存审批信息开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
        '第二步:写入资格审批待遇,并产生本产待遇文件
        If Save审批信息(lng病人ID) = False Then
            Exit Function
        End If
   End If
   WriteDebugDate_重庆渝北 "获取资格审批信息结束时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    
    '第三步:写入本次需明细结算的文本
    '   读取费用明细记录
    WriteDebugDate_重庆渝北 "获取明细信息开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If g病人身份_重庆渝北.结算标志 = 1 Then
        '由于是住院,需确定本次就诊的所有明细数据
        gstrSQL = Get明细记录(0)
    Else
        '只是本次结算的明细数据
        gstrSQL = Get明细记录(lng结帐ID)
    End If
    
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取处方明细")
    
    If rs明细.RecordCount = 0 Then
        ShowMsgbox "没有明细记录，可能相关项目未进行相应的对码"
        '需审批记录作废
        GoTo CancelRecordVerify:
        Exit Function
    End If
    WriteDebugDate_重庆渝北 "获取明细信息结束时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    
    WriteDebugDate_重庆渝北 "保存明细数据开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If g病人身份_重庆渝北.结算标志 = 1 Then
        '不保存相关明细记录
    Else
        If Save医保明细数据(rs明细) = False Then

            '需审批记录作废
            GoTo CancelRecordVerify:
            Exit Function
        End If
    End If
    WriteDebugDate_重庆渝北 "保存明细数据结束时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    
    WriteDebugDate_重庆渝北 "保存明细数据文本文件开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If Save费用明细文本文件(rs明细) = False Then
        '需审批记录作废
        GoTo CancelRecordVerify:
        Exit Function
    End If
    WriteDebugDate_重庆渝北 "保存明细数据文本文件结束时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")

    WriteDebugDate_重庆渝北 "保存历史费用结算结果开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    '第四步:保存历史的费用结算结果
    If g病人身份_重庆渝北.结算标志 = 1 Then
        rs明细.MoveFirst
        g病人身份_重庆渝北.lng主页ID = Nvl(rs明细!主页ID, 0)
        If g病人身份_重庆渝北.虚拟结算 Then
            If Save历史费用结算结果文本(g病人身份_重庆渝北.lng病人ID, Nvl(rs明细!主页ID, 0), False) = False Then
                    '需审批记录作废
                    GoTo CancelRecordVerify:
                    Exit Function
            End If
        End If
    Else
        If Save历史费用结算结果文本(0, 0) = False Then
                '需审批记录作废
                GoTo CancelRecordVerify:
                Exit Function
        End If
    End If
    WriteDebugDate_重庆渝北 "保存历史费用结算结果结束时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
        
    '第五步:需进行本地计算,即费用结算
    WriteDebugDate_重庆渝北 "本地费用结算开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If 病人费用结算(lng病人ID, 0) = False Then

        '需审批记录作废
        GoTo CancelRecordVerify:
        Exit Function
    End If
    WriteDebugDate_重庆渝北 "本地费用结算结束时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    
    '第六步:处方明细上传
    WriteDebugDate_重庆渝北 "处方上传开始时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    If g病人身份_重庆渝北.虚拟结算 Then
        '虚拟结算用不着上传明细
    Else
        If 处方明细上传(rs明细) = False Then
            ShowMsgbox "在进行处方明细上传时存在一条以上的明细上传失败,请以后注意补传!"
        End If
    End If
    病人结算 = True
    WriteDebugDate_重庆渝北 "处方上传结束时间:" & Format(Now, "yyyy-mm-dd HH:MM:SS")
    GoTo CancelRecordVerify:
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
CancelRecordVerify:
    If g病人身份_重庆渝北.结算标志 <> 1 Then
    
        '需审批记录作废
        Call 审批记录作废_重庆渝北
    Else
        If g病人身份_重庆渝北.虚拟结算 = False And 医保病人已经出院(lng病人ID) = True And 病人结算 = True Then
            '需审批记录作废
            Call 审批记录作废_重庆渝北
        End If
    End If
End Function
Public Sub WriteDebugInfor_重庆渝北(ByVal strInfor As String)
        '将调试信息写入文件中
        Dim objFile As New FileSystemObject
        Dim objText As TextStream
        Dim intDebug As Integer
        
        intDebug = GetSetting("ZLSOFT", "医保", "将串写入文本文件", 0)
        If intDebug <> 1 Then Exit Sub

        Dim strFile As String
        Dim rsTemp As New ADODB.Recordset
        strFile = App.Path & "\Test.log"
        
        If Not Dir(strFile) <> "" Then
            objFile.CreateTextFile strFile
        End If
        Set objText = objFile.OpenTextFile(strFile, ForAppending)
        objText.WriteLine strInfor
        objText.Close
        
End Sub

Public Sub WriteDebugDate_重庆渝北(ByVal strInfor As String)
        '将调试信息写入文件中
        Dim objFile As New FileSystemObject
        Dim objText As TextStream
        Dim intDebug As Integer
        
        intDebug = GetSetting("ZLSOFT", "医保", "测试时间", 0)
        If intDebug <> 1 Then Exit Sub

        Dim strFile As String
        Dim rsTemp As New ADODB.Recordset
        strFile = App.Path & "\Test.log"
        
        If Not Dir(strFile) <> "" Then
            objFile.CreateTextFile strFile
        End If
        Set objText = objFile.OpenTextFile(strFile, ForAppending)
        If InStr(1, strInfor, "==") <> 0 Then
            objText.WriteLine strInfor
        Else
            objText.WriteLine "姓名:" & g病人身份_重庆渝北.姓名 & vbTab & strInfor
        End If
        objText.Close
        
End Sub

Private Function insertInto子项(ByVal lng结果id As Long, ByVal XMLTEXT As String) As Boolean
    '功能:
    '过程参数:
    '   ZL_费用结算子项_INSERT
    '结果ID_IN       IN 费用结算子项.结果ID%TYPE,
    '序号_IN
    '下限金额_IN     IN 费用结算子项.下限金额%TYPE,
    '自付金额_IN     IN 费用结算子项.自付金额%TYPE,
    '支付金额_IN     IN 费用结算子项.支付金额%TYPE,
    '公务员补助_IN   IN 费用结算子项.公务员补助%TYPE,
    '先自付_IN       IN 费用结算子项.先自付%TYPE
    
    DebugTool "进入费用结算子项插入:XMLTEXT:" & XMLTEXT
    
    If Trim(XMLTEXT) = "" Then insertInto子项 = True: Exit Function
    
    insertInto子项 = False
    Set gobjXMLOutput = New MSXML2.DOMDocument

    If GetXML串(XMLTEXT) = False Then Exit Function
    
    DebugTool "GETXML串成功:XMLTEXT:" & XMLTEXT
    
    Dim lngCount As Long
    Dim lngRow As Long
    
    lngCount = GetOutXMLRows("SubRkn")
    
    
    Err = 0
    On Error GoTo errHand:
    For lngRow = 0 To lngCount - 1
        gstrSQL = "ZL_费用结算子项_INSERT("
        gstrSQL = gstrSQL & "" & lng结果id & ","
        gstrSQL = gstrSQL & "" & lngRow & ","
        
        'AKA160  number  14  2   子段下限金额
        gstrSQL = gstrSQL & Val(GetXMLOutput("aka160", , lngRow)) & ","
        'YKA106  number  14  2   自付金额
        gstrSQL = gstrSQL & Val(GetXMLOutput("yka106", , lngRow)) & ","
        'YKA 107 number  14  2   支付金额
        gstrSQL = gstrSQL & Val(GetXMLOutput("yka107", , lngRow)) & ","
        'YKA 063 number  14  2   公务员补助金额
        gstrSQL = gstrSQL & Val(GetXMLOutput("yka063", , lngRow)) & ","
        'YKA057  number  14  2   先行自付部分
        gstrSQL = gstrSQL & Val(GetXMLOutput("yka057", , lngRow)) & ")"
        DebugTool "gstrSQL=:" & gstrSQL
        gcnOracle_CQYB.Execute gstrSQL
    Next

    insertInto子项 = True
    Exit Function
errHand:
    DebugTool "执入子项错误" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & "错误信息:" & Err.Description
End Function




Public Function Save病情信息(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
        ByVal int性质 As Integer) As Boolean
    '功能:保存病情信息
    '参数:  lng病人id-病人id
    '       主页id:门诊可以无主页id
    '       int性质-0-门诊,1-入院,2-出院
    
    Err = 0: On Error GoTo errHand:
    DebugTool "进入保存病情!病人id:" & lng病人ID & " 主页id: " & lng主页ID & " 性质:" & int性质

    '保存相关病情:
    'ZL_病人诊断情况_91_UPDATE(
    gstrSQL = "ZL_病人诊断情况_91_UPDATE("
    '    病人ID_IN IN 病人诊断情况_425.病人ID%TYPE,
    gstrSQL = gstrSQL & "" & lng病人ID & ","
    '    主页ID_IN IN 病人诊断情况_425.主页ID%TYPE,
    gstrSQL = gstrSQL & "" & IIf(lng主页ID = 0, "NULL", lng主页ID) & ","
    '    性质_IN IN 病人诊断情况_425.性质%TYPE,
    gstrSQL = gstrSQL & "" & int性质 & ","
    '    序号_IN IN 病人诊断情况_425.序号%TYPE,
    gstrSQL = gstrSQL & "" & 1 & ","
    '    病情ID_IN IN 病人诊断情况_425.病情ID%TYPE,
    gstrSQL = gstrSQL & "" & IIf(g病人身份_重庆渝北.病情ID = 0, "NULL", g病人身份_重庆渝北.病情ID) & ","
    '   病情编码_IN IN 病人诊断情况_425.病情编码%TYPE,
    gstrSQL = gstrSQL & "" & IIf(g病人身份_重庆渝北.病情编码 = "", "NULL", "'" & g病人身份_重庆渝北.病情编码 & "'") & ","
    '    病情_IN IN 病人诊断情况_425.病情%TYPE
    gstrSQL = gstrSQL & "" & IIf(g病人身份_重庆渝北.病情名称 = "", "NULL", "'" & g病人身份_重庆渝北.病情名称 & "'") & ")"
    ExecuteProcedure_CQYB "保存病情1"
        
    'ZL_病人诊断情况_91_UPDATE(
    gstrSQL = "ZL_病人诊断情况_91_UPDATE("
    '    病人ID_IN IN 病人诊断情况_425.病人ID%TYPE,
    gstrSQL = gstrSQL & "" & lng病人ID & ","
    '    主页ID_IN IN 病人诊断情况_425.主页ID%TYPE,
    gstrSQL = gstrSQL & "" & IIf(lng主页ID = 0, "NULL", lng主页ID) & ","
    '    性质_IN IN 病人诊断情况_425.性质%TYPE,
    gstrSQL = gstrSQL & "" & int性质 & ","
    '    序号_IN IN 病人诊断情况_425.序号%TYPE,
    gstrSQL = gstrSQL & "" & 2 & ","
    '    病情ID_IN IN 病人诊断情况_425.病情ID%TYPE,
    gstrSQL = gstrSQL & "" & IIf(g病人身份_重庆渝北.病情1ID = 0, "NULL", g病人身份_重庆渝北.病情1ID) & ","
    '   病情编码_IN IN 病人诊断情况_425.病情编码%TYPE,
    gstrSQL = gstrSQL & "" & IIf(g病人身份_重庆渝北.病情编码 = "", "NULL", "'" & g病人身份_重庆渝北.病情编码 & "'") & ","
    '    病情_IN IN 病人诊断情况_425.病情%TYPE
    gstrSQL = gstrSQL & "" & IIf(g病人身份_重庆渝北.病情名称1 = "", "NULL", "'" & g病人身份_重庆渝北.病情名称1 & "'") & ")"
    ExecuteProcedure_CQYB "保存病情2"
        
    'ZL_病人诊断情况_91_UPDATE(
    gstrSQL = "ZL_病人诊断情况_91_UPDATE("
    '    病人ID_IN IN 病人诊断情况_425.病人ID%TYPE,
    gstrSQL = gstrSQL & "" & lng病人ID & ","
    '    主页ID_IN IN 病人诊断情况_425.主页ID%TYPE,
    gstrSQL = gstrSQL & "" & IIf(lng主页ID = 0, "NULL", lng主页ID) & ","
    '    性质_IN IN 病人诊断情况_425.性质%TYPE,
    gstrSQL = gstrSQL & "" & int性质 & ","
    '    序号_IN IN 病人诊断情况_425.序号%TYPE,
    gstrSQL = gstrSQL & "" & 3 & ","
    '    病情ID_IN IN 病人诊断情况_425.病情ID%TYPE,
    gstrSQL = gstrSQL & "" & IIf(g病人身份_重庆渝北.病情2ID = 0, "NULL", g病人身份_重庆渝北.病情2ID) & ","
    '   病情编码_IN IN 病人诊断情况_425.病情编码%TYPE,
    gstrSQL = gstrSQL & "" & IIf(g病人身份_重庆渝北.病情编码2 = "", "NULL", "'" & g病人身份_重庆渝北.病情编码2 & "'") & ","
    '    病情_IN IN 病人诊断情况_425.病情%TYPE
    gstrSQL = gstrSQL & "" & IIf(g病人身份_重庆渝北.病情名称2 = "", "NULL", "'" & g病人身份_重庆渝北.病情名称2 & "'") & ")"
    ExecuteProcedure_CQYB "保存病情3"
    
    DebugTool "保存病情成功!"
    Save病情信息 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    DebugTool "保存病情失败!"
End Function

Private Sub ExecuteProcedure_CQYB(ByVal strCaption As String)
'功能：执行SQL语句
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_CQYB.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Private Function 补传住院明细记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '补传相关明细记录
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput  As String, strOutput As String
    Dim strArr, strArr摘要
    Dim str单据 As String
    Dim rs单据 As New ADODB.Recordset
    Dim lng冲销ID As Long
    Dim lng结算标志 As Long
    
    lng结算标志 = g病人身份_重庆渝北.结算标志
    
    Err = 0
    On Error GoTo errHand:
    g病人身份_重庆渝北.结算标志 = 3

    补传住院明细记录 = False

    '读出未上传明细（排序，以便先上传正明细，再上传负明细）
    gstrSQL = "" & _
        "   Select distinct A.NO,A.记录性质,A.记录状态 " & _
        "   From 住院费用记录 A " & _
        "   Where A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页ID & " and A.记帐费用=1  and nvl(A.实收金额,0)<>0 and nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 " & _
        "   Order by A.NO,A.记录性质,Decode(A.记录状态,2,2,1)"
        
    
   zlDatabase.OpenRecordset rs单据, gstrSQL, "获取补传明细记录"
    '先检查是否存在退单情况，如果存在，看有否对应的记录单.
    '先传正单据
    
    With rs单据
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
             gstrSQL = Get明细记录(0, Nvl(!NO), Val(Nvl(!记录性质)), Val(Nvl(!记录状态)))
             Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "获取处方明细")
            Call 处方明细上传(rs明细)
           .MoveNext
        Loop
    End With
     g病人身份_重庆渝北.结算标志 = lng结算标志
    补传住院明细记录 = True
    Exit Function
errHand:
     g病人身份_重庆渝北.结算标志 = lng结算标志
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 病种更改_渝北(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal intinsure As Integer)
'*****************************************************************************
'调用者　　　　　　：被clsInsure 的  ChooseDisease  过程调用
'功能说明　　　　　：选择病人的出院病种
'调用过程清单及说明：
'　　【　　　】医保部件关闭
''*****************************************************************************
    '//TODO:病种选择，在医保前台程序中已有此功能
    Call frm补录病情_重庆渝北.ShowSelect(intinsure, lngPatiID, lngPageID, True)
End Function
