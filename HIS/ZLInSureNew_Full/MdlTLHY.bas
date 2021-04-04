Attribute VB_Name = "MdlTLHY"
Option Explicit
'变量命名规范:全局变量以g打头,模块级变量以m打头
'API函数定义示范
'Public Declare Function BJ_Hosp_Divide3 Lib "FYFJ.dll" Alias "Hosp_Divide3" (ByVal strIn As String) As Long

'可搜索"//TODO:增加自已的实现代码"，找到代码插入点，这些插入点都是必须填写代码的
'-------------------------------------------------------------------------------
'编程步骤说明
'1、为本接口部件命名，规则：zl9I_xxx，如北京医保部件，命名为：zl9I_BJYB，注意，类模块需要命名为：clsI_xxx
'2、如果需要单独保存医保相关的数据，请新建一个用户来处理，我们称之为中间库
'3、与医保相关的参数设置（含中间库的用户名、密码与主机串），请增加保险参数设置窗体，命名规则：frmSet医保名称，如：frmSet北京市
'4、如果中心提供的有医保项目清单、病种目录等，请在保险项目选择的项目更新按钮中填写代码，完成从文件或中心将相关下发数据更新到HIS库中
'5、编写代码完成医保项目对码的功能
'6、编写代码完成身份验证窗体
'7、填入以下函数或过程的主体代码，完成医保接口的主体功能
'8、根据接口性质，修改类模块中GetCapability()方法，相关参数请参见mdlInsure中的枚举变量"医院业务"
'9、根据需要修改类模块中其他方法的调用代码
'10、根据需要增加或修改公共窗体或模块

Public Declare Function GetMyLastError Lib "HopsInterface.dll" () As String '读取最后一次错误的内容
Public Declare Function FreeDllSession Lib "HopsInterface.dll" () As Long '断开与动态链接库的连接
Public Declare Function GetRyInfo Lib "HopsInterface.dll" _
                           (ByVal sCHRYBM As String, _
                            ByRef sXM As String, ByRef sXB As String, ByRef sSFZHM As String, _
                            ByRef sCHJTBH As String, ByRef sCSRQ As String, ByRef sHZXM As String, _
                            ByRef sHZXB As String, ByRef SHZSFZHM As String, ByRef sHKDZ As String, _
                            ByRef SYHZGXMC As String, ByRef cJTZHYE As Double) As Long '读取人员信息
Public Declare Function SetMZFYBXData Lib "HopsInterface.dll" _
                            (ByVal sCHRYBM As String, ByVal sJZYY As String, ByVal sJZRQ As String, _
                            ByVal sBXR As String, ByVal sJZYS As String, ByVal sJZKS As String, _
                            ByVal sZDJG As String, ByVal sSFTJ As String, ByVal sSFMZMXB As String, _
                            ByVal cFSZFY As Double, ByRef cJTZHZFJE As Double, ByRef cBCZFY As Double, _
                            ByRef cXJZFJE As Double, ByRef sMZSJH As String, _
                            ByVal sLYZTDM As String, ByVal sBZDM As String, ByVal sZLKSDM As String, ByVal sJZLXDM As String, _
                            ByVal cXYF As Double, ByVal cZYF As Double, ByVal cJCF As Double, ByVal cZLF As Double, ByRef cTCJJZF As Double) As Long '预结
Public Declare Function SaveMZFYBXData Lib "HopsInterface.dll" (ByVal sMZSJH As String) As Long '保存预结
Public Declare Function StricktheBalanceMZFYBX Lib "HopsInterface.dll" (ByVal sMZSJH As String) As Long '冲销门诊结帐
Public Declare Function InitDllSession Lib "HopsInterface.dll" () As Long '初始化与动态链接库的连接
Public Declare Function SetZYRegister Lib "HopsInterface.dll" _
                        (ByVal sCHRYBM As String, ByVal sJZYS As String, ByVal sJZKS As String, _
                        ByVal dRYRQ As String, ByVal sZYH As String, ByRef sZYJZDH As String, _
                        ByVal sJZLXDM As String, ByVal sLYZTDM As String, ByVal sYSZJDM As String, ByVal sRYZDXX As String, ByVal sRYDJLX As String) As Long '住院登记
Public Declare Function CancelRegister Lib "HopsInterface.dll" (ByVal ZYJZDH As String, ByVal CHRYBM As String) As Long '取消住院登记
Public Declare Function SetZYFYBXYPMX Lib "HopsInterface.dll" _
                        (ByVal ZYJZDH As String, ByVal NHYWBM As String, ByVal YYYWMC As String, _
                        ByVal sl As Double, ByVal je As Double, ByVal dj As Double, _
                        ByRef NBXH As Long) As Long '上传药品费用明细
Public Declare Function ModiZYFYBXYPMX Lib "HopsInterface.dll" (ByVal ZYJZDH As String, ByVal NBXH As Integer) As Long '药品费用明细冲销
Public Declare Function SetZYFYBXZLMX Lib "HopsInterface.dll" _
                        (ByVal ZYJZDH As String, ByVal NHZLBM As String, ByVal YYZLMC As String, _
                        ByVal sl As Double, ByVal je As Double, ByVal dj As Double, _
                        ByRef NBXH As Long) As Long '上传诊疗费用明细
Public Declare Function ModiZYFYBXZLMX Lib "HopsInterface.dll" (ByVal ZYJZDH As String, ByVal NBXH As Integer) As Long  '诊疗费用冲销
Public Declare Function SetZYFYBXCWMX Lib "HopsInterface.dll" _
                        (ByVal ZYJZDH As String, ByVal NHCWBM As String, ByVal YYCWMC As String, _
                        ByVal TS As Double, ByVal je As Double, ByVal dj As Double, _
                        ByRef NBXH As Long) As Long '上传床位费用明细
Public Declare Function ModiZYFYBXCWMX Lib "HopsInterface.dll" (ByVal ZYJZDH As String, ByVal NBXH As Integer) As Long  '床位费用明细冲销
Public Declare Function ZYcheckout Lib "HopsInterface.dll" _
                        (ByVal ZYJZDH As String, ByVal CHRYBM As String, ByVal dCYRQ As String, _
                        ByVal sBZDM As String, ByVal sSSMC As String, ByVal sBXR As String, _
                        ByRef cFSZFY As Double, ByRef cYPFZ As Double, ByRef cJCZLFZ As Double, _
                        ByRef cCWFZ As Double, ByRef cBCFWJE As Double, ByRef cKBYPF As Double, _
                        ByRef cKBJCZLF As Double, ByRef cKBCWF As Double, ByRef cQBXBZ As Double, _
                        ByRef cQBXSJZF As Double, ByRef cBXBL As Double, ByRef cYBJE As Double, _
                        ByRef cDSNLJBC As Double, ByRef cSJBXJE As Double, ByRef cGRZFZFY As Double, _
                        ByVal sSSMCDM As String, ByVal sCYJZLX As String, ByVal sCYZTDM As String, ByVal sZWYYDM As String, _
                        ByRef cJTZHZFJE As Double, ByRef cTCJJZF As Double) As Long '住院预结
Public Declare Function SaveZYFYBXalldata Lib "HopsInterface.dll" (ByVal ZYJZDH As String) As Long '保存住院预结
Public Declare Function StrickthebalanceZYFYBX Lib "HopsInterface.dll" (ByVal sZYJZDH As String) As Long '住院费用冲销包括入院登记
Public Declare Function GetCHRYBM Lib "HopsInterface.dll" _
                        (ByRef CHRYBM As String, ByVal frmY As Long, ByVal frmX As Long) As Long '取得合医编码
Public Declare Function GetBZDM Lib "HopsInterface.dll" _
                        (ByRef BZDM As String, ByRef BZMC As String, ByVal frmY As Long, _
                        ByVal frmX As Long) As Long '取得合医病种代码,名称
Public Declare Function GetJZYY Lib "HopsInterface.dll" (ByRef YLJGDM As String, ByRef YLJGMC As String, ByVal frmX As Long, ByVal frmY As Long) As Long '取医院代码
Public Declare Function GetSSMCDM Lib "HopsInterface.dll" (ByRef sSSMCDM As String, ByRef sSSMC As String, ByVal frmX As Long, ByVal frmY As Long) As Long '取手术代码
Public Declare Function ModiZYRegisterInfo Lib "HopsInterface.dll" (ByVal ZYJZDH As String, ByVal CHRYBM As String, ByVal dRYRQ As String, ByVal sZYH As String, ByVal sJZLXDM As String, ByVal sLYZTDM As String, ByVal sYSZJDM As String, ByVal sRYZDXX As String) As Long '修改入院信息

Public m医保初始化 As Boolean


Public Function 医保初始化_铜梁合医() As Boolean
'>Beging 医保初始化
Dim lngReturn As Long
    If m医保初始化 = False Then
        lngReturn = InitDllSession
        If lngReturn <> 1 Then
            MsgBox "错误信息:" & GetMyLastError & " 初始化失败,不能进行合医交易", vbInformation, "合医返回信息"
            
            Exit Function
        Else
            医保初始化_铜梁合医 = True
        End If
    End If
'>End 医保初始化

End Function


Public Function 医保终止_铜梁合医() As Boolean
Dim lngReturn As Long
    If m医保初始化 = True Then
        lngReturn = FreeDllSession
        If lngReturn <> 1 Then
            MsgBox GetMyLastError, vbInformation, "医保返回信息"
            Exit Function
        Else
            医保终止_铜梁合医 = True
        End If
        m医保初始化 = False
    End If
End Function

Public Function 身份标识_铜梁合医(Optional bytType As Byte, Optional lng病人ID As Long = 0, Optional ByRef intinsure As Integer = 0) As String
    '调用者  ：该方法由门诊费用部件、门诊挂号部件或入院登记部件调用
    '调用时机：在姓名处按回车时
    '功能说明：身份验证成功后，将病人信息串返回给主调程序
    
    Dim strReturn As String
    strReturn = frm身份验证_铜梁合医.GetIdentify(bytType, lng病人ID, type_铜梁合医)
    'if bytType=0 then
    '   完成门诊就诊登记功能
    'end if
    身份标识_铜梁合医 = strReturn
End Function

''Public Function 医保设置_北京(ByVal intInsure As Integer) As Boolean
''    '医保设置_北京 = frmSet北京.参数设置()
''End Function

Public Function 门诊挂号(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
    '调用者  ：该方法由门诊挂号部件调用
    '调用时机：点击门诊挂号窗体的确定按钮时
    '功能说明：通过调用医保商的门诊挂号接口，分解本次费用明细，得到结算结果（个人帐户多少、医保基金多少等）并保存
    '注意事项：需要调用过程zl_病人结算记录_Update对病人预交记录进行数据修正
    
End Function

Public Function 门诊挂号冲销(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
    '调用者  ：该方法由门诊挂号部件调用
    '调用时机：点击门诊挂号窗体的冲销按钮时
    '功能说明：通过调用医保商的门诊挂号冲销接口，完成门诊挂号结算的作废
    
End Function

Public Function 门诊虚拟结算_铜梁合医(rs明细 As ADODB.Recordset, str结算方式 As String, ByVal intinsure As Integer) As Boolean
    '调用者  ：该方法由门诊费用部件调用
    '调用时机：点击门诊收费窗体的预结算按钮时
    '功能说明：通过调用医保商的预结算方法，分解本次费用明细，得到结算结果（个人帐户多少、医保基金多少等），并将结算结果按格式保存在参数“str结算方式”中
    
    '步骤说明
    '1、如果接口需要，请调用费用明细上传接口，将本次明细上传
    '2、调用门诊预结算接口
    '3、将结算结果按规定格式返回
    
    '//TODO:增加自已的实现代码
    'rs明细记录集中是本次录入的门诊处方明细
    'str结算方式的格式说明：报销方式;金额;是否允许修改|....
'    str结算方式 = "医保基金;" & dbl医保基金 & ";0"
'    str结算方式 = str结算方式 & "|大额支付;" & dbl大额支付 & ";0"

Dim 总金额 As Double, 状态 As String, 就诊类型 As String, 开单人 As String, 医保号 As String, 合医信息
Dim R个人帐户 As Double, R医保基金 As Double, R补偿总费用 As Double, R自付费用 As Double, R流水号 As String * 32
Dim rsTmp As New ADODB.Recordset, 总西药费 As Double, 总中药费 As Double, 总检查费 As Double, 总治疗费 As Double

On Error GoTo errHandle

    If rs明细.RecordCount = 0 Then
        str结算方式 = "个人帐户;0;0"
        门诊虚拟结算_铜梁合医 = True
        Exit Function
    End If
    
    '取卡号
    gstrSQL = "select * from 保险帐户 where 险类=" & type_铜梁合医 & " and 病人id=" & rs明细("病人id")
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "取保险卡号")
    开单人 = rs明细("开单人")
    医保号 = rsTmp("医保号")
    合医信息 = Split(rsTmp("合医信息"), "|")
    
    Select Case 合医信息(4)
        Case "一般"
            状态 = "3"
        Case "危"
            状态 = "1"
        Case "急"
            状态 = "2"
        Case Else
            状态 = "4"
    End Select
    就诊类型 = IIf(合医信息(3) = 1, "4", IIf(合医信息(1) = 1, "3", "1"))
    
    '上传总费用及预结
    Do Until rs明细.EOF
        If Val(rs明细!数量) < 0 Or Val(rs明细!实收金额) < 0 Then
            MsgBox "本医保不支持负数记帐", vbInformation, "中联软件"
            门诊虚拟结算_铜梁合医 = False
            str结算方式 = ""
            Exit Function
        End If
        总金额 = 总金额 + rs明细("实收金额")
        rs明细.MoveNext
    Loop
    
    If SetMZFYBXData(医保号, 合医信息(0), CStr(Format(zlDatabase.Currentdate, "yyyy-mm-dd")), UserInfo.姓名, 开单人, "门诊", 合医信息(5), IIf(合医信息(1) = "0", "否", "是"), IIf(合医信息(2) = "0", "否", "是"), 总金额, R个人帐户, R补偿总费用, R自付费用, R流水号, 状态, Nvl(rsTmp!病种代码), "", 就诊类型, 总西药费, 总中药费, 总检查费, 总治疗费, R医保基金) <> 1 Then
        Err.Raise 9000, "合医返回信息", "错误信息" & GetMyLastError()
        str结算方式 = ""
        门诊虚拟结算_铜梁合医 = False
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        str结算方式 = "个人帐户;" & R个人帐户 & ";0|医保基金;" & R医保基金 & ";0"
        门诊虚拟结算_铜梁合医 = True
        Exit Function
    End If
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算_铜梁合医(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, ByVal intinsure As Integer) As Boolean
    '调用者  ：该方法由门诊费用部件调用
    '调用时机：点击门诊收费窗体的结算按钮时
    '功能说明：调用门诊结算接口
    
    '步骤说明
    '1、如果需要上传明细，则调处方明细上传接口
    '2、调用门诊结算接口
    '3、如果成功，则保存保险结算记录
    
    '//TODO:增加自已的实现代码

Dim 总金额 As Double, 状态 As String, 就诊类型 As String, 病种代码 As String, 开单人 As String, 医保号 As String, 病人ID As Long, 帐户余额 As Currency, 合医信息
Dim R个人帐户 As Double, R医保基金 As Double, R补偿总费用 As Double, R自付费用 As Double, R流水号 As String * 32
Dim rsTmp As New ADODB.Recordset, 总西药费 As Double, 总中药费 As Double, 总检查费 As Double, 总治疗费 As Double


On Error GoTo errHandle
    '取卡号
    gstrSQL = "select * from 保险帐户 where 险类=" & type_铜梁合医 & " and 医保号='" & strSelfNo & "'"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "取保险卡号")
    病人ID = rsTmp("病人id")
    医保号 = rsTmp("医保号")
    合医信息 = Split(rsTmp("合医信息"), "|")
    帐户余额 = Nvl(rsTmp("帐户余额"))
    
    Select Case 合医信息(4)
        Case "一般"
            状态 = "3"
        Case "危"
            状态 = "1"
        Case "急"
            状态 = "2"
        Case Else
            状态 = "4"
    End Select
    
    就诊类型 = IIf(合医信息(3) = 1, "4", IIf(合医信息(1) = 1, "3", "1"))
    病种代码 = Nvl(rsTmp!病种代码)
    gstrSQL = "select * from 门诊费用记录 where 结帐id=" & lng结帐ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "取结帐信息")
    开单人 = rsTmp("开单人")
    
    '上传总费用,分类总费用及预结
    Do Until rsTmp.EOF
            总金额 = 总金额 + rsTmp("实收金额")
            Select Case rsTmp!收据费目
                Case "西药费", "中成药费"
                    总西药费 = 总西药费 + rsTmp!实收金额
                Case "中药费"
                    总中药费 = 总中药费 + rsTmp!实收金额
                Case "检查费", "检验费"
                    总检查费 = 总检查费 + rsTmp!实收金额
                Case Else
                    总治疗费 = 总治疗费 + rsTmp!实收金额
            End Select
        rsTmp.MoveNext
    Loop
    
    If SetMZFYBXData(医保号, 合医信息(0), CStr(Format(zlDatabase.Currentdate, "yyyy-mm-dd")), UserInfo.姓名, 开单人, "门诊", 合医信息(5), IIf(合医信息(1) = "0", "否", "是"), IIf(合医信息(2) = "0", "否", "是"), 总金额, R个人帐户, R补偿总费用, R自付费用, R流水号, 状态, 病种代码, "", 就诊类型, 总西药费, 总中药费, 总检查费, 总治疗费, R医保基金) <> 1 Then
        Err.Raise 9000, "合医返回信息", "错误信息" & GetMyLastError()
        门诊结算_铜梁合医 = False
        Exit Function
    Else
        If SaveMZFYBXData(R流水号) <> 1 Then
            Err.Raise 9000, "合医返回信息", "错误信息" & GetMyLastError()
            门诊结算_铜梁合医 = False
            Exit Function
        Else
            gstrSQL = "zl_保险结算记录_insert(1," & _
                                                lng结帐ID & _
                                              "," & type_铜梁合医 & _
                                              "," & 病人ID & _
                                              "," & Format(zlDatabase.Currentdate, "YYYY") & _
                                              ",0" & _
                                              ",0" & _
                                              ",0" & _
                                              ",0" & _
                                              ",0" & _
                                              ",0" & _
                                              ",0" & _
                                              ",0," & 总金额 & _
                                              "," & R自付费用 & _
                                              "," & IIf(R补偿总费用 <> 0, 总金额 - R补偿总费用, R补偿总费用) & _
                                              ",0" & _
                                              "," & R医保基金 & _
                                              ",0" & _
                                              ",0" & _
                                              "," & R个人帐户 & _
                                              ",'" & Trim(MidUni(R流水号, 1, 32)) & "'" & _
                                              ",null" & _
                                              ",null" & _
                                              ",'院码:" & 合医信息(0) & "|体检:" & 合医信息(1) & "|慢病:" & 合医信息(2) & "|接种:" & 合医信息(3) & "|状态:" & 合医信息(4) & "|诊断:" & 合医信息(5) & "|" & Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss") & _
                                              "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
            帐户余额 = 帐户余额 - R个人帐户
            gstrSQL = "zl_保险帐户_更新信息(" & 病人ID & "," & type_铜梁合医 & ",'帐户余额','" & 帐户余额 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新帐户余额")
            门诊结算_铜梁合医 = True
            Exit Function
        End If
    End If
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_铜梁合医(ByVal lng结帐ID As Long, ByVal cur个人帐户 As Currency, ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean
    '调用者  ：该方法由门诊费用部件调用
    '调用时机：点击门诊收费主窗体的作废按钮时
    '功能说明：调用门诊结算作废接口

    '步骤说明
    '1、按接口规则判断是否必须从最后一次就诊的门诊单据开始退废
    '2、调用门诊结算作废接口
    '3、保存保险结算记录
    
    '//TODO:增加自已的实现代码

Dim rsTmp As New ADODB.Recordset
Dim lng冲销ID As Long

On Error GoTo errHand
'根据传入的结帐id查找冲销结帐id
    gstrSQL = "Select a.结帐id As 新结帐id From 病人预交记录 a /*新记录*/ ,病人预交记录 b Where a.No = b.No And a.记录性质 = b.记录性质 And a.记录性质 = 3 And a.记录状态 = 2 And b.结帐ID = " & lng结帐ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "取结帐ID")
    lng冲销ID = rsTmp!新结帐id
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=" & type_铜梁合医 & " and 病人id=" & lng病人ID & " and 记录id=" & lng结帐ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "取保险结算记录")
    
    If StricktheBalanceMZFYBX(rsTmp("支付顺序号")) <> 1 Then
        Err.Raise 9000, "合医返回信息", "错误信息" & GetMyLastError()
        rsTmp.Close
        门诊结算冲销_铜梁合医 = False
        Exit Function
    Else
        '保存保险结算记录
        gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & type_铜梁合医 & "," & rsTmp("病人id") & "," & _
            Format(zlDatabase.Currentdate, "YYYY") & ",0,0,0,0,0,0,0,0," & _
            -1 * Nvl(rsTmp!发生费用金额, 0) & "," & -1 * Nvl(rsTmp!全自付金额, 0) & "," & -1 * rsTmp!首先自付金额 & ",0," & -1 * rsTmp!统筹报销金额 & ",0,0," & _
            -1 * Nvl(rsTmp!个人帐户支付, 0) & ",'" & rsTmp!支付顺序号 & "',null,null,'" & Nvl(rsTmp!备注) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
        gstrSQL = "select * from 保险帐户 where 险类=" & type_铜梁合医 & " and 病人id=" & lng病人ID
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "查余额")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & type_铜梁合医 & ",'帐户余额','" & rsTmp!帐户余额 + cur个人帐户 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新帐户余额")
        门诊结算冲销_铜梁合医 = True
        Exit Function
    End If
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 入院登记_铜梁合医(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean
    '调用者  ：该方法由病人入院部件调用
    '调用时机：点击入院登记窗体的确定按钮时
    '功能说明：调用入院登记接口

    '步骤说明
    '1、从病案主页中提取入院日期（补充入院登记也是调用该接口，因此不能取当前日期做为入院日期上传）
    '2、调用入院登记接口
    '3、执行入院登记过程(zl_保险帐户_入院)，更改病人的当前状态
Dim 医保号 As String, R住院结算号 As String * 32, S就诊类型 As String, S来院状态 As String, S医生代码 As String, S入院诊断 As String, S入院类型 As String
Dim rsTmp As New ADODB.Recordset

On Error GoTo errHand
    
    gstrSQL = "select * from 保险帐户 where 险类=" & type_铜梁合医 & " and 病人id=" & lng病人ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "查合医卡号")
    医保号 = rsTmp!医保号
    
    gstrSQL = "select B.名称 as 入院科室,A.*,C.描述信息 from 病案主页 A,部门表 B,诊断情况 C where A.险类=" & type_铜梁合医 & " and A.病人id=" & lng病人ID & " and A.主页id=" & lng主页ID & " and A.入院科室id=B.id and a.病人id=c.病人id and a.主页id=c.主页id and c.诊断类型=1"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "查病主页中的入院时间")
    S就诊类型 = IIf(Nvl(rsTmp("住院目的")) = "分娩", 5, IIf(rsTmp("住院目的") = "治疗", 2, 9))
    S来院状态 = IIf(Nvl(rsTmp("入院病况")) = "一般", 3, IIf(rsTmp("入院病况") = "急", 2, 1))
    S医生代码 = ""
    S入院诊断 = Nvl(rsTmp("描述信息"), "无")
    S入院类型 = IIf(Nvl(rsTmp("入院方式")) = "转入", "Zr", "Nml")
    
    '调接口,传入的住院号为病人id&主页id
    If SetZYRegister(医保号, Nvl(rsTmp!住院医师, rsTmp!门诊医师), rsTmp!入院科室, CStr(Format(rsTmp!入院日期, "YYYY-MM-DD")), CStr(rsTmp!病人ID) & "_" & CStr(rsTmp!主页ID), R住院结算号, S就诊类型, S来院状态, S医生代码, S入院诊断, S入院类型) <> 1 Then
        Err.Raise 9000, "合医返回信息", "错误信息" & GetMyLastError() & "补充登记失败"
        入院登记_铜梁合医 = False
        Exit Function
    Else
        gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & type_铜梁合医 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & type_铜梁合医 & ",'顺序号','''" & Trim(MidUni(R住院结算号, 1, 32)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算单号")
        入院登记_铜梁合医 = True
        Exit Function
    End If
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 入院登记撤销_铜梁合医(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean
    '调用者  ：该方法由病人入院部件调用
    '调用时机：点击入院登记窗体的取消按钮时
    '功能说明：调用撤销入院登记或出院登记接口
Dim rsTmp As New ADODB.Recordset

On Error GoTo errHand

    gstrSQL = "select * from 保险帐户 where 险类=" & type_铜梁合医 & " and 病人id=" & lng病人ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "查合医住院结算号")
    If CancelRegister(rsTmp!顺序号, rsTmp!医保号) <> 1 Then
        Err.Raise 9000, "合医返回信息", "错误信息" & GetMyLastError()
        入院登记撤销_铜梁合医 = False
        Exit Function
    Else
        gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & type_铜梁合医 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更改病人的当前状态")
        gstrSQL = "ZL_病案主页_撤消医保入院(" & lng病人ID & "," & lng主页ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "撤消医保入院")
        
        gstrSQL = "select * from 住院费用记录 where 病人id=" & lng病人ID & " and 主页id=" & lng主页ID
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "查出病人费用记录并改上传为0")
        
        Do Until rsTmp.EOF
            DoEvents
            gstrSQL = "zl_病人费用记录_更新医保(" & rsTmp!ID & ",null,null,null,null,0,0)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            rsTmp.MoveNext
        Loop
        
        入院登记撤销_铜梁合医 = True
        Exit Function
    End If
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 出院登记_铜梁合医(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean
    '调用者  ：该方法由病人入出院部件调用
    '调用时机：点击出院窗体的确定按钮时
    '功能说明：调用出院登记接口

Dim R病种代码 As String, R病种名称 As String
Dim rsTmp As New ADODB.Recordset

On Error GoTo errHand

    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & type_铜梁合医 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保险出院")
    
    出院登记_铜梁合医 = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_铜梁合医(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean
    '调用者  ：该方法由病人入出院部件调用
    '调用时机：在出院病人区域，点击撤销出院菜单时
    '功能说明：调用撤销出院登记或入院登记接口

    '步骤说明
    '1、按接口规则进行检查
    '2、调用撤销出院登记或入院登记接口
    '3、执行入院登记过程(zl_保险帐户_入院)，更改病人的当前状态
    
    '//TODO:增加自已的实现代码
On Error GoTo errHand

    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保险帐户入院")
    出院登记撤销_铜梁合医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_铜梁合医(ByVal 病人ID As Long, ByVal str医保号 As String, ByVal intinsure As Integer) As Currency
    '调用者  ：该方法由门诊收费部件或住院结算部件调用
    '调用时机：其他部件需要了解当前医保病人的个人帐户余额的情况下
    '功能说明：调用个人帐户余额查询接口或直接从保险帐户表中提取个人帐户余额

    '步骤说明
    '1、调用查询接口获取个人帐户余额并更新保险帐户表
    '2、或者直接从保险帐户中提取个人帐户余额
    
    '//TODO:增加自已的实现代码
    '个人余额 = 0
Dim rsTmp As New ADODB.Recordset
    gstrSQL = "select 帐户余额 from 保险帐户 where 险类=" & type_铜梁合医 & " and 医保号='" & str医保号 & "' and 病人id=" & 病人ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "查个人帐户余额")
    个人余额_铜梁合医 = rsTmp("帐户余额")
    rsTmp.Close
End Function

Public Function 住院结算_铜梁合医(ByVal lng结帐ID As Long, ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean
    '调用者  ：该方法由住院结算部件调用
    '调用时机：点击住院结算窗体中的确定按钮时
    '功能说明：完成本次住院费用的医保结算

    '步骤说明
    '1、调用住院结算接口
    '2、如果住院结算返回的结算结果与住院预结算返回的不一致，需要调用zl_病人结算记录_Update过程进行修正
    
    '//TODO:增加自已的实现代码

Dim s住院结算单号 As String, s合医编码 As String, s出院日期 As String, s病种代码 As String, s手术名称 As String
Dim R发生总费用 As Double, R药品总费用 As Double, R检查治疗总费用 As Double, R床位总费用 As Double, R补偿范围金额 As Double
Dim R可报药品费 As Double, R可报检查治疗费 As Double, R可报床位费 As Double, R起报线标准 As Double, R实际支付起报线 As Double
Dim R报销比例 As Double, R应报金额 As Double, R当时年累计补偿金额 As Double, R实报金额 As Double, R自付费用 As Double, R个人帐户 As Double, R医保基金 As Double
Dim rsTmp As New ADODB.Recordset, S手术代码 As String, S出院类型 As String, S出院状态 As String, S转往医院代码 As String, S转往医院名称 As String

On Error GoTo errHandle

    S手术代码 = Space(10)
    s手术名称 = Space(100)
    S转往医院代码 = Space(20)
    S转往医院名称 = Space(100)
    If GetSSMCDM(S手术代码, s手术名称, 150, 150) <> 1 Then
        S手术代码 = ""
    End If
    
    gstrSQL = "select distinct A.顺序号,A.医保号,B.出院日期,A.病种代码,A.病种名称 ,B.主页id,B.出院方式 from 保险帐户 A,病案主页 B,病人预交记录 C where C.记录性质=2 and A.病人id=B.病人id and B.病人id=C.病人id and A.险类=" & type_铜梁合医 & " and B.主页id=C.主页id and C.结帐id=" & lng结帐ID & " and C.病人id=" & lng病人ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "查结算号,卡号,出院日期,病种代码,名称")
    s住院结算单号 = rsTmp!顺序号
    s合医编码 = CStr(rsTmp!医保号)
    
    If Nvl(Trim(rsTmp!病种代码)) = "" Or Nvl(Trim(rsTmp!病种名称)) = "" Then
        MsgBox "本病人为合医病人，必须先选择病种", vbInformation, "中联提示"
        住院结算_铜梁合医 = False
        Exit Function
    End If
    s病种代码 = CStr(Trim(rsTmp!病种代码))
    s手术名称 = Trim(rsTmp!病种名称)
    s出院日期 = Format(rsTmp!出院日期, "YYYY-MM-DD")
    S出院类型 = IIf(rsTmp!出院方式 = "转院", "Zc", "Nml")
    
    Select Case rsTmp!出院方式
        Case "正常"
            S出院状态 = 1 '治愈
        Case "好转"
            S出院状态 = 2 '好转
        Case "转院"
            S出院状态 = 3 '转院,转院必须要有转院代码
            If GetJZYY(S转往医院代码, S转往医院名称, 150, 150) <> 1 Then
                MsgBox "错误信息" & GetMyLastError() & vbCrLf & "出院方式为转院,必须要有转往医院代码,结算失败", vbInformation, "合医返回信息"
                住院结算_铜梁合医 = False
                Exit Function
            End If
        Case "死亡" '死亡
            S出院状态 = 4
    End Select
    
    '先接口预结,后接口结算,最后保存结算记录
    If ZYcheckout(s住院结算单号, s合医编码, s出院日期, s病种代码, s手术名称, UserInfo.姓名, R发生总费用, R药品总费用, R检查治疗总费用 _
                        , R床位总费用, R补偿范围金额, R可报药品费, R可报检查治疗费, R可报床位费, R起报线标准, R实际支付起报线, R报销比例 _
                        , R应报金额, R当时年累计补偿金额, R实报金额, R自付费用, S手术代码, S出院类型, S出院状态, S转往医院代码, R个人帐户, R医保基金) <> 1 Then
        Err.Raise 9000, "合医返回信息", "错误信息：" & GetMyLastError() & vbLf & "         预结失败，不能结算"
        住院结算_铜梁合医 = False
    ElseIf MsgBox("结算后如果要销帐将比麻烦你真的要结算吗?" & vbLf & "点[是]结算,点[否]取消", vbOKCancel Or vbQuestion, "中联软件") = vbOK Then
        If SaveZYFYBXalldata(s住院结算单号) <> 1 Then
            Err.Raise 9000, "合医返回信息", "错误信息" & GetMyLastError()
            住院结算_铜梁合医 = False
        Else
            gstrSQL = "zl_保险结算记录_insert(2," & _
                                        lng结帐ID & _
                                        "," & type_铜梁合医 & _
                                        "," & lng病人ID & _
                                        "," & Format(zlDatabase.Currentdate, "YYYY") & _
                                        ",0" & _
                                        ",0" & _
                                        ",0" & _
                                        "," & R当时年累计补偿金额 & _
                                        "," & rsTmp!主页ID & _
                                        "," & R起报线标准 & _
                                        ",0," & R实际支付起报线 & _
                                        "," & R发生总费用 & _
                                        "," & R自付费用 & _
                                        "," & R发生总费用 - R补偿范围金额 & _
                                        "," & R补偿范围金额 & _
                                        "," & R医保基金 & _
                                        ",0" & _
                                        ",0" & _
                                        "," & R个人帐户 & _
                                        ",'" & s住院结算单号 & "'" & _
                                        "," & rsTmp!主页ID & _
                                        ",0,'报比:" & CStr(R报销比例) & "|病码:" & MidUni(s病种代码, 1, 30) & "|病名:" & MidUni(s手术名称, 1, 200) & "|" & Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
            住院结算_铜梁合医 = True
        End If
    Else
        住院结算_铜梁合医 = False
    End If
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院虚拟结算_铜梁合医(ByVal rs预结明细 As ADODB.Recordset, ByVal lng病人ID As Long, ByVal intinsure As Integer) As String
'调用者  ：该方法由住院结算部件调用
'调用时机：输入病人信息或选择病人后
'功能说明：完成本次住院费用的医保预结算

Dim D医院费用 As Double, str提示 As String
Dim s住院结算单号 As String, s合医编码 As String, s出院日期 As String, s病种代码 As String, s手术名称 As String
Dim R发生总费用 As Double, R药品总费用 As Double, R检查治疗总费用 As Double, R床位总费用 As Double, R补偿范围金额 As Double
Dim R可报药品费 As Double, R可报检查治疗费 As Double, R可报床位费 As Double, R起报线标准 As Double, R实际支付起报线 As Double
Dim R报销比例 As Double, R应报金额 As Double, R当时年累计补偿金额 As Double, R实报金额 As Double, R自付费用 As Double, R个人帐户 As Double, R医保基金 As Double
Dim rs暂用 As New ADODB.Recordset, S手术代码 As String, S出院类型 As String, S出院状态 As String, S转往医院代码 As String, S转往医院名称 As String

On Error GoTo errHandle

'查住院结算号
    gstrSQL = "select A.顺序号,A.医保号,B.出院日期,A.病种代码,A.病种名称,B.出院方式 from 保险帐户 A,病案主页 B where A.病人id=B.病人id and A.险类=" & type_铜梁合医 & " and B.主页id=" & rs预结明细!主页ID & " and A.病人id=" & rs预结明细!病人ID
    Call zlDatabase.OpenRecordset(rs暂用, gstrSQL, "查结算号,卡号,出院日期,病种代码,名称")
    
    If rs暂用.EOF Or IsNull(rs暂用!顺序号) Then
        MsgBox "该病人可能合医补办尚未成功,请重新补办", vbInformation, "中联软件"
        住院虚拟结算_铜梁合医 = ""
        Exit Function
    End If
    s住院结算单号 = rs暂用!顺序号
    s合医编码 = CStr(rs暂用!医保号)
    s出院日期 = IIf(IsNull(rs暂用!出院日期) Or rs暂用!出院日期 = "", Format(zlDatabase.Currentdate, "YYYY-MM-DD"), Format(rs暂用!出院日期, "YYYY-MM-DD"))
    
    Select Case rs暂用!出院方式
        Case "正常"
            S出院状态 = 1 '治愈
        Case "好转"
            S出院状态 = 2 '好转
        Case "转院"
            S出院状态 = 3 '转院,转院必须要有转院代码
        Case "死亡" '死亡
            S出院状态 = 4
    End Select
    
    S出院类型 = "Nml" '预结时为不选转院代码,固定传正常结算Nml
    '没选病种的不允许结算
    If Nvl(Trim(rs暂用!病种代码)) = "" Or Nvl(Trim(rs暂用!病种名称)) = "" Then
        MsgBox "本病人为合医病人，必须先选择病种", vbInformation, gstrSysName
        住院虚拟结算_铜梁合医 = ""
        Exit Function
    End If
    s病种代码 = Trim(rs暂用!病种代码)
    s手术名称 = Trim(rs暂用!病种名称)
    
    Do Until rs预结明细.EOF
        D医院费用 = D医院费用 + rs预结明细!金额
        Call 处方上传_铜梁合医(rs预结明细!NO, rs预结明细!记录性质, rs预结明细!记录状态, str提示, rs预结明细!病人ID, type_铜梁合医)
        rs预结明细.MoveNext
    Loop
    
    '以下代码为判断是否返回预结结果
    If ZYcheckout(s住院结算单号, s合医编码, s出院日期, s病种代码, s手术名称, UserInfo.姓名, R发生总费用, R药品总费用, R检查治疗总费用 _
                    , R床位总费用, R补偿范围金额, R可报药品费, R可报检查治疗费, R可报床位费, R起报线标准, R实际支付起报线, R报销比例 _
                    , R应报金额, R当时年累计补偿金额, R实报金额, R自付费用, S手术代码, S出院类型, S出院状态, S转往医院代码, R个人帐户, R医保基金) <> 1 Then
        MsgBox "因为以下原因:" & GetMyLastError() & "预结失败", vbInformation, "合医返回信息"
        住院虚拟结算_铜梁合医 = ""
    ElseIf R发生总费用 = D医院费用 Then
        住院虚拟结算_铜梁合医 = "个人帐户;" & R个人帐户 & ";0|医保基金;" & R医保基金 & ";0"
    Else
        If MsgBox("可能部份费用上传失败,医院费用:  " & D医院费用 & " 与合医中心费用" & R发生总费用 & "不等" & vbLf & "点[是]继续,点[否]取消", vbOKCancel Or vbQuestion + vbDefaultButton2, "中联软件") = vbOK Then
            住院虚拟结算_铜梁合医 = "个人帐户;" & R个人帐户 & ";0|医保基金;" & R医保基金 & ";0"
        Else
            住院虚拟结算_铜梁合医 = ""
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_铜梁合医(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
'调用者  ：该方法由住院结算部件调用
'调用时机：对某次结算进行作废时
'功能说明：完成本次住院结算的作废
'张险华20050605实现步骤
'调接口成功后
'1.将全部费用记录上传标志清空
'2.更新保险帐户状态为0
'3.病案主页.病人信息中险类清空
'4.保存保险结算记录

Dim rs暂用 As New ADODB.Recordset, 新结帐id As Long

On Error GoTo errHand

    gstrSQL = "Select a.结帐id As 新结帐id From 病人预交记录 a /*新记录*/ ,病人预交记录 b Where a.No = b.No And a.记录性质 = 12 and b.记录性质=2 And b.结帐ID = " & lng结帐ID
    Call zlDatabase.OpenRecordset(rs暂用, gstrSQL, "查冲销记录结帐id")
    新结帐id = rs暂用!新结帐id
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=" & type_铜梁合医 & " and 记录id=" & lng结帐ID
    Call zlDatabase.OpenRecordset(rs暂用, gstrSQL, "查保险结算号")
    
    If StrickthebalanceZYFYBX(rs暂用!支付顺序号) <> 1 Then
        Err.Raise 9000, "合医返回信息", "错误信息:" & GetMyLastError()
        住院结算冲销_铜梁合医 = False
    Else
        gstrSQL = "Select ID From 住院费用记录 Where 记录性质=2 and 结帐id In (Select b.结帐id From 病人预交记录 a /*销帐记录*/,病人预交记录 b/*原结帐记录*/ Where a.结帐id=" & lng结帐ID & " And a.病人id=b.病人id And a.主页id=b.主页id And b.记录性质=2)"
        Call zlDatabase.OpenRecordset(rs暂用, gstrSQL, "查要销帐的费用记录")
        '将全部费用记录上传标志清空
        If rs暂用.EOF <> True Then rs暂用.MoveFirst
        
        Do While Not rs暂用.EOF
            gstrSQL = "zl_病人费用记录_更新医保(" & rs暂用!ID & ",null,null,null,null,0,0)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "清空病人费用记录中的全部保险信息")
            rs暂用.MoveNext
        Loop
        
        gstrSQL = "select * from 保险结算记录 where 险类=" & type_铜梁合医 & " and 性质=2 and 记录id=" & lng结帐ID
        Call zlDatabase.OpenRecordset(rs暂用, gstrSQL, "在保险结算记录中查病人id和主页id")
        '更新保险帐户状态为1
        gstrSQL = "zl_保险帐户_入院(" & rs暂用!病人ID & "," & type_铜梁合医 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新保险帐户中病人状态为0")
        '病案主页.病人信息中险类清空
        gstrSQL = "zl_病案主页_撤消医保入院(" & rs暂用!病人ID & "," & rs暂用!主页ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病案主页中险类=null")
        '保存保险结算记录
        gstrSQL = "zl_保险结算记录_insert(2," & _
                                                    新结帐id & _
                                                    "," & type_铜梁合医 & _
                                                    "," & rs暂用!病人ID & _
                                                    "," & Format(zlDatabase.Currentdate, "YYYY") & _
                                                    ",0" & _
                                                    ",0" & _
                                                    ",0" & _
                                                    "," & rs暂用!累计统筹报销 & _
                                                    "," & rs暂用!主页ID & _
                                                    "," & rs暂用!起付线 & _
                                                    ",0," & rs暂用!实际起付线 & _
                                                    "," & -1 * rs暂用!发生费用金额 & _
                                                    "," & -1 * rs暂用!全自付金额 & _
                                                    "," & -1 * (rs暂用!发生费用金额 - rs暂用!进入统筹金额) & _
                                                    "," & -1 * rs暂用!进入统筹金额 & _
                                                    "," & -1 * rs暂用!统筹报销金额 & _
                                                    ",0" & _
                                                    ",0" & _
                                                    "," & -1 * rs暂用!个人帐户支付 & _
                                                    ",'" & rs暂用!支付顺序号 & "'" & _
                                                    "," & rs暂用!主页ID & _
                                                    ",0,'" & rs暂用!备注 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
        住院结算冲销_铜梁合医 = True
    End If
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 处方上传_铜梁合医(ByVal str单据号 As String, ByVal int性质 As Integer, ByVal int状态 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0, Optional ByVal intinsure As Integer = 0) As Boolean
    '调用者  ：住院记帐或医嘱发送模块调用
    '调用时机：住院记帐保存时或保存后，根据参数决定（support...，可参见Getcapability）
    '功能说明：完成本次处方明细的上传

    '步骤说明
    '1、提取本单据的处方明细
    '2、仅上传本医保的病人处方
    '3、根据接口性质（每条单独上传或打包上传），将成功上传的明细打上上传标记
    
    '//TODO:增加自已的实现代码
On Error GoTo errHand
    If int状态 = 1 Then
    
        Dim s冲销流水号 As Long, s住院结算号 As String, 是否成功上传 As Boolean, 合医编码 As String
        Dim rs正常明细 As New ADODB.Recordset
        Dim rs临时 As New ADODB.Recordset
        Dim rs正常记帐表 As New ADODB.Recordset
        Dim Conn上传 As New ADODB.Connection
        Set Conn上传 = GetNewConnection
        '记帐表的情况先取病人id然后循环上传
        gstrSQL = "select distinct 病人id from 住院费用记录 where 记录性质=2 and  no='" & str单据号 & "'"
        Call zlDatabase.OpenRecordset(rs正常记帐表, gstrSQL, "查出记帐表中病人id")
        If rs正常记帐表.EOF <> True Then rs正常记帐表.MoveFirst
        
        Do While Not rs正常记帐表.EOF
            gstrSQL = "select A.id,A.NO,A.病人id,A.实收金额,A.标准单价,B.名称,A.计算单位,A.收费细目id,B.编码,A.数次*nvl(A.付数,1) as 数量 ,A.收费类别" & _
                        " from 住院费用记录 A,收费细目 B,病案主页 C " & _
                        "where A.no='" & str单据号 & "' and A.记录性质=" & int性质 & " and A.记录状态=1 And (A.是否上传=0 or A.是否上传 is null)" & _
                        " and A.病人ID=C.病人ID and A.主页ID=C.主页ID And C.险类=" & type_铜梁合医 & " and A.收费细目id=B.id" & _
                        " and A.病人ID=" & rs正常记帐表!病人ID
            Call zlDatabase.OpenRecordset(rs正常明细, gstrSQL, "查本次本医保记帐记录")
            
        
            处方上传_铜梁合医 = True
            If rs正常明细.EOF <> True Then rs正常明细.MoveFirst
            Do While Not rs正常明细.EOF
                gstrSQL = "select * from  保险帐户 where 险类=" & type_铜梁合医 & " and 病人id=" & rs正常明细!病人ID
                Call zlDatabase.OpenRecordset(rs临时, gstrSQL, "查住院结算号")
                '没有进行医保登记的病人不调上传接口
                If rs临时.EOF = True Then
                    str消息 = "病人ID为" & rs正常明细!病人ID & "的病人没有医保记录,请先进行医保登记"
                    Exit Function
                Else
                    s住院结算号 = rs临时!顺序号
                    If Nvl(rs正常明细!实收金额) <> 0 Then '金额为0的不上传
                        '对于负数记帐费用只给出提示,此数不做处理
                        If Val(rs正常明细!数量) < 0 Or Val(rs正常明细!实收金额) < 0 Then
                            str消息 = "本医保不支持负数记帐,请将单据号为" & rs正常明细!NO & "销帐"
                            处方上传_铜梁合医 = False
                        Else
                            gstrSQL = "select * from 保险支付项目 where 险类=" & type_铜梁合医 & " and 收费细目id=" & rs正常明细!收费细目ID
                            Call zlDatabase.OpenRecordset(rs临时, gstrSQL, "检查是否对码")
                            '没对码的不能上传
                            If rs临时.EOF Then
                                str消息 = rs正常明细!名称 & "未对码,上传失败"
                                处方上传_铜梁合医 = False
                            Else
                                合医编码 = rs临时!项目编码
                                '根据收费类别分别上传费用正常明细
                                Select Case rs正常明细!收费类别
                                Case "5", "6", "7"  '药品费用
                                    gstrSQL = "Select c.名称 From 药品目录 a,药品信息 b,药品剂型 c Where a.药名id=b.药名id And b.剂型=c.编码 And a.编码=" & rs正常明细!编码
                                    Call zlDatabase.OpenRecordset(rs临时, gstrSQL, "查剂型")
                                    If SetZYFYBXYPMX(s住院结算号, 合医编码, rs正常明细!名称 & "  |  " & rs临时!名称 & "  |  " & rs正常明细!计算单位, rs正常明细!数量, rs正常明细!实收金额, rs正常明细!标准单价, s冲销流水号) <> 1 Then
                                        是否成功上传 = False
                                    Else
                                        是否成功上传 = True
                                    End If
                                Case "J"    '床位费用
                                    If SetZYFYBXCWMX(s住院结算号, 合医编码, rs正常明细!名称 & "  |  " & rs正常明细!计算单位, rs正常明细!数量, rs正常明细!实收金额, rs正常明细!标准单价, s冲销流水号) <> 1 Then
                                        是否成功上传 = False
                                    Else
                                        是否成功上传 = True
                                    End If
                                Case Else    '检查费用
                                    If SetZYFYBXZLMX(s住院结算号, 合医编码, rs正常明细!名称 & "  |  " & rs正常明细!计算单位, rs正常明细!数量, rs正常明细!实收金额, rs正常明细!标准单价, s冲销流水号) <> 1 Then
                                        是否成功上传 = False
                                    Else
                                        是否成功上传 = True
                                    End If
                                End Select
                                '成功上传后打标记并保存冲销流水号
                                If 是否成功上传 Then
                                    处方上传_铜梁合医 = True
                                    gstrSQL = "zl_病人费用记录_更新医保(" & rs正常明细!ID & ",,,1,'" & 合医编码 & "',1,'" & s冲销流水号 & "')"
                                    Conn上传.Execute gstrSQL, , adCmdStoredProc
                                Else
                                    str消息 = GetMyLastError()
                                    处方上传_铜梁合医 = False
                                    If MsgBox("因为" & GetMyLastError() & vbLf & "单据" & rs正常明细!NO & rs正常明细!名称 & "上传失败" & "     是否继续上传", vbQuestion Or vbOKCancel, "合医返回信息") <> vbOK Then Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
                rs正常明细.MoveNext
            Loop
            rs正常记帐表.MoveNext
        Loop
    Else
        处方上传_铜梁合医 = 处方销帐_铜梁合医(str单据号, int性质, int状态, str消息, lng病人ID, type_铜梁合医)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 处方销帐_铜梁合医(ByVal str单据号 As String, ByVal int性质 As Integer, ByVal int状态 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0, Optional ByVal intinsure As Integer = 0) As Boolean

Dim s住院结算号 As String, 是否成功冲销 As Boolean
Dim rs冲销明细 As New ADODB.Recordset
Dim rs临时 As New ADODB.Recordset
Dim rs冲销记帐表 As New ADODB.Recordset
Dim Conn冲销 As New ADODB.Connection
Set Conn冲销 = GetNewConnection
On Error GoTo errHand:

    If int状态 = 2 Then
        gstrSQL = "select distinct 病人id from 住院费用记录 where 记录性质=2 and no='" & str单据号 & "'"
        Call zlDatabase.OpenRecordset(rs冲销记帐表, gstrSQL, "查出记帐表中病人id")
        If rs冲销记帐表.EOF <> True Then rs冲销记帐表.MoveFirst
        Do While Not rs冲销记帐表.EOF
        '查出要冲销的合医流水号
           gstrSQL = "Select a.Id," & _
                             "a.病人id," & _
                             "a.收费类别," & _
                             "b.实收金额," & _
                             "b.保险编码," & _
                             "b.摘要 As 冲销流水号" & _
                    " From 住院费用记录 a/*新记录*/, 住院费用记录 b/*原记录*/" & _
                    " Where b.No = '" & str单据号 & "'" & _
                            " And b.记录性质 = " & int性质 & _
                            " And b.记录状态 = 3" & _
                            " And b.是否上传 = 1" & _
                            " And a.No = b.No And a.记录性质 = b.记录性质" & _
                            " And a.序号 = b.序号 And a.记录状态 =" & int状态 & _
                            " And (a.是否上传 is null or a.是否上传=0)" & _
                            " and b.病人id=" & rs冲销记帐表!病人ID
            Call zlDatabase.OpenRecordset(rs冲销明细, gstrSQL, "查医保冲销流水号")
            
        
            If rs冲销明细.EOF <> True Then rs冲销明细.MoveFirst
            
            处方销帐_铜梁合医 = True
            Do While Not rs冲销明细.EOF
            
            gstrSQL = "select * from  保险帐户 where 险类=" & type_铜梁合医 & " and 病人id=" & rs冲销明细!病人ID
            Call zlDatabase.OpenRecordset(rs临时, gstrSQL, "查住院结算号")
            If rs临时.EOF = True Then
                str消息 = "病人ID为" & rs冲销明细!病人ID & "的病人没有医保记录,请先进行医保登记"
            Else
                s住院结算号 = rs临时!顺序号
            rs临时.Close
                If Nvl(rs冲销明细!实收金额) <> 0 Then
                    '根据收费类别分别调用接口冲销费用明细
                    Select Case rs冲销明细!收费类别
                        Case "5", "6", "7"  '药品费用
                            If ModiZYFYBXYPMX(s住院结算号, CInt(rs冲销明细!冲销流水号)) <> 1 Then
                                是否成功冲销 = False
                            Else
                                是否成功冲销 = True
                            End If
                        Case "J"    '床位费用
                            If ModiZYFYBXCWMX(s住院结算号, CInt(rs冲销明细!冲销流水号)) <> 1 Then
                                是否成功冲销 = False
                            Else
                                是否成功冲销 = True
                            End If
                        Case Else    '检查费用
                            If ModiZYFYBXZLMX(s住院结算号, CInt(rs冲销明细!冲销流水号)) <> 1 Then
                                是否成功冲销 = False
                            Else
                                是否成功冲销 = True
                            End If
                    End Select
                    '成功冲销后上传标记改为1并保存冲销流水号
                    If 是否成功冲销 Then
                        gstrSQL = "zl_病人费用记录_更新医保(" & rs冲销明细!ID & ",,,1,'" & rs冲销明细!保险编码 & "',1,'" & rs冲销明细!冲销流水号 & "')"
                        Conn冲销.Execute gstrSQL, , adCmdStoredProc
                    Else
                        str消息 = GetMyLastError()
                        处方销帐_铜梁合医 = False
                        If MsgBox("因为" & GetMyLastError() & vbLf & "单据" & str单据号 & "中" & rs冲销明细!名称 & "上传失败" & "       是否继续上传", vbQuestion Or vbOKCancel, "合医返回信息") <> vbOK Then Exit Function
                    End If
                End If
                End If
                是否成功冲销 = False
                rs冲销明细.MoveNext
            Loop
        rs冲销记帐表.MoveNext
        Loop
    End If
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function


Public Function 医保参数设置_铜梁合医(cap参数 As 医院业务) As Boolean
    '
    Select Case cap参数
        Case support门诊预算, _
             support门诊退费, _
             support门诊必须传递明细, _
             support记帐上传, _
             support记帐完成后上传, _
             support记帐作废上传, _
             support医嘱上传, _
             support撤销出院, _
             support未结清出院, _
             support结算使用个人帐户, _
             support出院结算必须出院, _
             support必须录入入出诊断, _
             support撤销出院, _
             support出院病人结算作废
            医保参数设置_铜梁合医 = True
    End Select

End Function

Public Function 病种选择_铜梁合医(lng病人ID As Long, intinsure As Integer) As Boolean
Dim R病种id As String, R病种名称 As String
    R病种id = Space(12)
    R病种名称 = Space(100)
    If GetBZDM(R病种id, R病种名称, 150, 150) <> 1 Then
        MsgBox GetMyLastError() & vbCrLf & "病种更新失败", vbInformation, "合医返回信息"
        病种选择_铜梁合医 = False
    Else
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'病种代码','''" & Trim(MidUni(R病种id, 1, 20)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存病种代码在保险帐户中")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'病种名称','''" & Trim(MidUni(R病种名称, 1, 200)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存病种名称在保险帐户中")
        病种选择_铜梁合医 = True
    End If
End Function
