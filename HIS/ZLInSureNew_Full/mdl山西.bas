Attribute VB_Name = "mdl山西"
Option Explicit

'EmployeeInfo
Private Type 个人基本信息
    个人编号00     As String  '01
    身份证号01     As String  '02
    姓名02         As String  '03
    性别03         As String  '04
    出生日期04     As String  '05
    卡号05         As String  '06
    医疗证号06     As String  '07
    单位编号07     As String  '08
    医疗人员类别08 As String  '09
    公务员标志09   As String  '10
    照顾人员标志10 As String  '11
  ' 扩展12       As String
  ' 扩展13       As String
  ' 扩展14       As String
  ' 扩展15       As String
    在院状态16     As String  '16
    密码          As String
End Type

 '(AccountInfo):
Private Type 帐户基本信息
    帐户结余金额00           As Currency   '01
    帐户年度01               As String  '02
    本年住院次数02           As Integer    '03
    本年总费用支出累计03     As Currency   '04
'   扩展05                 As String
    本年自费累计05           As Currency   '06
    本年自理累计06           As Currency   '07
    本年进入统筹累计07       As Currency   '08
    本年帐户支出累计08       As Currency   '09
'   扩展10                 As String
    本年统筹支付累计10       As Currency   '11
    本年现金支出累计11       As Currency   '12
'   扩展13                 As String
    本年公务员补助支出累计13 As Currency    '14
'   扩展15                 As String
'   扩展16                 As String
'   扩展17                 As String
'   扩展18                 As String
'   扩展19                 As String
'   扩展20                 As String
'   扩展21                 As String
'   扩展22                 As String
End Type

Type 结算信息_山西

'|费用总额|个人帐户支付|统筹支付|现金支付|公务员补助支付
'|自理金额|自费金额|住院人次|起付标准|转院连计费用
'|起付标准自付|起付标准公务员支付|分段1统筹支付|分段1公务员支付|分段1个人自付
'|分段2统筹支付|分段2公务员支付|分段2个人自付|分段3统筹支付|分段3公务员支付
'|分段3个人自付|超封顶公务员支付|公务员个人自付|超公务员封顶个人自付|隶属关系
'|单位类型|

    病人ID        As Long
    医保号        As Long
    费用总额      As Double
    个人帐户余额  As Double
    个人帐户      As Double
    统筹基金      As Double
    公务员补助    As Double
    自理金额      As Double
    自费金额      As Double
    
    
End Type

Public g个人基本信息 As 个人基本信息
Public g帐户基本信息 As 帐户基本信息
Public gcnSxDr As New ADODB.Connection '医保连接
Public g结算信息_山西 As 结算信息_山西

Private clsDR As Object   '测试类
Private mblnInit As Boolean '是否初始化
Private mlngReturn As Long  '接口返回值
Private mstrInput As String  '入参
Private mstrOutput As String '出参
Private mrsTMP As New ADODB.Recordset  '临时记录集
Private mstrSQL As String    '临时存放SQL语句

Public Declare Function InitDLL Lib "DBLIB.DLL" (ByVal intType As Long) As Long
Public Declare Function CommitTrans Lib "DBLIB.DLL" () As Long
Public Declare Function RollbackTrans Lib "DBLIB.DLL" () As Long

Public Declare Function CheckMTQ Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal strPersonNO As String, _
    ByVal strWorkunitNo As String, ByVal strMedKind As String, ByVal strSysdate As String, ByVal strDataBuffer As String) As Long
    
Public Declare Function ReadCard Lib "DBLIB.DLL" (ByVal strEmployeeInfo As String, ByVal strAccountInfo As String, _
    ByVal strDataBuffer As String, ByVal strPin As String, Optional ByVal strDestCardNO As String = vbNullString) As Long
    
Public Declare Function Registration Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal intRegType As Long, _
    ByVal strInHosNo As String, ByVal strApprNO As String, ByVal strMedType As String, _
    ByVal strDiseaseNO As String, ByVal strDiseaseName As String, ByVal strLHStatus As String, ByVal MainDocName As String, _
    ByVal strApprPerson As String, ByVal strTransactor As String, ByVal strTransDate As String, _
    ByVal strDataBuffer As String) As Long
    
Public Declare Function TreatInfoEntry Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal intRegType As Long, _
    ByVal strMedType As String, ByVal strInHosNo As String, ByVal strApprNO As String, _
    ByVal strTreatDate As String, ByVal strLeaveHosDt As String, _
    ByVal strDiseaseNO As String, ByVal strDiseaseName As String, _
    ByVal strLHDiseaseNO As String, ByVal strLHDiseaseName As String, ByVal strLHStatus As String, _
    ByVal strTrunHosKind As String, ByVal strmainDocName As String, _
    ByVal strApprPerson As String, ByVal strTransactor As String, ByVal strTransDate As String, _
    ByVal strDataBuffer As String) As Long
    
Public Declare Function FormularyEntry Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal strInHosNo As String, _
    ByVal intTransKind As Long, ByVal intItemKind As Long, ByVal strInternalCode As String, _
    ByVal strFormularyNo As String, ByVal strSysdate As String, ByVal strCenterCode As String, _
    ByVal strItemName As String, ByVal dblUnitPrice As Double, ByVal dblQuantity As Double, _
    ByVal dblAmount As Double, ByVal strDoseType As String, ByVal strDosage As String, _
    ByVal strFrequency As String, ByVal strUsage As String, ByVal strKeBie As String, _
    ByVal strExecDays As String, ByVal strFeeType As String, ByVal strDoctName As String, _
    ByVal strTransactor As String, ByVal strApprPerson As String, ByVal intIsOwnExpenses As Long, _
    ByVal strDataBuffer As String) As Long
    
Public Declare Function ExpenseCalc Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal intTransType As Long, _
    ByVal intInvoiceKind As Long, ByVal strInHosNo As String, ByVal strMedType As String, _
    ByVal strInvoiceNo As String, ByVal strUserName As String, ByVal dblAccCashPay As Double, _
    ByVal strDataBuffer As String) As Long
    
Public Declare Function PreExpenseCalc Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal strInHosNo As String, _
    ByVal strMedType As String, ByVal strDataBuffer As String) As Long
    
Public Declare Function ChangePinEx Lib "DBLIB.DLL" (ByVal strszOldPin As String, ByVal strszNewPin As String, _
    ByVal strDataBuffer As String) As Long

Public Function 医保初始化_山西() As Boolean
  Dim strUser As String, strServer As String, strPass As String  '医保连接用
    On Error GoTo errHand
  
    If mblnInit = False Then
        mstrSQL = "Select * From 保险参数 Where 险类=" & TYPE_山西
        Set mrsTMP = gcnOracle.Execute(mstrSQL)
        Do Until mrsTMP.EOF
            Select Case mrsTMP!参数名
                Case "医保用户名"
                    strUser = IIf(IsNull(mrsTMP("参数值")), "", mrsTMP("参数值"))
                Case "医保服务器"
                    strServer = IIf(IsNull(mrsTMP("参数值")), "", mrsTMP("参数值"))
                Case "医保用户密码"
                    strPass = IIf(IsNull(mrsTMP("参数值")), "", mrsTMP("参数值"))
                Case "医院等级"
            End Select
            mrsTMP.MoveNext
        Loop
        
        If OraDataOpen(gcnSxDr, strServer, strUser, strPass, False) = False Then
            MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
            Exit Function
        End If
        
        'Set clsDR = CreateObject("sxdr.clssxdr")
        
        mlngReturn = InitDLL(1)
        Call WriteBusinessLOG("InitDll", "", "")
        
        If mlngReturn = 0 Then
            mblnInit = True
            医保初始化_山西 = True
        Else
            MsgBox "初始化失败!", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        医保初始化_山西 = True
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 读卡审核(strInPin As String, Optional ByVal strCarNO As String = "NULL") As Boolean
'EmployeeInfo    OUT 个人基本信息
'AccountInfo     OUT 帐户基本信息
'DataBuffer      OUT 出错信息
'Pin IN  校验Pin个人密码
'DestCardNo  IN [OPTION] 卡号： 空 (默认值):真实读卡，返回读卡信息
'                              非空:模拟读卡，按此卡号从库中读出相关信息。
'                              （主要用于不读卡时的住院转方和预结）
    '医疗人员类别
    
    '编码  人员类别
    '11     在职
    '21     退休
    '33     二等乙级伤残军人
    '91     其他人员
    
    '姓名说明（暂无，自定义）
    '0   男
    '1   女
    
    '此处交易失败不回滚，外部调用时，把此函数当成开始事务，外部调用处根据反回的结果进行回滚

If mblnInit = False Then
    MsgBox "请先初始化医保接口！", vbInformation, gstrSysName
     Exit Function
End If

Dim retEmpInfo As String
Dim retAccInfo As String
Dim str封锁类型 As String
retEmpInfo = Space(600)
retAccInfo = Space(600)
mstrOutput = Space(600)

If strCarNO = "" Or strCarNO = "NULL" Then strCarNO = vbNullString

mlngReturn = ReadCard(retEmpInfo, retAccInfo, mstrOutput, strInPin, strCarNO)
Call WriteBusinessLOG("ReadCard", strInPin & "：" & strCarNO, Trim(mstrOutput))

If mlngReturn = 0 Then
    g个人基本信息.个人编号00 = Split(retEmpInfo, "|")(1)
    g个人基本信息.身份证号01 = Split(retEmpInfo, "|")(2)
    g个人基本信息.姓名02 = Split(retEmpInfo, "|")(3)
    g个人基本信息.性别03 = IIf(Split(retEmpInfo, "|")(4) = "0", "男", "女")
    g个人基本信息.出生日期04 = Split(retEmpInfo, "|")(5)
    g个人基本信息.卡号05 = Split(retEmpInfo, "|")(6)
    g个人基本信息.医疗证号06 = Split(retEmpInfo, "|")(7)
    g个人基本信息.单位编号07 = Split(retEmpInfo, "|")(8)
    g个人基本信息.医疗人员类别08 = Split(retEmpInfo, "|")(9)
    g个人基本信息.公务员标志09 = Split(retEmpInfo, "|")(10)
    g个人基本信息.照顾人员标志10 = Split(retEmpInfo, "|")(11)
    g个人基本信息.在院状态16 = Split(retEmpInfo, "|")(17)
    g个人基本信息.密码 = strInPin
    
    g帐户基本信息.帐户结余金额00 = Split(retAccInfo, "|")(1)
    g帐户基本信息.帐户年度01 = Split(retAccInfo, "|")(2)
    g帐户基本信息.本年住院次数02 = Split(retAccInfo, "|")(3)
    g帐户基本信息.本年总费用支出累计03 = Split(retAccInfo, "|")(4)
    g帐户基本信息.本年自费累计05 = Split(retAccInfo, "|")(6)
    g帐户基本信息.本年进入统筹累计07 = Split(retAccInfo, "|")(8)
    g帐户基本信息.本年帐户支出累计08 = Split(retAccInfo, "|")(9)
    g帐户基本信息.本年统筹支付累计10 = Split(retAccInfo, "|")(11)
    g帐户基本信息.本年现金支出累计11 = Split(retAccInfo, "|")(12)
    g帐户基本信息.本年公务员补助支出累计13 = Split(retAccInfo, "|")(14)
    
   
    With g个人基本信息
        mstrOutput = Space(600)
        mlngReturn = CheckMTQ(.卡号05, .个人编号00, .单位编号07, .医疗人员类别08, Format(zlDatabase.Currentdate, "yyyyMMdd"), mstrOutput)
        Call WriteBusinessLOG("CheckMTQ", .卡号05 & "," & .个人编号00 & "," & .单位编号07 & "," & .医疗人员类别08 & "," & Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
    End With
    If mlngReturn = 0 Then
        str封锁类型 = Split(mstrOutput, "|")(1)
        If Val(str封锁类型) = 0 Then
            读卡审核 = True
        ElseIf (Val(str封锁类型) >= 1 And Val(str封锁类型) <= 30) Or Val(str封锁类型) = 42 Then
            读卡审核 = False
            MsgBox "待遇封锁，不能执行医保交易！" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
            Call 回退_山西
        ElseIf Val(str封锁类型) > 30 And Val(str封锁类型) <> 42 Then
            读卡审核 = True
            MsgBox "待遇部分封锁，不能完全享受医保待遇！" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
        End If
    Else
        读卡审核 = False
        MsgBox "待遇审核失败！" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
        Call 回退_山西
    End If
Else
    读卡审核 = False
    MsgBox "读卡失败！" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
End If

End Function

Public Function 身份标识_山西(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim str就诊类别 As String
    Dim str门诊号 As String
    Dim str日期 As String
    
    If Not (bytType = 1 Or bytType = 0) Then Exit Function  '仅门诊收费，入院登记才调用
    
    Dim strIdeReturn As String  '接收返回信息
    
    strIdeReturn = frmIdentify山西.身份标识(bytType, lng病人ID) '''发起接口事务,readcard
    
    If strIdeReturn = "-1" Then
        身份标识_山西 = ""
    Else
        mstrSQL = "Select * from 保险帐户 where 病人ID=[1] and 险类=[2]"
        Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "病人信息", lng病人ID, TYPE_山西)
        
        If mrsTMP.EOF Then
            MsgBox "不是山西医保病人！，不能执行交易。", vbInformation, gstrSysName
            Exit Function
        Else
            If bytType = 0 Then
                str就诊类别 = Nvl(mrsTMP!就诊类别, "11")
                str门诊号 = mrsTMP!病人ID & "_0_" & IIf(IsNumeric(Nvl(mrsTMP!顺序号, 0)), mrsTMP!顺序号, 0) + 1
                
                gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_山西 & ",'顺序号','''" & IIf(IsNumeric(Nvl(mrsTMP!顺序号, 0)), Nvl(mrsTMP!顺序号, 0), 0) + 1 & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新顺序号")
            Else
                str就诊类别 = Nvl(mrsTMP!就诊类别, "21")
            End If
        End If
        
        mstrSQL = "select * from 保险病种 where ID=(" & _
                        "Select 病种ID from 保险帐户 where 病人ID=[1]" & _
                                                     " and 险类=[2]) " & _
                                           " and 险类=[2]"
        Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "病种信息", lng病人ID, TYPE_山西)
        
        If mrsTMP.EOF Then
            MsgBox "病种目录不正确！，不能执行交易。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If bytType = 0 Then
        
                    '门诊调用挂号登记交易
                    mstrOutput = Space(600)
                    str日期 = Format(zlDatabase.Currentdate, "yyyyMMdd")
                    mlngReturn = Registration(g个人基本信息.卡号05, _
                                                                        1, _
                                                              str门诊号, _
                                                            "", _
                                                              str就诊类别, _
                                                               mrsTMP!编码, _
                                                               mrsTMP!名称, _
                                                                "", "", _
                                                                "", _
                                                             UserInfo.姓名, _
                                                                str日期, _
                                                                mstrOutput)
                                                                
                    Call WriteBusinessLOG("Registration", g个人基本信息.卡号05 & "," & _
                                                                             "1," & _
                                                                    str门诊号 & "," & _
                                                                          "," & _
                                                              str就诊类别 & "," & _
                                                               mrsTMP!编码 & "," & _
                                                               mrsTMP!名称 & "," & _
                                                                            "0," & _
                                                                           "," & _
                                                             UserInfo.姓名 & "," & _
                                Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
                
            If mlngReturn = 0 Then
                提交_山西
            Else
                回退_山西
                MsgBox "挂号登记失败" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            Call 提交_山西
        End If
    
        '提交_山西
        身份标识_山西 = strIdeReturn
    End If

End Function

Public Function 提交_山西()
    CommitTrans
    Call WriteBusinessLOG("ComitTrans", "", "")
End Function

Public Function 回退_山西()
    RollbackTrans
    Call WriteBusinessLOG("RollbackTrans ", "", "")
End Function

Public Function 个人余额_山西(strSelfNo As String) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '功能: 提取参保病人个人帐户余额
    '参数: strSelfNO-病人个人编号
    '返回: 返回个人帐户余额的金额
    '如果是门诊，返回家庭帐户余额；住院返回个人帐户余额
    gstrSQL = "Select Nvl(帐户余额,0) AS 个人帐户 From 保险帐户 " & _
              " Where 医保号=[1] and 险类=[2]"
              
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人ID", strSelfNo, TYPE_山西)
    个人余额_山西 = rsTemp!个人帐户
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_山西(rs明细记录 As ADODB.Recordset, str结算方式 As String, Optional str结算 As String) As Boolean

Dim lng病人ID As Long
Dim str就诊类别 As String
Dim str门诊号 As String
Dim str处方明细流水号 As String
Dim dbl个人帐户 As Double, dbl统筹基金 As Double, dbl公务员补助 As Double
Dim rsTmpcd As New ADODB.Recordset '陈东用的临时数据集
Dim str保险项目编码 As String
On Error GoTo errHandle
If rs明细记录.RecordCount = 0 Then
    MsgBox "没有病人费用明细，不能进行医保操作", vbInformation, gstrSysName
    Exit Function
End If

'虚拟读卡审核
lng病人ID = rs明细记录!病人ID

mstrSQL = "Select * from 保险帐户 where 病人ID=[1] and 险类=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "病人信息", CLng(rs明细记录!病人ID), TYPE_山西)

If mrsTMP.EOF Then
    MsgBox "不是山西医保病人！，不能执行交易。", vbInformation, gstrSysName
    Exit Function
Else
    str就诊类别 = Nvl(mrsTMP!就诊类别, "11")
    str门诊号 = mrsTMP!病人ID & "_0_" & IIf(IsNumeric(Nvl(mrsTMP!顺序号, 0)), mrsTMP!顺序号, 0)
End If

'结算调用不读卡
If Nvl(str结算, 0) <> 9 Then
    If 读卡审核(mrsTMP!密码, mrsTMP!卡号) = False Then
        Exit Function
    End If
End If

'检查明细
rs明细记录.MoveFirst
Do Until rs明细记录.EOF
    '未对码的情况
    mstrSQL = "select * from 保险支付项目 where 险类=[1] and 收费细目ID=[2]"
    Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "医保明细", TYPE_山西, CLng(rs明细记录!收费细目ID))
    If mrsTMP.EOF Then
       ' Call 回退_山西
        mstrSQL = "Select * from 收费项目目录 where ID=[1]"
        Set rsTmpcd = zlDatabase.OpenSQLRecord(mstrSQL, "收费项目目录", CLng(rs明细记录!收费细目ID))
        MsgBox rsTmpcd!名称 & "(" & rsTmpcd!编码 & ")" & "未对码！" & vbCrLf & "请对码后再使用此功能！", vbInformation, gstrSysName
        Exit Function
    End If
    '对的码医保中心找不到的情况
    mstrSQL = "Select aka060 医保编码,aka061  项目名称,aka065  项目等级,aka069  自付比例,1 交易类型,aka063  收费类别" & _
              " From ka02 where aka060='" & mrsTMP!项目编码 & "' and aka061='" & mrsTMP!项目名称 & "' and " & Nvl(mrsTMP!附注, 0) & "=1" & _
              " union all " & _
              " Select aka090   医保编码,aka091 项目名称,aka065 项目等级,aka069 自付比例,2 交易类型,aka063  收费类别" & _
              " From ka03 where aka090='" & mrsTMP!项目编码 & "' and aka091='" & mrsTMP!项目名称 & "' and " & Nvl(mrsTMP!附注, 0) & "=2" & _
              " union all " & _
              "Select aka100   医保编码,aka102 项目名称,aka103 病床等级,0,3,aka063 收费类别 " & _
              " From ka04 where aka100='" & mrsTMP!项目编码 & "' and aka102='" & mrsTMP!项目名称 & "' and " & Nvl(mrsTMP!附注, 0) & "=3"
    str保险项目编码 = mrsTMP!项目编码
    Call OpenRecordset_OtherBase(mrsTMP, "医保明细", mstrSQL, gcnSxDr)
    If mrsTMP.EOF Then
      
      '  Call 回退_山西
        mstrSQL = "Select * from 收费项目目录 where ID=[1]"
        Set rsTmpcd = zlDatabase.OpenSQLRecord(mstrSQL, "收费项目目录", CLng(rs明细记录!收费细目ID))
        MsgBox rsTmpcd!名称 & "(" & rsTmpcd!编码 & ")" & "对码错误！" & vbCrLf & "请核对后再使用此功能！", vbInformation, gstrSysName
        Exit Function
    End If
    rs明细记录.MoveNext
Loop

'传明细
'

rs明细记录.MoveFirst

Do Until rs明细记录.EOF
    str处方明细流水号 = zlDatabase.GetNextID("人员表")
    mstrSQL = "select * from 保险支付项目 where 险类=[1] and 收费细目ID=[2]"
    Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "医保明细", TYPE_山西, CLng(rs明细记录!收费细目ID))
    
    mstrSQL = "Select aka060 医保编码,aka061  项目名称,aka065  项目等级,aka069  自付比例,1 交易类型,aka063  收费类别" & _
              " From ka02 where aka060='" & mrsTMP!项目编码 & "' and aka061='" & mrsTMP!项目名称 & "' and " & Nvl(mrsTMP!附注, 0) & "=1" & _
              " union all " & _
              "Select aka090   医保编码,aka091 项目名称,aka065 项目等级,aka069 自付比例,2 交易类型,aka063  收费类别" & _
              " From ka03 where aka090='" & mrsTMP!项目编码 & "' and aka091='" & mrsTMP!项目名称 & "' and " & Nvl(mrsTMP!附注, 0) & "=2" & _
              " union all " & _
              "Select aka100   医保编码,aka102 项目名称,aka103 病床等级,0,3,aka063 收费类别 " & _
              " From ka04 where aka100='" & mrsTMP!项目编码 & "' and aka102='" & mrsTMP!项目名称 & "' and " & Nvl(mrsTMP!附注, 0) & "=3"
  
    Call OpenRecordset_OtherBase(mrsTMP, "医保明细", mstrSQL, gcnSxDr)
    
'CardNo 卡号 N ,InHosNo IN  住院号,TransKind,  IN  交易类型(1  处方录入 -1 退处方上的一条明细)   N
'ItemKind    IN  项目类别(1:药品, 2:诊疗, 3服务设施),InternalCode    IN  收费项目医院内编码  N,FormularyNo IN  处方号（同一患者要保证唯一）    N
'SysDate IN  开方日期(yyyymmdd)  N,CenterCode  IN  收费项目医保中心编码    N,ItemName    IN  收费项目名称    N
'UnitPrice   IN  单价    N,Quantity    IN  数量    N,Amount  IN  金额    N,DoseType    IN  剂型,Dosage  IN  剂量
'Frequency   IN  频次,Usage   IN  用法,KeBie   IN  科别名称,ExecDays    IN  执行天数,FeeType IN  医保中心收费类别（见附录）  N
'DoctName    IN  开方医生,Transactor  IN  经办人  N,ApprPerson  IN  审批人（保留参数）,IsOwnExpenses   IN  全额自费标志( 0：非全额自费 1：全额自费)    N

'DataBuffer  OUT 出错信息/返回信息
'
'根据str结算 =9来判断是否是结算，结算是按记录性质+NO+序号做为流水号
If Nvl(str结算, 0) = 9 Then
   str处方明细流水号 = rs明细记录!记录性质 & rs明细记录!NO & rs明细记录!序号
End If
    mstrOutput = Space(600)
    mlngReturn = FormularyEntry(g个人基本信息.卡号05, _
                                                 str门诊号, _
                                                         1, _
                                           mrsTMP!交易类型, _
                                     rs明细记录!收费细目ID, _
                                         str处方明细流水号, _
                                       Format(rs明细记录!结算时间, "yyyyMMdd"), _
                                           mrsTMP!医保编码, _
                                           mrsTMP!项目名称, _
                 rs明细记录!实收金额 / (rs明细记录!数量), _
                                 rs明细记录!数量, _
                                           rs明细记录!实收金额, _
                                                        "", _
                                                        "", _
                                                        "", _
                                                        "", _
                                                        "", _
                                                        "", _
                                           mrsTMP!收费类别, _
                                                        "", _
                                             UserInfo.姓名, _
                                                        "", _
                             IIf(mrsTMP!自付比例 = 1, 1, 0), _
                                                mstrOutput)
    
    Call WriteBusinessLOG("FormularyEntry", g个人基本信息.卡号05 & "," & _
                                                       str门诊号 & "," & _
                                                            "1," & _
                                           mrsTMP!交易类型 & "," & _
                                     rs明细记录!收费细目ID & "," & _
                                               str处方明细流水号 & "," & _
                                       Format(rs明细记录!结算时间, "yyyyMMdd") & "," & _
                                           mrsTMP!医保编码 & "," & _
                                           mrsTMP!项目名称 & "," & _
               rs明细记录!实收金额 / (rs明细记录!数量) & "," & _
                                 rs明细记录!数量 & "," & _
                                           rs明细记录!实收金额 & "," & _
                                                             "," & _
                                                             "," & _
                                                             "," & _
                                                             "," & _
                                                             "," & _
                                                             "," & _
                                           mrsTMP!收费类别 & "," & _
                                                             "," & _
                                             UserInfo.姓名 & "," & _
                                                             "," & _
                             IIf(mrsTMP!自付比例 = 1, 1, 0) & "," _
                                                , Trim(mstrOutput))
    If mlngReturn = -1 Then
        Call 回退_山西
        MsgBox "上传" & mrsTMP!医保编码 & " " & mrsTMP!项目名称 & "失败，交易中止！" & _
                vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
        Exit Function
    End If
    '>beging 处理返回值
                '在费用记录中记录进入统筹金额
                '项目编码中保存项目类型（药品，诊疗）,摘要中保存自付比例,可根据比例得到甲类，乙类
        If Nvl(str结算, 0) = 9 Then
            mstrOutput = Replace(mstrOutput, "|", ";")
            
            gstrSQL = "ZL_病人费用记录_更新医保(" & rs明细记录!ID & "," & _
                    Split(mstrOutput, ";")(1) - Split(mstrOutput, ";")(3) - Split(mstrOutput, ";")(4) & _
                    ",NULL,1,'" & str保险项目编码 & "',NULL,NULL)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新医保字段")
            
        End If

    '>end  处理返回值
    
    rs明细记录.MoveNext
Loop

'调预结算交易
'CardNo  IN  卡号    N，InHosNo IN  住院号（门诊号）    N，MedType IN  医疗类别（同门诊挂号）  N
'DataBuffer  OUT 结算结果(结算执行成功)或出错原因(结算执行失败)  建议长度600以上

mstrOutput = Space(600)
'
If Nvl(str结算, 0) = 9 Then
    ''结算
    门诊虚拟结算_山西 = True
    str结算方式 = str门诊号 & "|" & str就诊类别
    Exit Function
Else
    ''预结算
    mlngReturn = PreExpenseCalc(g个人基本信息.卡号05, str门诊号, str就诊类别, mstrOutput)
    Call WriteBusinessLOG("PreExpenseCalc", g个人基本信息.卡号05 & "," & str门诊号 & "," & str就诊类别, Trim(mstrOutput))
End If

If mlngReturn = 0 Then
    '正确反回
    
    dbl个人帐户 = Val(Split(mstrOutput, "|")(2))
    dbl统筹基金 = Val(Split(mstrOutput, "|")(3))
    dbl公务员补助 = Val(Split(mstrOutput, "|")(5))
 '个人帐户允许前台修改，所以传为1位置在箭头所指 V
    str结算方式 = "个人帐户;" & dbl个人帐户 & ";1|统筹基金;" & dbl统筹基金 & ";0|公务员补助;" & dbl公务员补助 & ";0"
    门诊虚拟结算_山西 = True
Else
    门诊虚拟结算_山西 = False
End If
    
If Nvl(str结算, 0) <> 9 Then
    Call 回退_山西
End If
    
    Exit Function
errHandle:
    门诊虚拟结算_山西 = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call 回退_山西
End Function

Public Function 门诊结算_山西(lng结帐ID As Long, cur个帐支付 As Currency, str医保号 As String) As Boolean
'
    '根据结帐ID提取记录集,传递给虚拟结算接口，接中中根据传入的"str结算"，标识判断是否提交明细。
    Dim rs结算明细 As New ADODB.Recordset  ''传递结算明细记录
    Dim str预算信息 As String   '接收预结算返回的信息
    Dim lng病人ID As Long
    Dim cur余额 As Currency
On Error GoTo ErrH
    mstrSQL = "Select ID,NO,序号,记录性质,登记时间 as 结算时间,病人ID,收费类别,收据费目,计算单位,开单人, " & _
                     "收费细目ID,nvl(数次,0)*nvl(付数,0) as 数量,标准单价 as 单价, " & _
                     "实收金额,统筹金额,保险大类ID 保险支付大类ID, " & _
                     " 摘要,是否急诊 " & _
            "from 门诊费用记录 " & _
            "where 结帐ID=[1]"
    
    Set rs结算明细 = zlDatabase.OpenSQLRecord(mstrSQL, "门诊结算明细", lng结帐ID)
    lng病人ID = rs结算明细!病人ID
    
    
    If 读卡审核(g个人基本信息.密码, vbNullString) = False Then
        Exit Function
    End If
    
    If 门诊虚拟结算_山西(rs结算明细, str预算信息, 9) Then
        '交易ExpenseCalc
        mstrOutput = Space(600)
        rs结算明细.MoveFirst
        
        
        ' 门诊号                     ,医疗类别
        'ExpenseCalc( char* CardNo,  //卡号
        '                      int   TransType,      //交易类型
        '                      int   InvoiceKind,    //发票类型 0: 门诊 1:住院
        '                      char* InHosNo,        //住院门诊号
        '                      char* MedType,        //医疗类别
        '                      char* InvoiceNo,      //单据号
        '                      char* UserName,       //经办人
        '                      double AccCashPay,    //现金代交金额
        '                      char* DataBuffer );   //结算结果
        mstrOutput = Space(600)
        mlngReturn = ExpenseCalc(g个人基本信息.卡号05, 1, 0, Split(str预算信息, "|")(0), Split(str预算信息, "|")(1), _
                                       "1" & rs结算明细!NO, UserInfo.姓名, cur个帐支付, mstrOutput)
                                       
        Call WriteBusinessLOG("ExpenseCalc", g个人基本信息.卡号05 & "," & 1 & "," & 0 & "," & Split(str预算信息, "|")(0) & "," & Split(str预算信息, "|")(1) & "," & _
                                       "1" & rs结算明细!NO & "," & UserInfo.姓名 & "," & cur个帐支付, Trim(mstrOutput))
        '成功交易后提交
        If mlngReturn = -1 Then
            Call 回退_山西
            门诊结算_山西 = False
            Err.Raise 9000, gstrSysName, Trim(mstrOutput)
        Else
        g结算信息_山西.个人帐户 = cur个帐支付
        g结算信息_山西.统筹基金 = Val(Split(mstrOutput, "|")(3))
        g结算信息_山西.公务员补助 = Val(Split(mstrOutput, "|")(5))
        g结算信息_山西.费用总额 = Val(Split(mstrOutput, "|")(1))
        g结算信息_山西.自费金额 = Val(Split(mstrOutput, "|")(7))
        g结算信息_山西.自理金额 = Val(Split(mstrOutput, "|")(6))
            
        '保险结算记录
        gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_山西 & "," & _
                lng病人ID & "," & g帐户基本信息.帐户年度01 & ",0,0, " & _
                "" & _
                g帐户基本信息.本年统筹支付累计10 + g帐户基本信息.本年公务员补助支出累计13 & "," & g帐户基本信息.本年住院次数02 & ",NULL,NULL,NULL,0," & _
                g结算信息_山西.费用总额 & "," & g结算信息_山西.自费金额 & "," & g结算信息_山西.自理金额 & ",NULL," & g结算信息_山西.统筹基金 + g结算信息_山西.公务员补助 & ",NULL,NULL," & _
                cur个帐支付 & ",'" & Split(str预算信息, "|")(0) & "',NULL,NULL,'" & Split(str预算信息, "|")(1) & "')"
                '                  门诊号为str交易流水号                           医疗类别.
        Call zlDatabase.ExecuteProcedure(gstrSQL, "山西医保")
        
        cur余额 = 个人余额_山西(g个人基本信息.个人编号00) - cur个帐支付
        gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_山西 & ",'帐户余额','" & cur余额 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "山西医保")
            
        Call 提交_山西
            门诊结算_山西 = True
        End If
    End If
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_山西(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean

    Dim lng冲销ID As Long
    Dim str卡号 As String
    Dim str密码 As String
    Dim StrInput As String
    Dim rsTemp As New ADODB.Recordset
    Dim str门诊号 As String
    Dim str医疗类别  As String
    Dim strNO As String
    
    On Error GoTo errHand
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    '只能退最后一笔
    '取冲销记录的结帐ID，单据号
    
    '取卡证号码
    gstrSQL = "Select 卡号,密码 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读卡号密码", TYPE_山西, lng病人ID)
    str卡号 = Nvl(rsTemp!卡号)
    str密码 = Nvl(rsTemp!密码)
    gstrSQL = "select distinct A.结帐ID,A.NO from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!结帐ID
    strNO = rsTemp!NO
    '取结算流水号
    gstrSQL = "Select * From 保险结算记录 Where 性质=1 And 记录ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取结算流水号", lng结帐ID, TYPE_山西)
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "没有找到原始结算记录，无法进行门诊结算冲销！"
        Exit Function
    End If
    str门诊号 = Nvl(rsTemp!支付顺序号)
    str医疗类别 = Nvl(rsTemp!备注)
    
    If str门诊号 = "" Or str医疗类别 = "" Then
        Err.Raise 9000, gstrSysName, "到原始结算记录交易号不全，无法进行门诊结算冲销！"
        门诊结算冲销_山西 = False
        Exit Function
    End If
    
    ''真实读卡  根据情况，可能要加入密码录入窗口,现在暂不处理
    
    If 读卡审核(str密码) = False Then
        门诊结算冲销_山西 = False
        Exit Function
    End If
    
    '不清楚调不调挂号交易，现在不处理，如果要调，再添加
    
    '调用结算冲销
    mstrOutput = Space(600)
    mlngReturn = ExpenseCalc(str卡号, -1, 0, str门诊号, str医疗类别, _
                                   strNO, UserInfo.姓名, cur个人帐户, mstrOutput)
    Call WriteBusinessLOG("ExpenseCalc", str卡号 & ", -1, 0," & str门诊号 & "," & str医疗类别 & "," & _
                                   strNO & "," & UserInfo.姓名 & "," & cur个人帐户, Trim(mstrOutput))
    '成功后,保存本次结算情况
    
    If mlngReturn = 0 Then
        Call 提交_山西
        gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_山西 & "," & lng病人ID & "," & _
            Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
            0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
            -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & "," & -1 * Nvl(rsTemp!首先自付金额, 0) & "," & -1 * Nvl(rsTemp!进入统筹金额, 0) & "," & -1 * Nvl(rsTemp!统筹报销金额, 0) & ",0,0," & _
            -1 * Nvl(rsTemp!个人帐户支付, 0) & ",'" & Nvl(rsTemp!支付顺序号) & "',null,null,'" & Nvl(rsTemp!备注) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "门诊结算冲销")
        
        门诊结算冲销_山西 = True
    Else
        Call 回退_山西
        门诊结算冲销_山西 = False
        MsgBox "退费交易失败！" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
    End If
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function


Public Function 取消就诊_山西()

Dim str门诊号 As String, str就诊类别 As String
Dim lng病种ID As Long



mstrSQL = "select * from 保险帐户 where 卡号=[1] and 险类=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "保险帐户", g个人基本信息.卡号05, TYPE_山西)

str门诊号 = mrsTMP!病人ID & "_0_" & IIf(IsNumeric(Nvl(mrsTMP!顺序号, 0)), mrsTMP!顺序号, 0)
str就诊类别 = mrsTMP!就诊类别
lng病种ID = mrsTMP!病种ID

If 读卡审核(mrsTMP!密码, g个人基本信息.卡号05) = False Then
    Exit Function
End If

mstrSQL = "select * from 保险病种 where ID=[1] and 险类=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "保险病种", lng病种ID, TYPE_山西)

    '门诊调用退号登记交易
            mstrOutput = Space(600)
            mlngReturn = Registration(g个人基本信息.卡号05, _
                                                                -1, _
                                                      str门诊号, _
                                                               "", _
                                                      str就诊类别, _
                                                       mrsTMP!编码, _
                                                       mrsTMP!名称, _
                                                                  "", "", _
                                                                 "", _
                                                     UserInfo.姓名, _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), _
                                                        mstrOutput)
                                                        
            Call WriteBusinessLOG("Registration", g个人基本信息.卡号05 & "," & _
                                                                     "-1," & _
                                                            str门诊号 & "," & _
                                                                  "," & _
                                                      str就诊类别 & "," & _
                                                       mrsTMP!编码 & "," & _
                                                       mrsTMP!名称 & "," & _
                                                                    "0," & _
                                                                   "," & _
                                                     UserInfo.姓名 & "," & _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
    If mlngReturn = 0 Then
        提交_山西
    Else
        回退_山西
        Exit Function
    End If

End Function


Public Function 入院登记_山西(lngPatiID As Long, lngPageID As Long, str医保号 As String) As Boolean
'    参数      输入/输出   参数名  是否可空    长度
'CardNo           IN       卡号        16
'RegType          IN       登记类型
'                                -1  无费退院
'                                 1   入院登记
'                                 2   登记信息修改
'                                 3   出院登记
'MedType          IN       医疗类别(见附录)        3
'InHosNo          IN       住院号      15
'ApprNo           IN       审批编号        15

'Returns:
'   0 - SUCCESS
'   -1 - FAILURE
'Remarks:
'    入院登记前必须首先进行真实读卡，然后调用待遇次查检测函数，
' 如果该患者未完全封锁待遇方可进行入院住院处理。
Dim str卡号 As String
Dim str密码 As String
Dim str就诊类别 As String
Dim str入院日期 As String
Dim str病种编码 As String
Dim str病种名称 As String

'>Beging 提取病人入院信息
mstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,A.入院病床,A.住院医师,C.卡号," & _
          "C.密码,D.编码 As 科室编码,C.就诊类别 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
          "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
          "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]" & _
          " and C.险类=[3]"

Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "保险病人", lngPatiID, lngPageID, TYPE_山西)

If mrsTMP.EOF Then
    入院登记_山西 = False
    MsgBox "该病人未通过身份验证！不能办理医保入院。", vbInformation, gstrSysName
    Exit Function
End If

str密码 = mrsTMP!密码
str就诊类别 = mrsTMP!就诊类别
str卡号 = mrsTMP!卡号
str入院日期 = Format(mrsTMP!入院日期, "yyyyMMdd")
'>End

'>beging 提取病种
mstrSQL = "select * from 保险病种 where ID=(" & _
                "Select 病种ID from 保险帐户 where 病人ID=[1]" & _
                                             " and 险类=[2]) " & _
                                   " and 险类=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "病种信息", lngPatiID, TYPE_山西)

If mrsTMP.EOF Then
    MsgBox "病种目录不正确！，不能执行交易。", vbInformation, gstrSysName
    Exit Function
End If
str病种编码 = mrsTMP!编码
str病种名称 = mrsTMP!名称
'>End 提取病种

'>Beging 发起真实读卡
If 读卡审核(str密码) Then
    
    '>>beging 调入院登记
    mstrOutput = Space(600)
'    TreatInfoEntry ( char* CardNo,  //卡号
'                             int   RegType,         //登记类型(-1.无费退院 1.住院登记 2.信息修改 3. 出院登记)
'                             char* MedType,         //医疗类别
'                             char* InHosNo,         //住院门诊号
'                             char* ApprNo,          //审批编号
'                             char* TreatDate,       //入院日期(yyyymmdd)
'                             char* LeaveHosDt,      //出院日期(yyyymmdd)
'                             char* DiseaseNo,       //入院疾病编码
'                             char* DiseaseName,     //入院疾病名称
'                             char* LHDiseaseNo,     //出院疾病编码
'                             char* LHDiseaseName,   //出院疾病名称
'                             char* LHStatus,        //出院状态(1: 治愈 2: 好转 3: 死亡 9:其它)
'                             char* TrunHosKind,     //转院标志
'                             char* MainDocName,     //主管医师
'                             char* ApprPerson,      //审批人
'                             char* Transactor,      //经办人
'                             char* TransDate ,      //经办日期
'                             char* DataBuffer);     //出错信息

    mlngReturn = TreatInfoEntry(str卡号, 1, str就诊类别, lngPatiID & "_1_" & lngPageID, "", _
                         str入院日期, "", str病种编码, str病种名称, "" _
                         , "", "", 0, "", "", UserInfo.姓名, _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), mstrOutput)
                         
      '转院标志（见附录）      3 待定，目前传为0
     Call WriteBusinessLOG("TreatInfoEntry", str卡号 & ",1," & str就诊类别 & ", " & lngPatiID & "_" & lngPageID & ",," & _
                         str入院日期 & ", ," & str病种编码 & "," & str病种名称 & ", " & _
                         ", , , 0, ,," & UserInfo.姓名 & "," & _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
    '>>End 调入院登记
    If mlngReturn = -1 Then
        Call 回退_山西
        入院登记_山西 = False
        MsgBox "医保入院登记失败！" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
        Exit Function
    Else
        Call 提交_山西
        入院登记_山西 = True
        gstrSQL = "zl_保险帐户_入院(" & lngPatiID & "," & TYPE_山西 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "医保入院")
    
    End If
Else
    入院登记_山西 = False
End If
'>End 发起真实读卡


End Function

Public Function 撤销入院登记_山西(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim blnTrans As Boolean
    Dim str卡号 As String, str密码 As String, str就诊类别 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '是否存在未结费用，存在未结费用不允许撤销入院
    If 存在未结费用(lng病人ID, lng主页ID) Then
        MsgBox "该病人已发生费用，不允许撤销入院登记！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '发生费用不允许撤销入院
    gstrSQL = "Select 1 From 住院费用记录 Where 病人ID=[1] And 主页ID=[2] and Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "发生费用不允许撤销入院", lng病人ID, lng主页ID)
    If rsTemp.RecordCount <> 0 Then
        MsgBox "该病人已发生费用，不允许撤销入院登记！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '取保险病人相关信息
    gstrSQL = "Select * From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险病人相关信息", TYPE_山西, lng病人ID)
    str卡号 = rsTemp!卡号
    str密码 = Nvl(rsTemp!密码)
    str就诊类别 = Nvl(rsTemp!就诊类别, "21")
    
    '真实读卡，校验病人身份
    If Not 读卡审核(str密码) Then Exit Function
    '    TreatInfoEntry ( char* CardNo,  //卡号
    '                             int   RegType,         //登记类型(-1.无费退院 1.住院登记 2.信息修改 3. 出院登记)
    '                             char* MedType,         //医疗类别
    '                             char* InHosNo,         //住院门诊号
    '                             char* ApprNo,          //审批编号
    '                             char* TreatDate,       //入院日期(yyyymmdd)
    '                             char* LeaveHosDt,      //出院日期(yyyymmdd)
    '                             char* DiseaseNo,       //入院疾病编码
    '                             char* DiseaseName,     //入院疾病名称
    '                             char* LHDiseaseNo,     //出院疾病编码
    '                             char* LHDiseaseName,   //出院疾病名称
    '                             char* LHStatus,        //出院状态(1: 治愈 2: 好转 3: 死亡 9:其它)
    '                             char* TrunHosKind,     //转院标志
    '                             char* MainDocName,     //主管医师
    '                             char* ApprPerson,      //审批人
    '                             char* Transactor,      //经办人
    '                             char* TransDate ,      //经办日期
    '                             char* DataBuffer);     //出错信息
    blnTrans = True
    mstrOutput = Space(600)
    mlngReturn = TreatInfoEntry(str卡号, -1, str就诊类别, lng病人ID & "_1_" & lng主页ID, "", _
                         "", "", "", "", "" _
                         , "", "", 0, "", "", UserInfo.姓名, _
                        Format(zlDatabase.Currentdate, "yyyyMMddHHmmss"), mstrOutput)
                         
      '转院标志（见附录）      3 待定，目前传为0
     Call WriteBusinessLOG("TreatInfoEntry", str卡号 & ",1," & str就诊类别 & ", " & lng病人ID & "_" & lng主页ID & ",," & _
                         "" & ", ," & "" & "," & "" & ", " & _
                         ", , , 0, ,," & UserInfo.姓名 & "," & _
                        Format(zlDatabase.Currentdate, "yyyyMMddHHmmss"), Trim(mstrOutput))
    If mlngReturn = -1 Then
        MsgBox mstrOutput, vbInformation, gstrSysName
        Call 回退_山西
        Exit Function
    End If
    
    Call 提交_山西
    blnTrans = False
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_山西 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "撤销入院登记")
    
    撤销入院登记_山西 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call 回退_山西
End Function


Public Function 出院登记_山西(lngPatiID As Long, lngPageID As Long) As Boolean
'    参数      输入/输出   参数名  是否可空    长度
'CardNo           IN       卡号        16
'RegType          IN       登记类型
'                                -1  无费退院
'                                 3   出院登记
'MedType          IN       医疗类别(见附录)        3
'InHosNo          IN       住院号      15
'ApprNo           IN       审批编号        15

'Returns:
'   0 - SUCCESS
'   -1 - FAILURE
'Remarks:
'    入院登记前必须首先进行真实读卡，然后调用待遇次查检测函数，
    ' 如果该患者未完全封锁待遇方可进行入院住院处理。
    Dim bln结帐 As Boolean
    Dim blnTrans As Boolean
    Dim str卡号 As String
    Dim str密码 As String
    Dim str就诊类别 As String
    Dim str入院日期 As String
    Dim str出院日期 As String
    Dim str病种编码 As String
    Dim str病种名称 As String
    Dim str出院病种编码 As String
    Dim str出院病种名称 As String
    Dim lng病种选择情况 As Long '1 入院选择了，2入出院选择了，3入出院都没选择,
    Dim lngRegType As Long  '登记类型  (医保的参数)
    Dim str病种选择返回值 As String
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
    
    '> Beging 是否是无费退院
    lngRegType = 3
    If Not 存在未结费用(lngPatiID, lngPageID) Then
        '是否已结帐，已结帐病人不允许撤销入院
        bln结帐 = False
        gstrSQL = "Select 1 From 住院费用记录 Where 病人ID=[1] And 主页ID=[2] And Nvl(结帐ID,0)<>0 and Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否该调用就诊登记撤销", lngPatiID, lngPageID)
        If Not rsTemp.EOF Then
            bln结帐 = True
        End If
        
        If Not bln结帐 Then
            '无费退院=撤销入院登记
            lngRegType = -1
        End If
    End If
    '> End 是否是无费退院
    
    '>Beging 提取病人入院信息
    mstrSQL = "select A.入院日期,A.出院日期,B.住院号,D.名称 as 住院科室,A.入院病床,A.住院医师,C.卡号," & _
              "C.密码,D.编码 As 科室编码,C.就诊类别 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
              "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
              "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]" & _
              " and C.险类=[3]"
    Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "保险病人", lngPatiID, lngPageID, TYPE_山西)
    If mrsTMP.EOF Then
        MsgBox "该病人未通过身份验证！不能办理医保出院。", vbInformation, gstrSysName
        Exit Function
    End If
    
    str密码 = mrsTMP!密码
    str就诊类别 = mrsTMP!就诊类别
    str卡号 = mrsTMP!卡号
    str入院日期 = Format(mrsTMP!入院日期, "yyyyMMdd")
    str出院日期 = Format(mrsTMP!出院日期, "yyyyMMdd")
    '>End 提取病人入院信息
    '>
    
    '>beging 提取病种
    mstrSQL = "select * from 保险病种 where ID=(" & _
                    "Select 病种ID from 保险帐户 where 病人ID=[1]" & _
                                                 " and 险类=[2]) " & _
                                       " and 险类=[2]"
    Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "病种信息", lngPatiID, TYPE_山西)
    
    If mrsTMP.EOF Then
        lng病种选择情况 = 0
    Else
        lng病种选择情况 = 1
        str病种编码 = mrsTMP!编码
        str病种名称 = mrsTMP!名称
    End If
    
    mstrSQL = "select * from 保险病种 where ID=(" & _
                    "Select 出院病种ID from 保险帐户 where 病人ID=[1]" & _
                                                 " and 险类=[2]) " & _
                                       " and 险类=[2]"
    Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "病种信息", lngPatiID, TYPE_山西)
    If mrsTMP.EOF Then
        lng病种选择情况 = lng病种选择情况 - 2
    Else
        lng病种选择情况 = lng病种选择情况 + 2
        str出院病种编码 = mrsTMP!编码
        str出院病种名称 = mrsTMP!名称
    End If
    '>>Beging 判断是否有病种没选择，如果没有，则强制选择,然后根据返回值更新
    
    If lng病种选择情况 <> 3 Then
         If Not frm病种选择_山西.Select病种(lngPatiID, str病种编码, str病种名称, str出院病种编码, str出院病种名称) Then
             出院登记_山西 = False
             MsgBox "请选择病种后，再办理出院登记！", vbInformation, gstrSysName
             Exit Function
         End If
    End If
    '>>End 判断是否有病种没选择，如果没有，则强制选择,然后根据返回值更新
    
    '>End 提取病种
    
    '>Beging 发起虚拟读卡
    If 读卡审核(str密码) Then
        blnTrans = True
        '>>beging 调出院登记
        mstrOutput = Space(600)
        mlngReturn = TreatInfoEntry(str卡号, lngRegType, str就诊类别, lngPatiID & "_1_" & lngPageID, "", _
                             str入院日期, str出院日期, str病种编码, str病种名称, str出院病种编码 _
                             , str出院病种名称, "", 0, "", "", UserInfo.姓名, _
                            Format(zlDatabase.Currentdate, "yyyyMMdd"), mstrOutput)
                             
          '转院标志（见附录）      3 待定，目前传为0
         Call WriteBusinessLOG("TreatInfoEntry", str卡号 & "," & lngRegType & "," & str就诊类别 & ", " & lngPatiID & "_1_" & lngPageID & ",," & _
                             str入院日期 & "," & str出院日期 & "," & str病种编码 & "," & str病种名称 & ", " & str出院病种编码 & _
                             "," & str出院病种名称 & ", , 0, , ," & UserInfo.姓名 & "," & _
                            Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
        '>>End 调入院登记
        If mlngReturn = -1 Then
            Call 回退_山西
            出院登记_山西 = False
            MsgBox "医保出院登记失败！" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
            Exit Function
        Else
            Call 提交_山西
            blnTrans = False
            出院登记_山西 = True
            gstrSQL = "zl_保险帐户_出院(" & lngPatiID & "," & TYPE_山西 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
        End If
    Else
        出院登记_山西 = False
    End If
    '>End 发起虚拟读卡

    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call 回退_山西
End Function

Public Sub 更新病种_山西(lngPatiID As Long, lngPageID As Long)
Dim str入院病种编码 As String, str入院病种名称 As String
Dim str出院病种编码 As String, str出院病种名称 As String
   
   If Not frm病种选择_山西.Select病种(lngPatiID, str入院病种编码, str入院病种名称, _
    str出院病种编码, str出院病种名称) Then Exit Sub
    
End Sub



Public Function 住院虚拟结算_山西(rsExse As Recordset, ByVal lng病人ID As Long) As String
    Dim strPin As String                'IC卡密码
    Dim str就诊类别 As String           '就诊类别
    Dim str医保号 As String             '医疗证号
    Dim str住院号 As String             '住院号
    Dim lng主页ID As Long               '主页ID
    Dim blnTrans As Boolean             '是否开始医保事务
    Dim blnOut As Boolean               '病人是否已出院，决定是中途结算还是出院结算
    Dim dbl个人帐户 As Double, dbl医保基金 As Double, dbl公务员补助 As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset
    Dim cur发生费用 As Currency
    
    '明细上传相关变量
    Dim str流水号 As String, str类别 As String, str中心类别 As String
    Dim int项目类别 As Integer          '中心项目类别
    Dim dbl自付比例 As Double
    Dim str医院编码 As String, str医院名称 As String
    Dim str医保编码 As String, str医保名称 As String
    Dim str频次 As String, str用法 As String, str剂型 As String, str剂量 As String
    Dim str卡号 As String
    
    Const int费用总额 As Integer = 1
    Const int个人帐户 As Integer = 2
    Const int医保基金 As Integer = 3
    Const int公务员补助 As Integer = 5
    On Error GoTo errHand

    '读取病人IC卡密码
    gstrSQL = "Select 医保号,密码,卡号,就诊类别 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人IC卡密码", TYPE_山西, lng病人ID)
    strPin = Nvl(rsTemp!密码)
    str就诊类别 = Nvl(rsTemp!就诊类别, "11")
    str医保号 = rsTemp!医保号
    str卡号 = rsTemp!卡号
    
    '读取病人主页ID，出院日期
    gstrSQL = " Select A.主页ID,A.出院日期 From 病案主页 A,病人信息 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 and B.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人主页ID，出院日期", lng病人ID)
    blnOut = Not (IsNull(rsTemp!出院日期))
    lng主页ID = rsTemp!主页ID
    str住院号 = lng病人ID & "_1_" & lng主页ID
    
    '先读卡，检查是不是同一个病人的卡
    If Not 读卡审核(strPin, str卡号) Then Exit Function
    
    blnTrans = True
    If str医保号 <> g个人基本信息.个人编号00 Then
        MsgBox "当前IC卡不是该病人的！", vbInformation, gstrSysName
        Call 回退_山西
        Exit Function
    End If
    
    '取得总费用
    Do Until rsExse.EOF
        cur发生费用 = cur发生费用 + rsExse("金额")
        rsExse.MoveNext
    Loop
    cur发生费用 = Val(Format(cur发生费用, "#####0.00"))

    '上传未上传的费用明细
    gstrSQL = "  Select A.ID,A.NO,A.病人ID,A.收费类别,A.记录性质,A.记录状态,A.序号,A.收费细目ID,C.项目编码 AS 医保项目编码,B.编码,B.名称,A.实收金额 AS 金额" & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格,A.开单人 AS 医生,A.登记时间,E.名称 AS 开单部门 " & _
              "  From 住院费用记录 A,收费细目 B,保险支付项目 C,部门表 E" & _
              "  Where A.病人ID=[1] and A.主页ID=[2] and A.记帐费用=1 And A.操作员姓名 is not null AND A.实收金额 IS NOT NULL " & _
              "        and nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= " & TYPE_山西 & _
              "        and A.开单部门ID=E.ID " & _
              "  Order by A.病人ID,A.发生时间"
    Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次费用明细", lng病人ID, lng主页ID)
    With rsDetail
        Do While Not .EOF
            '入参说明
            'DLLFUNC int WINAPI FormularyEntry(
            '    char* CardNo,          //卡号          '    char* InHosNo,         //住院门诊号            '    int TransKind,         //交易类型          '    int ItemKind,          //项目类别
            '    char* InternalCode,    //收费项目医院内码          '    char* FormularyNo,     //处方号            '    char* SysDate,         //开方日期(yyyymmdd)            '    char* CenterCode,      //收费项目中心编码
            '    char* ItemName,        //收费项目名称          '    double UnitPrice,      //单价          '    double Quantity,       //数量          '    double Amount,         //金额
            '    char* DoseType,        //剂型          '    char* Dosage,          //剂量          '    char* Frequency,       //频次          '    char* Usage,           //用法
            '    char* KeBie,           //科别          '    float ExecDays,        //执行天数          '    char* FeeType,         //医保中心收费类别          '    char* DoctName,        //开方医生
            '    char* Transactor,      //经办人            '    char* ApprPerson,      //审批人            '    int  IsOwnExpenses,    //全额自费标志          '    char* DataBuffer   );

            '提取收费项目信息
            gstrSQL = " Select A.类别,A.ID AS 收费细目ID,A.编码 As 医院编码,A.名称 AS 医院名称,B.项目编码 As 医保编码,B.项目名称 AS 医保名称,B.附注" & _
                    " From 收费细目 A,(Select * From 保险支付项目 Where 险类=[1]) B" & _
                    " Where A.ID=[2] And A.ID=B.收费细目ID(+) "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取收费项目信息", TYPE_山西, CLng(!收费细目ID))
            If IsNull(rsTemp!医保编码) Then
                MsgBox "项目[" & rsTemp!医院编码 & "]" & rsTemp!医院名称 & "未对码！（保险项目管理）", vbInformation, gstrSysName
                Call 提交_山西
                Exit Function
            End If
            str类别 = rsTemp!类别
            int项目类别 = Nvl(rsTemp!附注, 0)
            str医院编码 = rsTemp!医院编码: str医院名称 = rsTemp!医院名称
            str医保编码 = Nvl(rsTemp!医保编码): str医保名称 = Nvl(rsTemp!医保名称)

            '提取项目的中心类别
            If int项目类别 = 1 Then
                gstrSQL = "Select aka063 类别,aka069 自付比例 From ka02 where aka060='" & str医保编码 & "'"
            ElseIf int项目类别 = 2 Then
                gstrSQL = "Select aka063 类别,aka069 自付比例 From ka03 where aka090='" & str医保编码 & "'"
            Else
                gstrSQL = "Select aka063 类别,0 自付比例 From ka04 where aka100='" & str医保编码 & "'"
            End If
            Call OpenRecordset_OtherBase(rsTemp, "提取中心项目比例、类别", gstrSQL, gcnSxDr)
            If rsTemp.RecordCount = 0 Then
                MsgBox "请重新为项目[" & str医院编码 & "]" & str医院名称 & "进行对码，中心项目已删除或取消！", vbInformation, gstrSysName
                Call 提交_山西
                Exit Function
            End If
            dbl自付比例 = Nvl(rsTemp!自付比例, 0)
            str中心类别 = rsTemp!类别
            
            '如果是药品，则提取该药品相关信息（频次、用法、剂型、剂量）
            str频次 = "": str用法 = "": str剂型 = "": str剂量 = ""
            If InStr(1, ",5,6,7,", rsTemp!类别) <> 0 Then
                gstrSQL = " Select A.频次,A.用法,D.名称 AS 剂型,C.剂量单位 " & _
                        " From (" & _
                        "   Select * from 药品收发记录 " & _
                        "   Where 单据 in (9,10) and NO=[2]) A,药品目录 B,药品信息 C,药品剂型 D" & _
                        "   Where A.费用ID=[1] And A.药品ID=B.药品ID And B.药名ID=C.药名ID And C.剂型=D.编码"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取药品相关信息", CLng(rsTemp!ID), CStr(rsTemp!NO))
                str频次 = Nvl(rsTemp!频次)
                str用法 = Nvl(rsTemp!用法)
                str剂型 = Nvl(rsTemp!剂型)
                str剂量 = Nvl(rsTemp!剂量单位)
            End If

            mstrOutput = Space(600)
            str流水号 = !记录性质 & !NO & 1 & !序号
            mlngReturn = FormularyEntry(g个人基本信息.卡号05, str住院号, IIf(!记录状态 = 2, -1, 1), int项目类别, _
                    str医院编码, str流水号, Format(!登记时间, "yyyyMMdd"), str医保编码, _
                    str医保名称, !价格, Abs(!数量), Abs(!金额), _
                    str剂型, str剂量, str频次, str用法, _
                    !开单部门, 0, str中心类别, Nvl(!医生), _
                    UserInfo.姓名, "", IIf(dbl自付比例 = 1, 1, 0), mstrOutput)
            Call WriteBusinessLOG("FormularyEntry", g个人基本信息.卡号05 & "," & str住院号 & "," & IIf(!记录状态 = 2, -1, 1) & "," & int项目类别 & "," & _
                    str医院编码 & "," & str流水号 & "," & Format(!登记时间, "yyyyMMdd") & "," & str医保编码 & "," & _
                    str医保名称 & "," & !价格 & "," & Abs(!数量) & "," & Abs(!金额) & "," & _
                    str剂型 & "," & str剂量 & "," & str频次 & "," & str用法 & "," & _
                    !开单部门 & "," & 0 & "," & str中心类别 & "," & Nvl(!医生) & "," & _
                    UserInfo.姓名 & ",," & IIf(dbl自付比例 = 1, 1, 0), mstrOutput)
            If mlngReturn = -1 Then
                MsgBox "上传处方[" & !NO & "]第" & !序号 & "行明细时出错！" & vbCrLf & "详细信息：" & vbCrLf & mstrOutput, vbInformation, gstrSysName
                Call 提交_山西
                Exit Function
            End If

            '更新费用明细中的统筹金额等信息
            gstrSQL = "ZL_病人费用记录_更新医保(" & !ID & "," & _
                    Split(mstrOutput, "|")(1) - Split(mstrOutput, "|")(3) - Split(mstrOutput, "|")(4) & _
                    ",NULL,1,'" & str医保编码 & "',1,'NULL')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新医保字段")
            .MoveNext
        Loop
    End With

    '住院预结算
    'DLLFUNC int WINAPI PreExpenseCalc ( char* CardNo,   //卡号,char* InHosNo,          //住院门诊号 ,char* MedType,            //医疗类别,char* DataBuffer );  //结算信息
    '返回参数说明
    '|费用总额|个人帐户支付|统筹支付|现金支付|公务员补助支付
    '|自理金额|自费金额|住院人次|起付标准|转院连计费用
    '|起付标准自付|起付标准公务员支付|分段1统筹支付|分段1公务员支付|分段1个人自付
    '|分段2统筹支付|分段2公务员支付|分段2个人自付|分段3统筹支付|分段3公务员支付
    '|分段3个人自付|超封顶公务员支付|公务员个人自付|超公务员封顶个人自付|隶属关系|单位类型|
    mstrOutput = Space(600)
    mlngReturn = PreExpenseCalc(g个人基本信息.卡号05, str住院号, str就诊类别, mstrOutput)
    Call WriteBusinessLOG("PreExpenseCalc", g个人基本信息.卡号05 & "," & str住院号 & "," & str就诊类别, mstrOutput)
    
    If mlngReturn = -1 Then
        MsgBox mstrOutput, vbInformation, gstrSysName
        Call 提交_山西
        Exit Function
    Else
        If cur发生费用 <> Val(Format(Split(mstrOutput, "|")(1), "#####0.00")) Then
            If MsgBox("医院的费用总金额(" & cur发生费用 & ")与医保中心的费用总额(" & Val(Format(Split(mstrOutput, "|")(1), "#####0.00")) & ")不等，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Call 提交_山西
                Exit Function
            End If
        End If

    End If

    '取结算信息
    dbl个人帐户 = Val(Split(mstrOutput, "|")(int个人帐户))
    dbl医保基金 = Val(Split(mstrOutput, "|")(int医保基金))
    dbl公务员补助 = Val(Split(mstrOutput, "|")(int公务员补助))

    '返回结算信息
    blnTrans = False
    住院虚拟结算_山西 = "个人帐户;" & dbl个人帐户 & ";1"
    住院虚拟结算_山西 = 住院虚拟结算_山西 & "|统筹基金;" & dbl医保基金 & ";1"
    住院虚拟结算_山西 = 住院虚拟结算_山西 & "|公务员补助;" & dbl公务员补助 & ";1"
    
    Call 提交_山西
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then Call 提交_山西
End Function

Public Function 住院结算_山西(ByVal lng结帐ID As Long) As Boolean
    Dim strNO As String, strPin As String
    Dim str就诊类别 As String, str医保号 As String, str住院号 As String
    Dim dbl个人帐户 As Double, dbl医保基金 As Double, dbl公务员补助 As Double, dbl总费用 As Double, dbl现金 As Double
    Dim lng病人ID As Long, lng主页ID As Long
    Dim blnTrans As Boolean
    Dim blnOut As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    Const int费用总额 As Integer = 1
    Const int个人帐户 As Integer = 2
    Const int医保基金 As Integer = 3
    Const int现金 As Integer = 4
    Const int公务员补助 As Integer = 5
    On Error GoTo errHand
    
    '提取结帐单号
    gstrSQL = "Select NO,病人ID From 病人结帐记录 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取结帐单号", lng结帐ID)
    strNO = "2" & rsTemp!NO
    lng病人ID = rsTemp!病人ID
    
    '提取帐户实际支付额
    gstrSQL = "Select Nvl(冲预交,0) AS 个人帐户 From 病人预交记录 Where 结帐ID=[1] And 记录性质 Not In (1,11) And 结算方式='个人帐户'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取帐户实际支付额", lng结帐ID)
    If Not rsTemp.EOF Then
        dbl个人帐户 = Nvl(rsTemp!个人帐户, 0)
    End If

    '读取病人IC卡密码
    gstrSQL = "Select 医保号,密码,就诊类别 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人IC卡密码", TYPE_山西, lng病人ID)
    strPin = Nvl(rsTemp!密码)
    str就诊类别 = Nvl(rsTemp!就诊类别, "11")
    str医保号 = rsTemp!医保号

    '读取病人主页ID，出院日期
    gstrSQL = " Select A.主页ID,A.出院日期 From 病案主页 A,病人信息 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 and B.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人主页ID，出院日期", lng病人ID)
    blnOut = Not (IsNull(rsTemp!出院日期))
    lng主页ID = rsTemp!主页ID
    str住院号 = lng病人ID & "_1_" & lng主页ID

    '先读卡，检查是不是同一个病人的卡
    If Not 读卡审核(strPin) Then Exit Function
    blnTrans = True
    If str医保号 <> g个人基本信息.个人编号00 Then
        Err.Raise 9000, gstrSysName, "当前IC卡不是该病人的！"
        Call 回退_山西
        Exit Function
    End If

    '入参说明
    'DLLFUNC int WINAPI ExpenseCalc(
        'char* CardNo,          //卡号      'int   TransType,       //交易类型      'int   InvoiceKind,     //发票类型 0: 门诊 1:住院
        'char* InHosNo,         //住院门诊号        'char* MedType,         //医疗类别      'char* InvoiceNo,       //单据号
        'char* UserName,        //经办人        'double AccCashPay,     //帐户支付金额      'char* DataBuffer );    //结算结果
    mstrOutput = Space(600)
    mlngReturn = ExpenseCalc(g个人基本信息.卡号05, IIf(blnOut, 1, 2), 1, _
        str住院号, str就诊类别, strNO, _
        UserInfo.姓名, dbl个人帐户, mstrOutput)
    Call WriteBusinessLOG("ExpenseCalc", g个人基本信息.卡号05 & "," & IIf(blnOut, 1, 2) & ",1," & _
        str住院号 & "," & str就诊类别 & "," & strNO & "," & _
        UserInfo.姓名 & "," & dbl个人帐户, mstrOutput)
    If mlngReturn = -1 Then
        Err.Raise 9000, gstrSysName, mstrOutput
        Call 回退_山西
        Exit Function
    End If

    '取结算信息
    dbl总费用 = Val(Split(mstrOutput, "|")(int费用总额))
    dbl现金 = Val(Split(mstrOutput, "|")(int现金))
    dbl医保基金 = Val(Split(mstrOutput, "|")(int医保基金))
    dbl公务员补助 = Val(Split(mstrOutput, "|")(int公务员补助))
    
    '费用总额比较，
    
    Call 提交_山西
    blnTrans = False
    
    '保存保险结算记录
    '大病自付=大病补助;超限自付=公务员补助
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_山西 & "," & lng病人ID & "," & _
        Year(zlDatabase.Currentdate()) & ",0,0,0,0,0,0,0,0," & dbl总费用 & "," & dbl现金 & ",0," & _
        dbl医保基金 & "," & dbl医保基金 & ",0," & dbl公务员补助 & "," & dbl个人帐户 & ",'" & str住院号 & "'," & lng主页ID & "," & IIf(blnOut, 1, 0) & ",'" & str就诊类别 & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    住院结算_山西 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then Call 回退_山西
End Function


Public Function 出院登记撤销_山西(lng病人ID As Long, lng主页ID As Long) As Boolean

Dim str卡号 As String
Dim str密码 As String
Dim str就诊类别 As String
Dim str入院日期 As String
Dim str出院日期 As String
Dim str病种编码 As String
Dim str病种名称 As String
Dim str出院病种编码 As String
Dim str出院病种名称 As String
Dim lng病种选择情况 As Long '1 入院选择了，2入出院选择了，3入出院都没选择,
Dim lngRegType As Long  '登记类型  (医保的参数)
Dim str病种选择返回值 As String

'//////////////////////////  出院登记，如果无费用，调用无费退费，改为非医保,不管是否在院，
'                              如有费用，并且已出院，调出院登记
lngRegType = 2   '设标志为修改记录

'>Beging 提取病人入院信息
mstrSQL = "select A.入院日期,A.出院日期,B.住院号,D.名称 as 住院科室,A.入院病床,A.住院医师,C.卡号," & _
          "C.密码,D.编码 As 科室编码,C.就诊类别 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
          "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
          "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]" & _
          " and C.险类=[3]"

Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "保险病人", lng病人ID, lng主页ID, TYPE_山西)
If mrsTMP.EOF Then
    出院登记撤销_山西 = False
    MsgBox "该病人未通过身份验证！不能办理撤消出院。", vbInformation, gstrSysName
    Exit Function
End If

str密码 = mrsTMP!密码
str就诊类别 = mrsTMP!就诊类别
str卡号 = mrsTMP!卡号
str入院日期 = Format(mrsTMP!入院日期, "yyyyMMdd")
str出院日期 = "" '撤消出院，所以出院日期改为空
'>End 提取病人入院信息



'>

'>beging 提取病种
mstrSQL = "select * from 保险病种 where ID=(" & _
                "Select 病种ID from 保险帐户 where 病人ID=[1]" & _
                                             " and 险类=[2]) " & _
                                   " and 险类=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "病种信息", lng病人ID, TYPE_山西)

If mrsTMP.EOF Then
    lng病种选择情况 = 0
Else
    lng病种选择情况 = 1
    str病种编码 = mrsTMP!编码
    str病种名称 = mrsTMP!名称
End If

mstrSQL = "select * from 保险病种 where ID=(" & _
                "Select 出院病种ID from 保险帐户 where 病人ID=[1]" & _
                                             " and 险类=[2]) " & _
                                   " and 险类=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "病种信息", lng病人ID, TYPE_山西)
If mrsTMP.EOF Then
    lng病种选择情况 = lng病种选择情况 - 2
Else
    lng病种选择情况 = lng病种选择情况 + 2
    str出院病种编码 = mrsTMP!编码
    str出院病种名称 = mrsTMP!名称
End If
'>>Beging 判断是否有病种没选择，如果没有，则强制选择,然后根据返回值更新

If lng病种选择情况 <> 3 Then
     If Not frm病种选择_山西.Select病种(lng病人ID, str病种编码, str病种名称, str出院病种编码, str出院病种名称) Then
         出院登记撤销_山西 = False
         MsgBox "请选择病种后，再办理出院登记！", vbInformation, gstrSysName
         Exit Function
     End If
End If
'>>End 判断是否有病种没选择，如果没有，则强制选择,然后根据返回值更新

'>End 提取病种

'>Beging 发起虚拟读卡 医保必须真实读卡
If 读卡审核(str密码) Then
    
    '>>beging 调接口
    mstrOutput = Space(600)
    mlngReturn = TreatInfoEntry(str卡号, lngRegType, str就诊类别, lng病人ID & "_1_" & lng主页ID, "", _
                         str入院日期, str出院日期, str病种编码, str病种名称, str出院病种编码 _
                         , str出院病种名称, "", 0, "", "", UserInfo.姓名, _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), mstrOutput)
                         
      '转院标志（见附录）      3 待定，目前传为0
     Call WriteBusinessLOG("TreatInfoEntry", str卡号 & "," & lngRegType & "," & str就诊类别 & ", " & lng病人ID & "_1_" & lng主页ID & ",," & _
                         str入院日期 & "," & str出院日期 & "," & str病种编码 & "," & str病种名称 & ", " & str出院病种编码 & _
                         "," & str出院病种名称 & ", , 0, , ," & UserInfo.姓名 & "," & _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
    '>>End 调接口
    If mlngReturn = -1 Then
        Call 回退_山西
        出院登记撤销_山西 = False
        MsgBox "医保撤消出院失败！" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
        Exit Function
    Else
        Call 提交_山西
        出院登记撤销_山西 = True
        gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_山西 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "撤消出院登记")
    End If
Else
    出院登记撤销_山西 = False
End If
'>End 发起虚拟读卡

End Function


Public Function 记帐上传_山西(int性质 As Integer, int状态 As Integer, str单据号 As String) As Boolean

    Dim strPin As String                'IC卡密码
    Dim str就诊类别 As String           '就诊类别
    'Dim str医保号 As String             '医疗证号
    Dim str住院号 As String             '住院号
    Dim lng主页ID As Long               '主页ID
    Dim blnOut As Boolean               '病人是否已出院，决定是中途结算还是出院结算
    Dim lng病人ID As Long               '病人ID
    Dim dbl个人帐户 As Double, dbl医保基金 As Double, dbl公务员补助 As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset
    Dim rsCd As New ADODB.Recordset
    
    '明细上传相关变量
    Dim str流水号 As String, str类别 As String, str中心类别 As String
    Dim int项目类别 As Integer          '中心项目类别
    Dim dbl自付比例 As Double
    Dim str医院编码 As String, str医院名称 As String
    Dim str医保编码 As String, str医保名称 As String
    Dim str频次 As String, str用法 As String, str剂型 As String, str剂量 As String
    
    Const int费用总额 As Integer = 1
    Const int个人帐户 As Integer = 2
    Const int医保基金 As Integer = 3
    Const int公务员补助 As Integer = 5
    On Error GoTo errHand
    
    ' 如果记录状态为1的单据，有负数记录，则不允许保存单据
    If int状态 = 1 Then
        gstrSQL = "Select distinct  A.病人ID from 住院费用记录 A,保险帐户 B " & _
            "where A.病人ID=B.病人ID And A.记录性质=[1]" & _
            " And A.记录状态=[2] And A.NO=[3] " & _
            " And B.险类=[4] And A.实收金额<0"
        Set rsCd = zlDatabase.OpenSQLRecord(gstrSQL, "是否有负数记录", int性质, int状态, str单据号, TYPE_山西)
        If Not rsCd.EOF Then
            MsgBox "本医保不支持负数记录！请调整。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '根据NO号,提取病人ID
    gstrSQL = "Select distinct  A.病人ID from 住院费用记录 A,保险帐户 B " & _
            "where A.病人ID=B.病人ID And A.记录性质=[1]" & _
            " And A.记录状态=[2] And A.NO=[3]" & _
            " And B.险类=[4]"
    Set rsCd = zlDatabase.OpenSQLRecord(gstrSQL, "取病人ID", int性质, int状态, str单据号, TYPE_山西)
    
    
    记帐上传_山西 = True
    '> Beging 按病人传明细.
    Do Until rsCd.EOF
        '>> Beging 读卡审核
        lng病人ID = rsCd!病人ID
        gstrSQL = "Select * from 保险帐户 where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取卡号", lng病人ID)
        If 读卡审核(rsTemp!密码, rsTemp!卡号) Then
            gstrSQL = "Select A.*,nvl(A.付数,1)*nvl(A.数次,0) as 数量,A.实收金额 as 金额," & _
                              "nvl(A.实收金额,0)/(nvl(A.付数,1)*nvl(A.数次,0)) as 价格,A.开单人 as 医生,C.名称 as 开单部门,B.* " & _
                      " from 住院费用记录 A,保险帐户 B,部门表 C" & _
                      " where A.NO=[1]" & _
                            " And A.记录性质=[2]" & _
                            " And A.记录状态=[3]" & _
                            " And nvl(A.是否上传,0)=0 " & _
                            " And A.病人ID=B.病人ID " & _
                            " and B.险类=[4]" & _
                            " and A.开单部门ID=C.ID " & _
                            " ANd A.病人ID=[5]"
        
            Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次费用明细", str单据号, int性质, int状态, TYPE_山西, lng病人ID)
            With rsDetail
                Do While Not .EOF
                    '入参说明
                    'DLLFUNC int WINAPI FormularyEntry(
                    '    char* CardNo,          //卡号          '    char* InHosNo,         //住院门诊号            '    int TransKind,         //交易类型          '    int ItemKind,          //项目类别
                    '    char* InternalCode,    //收费项目医院内码          '    char* FormularyNo,     //处方号            '    char* SysDate,         //开方日期(yyyymmdd)            '    char* CenterCode,      //收费项目中心编码
                    '    char* ItemName,        //收费项目名称          '    double UnitPrice,      //单价          '    double Quantity,       //数量          '    double Amount,         //金额
                    '    char* DoseType,        //剂型          '    char* Dosage,          //剂量          '    char* Frequency,       //频次          '    char* Usage,           //用法
                    '    char* KeBie,           //科别          '    float ExecDays,        //执行天数          '    char* FeeType,         //医保中心收费类别          '    char* DoctName,        //开方医生
                    '    char* Transactor,      //经办人            '    char* ApprPerson,      //审批人            '    int  IsOwnExpenses,    //全额自费标志          '    char* DataBuffer   );
        
                    '提取收费项目信息
                    gstrSQL = " Select A.类别,A.ID AS 收费细目ID,A.编码 As 医院编码,A.名称 AS 医院名称,B.项目编码 As 医保编码,B.项目名称 AS 医保名称,B.附注" & _
                            " From 收费细目 A,(Select * From 保险支付项目 Where 险类=[1]) B" & _
                            " Where A.ID=[2] And A.ID=B.收费细目ID(+) "
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取收费项目信息", TYPE_山西, CLng(!收费细目ID))
                    If IsNull(rsTemp!医保编码) Then
                        MsgBox "项目[" & rsTemp!医院编码 & "]" & rsTemp!医院名称 & "未对码！（保险项目管理）", vbInformation, gstrSysName
                        Call 提交_山西
                        Exit Function
                    End If
                    str类别 = rsTemp!类别
                    int项目类别 = Nvl(rsTemp!附注, 0)
                    str医院编码 = rsTemp!医院编码: str医院名称 = rsTemp!医院名称
                    str医保编码 = Nvl(rsTemp!医保编码): str医保名称 = Nvl(rsTemp!医保名称)
        
                    '提取项目的中心类别
                    If int项目类别 = 1 Then
                        gstrSQL = "Select aka063 类别,aka069 自付比例 From ka02 where aka060='" & str医保编码 & "'"
                    ElseIf int项目类别 = 2 Then
                        gstrSQL = "Select aka063 类别,aka069 自付比例 From ka03 where aka090='" & str医保编码 & "'"
                    Else
                        gstrSQL = "Select aka063 类别,0 自付比例 From ka04 where aka100='" & str医保编码 & "'"
                    End If
                    Call OpenRecordset_OtherBase(rsTemp, "提取中心项目比例、类别", gstrSQL, gcnSxDr)
                    If rsTemp.RecordCount = 0 Then
                        MsgBox "请重新为项目[" & str医院编码 & "]" & str医院名称 & "进行对码，中心项目已删除或取消！", vbInformation, gstrSysName
                        Call 提交_山西
                        Exit Function
                    End If
                    dbl自付比例 = Nvl(rsTemp!自付比例, 0)
                    str中心类别 = rsTemp!类别
                    
                    '如果是药品，则提取该药品相关信息（频次、用法、剂型、剂量）
                    str频次 = "": str用法 = "": str剂型 = "": str剂量 = ""
                    If InStr(1, ",5,6,7,", rsTemp!类别) <> 0 Then
                        gstrSQL = " Select A.频次,A.用法,D.名称 AS 剂型,C.剂量单位 " & _
                                " From (" & _
                                "   Select * from 药品收发记录 " & _
                                "   Where 单据 in (9,10) and NO='" & rsTemp!NO & "') A,药品目录 B,药品信息 C,药品剂型 D" & _
                                "   Where A.费用ID=" & rsTemp!ID & " And A.药品ID=B.药品ID And B.药名ID=C.药名ID And C.剂型=D.编码"
                        Call OpenRecordset(rsTemp, "提取药品相关信息")
                        str频次 = Nvl(rsTemp!频次)
                        str用法 = Nvl(rsTemp!用法)
                        str剂型 = Nvl(rsTemp!剂型)
                        str剂量 = Nvl(rsTemp!剂量单位)
                    End If
        
        
                    mstrOutput = Space(600)
                    str流水号 = !记录性质 & !NO & 1 & !序号
                    str住院号 = !病人ID & "_1_" & !主页ID
                    mlngReturn = FormularyEntry(g个人基本信息.卡号05, str住院号, IIf(!记录状态 = 2, -1, 1), int项目类别, _
                            str医院编码, str流水号, Format(!登记时间, "yyyyMMdd"), str医保编码, _
                            str医保名称, !价格, Abs(!数量), Abs(!金额), _
                            str剂型, str剂量, str频次, str用法, _
                            !开单部门, 0, str中心类别, Nvl(!医生), _
                            UserInfo.姓名, "", IIf(dbl自付比例 = 1, 1, 0), mstrOutput)
                    Call WriteBusinessLOG("FormularyEntry", g个人基本信息.卡号05 & "," & str住院号 & "," & IIf(!记录状态 = 2, -1, 1) & "," & int项目类别 & "," & _
                            str医院编码 & "," & str流水号 & "," & Format(!登记时间, "yyyyMMdd") & "," & str医保编码 & "," & _
                            str医保名称 & "," & !价格 & "," & Abs(!数量) & "," & Abs(!金额) & "," & _
                            str剂型 & "," & str剂量 & "," & str频次 & "," & str用法 & "," & _
                            !开单部门 & "," & 0 & "," & str中心类别 & "," & Nvl(!医生) & "," & _
                            UserInfo.姓名 & ",," & IIf(dbl自付比例 = 1, 1, 0), mstrOutput)
                    If mlngReturn = -1 Then
                        MsgBox "上传处方[" & !NO & "]第" & !序号 & "行明细时出错！" & vbCrLf & "详细信息：" & vbCrLf & mstrOutput, vbInformation, gstrSysName
                        Call 提交_山西
                        Exit Function
                    End If
        
                    '更新费用明细中的统筹金额等信息
                    
                    gstrSQL = "ZL_病人费用记录_更新医保(" & !ID & "," & _
                            Split(mstrOutput, "|")(1) - Split(mstrOutput, "|")(3) - Split(mstrOutput, "|")(4) & _
                            ",NULL,1,'" & str医保编码 & "',1,NULL)"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新医保字段")
                    
                    '>>>更改上传标志
                    
                    .MoveNext
                Loop
            End With
            Call 提交_山西
        End If
        '>> End 读卡审核
        
        rsCd.MoveNext
    Loop
    '> END 按病人传明细.
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    Call 提交_山西
Resume
End Function


Public Function 住院结算冲销_山西(lng结帐ID As Long) As Boolean
    Dim lng冲销ID As Long
    Dim str卡号 As String
    Dim str密码 As String
    Dim StrInput As String
    Dim rsTemp As New ADODB.Recordset
    Dim str住院号 As String  '医保住院号,格式为 病人ID_1_主页ID
    Dim str医疗类别  As String
    Dim strNO As String
    Dim lng病人ID As Long, lng主页ID As Long
    Dim dbl个人帐户 As Double
    
    On Error GoTo errHand
    
    '取结算流水号
    gstrSQL = "Select * From 保险结算记录 Where 性质=2 And 记录ID=" & lng结帐ID & " and 险类=" & TYPE_山西
    Call OpenRecordset(rsTemp, "取结算流水号")
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "没有找到原始结算记录，无法进行住院结算冲销！"
        Exit Function
    End If
    
    str医疗类别 = Nvl(rsTemp!备注)
    lng病人ID = rsTemp!病人ID
    
    '读取病人主页ID，出院日期
    gstrSQL = " Select A.主页ID,A.出院日期 From 病案主页 A,病人信息 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 and B.病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "读取病人主页ID，出院日期")
    
    lng主页ID = rsTemp!主页ID
    str住院号 = lng病人ID & "_1_" & lng主页ID
    
    If str住院号 = "" Or str医疗类别 = "" Then
        Err.Raise 9000, gstrSysName, "原始结算记录交易号不全，无法进行门诊结算冲销！"
        住院结算冲销_山西 = False
        Exit Function
    End If
    
    '取卡证号码
    gstrSQL = "Select 卡号,密码 From 保险帐户 Where 险类=" & TYPE_山西 & " And 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "读卡号密码")
    str卡号 = Nvl(rsTemp!卡号)
    str密码 = Nvl(rsTemp!密码)
    
    '取冲销记录的结帐ID，单据号
    gstrSQL = "select distinct A.NO,A.ID from 病人结帐记录 A,病人结帐记录 B where A.NO=B.NO and  A.记录状态=2 and B.ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "读新产生的结帐ID")
    lng冲销ID = rsTemp!ID
    strNO = "2" & rsTemp!NO
    
    
        '提取帐户实际支付额
    gstrSQL = "Select Nvl(冲预交,0) AS 个人帐户 From 病人预交记录 Where 结帐ID=" & lng结帐ID & " And 记录性质 Not In (1,11) And 结算方式='个人帐户'"
    Call OpenRecordset(rsTemp, "提取帐户实际支付额")
    If Not rsTemp.EOF Then
        dbl个人帐户 = Nvl(rsTemp!个人帐户, 0)
    End If

    ''真实读卡  根据情况，可能要加入密码录入窗口,现在暂不处理
    
    If 读卡审核(str密码) = False Then
        住院结算冲销_山西 = False
        Exit Function
    End If
    
  
    '调用结算冲销
    mstrOutput = Space(600)
    mlngReturn = ExpenseCalc(str卡号, -2, 1, str住院号, str医疗类别, _
                                   strNO, UserInfo.姓名, dbl个人帐户, mstrOutput)
    Call WriteBusinessLOG("ExpenseCalc", str卡号 & ", -2, 1," & str住院号 & "," & str医疗类别 & "," & _
                                   strNO & "," & UserInfo.姓名 & "," & dbl个人帐户, Trim(mstrOutput))
    '成功后,保存本次结算情况
    
    If mlngReturn = 0 Then
        Call 提交_山西
        gstrSQL = "Select * From 保险结算记录 Where 性质=2 And 记录ID=" & lng结帐ID & " and 险类=" & TYPE_山西
        Call OpenRecordset(rsTemp, "取结算信息")
        
        gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_山西 & "," & lng病人ID & "," & _
            Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
            0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
            -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & "," & -1 * Nvl(rsTemp!首先自付金额, 0) & "," & -1 * Nvl(rsTemp!进入统筹金额, 0) & "," & -1 * Nvl(rsTemp!统筹报销金额, 0) & ",0,0," & _
            -1 * Nvl(rsTemp!个人帐户支付, 0) & ",'" & Nvl(rsTemp!支付顺序号) & "'," & lng主页ID & "," & rsTemp!中途结帐 & ",'" & Nvl(rsTemp!备注) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "住院结算冲销")
        
        住院结算冲销_山西 = True
    Else
        Call 回退_山西
        住院结算冲销_山西 = False
        Err.Raise 9000, gstrSysName, "退费交易失败！" & vbCrLf & Trim(mstrOutput)
    End If
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 修改密码_山西(strOldPwd As String, strNewPwd As String) As Boolean
    修改密码_山西 = False
    mstrOutput = Space(600)
    
    mlngReturn = ChangePinEx(strOldPwd, strNewPwd, mstrOutput)
    If mlngReturn = 0 Then
        修改密码_山西 = True
        MsgBox "密码修改成功!" & vbCrLf, vbInformation, gstrSysName
    Else
        修改密码_山西 = False
        MsgBox "密码修改失败!" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
    End If
End Function


