Attribute VB_Name = "mdl冕宁"
Option Explicit
'//根据医院系统用户编号查询医保系统用户编号
Public Declare Function finduser Lib "YBXT.dll" (ByVal yyid As String) As Boolean
'//新增用户信息
Public Declare Function inuser Lib "YBXT.dll" (ByVal yyid As String, ByVal user_name As String) As Double
'//修改用户信息
Public Declare Function edituser Lib "YBXT.dll" (ByVal yyid As String, ByVal user_name As String) As Double
'//删除用户信息
Public Declare Function deluser Lib "YBXT.dll" (ByVal yyid As String) As Double
'//根据医院系统科别编号查询医保系统科别编号
Public Declare Function findkb Lib "YBXT.dll" (ByVal yykbbh As String) As Boolean
'//新增科别信息
Public Declare Function inkb Lib "YBXT.dll" (ByVal yykbbh As String, ByVal mc As String) As Double
'//修改科别信息
Public Declare Function editkb Lib "YBXT.dll" (ByVal yykbbhas As String, ByVal mc As String) As Double
'//删除科别信息
Public Declare Function delkb Lib "YBXT.dll" (ByVal yykbbh As String) As Double
'//根据医院系统医生编号查询医保系统医生编号
Public Declare Function findys Lib "YBXT.dll" (ByVal yyysbh As String) As Boolean
'//新增医生信息
Public Declare Function inys Lib "YBXT.dll" (ByVal yyysbh As String, ByVal yykbbh As String, ByVal xm As String) As Double
'//修改医生信息
Public Declare Function editys Lib "YBXT.dll" (ByVal yyysbh As String, ByVal yykbbh As String, ByVal xm As String) As Double
'//删除医生信息
Public Declare Function delys Lib "YBXT.dll" (ByVal yyysbh As String) As Double
'//根据医院项目DM查询医保项目DM
Public Declare Function findxm Lib "YBXT.dll" (ByVal dm As String) As String
'//新增项目
Public Declare Function inxm Lib "YBXT.dll" (ByVal yyxmdm As String, ByVal yyfldm As String, ByVal xmzl As String, _
                                            ByVal zlxmmc As String, ByVal zlflmc As String, ByVal pybh As String, _
                                            ByVal dj As String, ByVal jldw As String) As Double
'//修改项目
Public Declare Function editxm Lib "YBXT.dll" (ByVal yyxmdm As String, ByVal zlxmmc As String, ByVal pybh As String, _
                                              ByVal dj As String, ByVal jldw As String) As Double
'//删除项目
Public Declare Function delxm Lib "YBXT.dll" (ByVal yyxmdm As String) As Double
'//查询指定项目对应的医保项目信息。
Public Declare Function xmcx Lib "YBXT.dll" (ByVal yyxmdm As String) As Double

'//入院初始化
Public Declare Function rycsh Lib "YBXT.dll" (ByRef sbh As String, ByRef zhzt As String, ByRef zffs As String, _
                                             ByRef net As String, ByRef Zhye As Double, ByRef tcqfx As Double) As Double
'//支付分配[入院、补缴费用]
Public Declare Function zffp Lib "YBXT.dll" (ByVal zhzt As String, ByVal yjje As Double, ByVal Zhye As Double) As Double
'//写入院表
Public Declare Function ryappend Lib "YBXT.dll" (ByVal sbh As String, ByVal yyzyh As String, ByVal yykb As String, _
                                                ByVal yyzdys As String, ByVal yysfydm As String, ByVal ryzd As String, _
                                                ByVal lxdh As String, ByVal jtzz As String, ByVal Bz As String, _
                                                ByVal net As String, ByVal yjje As Double, ByVal zhyjje As Double) As Double
'床位添加
Public Declare Function cwappend Lib "YBXT.dll" (ByVal yyzyh As String, ByVal yyczydm As String, ByVal cwh As String, ByVal djrq As String) As Double
'床位删除
Public Declare Function cwdel Lib "YBXT.dll" (ByVal yyzyh As String, ByVal cwh As String, ByVal djrq As String) As Double

'//补缴费用初始化
Public Declare Function bjcsh Lib "YBXT.dll" (ByVal sbh As String, ByVal zhzt As String, ByVal zffs As String, _
                                            ByVal net As String, ByVal yyzyh As String, ByVal Zhye As Double) As Double
'//写补缴费用表
Public Declare Function bjappend Lib "YBXT.dll" (ByVal sbh As String, ByVal sjh As String, ByVal yyzyh As String, _
                                                ByVal yysfydm As String, ByVal Bz As String, ByVal net As String, _
                                                ByVal bjje As Double, ByVal zhbjje As Double) As Double

'//添加记帐信息
Public Declare Function jzappend Lib "YBXT.dll" (ByVal yyzyh As String, ByVal yyxmdm As String, ByVal yyczydm As String, _
                                                ByVal yycfys As String, ByVal djh As String, ByVal Bz As String, _
                                                ByVal sl As String, ByVal je As String, ByVal cfrq As String) As Double
'//删除记帐信息
Public Declare Function jzdel Lib "YBXT.dll" (ByVal yyzyh As String, ByVal Bz As String) As Double

'//计算病人当前报销费用信息
Public Declare Function jsbxf Lib "YBXT.dll" (ByVal yyzyh As String, ByRef qfx As Double, ByRef ylf As Double, _
                                            ByRef Tcje As Double, ByRef bxje As Double, ByRef Bcbxje As Double, _
                                            ByRef Jbbcbxje As Double, ByRef Gwybcbxje As Double, ByRef Tsbxje As Double) As Double

'//出院初始化
Public Declare Function cycsh Lib "YBXT.dll" (ByRef sbh As String, ByRef zhzt As String, ByRef zffs As String, _
                                            ByRef net As String, ByRef yyzyh As String, ByRef Mxbbz As String, _
                                            ByRef rzhye As Double, ByRef rylf As Double, ByRef rqfx As Double, _
                                            ByRef rtcje As Double, ByRef ryjje As Double, ByRef rzhyjje As Double, _
                                            ByRef rbxje As Double, ByRef rbcbxje As Double, ByRef rpbxje As Double, _
                                            ByRef rnbxje As Double, ByRef rpbcbxje As Double, ByRef rnbcbxje As Double, _
                                            ByRef rbcjs1 As Double, ByRef rbcjs2 As Double, ByRef rzhzf As Double, _
                                            ByRef rxjzf As Double, ByRef rzhtk As Double, ByRef rxjtk As Double, _
                                            ByRef rjbbcbxje As Double, ByRef rgwybcbxje As Double, ByRef rtsbxje As Double, _
                                            ByRef rpjbbcbxje As Double, ByRef rnjbbcbxje As Double, ByRef rpgwybcbxje As Double, _
                                            ByRef rngwybcbxje As Double, ByRef rptsbxje As Double, ByRef rntsbxje As Double) As Double
'//写出院表
Public Declare Function cyappend Lib "YBXT.dll" (ByVal yyzyh As String, ByVal bah As String, ByVal Cyzd As String, _
                                                ByVal djh As String, ByVal yysfy As String, ByVal Bz As String, _
                                                ByVal Bl As String, ByVal Mxbbz As String, ByVal qfx As Double, _
                                                ByVal Tcje As Double, ByVal Zhzf As Double, ByVal Grzf As Double, _
                                                ByVal Zhtk As Double, ByVal Xjtk As Double, ByVal bxje As Double, _
                                                ByVal Bcbxje As Double, ByVal Pbxje As Double, ByVal Nbxje As Double, _
                                                ByVal Pbcbxje As Double, ByVal Nbcbxje As Double, ByVal Bcjs1 As Double, _
                                                ByVal Bcjs2 As Double, ByVal Jbbcbxje As Double, ByVal Gwybcbxje As Double, _
                                                ByVal Tsbxje As Double, ByVal Pjbbcbxje As Double, ByVal Njbbcbxje As Double, _
                                                ByVal Pgwybcbxje As Double, ByVal Ngwybcbxje As Double, ByVal Ptsbxje As Double, _
                                                ByVal Ntsbxje As Double) As Double
        
'//门诊初始化
Public Declare Function mzcsh Lib "YBXT.dll" (ByRef sbh As String, ByRef net As String, _
                                     ByRef rylx As String, ByRef zhzt As String, _
                                     ByRef Zhye As Double) As Double
'//计算特殊门诊病人报销费用（计算每笔明细进入统筹部分）
Public Declare Function mzfyjs Lib "YBXT.dll" Alias "mzjs" (ByVal sbh As String, ByVal yyxmdm As String, _
                                            ByVal sl As Long, ByVal je As Double, _
                                            ByRef bxbl As Double, ByRef bxje As Double, _
                                            ByRef bsbhddj As Double) As Double
'门诊特殊病人报销费用计算基金支付部分(根据明细的进入统筹合计，计算出基金报销部分）
Public Declare Function mztsjs Lib "YBXT.dll" (ByVal sbh As String, ByVal yydjh As String, ByVal bxje As Double) As Double

'//写门诊支付明细
Public Declare Function mzzfmxappend Lib "YBXT.dll" (ByVal sbh As String, ByVal yyxmdm As String, _
                                                    ByVal yydjh As String, ByVal yysfydm As String, _
                                                    ByVal rylx As String, ByVal sl As Long, _
                                                    ByVal je As Double, ByVal dj As Double, _
                                                    ByVal bxbl As Double, ByVal bxje As Double) As Double
'//写门诊支付表
Public Declare Function mzappend Lib "YBXT.dll" (ByVal sbh As String, ByVal yydjh As String, _
                                                ByVal yysfydm As String, ByVal yyysdm As String, _
                                                ByVal rylx As String, ByVal net As String, _
                                                ByVal Zhzf As Double, ByVal xjzf As Double, _
                                                ByVal bxje As Double) As Double

'//住院与社保结算
Public Declare Sub zyjs Lib "YBMOD.dll" (ByVal yyczy As String)
'//门诊与社保结算
Public Declare Sub mzjs Lib "YBMOD.dll" (ByVal yyczy As String)
'//床位登记
Public Declare Sub cwdj Lib "YBMOD.DLL " (ByVal yyczy As String)
'//医保病人住院信息查询
Public Declare Sub ybcx Lib "YBMOD.dll" (ByVal yyczy As String)
'//查询人员医保信息
Public Declare Function getyhxx_vb Lib "YBXT.dll" (ByVal sbh As String, ByRef fhz As String) As Double
'//查询查询病人住院是否通过审批
Public Declare Function getspxx Lib "YBXT.dll" (ByVal yyzyh As String) As Double

Public Declare Function mzzffp Lib "YBXT.dll" (ByVal zhzt As String, ByVal yjje As Double, ByVal Zhye As Double, ByVal yydjh As String) As Double

'//删除门诊预结上传的明细
Public Declare Function mzdeltmp Lib "YBXT.dll" (ByVal yydjh As String) As Boolean

Private mrsCdTmp As New ADODB.Recordset   '临时记录集
Private mstrCdSql As String    '临时存放SQL语句Public Sub 更新病种_冕宁(lngPatiID, lngPageID)
Public gstrPuser_id As String '操作员代码

Private mblnInit As Boolean

'/////出院登记要传的参数太多，保存到数据库中不方便，所以定义为变量。方便结算
Private m_yyzyh As String * 18 '住院号，是病人住院的唯一标识号。（可以由出院初始化函数返回值得到，不能为空）
Private m_bah As String '病案号 (可以为空)
Private m_cyzd As String ' 出院诊断 (可以为空)
Private m_djh As String ' 收据号 (可以为空)
Private m_yysfy As String ': 收费员代码? (不能为空)

Private m_Bz As String ': 备注? (可以为空)
Private m_Bl As String ': 病例号? (可以为空)
Private m_Mxbbz As String * 1 ': 慢性病标志? (由出院初始化函数返回值得到)
Private m_Qfx As Double ': 统筹起伏线? (由出院初始化函数返回值得到)
Private m_Tcje As Double '：统筹金额（符合报销条件的金额）。（由出院初始化函数返回值得到）

Private m_Zhzf As Double ': 出院帐户支付金额? (由出院初始化函数返回值得到)
Private m_Grzf As Double '：出院现金支付金额（xjzf）。（由出院初始化函数返回值得到）
Private m_Zhtk As Double ' 出院帐户退款金额? (由出院初始化函数返回值得到)
Private m_Xjtk As Double ': 出院现金退款金额? (由出院初始化函数返回值得到)
Private m_Bxje As Double ': 基本报销金额? (由出院初始化函数返回值得到)

Private m_Bcbxje As Double ': 高额补充报销金额? (由出院初始化函数返回值得到)
Private m_Pbxje As Double ': 上年度基本报销金额? (由出院初始化函数返回值得到)
Private m_Nbxje As Double ': 本年度基本报销金额? (由出院初始化函数返回值得到)
Private m_Pbcbxje As Double ': 上年度高额补充报销金额? (由出院初始化函数返回值得到)
Private m_Nbcbxje As Double ': 本年度高额补充报销金额? (由出院初始化函数返回值得到)

Private m_Bcjs1 As Double ': 补充基数1? (由出院初始化函数返回值得到)
Private m_Bcjs2 As Double ': 补充基数2? (由出院初始化函数返回值得到)
Private m_Jbbcbxje As Double ': 基本补充报销金额? (由出院初始化函数返回值得到)
Private m_Gwybcbxje As Double ' : 公务员补充报销金额? (由出院初始化函数返回值得到)
Private m_Tsbxje As Double ': 特殊报销金额? (由出院初始化函数返回值得到)

Private m_Pjbbcbxje As Double ': 上年度基本补充报销金额? (由出院初始化函数返回值得到)
Private m_Njbbcbxje As Double ': 本年度基本补充报销金额? (由出院初始化函数返回值得到)
Private m_Pgwybcbxje As Double ': 上年度公务员补充报销金额? (由出院初始化函数返回值得到)
Private m_Ngwybcbxje As Double ': 本年度公务员补充报销金额? (由出院初始化函数返回值得到)
Private m_Ptsbxje As Double ': 上年度特殊报销金额? (由出院初始化函数返回值得到)

Private m_Ntsbxje As Double ': 本年度特殊报销金额? (由出院初始化函数返回值得到)

Private m_rylf As Double   '医疗费总额,结算后要保存
Private m_rxjzf As Double  '现金支付金额,结算后要保存
'/////出院登记要传的参数太多，保存到数据库中不方便，所以定义为变量。方便结算

Public Function 门诊结算冲销_冕宁(lngStlID, curMoney, lng病人ID) As Boolean
'被clsInsure 的  ClinicDelSwap 过程调用
'功能说明:本医保不允许冲销
On Error GoTo ErrH
    Err.Raise 9000, gstrSysName, "冕宁医保规定:不允许冲销已结算单据!", vbInformation, gstrSysName
    门诊结算冲销_冕宁 = False
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊虚拟结算_冕宁(rs明细记录 As ADODB.Recordset, str结算方式 As String, Optional str结算 As String) As Boolean
'被clsInsure 的 ClinicPreSwap 过程调用
'功能说明:完成门诊预结算功能
'调用函数清单及说明
'
'mzfyjs       门诊费用计算
'mztsjs       门诊特殊计算
'mzzfmxappend 门诊明细增加
'mzzffp       门诊支付分配
'mzdeltmp     门诊临时记录删除
'
Dim sbh As String  '社保号 (传入)
Dim yyxmdm As String '医院项目代码 (传入)
Dim sl  As Long      '数量(传入)
Dim dj As Double   '单价 (传入)
Dim je As Double    '金额 ( 传入)
Dim tmp As Double '接收对方返回的交易状态
Dim bxbl As Double '报销比例 (反回)
Dim bxje As Double  '报销金额 (反回)
Dim sbhddj As Double '社报核定价 (反回)
Dim lng病人ID As Long
Dim bxjeHj As Double  '报销金额合计
Dim ssjehj As Double    '实收金额合计
Dim rylx As String    '消费类型(传入)
Dim zhzt As String '帐户状态(传入)
Dim yjje As Double  '缴费金额(传入) =实收金额合计-报销金额合计
Dim Zhye As Double  '帐户余额
Dim yydjh As String  '医院单据号 (传入) 格式:1 & NO
Dim yysfydm As String '医院收费员代码(传入)
Dim blReturn As Boolean  ' 接收删除明细的返回值
Dim dbl个人帐户 As Double
Dim dbl统筹基金 As Double

On Error GoTo errHandle
    If rs明细记录.RecordCount = 0 Then
        MsgBox "没有病人费用明细，不能进行医保操作", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Beging 取医保号
    lng病人ID = rs明细记录!病人ID
    mstrCdSql = "Select * from 保险帐户 where 病人ID=[1] And 险类=[2]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "病人信息", lng病人ID, TYPE_冕宁)
    
    If mrsCdTmp.EOF Then
        MsgBox "不是冕宁保病人！，不能执行交易。", vbInformation, gstrSysName
        Exit Function
    End If
    sbh = mrsCdTmp.Fields!医保号
    rylx = mrsCdTmp.Fields!就诊类别
    Zhye = mrsCdTmp.Fields!帐户余额
    zhzt = mrsCdTmp.Fields!人员身份
    
    'Ene 取医保号
        
    'Beging 检查明细记录是否对码,对的码在医保中心是否存在
    Do Until rs明细记录.EOF
        mstrCdSql = "select * from 保险支付项目 where 险类=[1] and 收费细目ID=[2]"
        Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "医保明细", TYPE_冕宁, CLng(rs明细记录!收费细目ID))
        If mrsCdTmp.EOF Then
            mstrCdSql = "Select * from 收费项目目录 where ID=[1]"
            Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "收费项目目录", CLng(rs明细记录!收费细目ID))
            MsgBox mrsCdTmp!名称 & "(" & mrsCdTmp!编码 & ")" & "未对码！" & vbCrLf & "请对码后再使用此功能！", vbInformation, gstrSysName
            Exit Function
        End If
        yyxmdm = mrsCdTmp!项目编码
        '检查在医保是否存在
        'findxm (yyxmdm)
        rs明细记录.MoveNext
    Loop
    'End 检查明细记录是否对码,对的码在医保中心是否存在
    
    
    'Beging 逐笔上传明细,并计算每笔明细的报销比例,报销金额,
    rs明细记录.MoveFirst
    
    bxbl = 0
    bxje = 0
    sbhddj = 0
    bxjeHj = 0
    ssjehj = 0
    Do Until rs明细记录.EOF
        mstrCdSql = "select * from 保险支付项目 where 险类=[1] and 收费细目ID=[2]"
        Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "医保明细", TYPE_冕宁, CLng(rs明细记录!收费细目ID))
        yyxmdm = mrsCdTmp!项目编码
        sl = rs明细记录!数量
        dj = rs明细记录!单价
        je = rs明细记录!实收金额
        If Nvl(str结算, 0) = 9 Then
            yydjh = "1" & rs明细记录!NO
        Else
            yydjh = lng病人ID & Format(rs明细记录!结算时间, "yyMMddHHmmdd")
        End If
        ssjehj = ssjehj + je
        yysfydm = gstrPuser_id
        '人员类型为3 （特殊病人）,要计算报销金额
        If rylx = 3 Then
            tmp = mzfyjs(sbh, yyxmdm, sl, je, bxbl, bxje, sbhddj)
            Call WriteBusinessLOG("mzfyjs", "sbh, yyxmdm, sl, je, bxbl, bxje, sbhddj", tmp & "," & sbh & "," & yyxmdm & "," & sl & "," & je & "," & bxje & "," & sbhddj)
            bxjeHj = bxjeHj + bxje
            
        End If
        
        
        tmp = mzzfmxappend(sbh, yyxmdm, yydjh, yysfydm, rylx, sl, je, dj, bxbl, bxje)
        Call WriteBusinessLOG("mzzfmxappend", "sbh, yyxmdm, yydjh, yysfydm, rylx, sl, je, dj, bxbl, bxje", tmp & "," & sbh & "," & yyxmdm & "," & yydjh & "," & yysfydm & "," & rylx & "," & sl & "," & je & "," & dj & "," & bxbl & "," & bxje)
        
        Select Case tmp
            Case 0
                If Nvl(str结算, 0) = 9 Then
                    gstrSQL = "ZL_病人费用记录_更新医保(" & rs明细记录!ID & "," & _
                            bxje & _
                            ",NULL,1,NULL,1," & bxbl & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新医保字段")
                    
                End If
            Case 1
                MsgBox "单据号（处方号）重复", vbInformation, gstrSysName
                Exit Function
            Case 2
                MsgBox "个人信息库中没有该参保人员", vbInformation, gstrSysName
                Exit Function
            Case 3
                MsgBox "医保项目库没有该项目", vbInformation, gstrSysName
                Exit Function
            Case 99
                MsgBox "错误", vbInformation, gstrSysName
                Exit Function
        End Select
    
    
      rs明细记录.MoveNext
    Loop
    'end 逐笔上传明细
    
    '
    
    yjje = ssjehj - bxjeHj '实收金额-报销金额
    dbl个人帐户 = mzzffp(zhzt, yjje, Zhye, yydjh)
    Call WriteBusinessLOG("Mzzffp", "zhzt, yjje, zhye, yydjh", dbl个人帐户 & "," & zhzt & "," & yjje & "," & Zhye & "," & yydjh)
    
    dbl统筹基金 = bxjeHj
    '人员类型=3,报销金额>0
    If rylx = 3 And bxjeHj > 0 Then
        dbl统筹基金 = mztsjs(sbh, yydjh, bxjeHj)
        Call WriteBusinessLOG("mztsjs", "sbh, yydjh, bxjehj", dbl统筹基金 & "," & sbh & "," & yydjh & "," & bxjeHj)
    End If
    
    'beging 是预结算,要调用删除明细
    If Nvl(str结算, 0) <> 9 Then
        blReturn = mzdeltmp(yydjh)
        Call WriteBusinessLOG("mzdeltmp", yydjh, IIf(blReturn, "True", "False"))
    End If
    'end 是预结算
    
        '返回值到HIS
    str结算方式 = "个人帐户;" & dbl个人帐户 & ";1|统筹基金;" & dbl统筹基金 & ";0|公务员补助;" & 0 & ";0"
    门诊虚拟结算_冕宁 = True
    
    Exit Function

errHandle:
    门诊虚拟结算_冕宁 = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_冕宁(lng结帐ID, cur个帐支付, str医保号) As Boolean
    '被clsInsure 的 ClinicSwap 过程调用
    '功能说明:完成门诊结算
    '调用过程清单及功能说明
    '    门诊虚拟结算_冕宁: 完成门诊预结算功能
    '    医院医师代码_冕宁: 查询是否有指定医师,如果没有则添加,返回医师编号
    '    mzappend         : 门诊结算登记
        
    Dim sbh As String  '社保号 (传入)
    Dim yydjh As String  '医院单据号 (传入) 格式:1 & NO
    Dim yysfydm As String '医院收费员代码(传入)
    Dim yyysdm As String  '医院医师代码(传入)
    Dim rylx As String    '消费类型(传入)
    Dim net As String '网络状态(传入)
    Dim Zhzf As Double '帐户支付
    Dim xjzf As Double  '现金支付
    Dim bxje As Double  '报销金额 (传入)
    Dim str预算信息 As String '接收预算返回的信息
    Dim tmp As Double '接收对方返回的交易状态
    Dim lng病人ID As Long
    
    Dim rscd明细记录 As New ADODB.Recordset
On Error GoTo ErrH
    mstrCdSql = "Select ID,NO,序号,记录性质,登记时间 as 结算时间,病人ID,收费类别,收据费目,计算单位,开单人, " & _
                "收费细目ID,nvl(数次,0)*nvl(付数,0) as 数量,标准单价 as 单价, " & _
                "实收金额,统筹金额,保险大类ID 保险支付大类ID, " & _
                "摘要,是否急诊 " & _
                "from 门诊费用记录 " & _
                "where 结帐ID=[1]"
    Set rscd明细记录 = zlDatabase.OpenSQLRecord(mstrCdSql, "费用记录", lng结帐ID)
    
    yyysdm = 医院医师代码_冕宁(rscd明细记录!开单人)
    lng病人ID = rscd明细记录!病人ID
    yydjh = 1 & rscd明细记录!NO
    yysfydm = gstrPuser_id
    
    '调上传明细
    If 门诊虚拟结算_冕宁(rscd明细记录, str预算信息, 9) = False Then
        Exit Function
    End If
    
    mstrCdSql = "Select * from 保险帐户 where 险类=[1] And 病人ID=[2]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "保险帐户", TYPE_冕宁, lng病人ID)
    sbh = mrsCdTmp.Fields!医保号
    rylx = mrsCdTmp.Fields!就诊类别
    net = mrsCdTmp.Fields!密码
    Zhzf = cur个帐支付
    
    mstrCdSql = "Select sum(nvl(冲预交,0)) as 金额 from 病人预交记录 where 结算方式='现金' And 结帐ID=[1]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "保险帐户", lng结帐ID)
    xjzf = Nvl(mrsCdTmp!金额, 0)
    
    mstrCdSql = "Select sum(nvl(冲预交,0)) as 金额 from 病人预交记录 where 结算方式='统筹基金' And 结帐ID=[1]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "保险帐户", lng结帐ID)
    bxje = Nvl(mrsCdTmp!金额, 0)
    
    tmp = mzappend(sbh, yydjh, yysfydm, yyysdm, rylx, net, Zhzf, xjzf, bxje)
    Call WriteBusinessLOG("mzappend", "sbh,yydjh,yysfydm,yyysdm,rylx,net,zhzf,xjzf,bxje", tmp & "," & sbh & "," & yydjh & "," & yysfydm & "," & yyysdm & "," & rylx & "," & net & "," & Zhzf & "," & xjzf & "," & bxje)
      
    Select Case tmp
        Case 0
            门诊结算_冕宁 = True
        Case 1
            Err.Raise 9000, gstrSysName, "单据号（处方号）重复"
            Exit Function
        Case 2
            Err.Raise 9000, gstrSysName, "个人信息库中没有该参保人员"
            Exit Function
        Case 3
            Err.Raise 9000, gstrSysName, "网络或刷卡错误"
            Exit Function
        Case Else
            Err.Raise 9000, gstrSysName, "错误"
            Exit Function
    End Select
      
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_冕宁 & "," & _
            lng病人ID & "," & Format(Now(), "YYYY") & ",0,0, " & _
            "" & _
            0 & ",NULL,NULL,NULL,NULL,0," & _
            Zhzf + xjzf + bxje & "," & xjzf & ",NULL,NULL," & bxje & ",NULL,NULL," & _
            cur个帐支付 & ",NULL,NULL,NULL,'" & rylx & "')"
            '                                 消费类型
    Call zlDatabase.ExecuteProcedure(gstrSQL, "冕宁医保")
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 撤销入院登记_冕宁(lngPatiID, lngPageID) As Boolean
'被clsInsure 的 ComeInDelSwap  过程调用
'功能: 不支持此交易
    撤销入院登记_冕宁 = False
    MsgBox "冕宁医保不支持撤消入院", vbInformation, gstrSysName
End Function

Public Function 入院登记_冕宁(lngPatiID, lngPageID, str医保号) As Boolean
'被clsInsure 的 ComeInSwap 过程调用
'功能:完成医保病人入院登记
'调用过程清单及说明:
'    医院科室代码_冕宁 : 查询是否有指定科室,如果没有则添加,返回科室编号
'    医院医师代码_冕宁 : 查询是否有指定医师,如果没有则添加,返回医师编号
'    ryappend          : 医保入院登记交易

    Dim zt  As Double '接收返回值
    Dim strMsg As String '提示信息
    Dim sbh As String    '社保号
    Dim yyzyh As String  '医院住院号
    Dim yykb As String   '医生所在科别
    Dim yyzdys As String '医院诊断医师
    Dim yysfydm As String '医院收费员代码
    Dim ryzd As String '入院诊断
    Dim lxdh  As String ' 联系电话
    Dim jtzz As String '家庭住址
    Dim Bz As String '备注
    Dim net As String '网络状态
    Dim iyjje As Double '预交金额
    Dim izhyj As Double  '帐户预交
    
    mstrCdSql = "Select * from 保险帐户 where 险类=[1] And 病人ID=[2]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "保险帐户", TYPE_冕宁, lngPatiID)
    sbh = mrsCdTmp!医保号
    yyzyh = lngPatiID & "_" & lngPageID
    net = mrsCdTmp.Fields!密码
    
    mstrCdSql = "Select * from 病案主页 where 险类=[1] And 病人ID=[2] And 主页ID=[3]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "病案主页", TYPE_冕宁, lngPatiID, lngPageID)
    
    yykb = 医院科室代码_冕宁(mrsCdTmp!入院科室ID)
    
    yyzdys = 医院医师代码_冕宁(mrsCdTmp!门诊医师)
    yysfydm = gstrPuser_id
    
    lxdh = Nvl(mrsCdTmp!联系人电话, "")
    jtzz = Nvl(mrsCdTmp!家庭地址, "")
    Bz = Nvl(mrsCdTmp!备注, "")
    iyjje = 0
    izhyj = 0
    zt = 99
    mstrCdSql = "select 诊断类型,描述信息 From 诊断情况 Where 病人ID=[1] And 主页ID=[2] And 诊断类型>=1 and 诊断类型<=2 order by 诊断类型 desc"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "入院诊断", lngPatiID, lngPageID)
    
    If mrsCdTmp.RecordCount > 0 Then
        ryzd = Trim(Nvl(mrsCdTmp!描述信息, "空"))
    Else
        ryzd = "无"
    End If
    
    zt = ryappend(sbh, yyzyh, yykb, yyzdys, yysfydm, ryzd, lxdh, jtzz, Bz, net, iyjje, izhyj)
    Call WriteBusinessLOG("ryappend(登记入院信息)", "sbh:" & sbh & ",yyzyh:" & yyzyh & _
                                                    ",yykb:" & yykb & ",yyzdys:" & yyzdys & _
                                                    ",yysfydm:" & yysfydm & ",ryzd:" & ryzd & _
                                                    ",lxdh:" & lxdh & ",jtzz" & jtzz & _
                                                    ",bz:" & Bz & ",net:" & net & _
                                                    ",iyjje:" & iyjje & ",izhyj:" & izhyj, zt)
                                                    
    
    Select Case zt
        Case 0
            入院登记_冕宁 = True
            gstrSQL = "zl_保险帐户_入院(" & lngPatiID & "," & TYPE_冕宁 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "医保入院")
        Case 1
            strMsg = "请输入完整参数信息!"
        Case 2
            strMsg = "该病人已经入院!"
        Case 3
            strMsg = "医保住院号重复，不能登记入院!"
        Case 4
            strMsg = "无该社保号人员!"
        Case 5
            strMsg = "数据上传错误或IC卡操作错误!"
        Case 99
            strMsg = "错误!"
    End Select
        
    If zt <> 0 Then
        MsgBox strMsg, vbInformation, gstrSysName
        入院登记_冕宁 = False
    End If
            
End Function

Public Function 身份标识_冕宁(bytType As Byte, lng病人ID As Long) As String
'被clsInsure 的 Identify  过程调用
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
'调用过程清单及说明:
'     frmIdentify冕宁 : 完成身份验证,返回病人信息

    Dim str就诊类别 As String
    Dim str门诊号 As String
    Dim str日期 As String
    Dim strIdeReturn As String  '接收返回信息
    
    If Not (bytType = 1 Or bytType = 0) Then Exit Function  '仅门诊收费，入院登记才调用
    
    strIdeReturn = frmIdentify冕宁.身份标识(bytType, lng病人ID) '''发起接口事务,readcard
    If strIdeReturn = "99" Then
        身份标识_冕宁 = ""
    Else
        身份标识_冕宁 = strIdeReturn
    End If
    
End Function

Public Sub 取消就诊_冕宁()
'被clsInsure 的 IdentifyCancel  过程调用
'功能:
'调用过程清单及说明:(无)
End Sub

Public Function 医保初始化_冕宁() As Boolean
'被clsInsure 的  InitInsure  过程调用
'功能: 医保初始化
'调用过程清单及说明:
'    收费员代码_冕宁: 查询当前人员是否在医保中心注册,如果没有则添加,返回人员编号
    医保初始化_冕宁 = True
    mblnInit = True
    gstrPuser_id = 收费员代码_冕宁()
End Function

Public Function 出院登记撤销_冕宁(lngPatiID, lngPageID) As Boolean
'被clsInsure 的 LeaveDelSwap 过程调用
'功能: 参保病人撤消出院
'调用过程清单及说明:
' (无)

Dim rsCd As New ADODB.Recordset

mstrCdSql = "select A.入院日期,A.出院日期,B.住院号,D.名称 as 住院科室,A.入院病床,A.住院医师,C.卡号," & _
          "C.密码,D.编码 As 科室编码,C.就诊类别 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
          "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
          "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]" & _
          " and C.险类=" & TYPE_冕宁

Set rsCd = zlDatabase.OpenSQLRecord(mstrCdSql, "保险病人", lngPatiID, lngPageID)
If rsCd.EOF Then
    出院登记撤销_冕宁 = False
    MsgBox "该病人未通过身份验证！不能办理撤消出院。", vbInformation, gstrSysName
    Exit Function
End If

出院登记撤销_冕宁 = True
gstrSQL = "zl_保险帐户_入院(" & lngPatiID & "," & TYPE_冕宁 & ")"
Call zlDatabase.ExecuteProcedure(gstrSQL, "撤消出院登记")

End Function

Public Function 出院登记_冕宁(lngPatiID, lngPageID) As Boolean
'被clsInsure 的 LeaveSwap 过程调用
'功能: 办理医保病人出院
'调用过程清单及说明:
' (无)

    出院登记_冕宁 = True
    gstrSQL = "zl_保险帐户_出院(" & lngPatiID & "," & TYPE_冕宁 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
End Function

Public Function 病人变动记录上传_冕宁(lngPatiID, lngPageID) As Boolean
'被clsInsure 的 ModiPatiSwap 过程调用
'功能: 上传床位信息到医保中心
'调用过程清单及说明
'    cwappend: 床位登记
    Dim tmp As Double
    Dim yyzyh As String
    Dim yyczydm As String
    Dim cwh As String
    Dim djrq As String
    
    gstrSQL = "Select B.名称||' '||A.床号||'床' as 床位号,to_char(开始时间,'YYYY-MM-DD') as 开始日期" & _
             " from 病人变动记录 A,部门表 B " & _
            "where A.科室ID=B.ID And A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.终止时间 is null And A.床号 is not null"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(gstrSQL, "变动记录", lngPatiID, lngPageID)
    
    Do Until mrsCdTmp.EOF
        yyzyh = lngPatiID & "_" & lngPageID
        yyczydm = 收费员代码_冕宁()
        cwh = mrsCdTmp!床位号
        djrq = mrsCdTmp!开始日期
        
        tmp = cwappend(yyzyh, yyczydm, cwh, djrq)
        Call WriteBusinessLOG("cwappend(床位变动)", "yyzyh:" & yyzyh & ",yyczydm:" & yyczydm & ",cwh:" & cwh & "djrq:" & djrq, tmp)
        mrsCdTmp.MoveNext
    Loop
    病人变动记录上传_冕宁 = True
End Function

Public Function 个人余额_冕宁(str医保号) As Currency
'被clsInsure 的 SelfBalance 过程调用
'功能: 提取参保病人个人帐户余额
'调用过程清单及说明:
' (无)

    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    gstrSQL = "Select Nvl(帐户余额,0) AS 个人帐户 From 保险帐户 " & _
              " Where 医保号=[1] and 险类=[2]"
              
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人ID", str医保号, TYPE_冕宁)
    个人余额_冕宁 = rsTemp!个人帐户
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function 住院结算冲销_冕宁(lng结帐ID) As Boolean
'被clsInsure 的 SettleDelSwap  过程调用
'功能:冕宁医保不支持结算冲销
On Error GoTo ErrH
    Err.Raise 9000, gstrSysName, "冕宁医保不支持结算冲销!"
    住院结算冲销_冕宁 = False
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算_冕宁(lng结帐ID) As Boolean
'被clsInsure 的 SettleSwap 过程调用
'功能：完成住院登记交易
'调用过程清单及说明
'    收费员代码_冕宁:查询当前人员是否在医保中心注册,如果没有则添加,返回人员编号
'    cyappend       :写出院表

    Dim lng病人ID As Long, lng主页ID As Long
    Dim blnOut As Boolean  '是否中途结算
    Dim strNO As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMsg As String
    Dim tmp As Double
    
On Error GoTo ErrH

    m_bah = "" '病案号
    m_djh = "" '单据号
    m_yysfy = 收费员代码_冕宁()
    m_Bz = ""
    m_Bl = ""
    
'其他参数由住院初始化返回
    '提取结帐单号
    gstrSQL = "Select NO,病人ID From 病人结帐记录 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取结帐单号", lng结帐ID)
    strNO = "2" & rsTemp!NO
    lng病人ID = rsTemp!病人ID
    

    '读取病人主页ID，出院日期
    gstrSQL = " Select A.主页ID,A.出院日期 From 病案主页 A,病人信息 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 and B.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人主页ID，出院日期", lng病人ID)
    blnOut = Not (IsNull(rsTemp!出院日期))
    lng主页ID = rsTemp!主页ID
    
    '2005 0520 添加出院诊断
    gstrSQL = "select * From 诊断情况 Where 病人ID=[1] And 主页ID=[2] And 诊断类型=3"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "诊断情况", lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
        m_cyzd = Trim(Nvl(rsTemp!描述信息, "无"))
    Else
        m_cyzd = "无"
    End If
    
    If m_cyzd = "无" Then
        gstrSQL = "select * From 诊断情况 Where 病人ID=[1] And 主页ID=[2] And 诊断类型=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "诊断情况", lng病人ID, lng主页ID)
        If Not rsTemp.EOF Then
            m_cyzd = Trim(Nvl(rsTemp!描述信息, "无"))
        Else
            m_cyzd = "无"
        End If
        m_cyzd = InputBox("请输入出院诊断", "出院诊断录入", m_cyzd)
    End If
    
    If Trim(m_cyzd) = "" Then m_cyzd = "无"
    
    '2005 0520 添加出院诊断
    
    tmp = cyappend(m_yyzyh, m_bah, m_cyzd, m_djh, m_yysfy, _
                  m_Bz, m_Bl, m_Mxbbz, m_Qfx, m_Tcje, _
                  m_Zhzf, m_Grzf, m_Zhtk, m_Xjtk, m_Bxje, _
                  m_Bcbxje, m_Pbxje, m_Nbxje, m_Pbcbxje, m_Nbcbxje, _
                  m_Bcjs1, m_Bcjs2, m_Jbbcbxje, m_Gwybcbxje, m_Tsbxje, _
                  m_Pjbbcbxje, m_Njbbcbxje, m_Pgwybcbxje, m_Ngwybcbxje, m_Ptsbxje, _
                  m_Ntsbxje)
       
    Call WriteBusinessLOG("cyappend(写出院表)", "m_yyzyh:" & m_yyzyh & ",m_bah:" & m_bah & ",m_cyzd:" & m_cyzd & ",m_djh:" & m_djh & ",m_yysfy:" & m_yysfy & "," & _
                  "m_Bz:" & m_Bz & ",m_Bl:" & m_Bl & ",m_Mxbbz:" & m_Mxbbz & ",m_Qfx:" & m_Qfx & ",m_Tcje" & m_Tcje & "," & _
                  "m_Zhzf:" & m_Zhzf & ",m_Grzf:" & m_Grzf & ",m_Zhtk:" & m_Zhtk & ",m_Xjtk:" & m_Xjtk & ",m_Bxje:" & m_Bxje & "," & _
                  "m_Bcbxje" & m_Bcbxje & ",m_Pbxje:" & m_Pbxje & ",m_Nbxje:" & m_Nbxje & ",m_Pbcbxje:" & m_Pbcbxje & ",m_Nbcbxje:" & m_Nbcbxje & "," & _
                  "m_Bcjs1:" & m_Bcjs1 & ",m_Bcjs2:" & m_Bcjs2 & ",m_Jbbcbxje:" & m_Jbbcbxje & ",m_Gwybcbxje:" & m_Gwybcbxje & ",m_Tsbxje:" & m_Tsbxje & "," & _
                  "m_Pjbbcbxje:" & m_Pjbbcbxje & ",m_Njbbcbxje:" & m_Njbbcbxje & ",m_Pgwybcbxje:" & m_Pgwybcbxje & ",m_Ngwybcbxje:" & m_Ngwybcbxje & ",m_Ptsbxje:" & m_Ptsbxje & "," & _
                  "m_Ntsbxje:" & m_Ntsbxje, tmp)

    '保存保险结算记录
    '大病自付=大病补助;超限自付=公务员补助
    Select Case tmp
        Case 0
            gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_冕宁 & "," & lng病人ID & "," & _
                Year(zlDatabase.Currentdate()) & ",0,0,0,0,0,0,0,0," & m_rylf & "," & m_rxjzf & ",0," & _
                m_Tcje & "," & m_Bxje + m_Bcbxje + m_Jbbcbxje + m_Tsbxje & ",0," & m_Gwybcbxje & "," & m_Zhzf & _
                ",'" & lng病人ID & "_" & lng主页ID & "'," & lng主页ID & "," & IIf(blnOut, 1, 0) & _
                ",'" & m_Bxje & ";" & m_Bcbxje & ";" & m_Jbbcbxje & ";" & m_Tsbxje & "')"
                 '  ^报销金额       ^补充报销        ^基本补充报销      ^ 特殊报销金额
            'gcnOracle.Execute gstrSQL, , adCmdStoredProc
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新保险结算记录")
            住院结算_冕宁 = True
        Case 1
            strMsg = "已经出院。"
        Case 2
            strMsg = "网络不通或写卡错误。"
        Case 3
            strMsg = "参数信息不全。"
        Case Else
            strMsg = "其他错误。"
    End Select
    
    If tmp <> 0 Then
        住院结算_冕宁 = False
        Err.Raise 9000, gstrSysName, strMsg, vbInformation, gstrSysName
    End If
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 记帐上传_冕宁(int性质, int状态, str单据号) As Boolean
'被clsInsure 的 TranChargeDetail 过程调用
'功能:住院费用上传及删除
'调用过程清单及说明
'    医院医师代码_冕宁: 查找医保中心是否有指定编号医师
'    jzappend         : 记帐费用添加
'    jzdel            : 记帐费用删除
    Dim rsCd As New ADODB.Recordset
    Dim rs记帐明细 As New ADODB.Recordset
    Dim lng病人ID As Long
    
    Dim yyzyh As String '住院号
    Dim yyxmdm As String '项目代码
    Dim yyczydm As String '操作员代码
    Dim yycfys As String  '处方医师
    Dim djh As String  '单据号
    Dim Bz As String '备注    流水号,唯一标识
    Dim sl As String '数量
    Dim je As String  '金额
    Dim cfrq As String '处方日期  格式 :YYYY-MM-DD
    
    Dim tmp As Double '接收返回值
    Dim strMsg As String '存放提示信息
    
    ' 如果记录状态为1的单据，有负数记录，则不允许保存单据
    
    If int状态 = 1 Then
    gstrSQL = "Select distinct  A.病人ID from 住院费用记录 A,保险帐户 B " & _
            "where A.病人ID=B.病人ID And A.记录性质=[1]" & _
            " And A.记录状态=[2] And A.NO=[3] " & _
            " And B.险类=[4] And A.实收金额<0"
        Set rsCd = zlDatabase.OpenSQLRecord(gstrSQL, "是否有负数记录", int性质, int状态, str单据号, TYPE_冕宁)
        If Not rsCd.EOF Then
            MsgBox "本医保不支持负数记录！请调整。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '根据NO号,提取病人ID
    gstrSQL = "Select distinct  A.病人ID from 住院费用记录 A,保险帐户 B " & _
            "where A.病人ID=B.病人ID And A.记录性质=[1]" & _
            " And A.记录状态=[2] And A.NO=[3] " & _
            " And B.险类=[4]"
    Set rsCd = zlDatabase.OpenSQLRecord(gstrSQL, "取病人ID", int性质, int状态, str单据号, TYPE_冕宁)
    
    
    记帐上传_冕宁 = True
    
    '>beging 如果是记帐表,需要将记帐表上的医保病人逐个上传
    Do Until rsCd.EOF
        lng病人ID = rsCd!病人ID
        'If int状态 = 1 Then
            gstrSQL = "Select A.*,D.项目编码,nvl(A.付数,1)*nvl(A.数次,0) as 数量,A.实收金额 as 金额," & _
                                      "nvl(A.实收金额,0)/(nvl(A.付数,1)*nvl(A.数次,0)) as 价格,A.开单人 as 医生,C.名称 as 开单部门,B.* " & _
                              " from 住院费用记录 A,保险帐户 B,部门表 C,保险支付项目 D " & _
                              " where A.NO=[1]" & _
                                    " And A.记录性质=[2]" & _
                                    " And A.记录状态=[3]" & _
                                    " And nvl(A.是否上传,0)=0 " & _
                                    " And A.病人ID=B.病人ID " & _
                                    " and B.险类=[4]" & _
                                    " and A.开单部门ID=C.ID " & _
                                    " ANd A.病人ID=[5]" & _
                                    " And A.收费细目ID=D.收费细目ID " & _
                                    " And D.险类=[4]"
                                    
                                    
            Set rs记帐明细 = zlDatabase.OpenSQLRecord(gstrSQL, "记帐明细", str单据号, int性质, int状态, TYPE_冕宁, lng病人ID)
            Do Until rs记帐明细.EOF
                yyzyh = rs记帐明细!病人ID & "_" & rs记帐明细!主页ID
                yyxmdm = rs记帐明细!项目编码
                yyczydm = gstrPuser_id
                yycfys = Nvl(医院医师代码_冕宁(rs记帐明细!医生), "")
                djh = rs记帐明细!记录性质 & rs记帐明细!NO
                Bz = rs记帐明细!记录性质 & rs记帐明细!NO & "_" & rs记帐明细!序号
                sl = rs记帐明细!数量
                je = Format(rs记帐明细!金额, "0.00")
                cfrq = Format(rs记帐明细!发生时间, "yyyy-MM-dd")
                '>>beging 正常单据
                If int状态 = 1 Then
                    tmp = jzappend(yyzyh, yyxmdm, yyczydm, yycfys, djh, Bz, sl, je, cfrq)
                    Call WriteBusinessLOG("jzappend(记帐登记)", "yyzyh:" & yyzyh & ",yyxmdm:" & yyxmdm & _
                                                              ",yyczydm:" & yyczydm & ",yycfys:" & yycfys & _
                                                               ",djh:" & djh & ",bz:" & Bz & _
                                                               ",sl:" & sl & ",je:" & je & _
                                                               ",cfrq:" & cfrq, tmp)
                
                End If
                '>> end 正常单据

                '>>beging 冲销单据
                If int状态 = 2 Then
                    tmp = jzdel(yyzyh, Bz)
                    Call WriteBusinessLOG("jzdel(记帐删除)", "yyzyh:" & yyzyh & ",bz:" & Bz, tmp)
                    
                End If
                '>>end 冲销单据

                If tmp = 0 Then
                    '上传成功,则打上传标志
                    gstrSQL = "ZL_病人费用记录_更新医保(" & rs记帐明细!ID & "," & _
                            "0" & _
                            ",NULL,1,NULL,1,'" & Bz & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新医保字段")
                Else
                    strMsg = "上传处方[" & rs记帐明细!NO & "]第" & rs记帐明细!序号 & "行明细时出错！" & vbCrLf & "详细信息：" & vbCrLf
                    Select Case tmp
                    Case 1
                        If int状态 = 1 Then
                            strMsg = strMsg & "没有住院信息?"
                        Else
                            strMsg = strMsg & "找不到对应记帐信息?"
                        End If
                    Case 2
                         strMsg = strMsg & "找不到对应项目?"
                    Case Else
                         strMsg = strMsg & "错误"
                    End Select
                    MsgBox strMsg, vbInformation, gstrSysName
                    'Exit Function
                End If
                rs记帐明细.MoveNext
            Loop
        'End If
        
    rsCd.MoveNext
    Loop
    '>end 如果是记帐表,需要将记帐表上的医保病人逐个上传
                                           
End Function

Public Function 住院虚拟结算_冕宁(rsExse As Recordset, ByVal lng病人ID As Long, str密码 As String) As String
'被clsInsure 的 WipeoffMoney 过程调用
'功能说明:完成住院预结算功能
'调用过程清单及功能说明:
'    jsbxf          :计算病人当前报销费用信息
'    记帐上传_冕宁  :住院费用上传及删除
'    cycsh          :出院初始化

    Dim yyzyh As String '住院号
    Dim qfx As Double '起付线
    Dim ylf As Double '医疗费
    Dim Tcje As Double '符合报销条件金额（统筹金额）
    Dim bxje As Double '基本报销金额
    Dim Bcbxje As Double '高额补充报销金额
    Dim Jbbcbxje As Double '基本补充报销金额
    Dim Gwybcbxje As Double '公务员补充报销金额
    Dim Tsbxje As Double    '特殊报销金额
    
    Dim sbh As String * 20   '社保号
    Dim zhzt As String * 3 '帐户状态
    Dim zffs As String * 1 '帐户支付方式
    Dim net As String * 1  '网络状态
    Dim rzhye As Double  '帐户余额
    Dim ryjje As Double  '预缴金额
    Dim rzhyjje As Double '帐户预缴金额
    
    Dim tmp As Double '返回值
    Dim strMsg  As String '反回信息
    
    gstrSQL = "select * from 病人信息 where 病人ID=[1] And 险类=[2]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(gstrSQL, "病人信息", lng病人ID, TYPE_冕宁)
    If mrsCdTmp.EOF Then
        MsgBox "不是冕宁医保病人,不能执行此操作!", vbInformation, gstrSysName
        Exit Function
    End If
    
    '>beging 补传未上传记录
       ' 查出未上传记录的,记录性质 , 记录状态, NO 调用记帐上传
       gstrSQL = "Select distinct 记录性质,记录状态,NO From 住院费用记录 A,保险帐户 B,病人信息 C " & _
                 " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID And A.主页ID=C.住院次数" & _
                 " And nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 and A.记帐费用=1 And A.操作员姓名 is not null " & _
                 " AND A.实收金额 IS NOT NULL And B.病人ID=[1] And B.险类=[2]"
        Set mrsCdTmp = zlDatabase.OpenSQLRecord(gstrSQL, "病人费用记录", lng病人ID, TYPE_冕宁)
        Do Until mrsCdTmp.EOF
            Call 记帐上传_冕宁(mrsCdTmp!记录性质, mrsCdTmp!记录状态, mrsCdTmp!NO)
            mrsCdTmp.MoveNext
        Loop
    
    '>end 补传未上传记录
    gstrSQL = "Select * from 病人信息 where 病人ID=[1] And 险类=[2]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(gstrSQL, "病人费用记录", lng病人ID, TYPE_冕宁)
    yyzyh = lng病人ID & "_" & mrsCdTmp!住院次数
    '虚拟计算
    tmp = jsbxf(yyzyh, qfx, ylf, Tcje, bxje, Bcbxje, Jbbcbxje, Gwybcbxje, Tsbxje)
    Call WriteBusinessLOG("jsbxf(住院虚拟结算)", "yyzyh:" & yyzyh & ",qfx:" & qfx & ",ylf:" & ylf & ",tcje:" & Tcje & ",bxje:" & bxje & ",bcbxje:" & Bcbxje & ",jbbcbxje:" & Jbbcbxje & ",gwybcbxje:" & Gwybcbxje & ",tsbxje:" & Tsbxje, tmp)
    If tmp = 0 Then
        '查总费用是否与医保中心相等
        
        gstrSQL = "Select sum(nvl(实收金额,0))-sum(nvl(结帐金额,0)) as 未结费用 From 住院费用记录 Where nvl(记录状态,0)<>0 and 记帐费用=1 And 病人ID= [1]"
        Set mrsCdTmp = zlDatabase.OpenSQLRecord(gstrSQL, "未结费用", lng病人ID)
        
        If Val(Nvl(mrsCdTmp.Fields!未结费用, 0)) <> Val(Format(ylf, "0.00")) Then
            If MsgBox("医院的费用总金额(" & Nvl(mrsCdTmp.Fields!未结费用, 0) & ")与医保中心的费用总额(" & Val(Format(ylf, "#####0.00")) & ")不等，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        '冕宁医保预结算不能计算出个人帐户金额,必须出院刷卡才能得到.
        '结帐处调用(str密码 = 1)，则要刷卡。
        If str密码 = 1 Then
            '结算初始化
                sbh = Space(20): zhzt = Space(3): zffs = Space(1): net = Space(1)
                m_yyzyh = Space(18): m_Mxbbz = Space(1)
                tmp = cycsh(sbh, zhzt, zffs, net, m_yyzyh, _
                            m_Mxbbz, rzhye, m_rylf, m_Qfx, m_Tcje, _
                            ryjje, rzhyjje, m_Bxje, m_Bcbxje, m_Pbxje, _
                            m_Nbxje, m_Pbcbxje, m_Nbcbxje, m_Bcjs1, m_Bcjs2, _
                            m_Zhzf, m_rxjzf, m_Zhtk, m_Xjtk, m_Jbbcbxje, _
                            m_Gwybcbxje, m_Tsbxje, m_Pjbbcbxje, m_Njbbcbxje, m_Pgwybcbxje, _
                            m_Ngwybcbxje, m_Ptsbxje, m_Ntsbxje)
            
                Call WriteBusinessLOG("crcsh(出院初始化)", "sbh:" & sbh & ",zhzt:" & zhzt & ", zffs:" & zffs & " , net:" & net & ", yyzyh:" & m_yyzyh & " ," & _
                            "mxbbz:" & m_Mxbbz & ", rzhye:" & rzhye & ", rylf:" & m_rylf & ", rqfx:" & m_Qfx & ", rtcje:" & m_Tcje & "," & _
                            "ryjje:" & ryjje & ", rzhyjje:" & rzhyjje & ", rbxje:" & m_Bxje & ", rbcbxje:" & m_Bcbxje & " , rpbxje:" & m_Pbxje & "," & _
                            "rnbxje:" & m_Nbxje & ", rpbcbxje:" & m_Pbcbxje & ", rnbcbxje:" & m_Nbcbxje & ", rbcjs1:" & m_Bcjs1 & ", rbcjs2:" & m_Bcjs2 & "," & _
                            "rzhzf:" & m_Zhzf & ", rxjzf:" & m_rxjzf & ", rzhtk:" & m_Zhtk & ", rxjtk:" & m_Xjtk & ", rjbbcbxje:" & m_Jbbcbxje & "," & _
                            "rgwybcbxje:" & m_Gwybcbxje & ", rtsbxje:" & m_Tsbxje & ", rpjbbcbxje:" & m_Pjbbcbxje & ", rnjbbcbxje:" & m_Njbbcbxje & ", rpgwybcbxje:" & m_Pgwybcbxje & "," & _
                            "rngwybcbxje:" & m_Ngwybcbxje & ", rptsbxje:" & m_Ptsbxje & ", rntsbxje:" & m_Ntsbxje, tmp)
 
                Select Case tmp
                    Case 0
                        住院虚拟结算_冕宁 = "个人帐户;" & m_Zhzf & ";1"
                        住院虚拟结算_冕宁 = 住院虚拟结算_冕宁 & "|统筹基金;" & m_Bxje + m_Bcbxje + m_Jbbcbxje + m_Tsbxje & ";1"
                        住院虚拟结算_冕宁 = 住院虚拟结算_冕宁 & "|公务员补助;" & m_Gwybcbxje & ";1"
                    Case 1
                        strMsg = "个人信息无该社保号?"
                    Case 2
                        strMsg = "没有病人的住院信息?"
                    Case 3
                        strMsg = "网络不通?"
                    Case 4
                        strMsg = "费用计算错误?"
                    Case Else
                        strMsg = " 其它错误?"
                End Select
        Else
            住院虚拟结算_冕宁 = "个人帐户;" & 0 & ";1"
            住院虚拟结算_冕宁 = 住院虚拟结算_冕宁 & "|统筹基金;" & bxje + Bcbxje + Jbbcbxje + Tsbxje & ";1"
            住院虚拟结算_冕宁 = 住院虚拟结算_冕宁 & "|公务员补助;" & Gwybcbxje & ";1"
            
        End If
    End If
    
End Function


Public Function 收费员代码_冕宁() As String
'功能说明:查询当前人员是否在医保中心注册,如果没有则添加,返回人员编号
'调用过程清单及说明:
'    finduser  :查找医保中心是否有指定编号人员
'    inuser    :添加医师

Dim tmp As Double
Dim rsRy As New ADODB.Recordset '人员表
Dim blnReturn As Boolean '返回值

Dim str_SQL As String
str_SQL = "Select 接受培训 from 人员表 where id=[1]"
Set rsRy = zlDatabase.OpenSQLRecord(str_SQL, "人员表", UserInfo.ID)

If Nvl(rsRy.Fields!接受培训, "9") <> "1" Then
    blnReturn = finduser(UserInfo.编号)
    Call WriteBusinessLOG("finduser(查找人员)", "yyid:" & UserInfo.编号, IIf(blnReturn, "True", "False"))
    If blnReturn = False Then
        tmp = inuser(UserInfo.编号, UserInfo.姓名)
        Call WriteBusinessLOG("inuser(添加人员)", "yyid:" & UserInfo.编号 & ",user_name:" & UserInfo.姓名, tmp)
        str_SQL = "Update 人员表 set 接受培训='1' where id=" & UserInfo.ID
        gcnOracle.Execute str_SQL
        gcnOracle.Execute "Commit"
    End If
End If
收费员代码_冕宁 = UserInfo.编号

End Function



Public Function 医院医师代码_冕宁(STR姓名 As String) As String
'功能说明:查询是否有指定医师,如果没有则添加,返回医师编号
'调用过程清单及说明:
'    findys  :查找医保中心是否有指定编号医师
'    inys    :添加医师
'    findkb  :查找医保中心是否有指定编号科别
'    inkb    :添加科别

    Dim rsTemp As New ADODB.Recordset
    Dim tmp As Double
    Dim str医师编码 As String
    Dim str医师姓名 As String
    Dim lng人员ID As Long
    Dim str科别编号 As String
    Dim blnReturn As Boolean '返回值
    
    mstrCdSql = "Select ID,编号,姓名,个人简介 from 人员表 where 姓名=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrCdSql, "人员表", STR姓名)
    lng人员ID = rsTemp!ID
    str医师编码 = rsTemp!编号
    str医师姓名 = rsTemp!姓名
    
    If Nvl(rsTemp!个人简介, "9") <> "1" Then
    
        mstrCdSql = "select * from 部门表 where ID in (select 部门ID from 部门人员 where 缺省=1 and 人员ID=[1])"
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrCdSql, "部门表", lng人员ID)
        
        If Nvl(rsTemp!位置, "9") <> "1" Then
            blnReturn = findkb(rsTemp!编码)
            Call WriteBusinessLOG("findkb(查找科别", "yykbbh:" & rsTemp!编码, IIf(blnReturn, "True", "False"))
            If blnReturn = False Then
                tmp = inkb(rsTemp!编码, rsTemp!名称)
                Call WriteBusinessLOG("inkb(添加科室)", "yykbbh:" & rsTemp!编码 & ",mc:" & rsTemp!名称, tmp)
                mstrCdSql = "Update 部门表 set 位置='1' where id= " & rsTemp!ID
                gcnOracle.Execute mstrCdSql
                gcnOracle.Execute "Commit"
            End If
        End If
    
        str科别编号 = rsTemp!编码
        blnReturn = findys(str医师编码)
        Call WriteBusinessLOG("findys(查找医师)", "yyysbh:" & str医师编码, IIf(blnReturn, "True", "False"))
        If blnReturn = False Then
            tmp = inys(str医师编码, str科别编号, str医师姓名)
            Call WriteBusinessLOG("inys(添加医师)", "yyysbh:" & str医师编码 & ",yykbbh:" & str科别编号 & ",xm:" & str医师姓名, tmp)
            mstrCdSql = "Update 人员表 set 个人简介=1 where id=" & lng人员ID
            gcnOracle.Execute mstrCdSql
            gcnOracle.Execute "Commit"
        End If
        
    End If
    医院医师代码_冕宁 = str医师编码
    
End Function

Public Function 医院科室代码_冕宁(ByVal lng科室ID As Long) As String
'功能说明:查询是否有指定科室,如果没有则添加,返回科室编号
'调用过程清单及说明:
'    findkb  :查找医保中心是否有指定编号科别
'    inkb    :添加科别
    Dim str科别编号 As String
    Dim blnReturn As Boolean '返回值
    Dim tmp  As Double  '返回值
    
    Dim rsTemp As New ADODB.Recordset
    
    mstrCdSql = "select * from 部门表 where ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrCdSql, "部门表", lng科室ID)
    
    If Nvl(rsTemp!位置, "9") <> "1" Then
        blnReturn = findkb(rsTemp!编码)
        Call WriteBusinessLOG("findkb(查找科别", "yykbbh:" & rsTemp!编码, IIf(blnReturn, "True", "False"))
        If blnReturn = False Then
            tmp = inkb(rsTemp!编码, rsTemp!名称)
            Call WriteBusinessLOG("inkb(添加科室)", "yykbbh:" & rsTemp!编码 & ",mc:" & rsTemp!名称, tmp)
            mstrCdSql = "Update 部门表 set 位置='1' where id= " & lng科室ID
            gcnOracle.Execute mstrCdSql
        End If
    End If
    
    医院科室代码_冕宁 = rsTemp!编码

End Function












