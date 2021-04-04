Attribute VB_Name = "mdl涪陵"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)
 
Public Type str_Out             '出口参数结构
    errtext As String           '错误信息
    out1 As String              '第1个出口值
    out2 As String              '..2........
    out3 As String              '  .
    out4 As String              '  .
    out5 As String              '  .
    out6 As String              '  .
    out7 As String              '  .
    out8 As String              '  .
    out9 As String              '  .
    out10 As String             '  .
    out11 As String             '  .
    out12 As String             '  .
    out13 As String             '  .
    out14 As String             '  .
    out15 As String             '  .
    out16 As String             '  .
    out17 As String             '  .
    out18 As String             '  .
    out19 As String             '  .
    out20 As String             '  .
    out21 As String             '  .
    out22 As String             '  .
    out23 As String             '  .
    out24 As String             '  .
    out25 As String             '  .
    out26 As String             '  .
    out27 As String             '  .
    out28 As String             '  .
    out29 As String             '第29个出口值
    out30 As String             '第30个出口值
End Type

'所有入口参数类型均为字符串类型
'所有间接出口值均应由str_Out结构中获取
'所有涉及到日期或时间的参数均应写为"yyyy-MM-dd HH24:MI:SS"形式的字符串

'=========================================================================================================
'功能说明:查询药品,诊治项目,床位,疾病等自付比例或可报销额。
'入口参数:医保机构编码,医院编号,医院项目编码,查询类别,人员类别
'间接出口参数:自付比例或可报销额,标志
'    标志说明:1---自付比例,2---可报销额(表示最多可报销额为该金额,当该项费用小于可报销时,为该项费用)
'             3---自付比例|可报销额(表示自付比例为该项比例并且最多报销费用额为可报销额,大于部分全部自费)
'             4---没有匹配(全部自费),5---医院没有该项目(全部自费)
'参数含义说明:
'    查询类别: 1---药品,2---诊治项目,3---床位,4---疾病
'    人员类别: 01---在职人员,02---退休人员
'=========================================================================================================
Public Declare Function fl_readzfbl Lib "cxybclient.dll" Alias "readzfbl" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal stryyxmbm As String, ByVal strcxlb As String, ByVal strrrlb As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'功能说明:刷新IC卡,并显示ic卡的详细信息
'入口参数:医保机构编码,医院编号,显示标志
'间接出口参数:医保机构编码,个人帐号,ic卡号,身份证号码,姓名,性别,单位编码,单位名称,出生日期,人员类别,
'             IC卡余额,起付标准,基本医疗最高限额,基本医疗累计支出,大病医疗累计支出
'参数含义说明:
'        性别: 1---女,0---男
'    人员类别: 01---在职,02---退休,03---下岗
'    显示标志: 1---显示,0---不显示(显示时,接口客户端将弹出对话框显示IC卡中信息,包含简介出口值中所有信息)
'=========================================================================================================
Public Declare Function fl_readicxx Lib "cxybclient.dll" Alias "readicxx" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strxxbz As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'功能说明:医保病人住院如果需要审批可以调用此函数,将病人信息传到医保,由医保审批。
'入口参数:医保机构编码,医院编号,医保就诊编号,医院疾病编码,医院疾病名称,申请日期,原因
'         急诊标志改为转诊转院, 医生姓名,特病标志
'间接出口参数:无
'参数含义说明:
'    急诊标志: 0---否,1---是
'    特病标志: 0---否,1---是
'=========================================================================================================
Public Declare Function fl_request Lib "cxybclient.dll" Alias "request" (ByVal strybjgbm As String, _
    ByVal stryybm As String, ByVal strybjzbh As String, ByVal stryyjbbm As String, _
    ByVal stryyjbmc As String, ByVal strsjrq As String, ByVal strsjyy As String, _
    ByVal strzybz As String, ByVal strysxm As String, ByVal strtbbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:刷新IC卡,显示IC卡信息,门诊或住院登记,并返回病人入院唯一标识编号(医保就诊编号)。
'入口参数:医保机构编码,医院编号,标志,操作员名称,生成日期,是否大出血
'间接出口参数:医保就诊编号,医保机构编码,个人帐号,ic卡号,身份证号码,姓名,性别,单位编码,单位名称,出生日期,
'             人员类别,IC卡余额,起付标准,基本医疗最高限额,基本医疗累计支出,大病医疗累计支出
'参数含义说明:
'        标志:0---门诊,1---住院
'  是否大出血:接口函数未说明
'=========================================================================================================
Public Declare Function fl_reg Lib "cxybclient.dll" Alias "reg" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strbz As String, ByVal strczymc As String, ByVal strscrq As String, ByVal strsfdcx As String, _
    strOut As str_Out) As Boolean
    
'=========================================================================================================
'功能说明:写处方信息
'入口参数:医保机构编码,医院编号,医保就诊编号,处方编号,医院疾病编码,医院疾病名称,特病标志,医生姓名
'         录入员姓名,科室编号,科室名称,处方日期
'间接出口参数:无
'参数含义说明:
'    特病标志:0---否,1---是
'=========================================================================================================
Public Declare Function fl_wrecipe Lib "cxybclient.dll" Alias "wrecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal stryyjbbm As String, ByVal stryyjbmc As String, _
    ByVal strtbbz As String, ByVal strysxm As String, ByVal strlryxm As String, ByVal strksbh As String, _
    ByVal strksmc As String, ByVal strcfrq As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'功能说明:修改处方信息，以医保就诊编号，处方编号为条件修改(更新)记录
'入口参数:医保机构编码,医院编号,医保就诊编号,处方编号,医院疾病编码,医院疾病名称,特病标志,医生姓名
'         录入员姓名,科室编号,科室名称,处方日期
'间接出口参数:无
'参数含义说明:
'    特病标志:0---否,1---是
'=========================================================================================================
Public Declare Function fl_urecipe Lib "cxybclient.dll" Alias "urecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal stryyjbbm As String, ByVal stryyjbmc As String, _
    ByVal strtbbz As String, ByVal strysxm As String, ByVal strlryxm As String, ByVal strksbh As String, _
    ByVal strksmc As String, ByVal strcfrq As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:删除处方信息
'入口参数:医保机构编码,医院编号,医保就诊编号,处方编号
'间接出口参数:无
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_drecipe Lib "cxybclient.dll" Alias "drecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:写处方明细信息
'入口参数:医保机构编码,医院编号,医保就诊编号,处方编号,明细序号,医院明细编码,医院明细名称,产地,规格,类别,
'         单位,单价,数量,时间,录入人,标志
'间接出口参数:无
'参数含义说明:
'        标志:1---药品,2---诊治项目,3---床位费
'医院明细编码:为医院药品,诊治项目,床位费编码(对应标志)
'医院明细名称:为医院药品,诊治项目,床位费名称(对应标志)
'=========================================================================================================
Public Declare Function fl_wdetails Lib "cxybclient.dll" Alias "wdetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal strmxxh As String, ByVal stryymxbm As String, _
    ByVal stryymxmc As String, ByVal strcd As String, ByVal strgg As String, ByVal strlb As String, _
    ByVal strdw As String, ByVal strdj As String, ByVal strsl As String, ByVal strsj As String, _
    ByVal strlrr As String, ByVal strbz As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'功能说明:修改处方明细信息,以医保就诊编号,处方编号,明细序号为主键修改
'入口参数:医保机构编码,医院编号,医保就诊编号,处方编号,明细序号,医院明细编码,医院明细名称,产地,规格,类别,
'         单位,单价,数量,时间,录入人,标志
'间接出口参数:无
'参数含义说明:(同上)
'=========================================================================================================
Public Declare Function fl_udetails Lib "cxybclient.dll" Alias "udetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal strmxxh As String, ByVal stryymxbm As String, _
    ByVal stryymxmc As String, ByVal strcd As String, ByVal strgg As String, ByVal strlb As String, _
    ByVal strdw As String, ByVal strdj As String, ByVal strsl As String, ByVal strsj As String, _
    ByVal strlrr As String, ByVal strbz As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'功能说明:删除处方明细信息
'入口参数:医保机构编码,医院编号,医保就诊编号,处方编号,明细序号
'间接出口参数:无
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_ddetails Lib "cxybclient.dll" Alias "ddetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal strmxxh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:对医保病人的费用进行预结算
'入口参数:医保机构编码,医院编号,医保就诊编号,出院日期,操作员,结算标志,显示标志
'间接出口参数:费用合计,特殊病种费用,本次本年帐户支付,本次历年帐户支付,累计分段自付,统筹金支付,起付段支付,
'             单位支付,自费费用,特检先自付,特治先自付,特检费用,特治费用,补充医疗保险支付,本次统筹记入累计,
'             补充医疗记入累计,门诊统筹记入累计,未报销费用,医保支付,个人现金支付
'参数含义说明:
'    显示标志:0---不显示,1---显示
'    结算标志:1---试结算,2---中途结算
'=========================================================================================================
Public Declare Function fl_pcalc Lib "cxybclient.dll" Alias "pcalc" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcyrq As String, ByVal strczy As String, ByVal strjsbz As String, _
    ByVal strxsbz As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'功能说明:正式结算
'入口参数:医保机构编码,医院编号,医保就诊编号,出院日期,操作员,显示标志
'间接出口参数:费用合计,特殊病种费用,本次本年帐户支付,本次历年帐户支付,累计分段自付,统筹金支付,起付段支付,
'             单位支付,自费费用,特检先自付,特治先自付,特检费用,特治费用,补充医疗保险支付,本次统筹记入累计,
'             补充医疗记入累计,门诊统筹记入累计,未报销费用,医保支付,个人现金支付,个人帐户余额
'参数含义说明:
'    显示标志:0---不显示,1---显示
'=========================================================================================================
Public Declare Function fl_calc Lib "cxybclient.dll" Alias "calc" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcyrq As String, ByVal strczy As String, _
    ByVal strxsbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:住院情况下,取消正式结算,返回到结算前状态;门诊情况下,生成红字单据,冲掉门诊记录
'入口参数:医保机构编码,医院编号,医保就诊编号,显示标志
'间接出口参数:无
'参数含义说明:
'    显示标志:0---不显示,1---显示
'=========================================================================================================
Public Declare Function fl_rollbackcalc Lib "cxybclient.dll" Alias "rollbackcalc" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strxsbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:删除该医保就诊编号的所有信息,包括入院登记,处方,处方明细等。但是如果已正式结算,则不能使用该函数
'入口参数:医保机构编码,医院编号,医保就诊编号
'间接出口参数:无
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_dall Lib "cxybclient.dll" Alias "dall" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:检测该处方是否能删除或修改
'入口参数:医保机构编码,医院编号,医保就诊编号,处方编号
'间接出口参数:无
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_canupdaterecipe Lib "cxybclient.dll" Alias "canupdaterecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:检测该处方明细是否能删除或修改
'入口参数:医保机构编码,医院编号,医保就诊编号,处方编号,明细序号
'间接出口参数:无
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_canupdatedetails Lib "cxybclient.dll" Alias "canupdatedetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal strmxbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:检测是否能够回滚,住院情况必须使用此函数判断
'入口参数:医保机构编码,医院编号,医保就诊编号
'间接出口参数:无
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_canrollback Lib "cxybclient.dll" Alias "canrollback" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:检测是否更新了医保自负比例,检索是否有更新时间大于最后一次检索时间的药品,诊治项目,疾病,床位
'入口参数:医保机构编码,医院编号,类型标志
'间接出口参数:无
'参数含义说明:
'    类型标志:1---药品,2---诊治项目,3---疾病,4---床位
'=========================================================================================================
Public Declare Function fl_havenewzfbl Lib "cxybclient.dll" Alias "havenewzfbl" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strlxbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:返回医保服务器时间
'入口参数:无
'间接出口参数:医保服务器时间
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_getsystime Lib "cxybclient.dll" Alias "getsystime" (strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:返回医保机构编码(需医保中心系统卡或病人IC卡)
'入口参数:无
'间接出口参数:医保机构编码,医院编码
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_getybjgbm Lib "cxybclient.dll" Alias "getybjgbm" (strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:返回在院医疗单号(需医保中心系统卡或病人IC卡)
'入口参数:医保机构编码,医院编号,个人帐号
'间接出口参数:在院医疗单号,医院编码
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_getlastzyxx Lib "cxybclient.dll" Alias "getlastzyxx" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strgrzh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:写或修改与医保关联的医院基本信息
'入口参数:类别,医院信息编码,医院信息名称,其它
'间接出口参数:无
'参数含义说明:
'        类别:1---药品,2---诊治项目,3---床位费,0---疾病
'        其它:若类别为1,该项为(01---国产,02---合资,03---进口);
'             若类别为2,该项为国家编码;若类别为其它,该项为空
'=========================================================================================================
Public Declare Function fl_wyyglxx Lib "cxybclient.dll" Alias "wyyglxx" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strlb As String, ByVal stryyxxbm As String, _
    ByVal stryyxxmc As String, ByVal strqt As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:修改用户的IC卡密码
'入口参数:无
'间接出口参数:无
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_changePassword Lib "cxybclient.dll" Alias "changepassword" (ByVal strybjgbm As String, _
    ByVal stryybm As String, strOut As str_Out) As Boolean

'=========================================================================================================
'功能说明:校验系统卡
'入口参数:无
'间接出口参数:无
'参数含义说明:
'=========================================================================================================
Public Declare Function fl_checkxtk Lib "cxybclient.dll" Alias "checkxtk" (strOut As str_Out) As Boolean

Public gstrOutPara As str_Out
Private mblnReturn As Boolean

Public Sub initType()
    With gstrOutPara
        .errtext = Space(1024)
        .out1 = Space(50)
        .out2 = Space(50)
        .out3 = Space(50)
        .out4 = Space(50)
        .out5 = Space(50)
        .out6 = Space(50)
        .out7 = Space(50)
        .out8 = Space(50)
        .out9 = Space(50)
        .out10 = Space(50)
        .out11 = Space(50)
        .out12 = Space(50)
        .out13 = Space(50)
        .out14 = Space(50)
        .out15 = Space(50)
        .out16 = Space(50)
        .out17 = Space(50)
        .out18 = Space(50)
        .out19 = Space(50)
        .out20 = Space(50)
        .out21 = Space(50)
        .out22 = Space(50)
        .out23 = Space(50)
        .out24 = Space(50)
        .out25 = Space(50)
        .out26 = Space(50)
        .out27 = Space(50)
        .out28 = Space(50)
        .out29 = Space(50)
        .out30 = Space(50)
    End With
End Sub

Public Sub TrimType()
    With gstrOutPara
        .errtext = Trim(.errtext)
        .out1 = Trim(.out1)
        .out2 = Trim(.out2)
        .out3 = Trim(.out3)
        .out4 = Trim(.out4)
        .out5 = Trim(.out5)
        .out6 = Trim(.out6)
        .out7 = Trim(.out7)
        .out8 = Trim(.out8)
        .out9 = Trim(.out9)
        .out10 = Trim(.out10)
        .out11 = Trim(.out11)
        .out12 = Trim(.out12)
        .out13 = Trim(.out13)
        .out14 = Trim(.out14)
        .out15 = Trim(.out15)
        .out16 = Trim(.out16)
        .out17 = Trim(.out17)
        .out18 = Trim(.out18)
        .out19 = Trim(.out19)
        .out20 = Trim(.out20)
        .out21 = Trim(.out21)
        .out22 = Trim(.out22)
        .out23 = Trim(.out23)
        .out24 = Trim(.out24)
        .out25 = Trim(.out25)
        .out26 = Trim(.out26)
        .out27 = Trim(.out27)
        .out28 = Trim(.out28)
        .out29 = Trim(.out29)
        .out30 = Trim(.out30)
    End With
End Sub

Public Function 医保初始化_涪陵() As Boolean
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取参数", TYPE_涪陵)
    
    With rsTemp
        Do While Not .EOF
            If !参数名 = "医保机构编码" Then
                gstr医保机构编码 = Nvl(!参数值)
            ElseIf !参数名 = "医院编码" Then
                gstr医院编码 = Nvl(!参数值)
            End If
            .MoveNext
        Loop
    End With
    
    If gstr医保机构编码 = "" Then
        MsgBox "请运行保险类别管理并设置保险参数后再使用本接口！[医保机构编码]", vbInformation, gstrSysName
        Exit Function
    End If
    
    医保初始化_涪陵 = True
End Function

Public Function 身份标识_涪陵(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmIDentified As New frmIdentify涪陵
    Dim strPatiInfo As String, cur余额 As Currency, str就诊编号 As String
    Dim arr, datCurr As Date, str门诊号 As String
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    If lng病人ID = 0 Then
        strTemp = "0"
    Else
        gstrSQL = "Select * From 保险帐户 where 病人id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
        If rsTemp.EOF Then
            strTemp = "0"
        Else
            strTemp = Nvl(rsTemp!退休证号, 0)
        End If
    End If
    
'    WriteInfo "调用窗体，获取医保病人信息"
    strPatiInfo = frmIDentified.GetPatient(bytType, strTemp)
'    WriteInfo "医保病人信息：" & strPatiInfo
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '建立病人档案信息，传入格式：
        '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
        '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
        '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
'        WriteInfo "写医保病人信息"
        lng病人ID = BuildPatiInfo(bytType, strPatiInfo, lng病人ID, TYPE_涪陵)
'        WriteInfo "写医保病人信息完成"
        '返回格式:中间插入病人ID
        strPatiInfo = frmIDentified.mstrPatient & lng病人ID & ";" & frmIDentified.mstrOther
        str就诊编号 = frmIDentified.mstr就诊编号
        '写入就诊编号
'        WriteInfo "写保险帐户信息"
        If bytType = 0 Or bytType = 5 Then
            gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_涪陵 & ",'顺序号','''" & str就诊编号 & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_涪陵")
            gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_涪陵 & ",'退休证号','''" & CLng(strTemp) + 1 & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_涪陵")
        End If
        gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_涪陵 & ",'灰度级','''" & frmIDentified.mstr起付标准 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_涪陵")
'        WriteInfo "完成身份验证"
        Unload frmIDentified
    Else
        身份标识_涪陵 = ""
        MsgBox "医保病人信息提取失败", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    身份标识_涪陵 = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_涪陵 = ""
End Function

Public Function 个人余额_涪陵(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select 帐户余额 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取个人帐户余额", lng病人ID, TYPE_涪陵)
    
    If rsTemp.EOF Then
        个人余额_涪陵 = 0
    Else
        个人余额_涪陵 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
    End If
End Function

Public Function 门诊虚拟结算_涪陵(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'参数：rsDetail     费用明细(传入)
'    病人ID         adBigInt, 19, adFldIsNullable
'    收费类别       adVarChar, 2, adFldIsNullable
'    收据费目       adVarChar, 20, adFldIsNullable
'    计算单位       adVarChar, 6, adFldIsNullable
'    开单人         adVarChar, 20, adFldIsNullable
'    收费细目ID     adBigInt, 19, adFldIsNullable
'    数量           adSingle, 15, adFldIsNullable
'    单价           adSingle, 15, adFldIsNullable
'    实收金额       adSingle, 15, adFldIsNullable
'    统筹金额       adSingle, 15, adFldIsNullable
'    保险支付大类ID adBigInt, 19, adFldIsNullable
'    是否医保       adBigInt, 19, adFldIsNullable
'    摘要           adVarChar, 200, adFldIsNullable
'    是否急诊       adBigInt, 19, adFldIsNullable
'    str结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim cur自付 As Currency, cur报销 As Currency, cur余额 As Currency, lngErr As Long
    Dim lng病人ID As Long, rsTemp As New ADODB.Recordset, str报销明细 As String
    Dim strTemp As String, curTemp As Currency, str自付比例 As String, str可报销额 As String
    
    On Error GoTo errHandle
    If rs明细.RecordCount = 0 Then
        MsgBox "没有费用，不能进行预结算。", vbInformation, gstrSysName
        门诊虚拟结算_涪陵 = False
        Exit Function
    End If
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID"): lngErr = 1
    cur自付 = 0: cur报销 = 0: lngErr = 2
    gstrSQL = "Select * from 保险帐户 where 病人id=[1]"
    lngErr = 3
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保预结算", lng病人ID)
    lngErr = 4
    cur余额 = rsTemp!帐户余额: lngErr = 5
    strTemp = IIf(rsTemp!在职 = "2", "02", "01"): lngErr = 4
    str报销明细 = ""
    While Not rs明细.EOF
        gstrSQL = "select * from 收费细目 where id=[1]"
        lngErr = 6
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保预结算", CLng(rs明细!收费细目ID))
        lngErr = 7
        
        '获取收费细目的自付比例
        initType
        mblnReturn = fl_readzfbl(gstr医保机构编码, gstr医院编码, rsTemp!类别 & "_" & rsTemp!ID, _
            IIf(rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7", "1", IIf(rsTemp!类别 = "J", "3", "2")), _
            strTemp, gstrOutPara): lngErr = 8
        TrimType
        
        If mblnReturn = False Then
            MsgBox "在获取项目[" & rsTemp!ID & "]的自付比例时，医保接口返回以下错误：" & Chr(13) & Chr(10) & gstrOutPara.errtext
            门诊虚拟结算_涪陵 = False
            Exit Function
        End If
        Select Case gstrOutPara.out2
            Case "1"            '返回为自付比例
                curTemp = rs明细!实收金额 * (1 - CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0))): lngErr = 9
            Case "2"            '返回为报销限额
                curTemp = IIf(rs明细!实收金额 > CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0)), CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0)), rs明细!实收金额): lngErr = 10
            Case "3"            '按自付比例计算报销金额，若大于可报销额，则取可报销额
                str自付比例 = Left(gstrOutPara.out1, InStr(gstrOutPara.out1, "|") - 1): lngErr = 11
                str可报销额 = Mid(gstrOutPara.out1, InStr(gstrOutPara.out1, "|") + 1): lngErr = 12
                str自付比例 = IIf(IsNumeric(str自付比例), str自付比例, 0): lngErr = 13
                str可报销额 = IIf(IsNumeric(str可报销额), str可报销额, 0): lngErr = 14
                curTemp = rs明细!实收金额 * (1 - CCur(str自付比例)): lngErr = 15
                curTemp = IIf(curTemp > CCur(str可报销额), CCur(str可报销额), curTemp): lngErr = 16
            Case "4", "5"       '自付比例为100%
                curTemp = 0
        End Select
        str报销明细 = str报销明细 & "项目名称:" & rsTemp!名称 & "[" & rsTemp!类别 & "_" & rsTemp!ID & "]　　自付比例:[" & _
            gstrOutPara.out2 & "]" & gstrOutPara.out1 & "　　报销金额:" & curTemp & Chr(13) & Chr(10)
        
        cur报销 = cur报销 + curTemp: lngErr = 17
        cur自付 = rs明细!实收金额 - curTemp: lngErr = 18
        rs明细.MoveNext: lngErr = 19
    Wend
    
    '如果报销额大于帐户余额，则允许从帐户中支付的最大额为帐户余额，多余部分计入现金支付
    '余额需要保留1元作为保底
    If cur报销 > cur余额 - 1 Then
        curTemp = cur报销 - (cur余额 - 1): lngErr = 20
        cur报销 = cur余额 - 1: lngErr = 21
        cur自付 = cur自付 + curTemp: lngErr = 22
    End If
    
'    MsgBox str报销明细, vbInformation, "报销明细"
    
    str结算方式 = "个人帐户;" & cur报销 & ";0": lngErr = 23
    门诊虚拟结算_涪陵 = True
    Exit Function
errHandle:
    MsgBox "错误出现在[门诊预结算]模块，第" & lngErr & "行，错误信息：" & Chr(13) & Chr(10) & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Public Function 门诊结算_涪陵(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency, Optional ByRef strAdvance As String) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
'        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str操作员 As String, datCurr As Date, str就诊编号 As String, str结算方式 As String
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur起付线 As Currency, cur基本统筹限额 As Currency, cur公务员补助 As Currency
    Dim cur大额统筹限额 As Currency, cur基数自付 As Currency, cur余额 As Currency
    Dim cur发生费用 As Currency, cur先自付 As Currency, lng病种ID As Long, str公务员类别 As String
    Dim blnOld As Boolean
    
    If gstr医保机构编码 = "" Then
        Err.Raise 9000, gstrSysName, "准备读取医保机构编码，请插入系统卡或病人卡！"
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    
    On Error GoTo errHandle
    gstrSQL = "Select * From 门诊费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    If rs明细.EOF = True Then
        Err.Raise 9000, gstrSysName, "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str操作员 = ToVarchar(IIf(IsNull(rs明细("操作员姓名")), UserInfo.姓名, rs明细("操作员姓名")), 20)
    '强制选择病种
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & TYPE_涪陵

    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种")
    If rsTemp.State = 1 Then
        lng病种ID = rsTemp("ID")
    Else
        门诊结算_涪陵 = False
        Exit Function
    End If
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_涪陵 & ",'病种ID'," & lng病种ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_涪陵")

    '需要先上传费用明细
    If 费用明细传递_涪陵(lng结帐ID) = False Then
        Err.Raise 9000, gstrSysName, "记录上传时发生错误", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select nvl(顺序号,0) as 顺序号,病种id From 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_涪陵
    Set rsTemp = gcnOracle.Execute(gstrSQL)
    lng病种ID = rsTemp!病种ID
    str就诊编号 = rsTemp!顺序号
    
    '医保机构编码, 医院编号, 医保就诊编号， 出院日期，操作员，显示标志
    datCurr = zlDatabase.Currentdate
    initType
'    mblnReturn = fl_pcalc(gstr医保机构编码, gstr医院编码, str就诊编号, Format(datCurr, "yyyy-MM-dd"), str操作员, "1", "0", gstrOutPara)
    mblnReturn = fl_calc(gstr医保机构编码, gstr医院编码, str就诊编号, Format(datCurr, "yyyy-MM-dd"), str操作员, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext, vbInformation, gstrSysName
        门诊结算_涪陵 = False
        Exit Function
    End If
    '间接出口参数:1费用合计,2特殊病种费用,3本次本年帐户支付,4本次历年帐户支付,5累计分段自付,6统筹金支付,7起付段支付,
    '             8单位支付,9自费费用,10特检先自付,11特治先自付,12特检费用,13特治费用,14补充医疗保险支付,15本次统筹记入累计,
    '             16补充医疗记入累计,17门诊统筹记入累计,18未报销费用,19医保支付,20个人现金支付,21个人帐户余额
    '增加出口参数:22公务员类别;23公务员补助
    '获取个人帐户支付和个人现金支付
    cur个人帐户 = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
    cur余额 = CCur(gstrOutPara.out21)
    cur全自付 = CCur(gstrOutPara.out20) + CCur(cur个人帐户)
    cur发生费用 = CCur(gstrOutPara.out1)
    cur先自付 = CCur(gstrOutPara.out10) + CCur(gstrOutPara.out11)
    cur公务员补助 = CCur(gstrOutPara.out23)
    str公务员类别 = gstrOutPara.out22
    
    '强制赋值，因为无预结算功能，避免程序显示个人帐户支付，正式结算时按此结果保存预交记录，而正式结算时无个人帐户支付确未修正的情况
    str结算方式 = "||个人帐户|" & cur个人帐户 & "||公务员补助|" & cur公务员补助
    
    '如果存在
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        #If gverControl < 2 Then
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
        #Else
            strAdvance = str结算方式
            gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
    End If
    #If gverControl < 2 Then
        blnOld = True
        frm结算信息.ShowME (lng结帐ID)
    #End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_涪陵, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & Get病人ID(CStr(str医保号), CStr(TYPE_涪陵)) & _
            "," & TYPE_涪陵 & "," & Year(datCurr) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & _
            cur起付线 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_涪陵 & "," & _
            Get病人ID(CStr(str医保号), CStr(TYPE_涪陵)) & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            cur发生费用 & "," & cur全自付 & "," & cur先自付 & ",NULL,NULL," & cur公务员补助 & ",NULL," & _
            cur个人帐户 & ",'" & str公务员类别 & "',NULL,NULL,'" & str就诊编号 & "'" & IIf(blnOld, "", ",1") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    门诊结算_涪陵 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 费用明细传递_涪陵(lng结帐ID As Long, Optional rs明细IN As ADODB.Recordset = Nothing) As Boolean
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str操作员 As String, cur发生费用, str就诊编号 As String, strBillNO As String
    Dim lng病种ID As Long, str病种名称 As String, str病种编码 As String, int特病标志 As Integer
    Dim str科室编号 As String, str科室名称 As String, lng科室ID As Long
    Dim str明细编码 As String, str明细名称 As String, str处方号 As String
    Dim strTemp As String, iLoop As Long
'    病人ID         adBigInt, 19, adFldIsNullable
'    收费类别       adVarChar, 2, adFldIsNullable
'    收据费目       adVarChar, 20, adFldIsNullable
'    计算单位       adVarChar, 6, adFldIsNullable
'    开单人         adVarChar, 20, adFldIsNullable
'    收费细目ID     adBigInt, 19, adFldIsNullable
'    数量           adSingle, 15, adFldIsNullable
'    单价           adSingle, 15, adFldIsNullable
'    实收金额       adSingle, 15, adFldIsNullable
'    统筹金额       adSingle, 15, adFldIsNullable
'    保险支付大类ID adBigInt, 19, adFldIsNullable
'    是否医保       adBigInt, 19, adFldIsNullable
'    摘要           adVarChar, 200, adFldIsNullable
'    是否急诊       adBigInt, 19, adFldIsNullable
'    str结算方式  "报销方式;金额;是否允许修改|...."
    
    On Error GoTo errHandle
    
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    If rs明细IN Is Nothing Then
        gstrSQL = "Select * From 门诊费用记录 Where 记录状态<>0 And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 and 结帐ID=[1]"
        Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    Else
        Set rs明细 = rs明细IN.Clone
    End If
    If rs明细.EOF = True Then
'        MsgBox "没有需要上传的收费记录", vbExclamation, gstrSysName
        费用明细传递_涪陵 = True
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str操作员 = ToVarchar(UserInfo.姓名, 20)
    
'    gstrSQL = "select max(主页ID) as 主页ID from 病案主页 where 病人ID =" & lng病人ID
'    Call OpenRecordset(rsTemp, gstrsysname)
'    strBillNo = CStr(lng病人ID) & "_" & CStr(rsTemp("主页ID"))
    gstrSQL = "Select nvl(顺序号,0) as 顺序号,病种ID,中心,退休证号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_涪陵)
    str处方号 = rsTemp!退休证号
    str就诊编号 = rsTemp!顺序号
    lng病种ID = Nvl(rsTemp!病种ID, 0)
'    gstr医保机构编码 = rsTemp!中心
    gstrSQL = "Select * From 保险病种 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病种ID)
    If rsTemp.EOF Then
        str病种名称 = "未知"
        str病种编码 = "0"
        int特病标志 = 0
    Else
        str病种名称 = rsTemp!名称
        str病种编码 = rsTemp!ID
        int特病标志 = IIf(rsTemp!类别 = 2, 1, 0)
    End If
    lng科室ID = rs明细!开单部门ID
    gstrSQL = "Select * From 部门表 where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng科室ID)
    str科室编号 = rsTemp!编码
    str科室名称 = rsTemp!名称
    
'    str处方号 = NVL(rs明细!主页ID, 0) & Right(rs明细!NO, 2)
    '写处方信息
    initType
    mblnReturn = fl_wrecipe(gstr医保机构编码, gstr医院编码, str就诊编号, str处方号, str病种编码, str病种名称, _
                         int特病标志, Nvl(rs明细!开单人, rs明细!划价人), Nvl(rs明细!操作员姓名, UserInfo.姓名), str科室编号, _
                         str科室名称, Format(rs明细!登记时间, "yyyy-MM-dd"), gstrOutPara)
    TrimType
    If mblnReturn = False Then
        If InStr(gstrOutPara.errtext, "(YBYY.PRI_QTYL") > 0 Then
            费用明细传递_涪陵 = True
        Else
            MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
            费用明细传递_涪陵 = False
            Exit Function
        End If
    End If
    
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_涪陵 & ",'退休证号','" & CLng(str处方号) + 1 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    iLoop = 1
    '写处方明细
    Do Until rs明细.EOF
        gstrSQL = "Select * From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        str明细编码 = rsTemp!ID
        str明细名称 = rsTemp!名称
        initType
        If InStr(Nvl(rsTemp!规格, " "), "┆") > 0 Then
            strTemp = Left(rsTemp!规格, InStr(rsTemp!规格, "┆") - 1)
        Else
            strTemp = Nvl(rsTemp!规格, " ")
        End If
'入口参数:医保机构编码,医院编号,医保就诊编号,处方编号,明细序号,医院明细编码,医院明细名称,产地,规格,类别,
'         单位,单价,数量,时间,录入人,标志
        If (IsNull(rs明细!是否上传) Or rs明细!是否上传 = 0) And rs明细!实收金额 <> 0 And rs明细!付数 * rs明细!数次 <> 0 Then
            mblnReturn = fl_wdetails(gstr医保机构编码, gstr医院编码, str就诊编号, str处方号, iLoop, _
                rsTemp!类别 & "_" & rsTemp!ID, rsTemp!名称, " ", ToVarchar(strTemp, 10), Nvl(rsTemp!费用类型, " "), Nvl(rsTemp!计算单位, " "), Format(rs明细!实收金额 / (rs明细!付数 * rs明细!数次), "#.####"), _
                rs明细!付数 * rs明细!数次, Format(rs明细!登记时间, "yyyy-MM-dd"), Nvl(rs明细!操作员姓名, UserInfo.姓名), _
                IIf(rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7", "1", IIf(rsTemp!类别 = "J", "3", "2")), gstrOutPara)
        End If
        TrimType
        If mblnReturn = False Then
            MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
            费用明细传递_涪陵 = False
            Exit Function
        End If
        
        gstrSQL = "zl_病人记帐记录_上传 ('" & rs明细!ID & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        rs明细.MoveNext
        iLoop = iLoop + 1
    Loop
    费用明细传递_涪陵 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_涪陵(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur票据总金额 As Currency, lngErr As Long
    Dim datCurr As Date
    
    If gstr医保机构编码 = "" Then
        Err.Raise 9000, gstrSysName, "准备读取医保机构编码，请插入系统卡或病人卡"
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 门诊费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    lngErr = 1
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "Select * from 保险帐户 where 病人ID=[2]"
    lngErr = 2
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    str就诊编号 = Nvl(rsTemp!顺序号, "0")
'    gstr医保机构编码 = rsTemp!中心
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[3]"
    lngErr = 3
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    lng冲销ID = rsTemp("结帐ID")
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    lngErr = 4
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_涪陵, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    If IsNull(rsTemp!备注) Then
        Err.Raise 9000, gstrSysName, "该单据的就诊编号丢失，不能作废。"
        Exit Function
    End If
    str就诊编号 = rsTemp!备注
    '调用接口数冲销
    initType
    mblnReturn = fl_canrollback(gstr医保机构编码, gstr医院编码, str就诊编号, gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, "判断是否可以冲销时，医保端返回以下信息，退费不能继续。" & Chr(13) & Chr(10) & gstrOutPara.errtext
        Exit Function
    End If
    initType
    mblnReturn = fl_rollbackcalc(gstr医保机构编码, gstr医院编码, str就诊编号, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext
        Exit Function
    End If
    '帐户年度信息
    Call Get帐户信息(TYPE_涪陵, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计): lngErr = 5
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_涪陵 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")": lngErr = 6
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_涪陵 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        cur个人帐户 * -1 & ",NULL,NULL,NULL,'" & str就诊编号 & "')": lngErr = 7
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    门诊结算冲销_涪陵 = True
    Exit Function
errHandle:
    ErrMsgBox "错误发生在[门诊结算冲销]模块，第" & lngErr & "行，错误信息：" & Chr(13) & Chr(10) & Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 入院登记_涪陵(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str病种 As String, str病种编码 As String
    Dim rsTmp As New ADODB.Recordset, str就诊编号 As String, datCurr As Date
    Dim lng病种ID As Long
    Dim str转入 As String   '0-否;1-是
    
    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.入院日期,A.入院方式,B.住院号,D.名称 as 住院科室,A.入院病床,A.住院医师,C.卡号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTmp.RecordCount <> 0 Then
        datCurr = rsTmp!入院日期
    End If
    str转入 = IIf(Nvl(rsTmp!入院方式, "普通") = "转入", "1", "0")
    Call DebugTool("入院方式[" & Nvl(rsTmp!入院方式, "普通") & "]")
    
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID)   '入院诊断
    If rsTmp.BOF Then 入院登记_涪陵 = False: Exit Function
    '强制选择病种
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & TYPE_涪陵
    
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种")
    If rsTemp.State = 1 Then
        lng病种ID = rsTemp("ID")
        str病种 = rsTemp!名称
        str病种编码 = rsTemp!ID
    Else
        入院登记_涪陵 = False
        Exit Function
    End If
    
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If

    initType
    mblnReturn = fl_reg(gstr医保机构编码, gstr医院编码, 1, UserInfo.姓名, Format(datCurr, "yyyy-MM-dd"), "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        入院登记_涪陵 = False
        Exit Function
    End If
    str就诊编号 = gstrOutPara.out1
    
    initType
'入口参数:医保机构编码,医院编号,医保就诊编号,医院疾病编码,医院疾病名称,申请日期,原因
'         急诊标志, 医生姓名,特病标志
    '进行入院请求
    Call DebugTool("Request函数入参：" & gstr医保机构编码 & "," & gstr医院编码 & "," & str就诊编号 & "," & str病种编码 & "," & str病种 & "," & Format(datCurr, "yyyy-MM-dd") & "," & _
            strInNote & "," & str转入 & "," & UserInfo.姓名 & "," & "0")
    mblnReturn = fl_request(gstr医保机构编码, gstr医院编码, str就诊编号, str病种编码, str病种, Format(datCurr, "yyyy-MM-dd"), _
            strInNote, str转入, UserInfo.姓名, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        入院登记_涪陵 = False
        Exit Function
    End If
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_涪陵 & ",'顺序号'," & str就诊编号 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_涪陵")
    
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_涪陵 & ",'病种ID'," & lng病种ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_涪陵")
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_涪陵 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_涪陵 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_涪陵 = False
End Function

Public Function 住院结算冲销_涪陵(lng结帐ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额

    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String, lng病人ID As Long
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur票据总金额 As Currency
    Dim datCurr As Date, cur个人帐户 As Currency

    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate

    gstrSQL = "Select 病人ID,结帐金额 From 住院费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)

    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop

    gstrSQL = "Select * from 保险帐户 where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_涪陵)
    str就诊编号 = Nvl(rsTemp!顺序号, "0")

    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B" & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng冲销ID = rsTemp("ID")

    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_涪陵, lng结帐ID)

    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If

    If IsNull(rsTemp!备注) Then
        Err.Raise 9000, gstrSysName, "该单据的就诊编号丢失，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If

    str就诊编号 = rsTemp!备注
    cur个人帐户 = rsTemp!个人帐户支付

    '调用接口数冲销
    initType
    mblnReturn = fl_canrollback(gstr医保机构编码, gstr医院编码, str就诊编号, gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If

    initType
    mblnReturn = fl_rollbackcalc(gstr医保机构编码, gstr医院编码, str就诊编号, "0", gstrOutPara)

    '帐户年度信息
    Call Get帐户信息(TYPE_涪陵, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)

    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_涪陵 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_涪陵 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        cur个人帐户 * -1 & ",NULL,NULL,NULL,'" & str就诊编号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    住院结算冲销_涪陵 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算_涪陵(lng结帐ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
'        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str操作员 As String, datCurr As Date, str就诊编号 As String, cur起付段支付 As Currency
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur个人帐户 As Currency, cur起付线 As Currency, cur基本统筹限额 As Currency
    Dim cur大额统筹限额 As Currency, cur基数自付 As Currency, cur余额 As Currency
    Dim cur发生费用 As Currency, cur统筹支付 As Currency, cur补充医保 As Currency, cur全自付 As Currency, cur先自付 As Currency, cur公务员补助 As Currency
    Dim lng主页ID As Long, cur帐户支付 As Currency
    Dim str出院日期 As String, str公务员类别 As String
    
    On Error GoTo errHandle
    '需要先上传费用明细
'    费用明细传递_涪陵 lng结帐ID
    
    If gstr医保机构编码 = "" Then
        Err.Raise 9000, gstrSysName, "准备读取医保机构编码，请插入系统卡或病人卡！"
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    
    '求个人帐户支付金额
    gstrSQL = "Select Nvl(冲预交,0) as 金额 From 病人预交记录 Where 结算方式='个人帐户' And 记录性质=2 And 结帐ID=" & lng结帐ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "模拟医保", lng结帐ID)
    If rsTemp.RecordCount > 0 Then
        cur帐户支付 = rsTemp("金额")
    End If
    
    gstrSQL = "Select * From 住院费用记录 Where 结帐ID=[1] And Rownum<2"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    lng病人ID = rs明细("病人ID")
    lng主页ID = rs明细("主页ID")
    str操作员 = UserInfo.姓名
    
    '取病案主页中的出院日期
    gstrSQL = " Select A.出院日期 From 病案主页 A,病人信息 B" & _
              " Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 And A.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    If rsTemp.RecordCount <> 0 Then
        str出院日期 = Format(rsTemp!出院日期, "yyyy-MM-dd")
    Else
        str出院日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End If
    
    gstrSQL = "Select nvl(顺序号,0) as 顺序号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_涪陵)
    str就诊编号 = rsTemp!顺序号
    '医保机构编码, 医院编号, 医保就诊编号， 出院日期，操作员，显示标志
    datCurr = zlDatabase.Currentdate
    initType
    mblnReturn = fl_calc(gstr医保机构编码, gstr医院编码, str就诊编号, str出院日期, str操作员, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext
        住院结算_涪陵 = False
        Exit Function
    End If
'间接出口参数:1费用合计,2特殊病种费用,3本次本年帐户支付,4本次历年帐户支付,5累计分段自付,6统筹金支付,7起付段支付,
'             8单位支付,9自费费用,10特检先自付,11特治先自付,12特检费用,13特治费用,14补充医疗保险支付,15本次统筹记入累计,
'             16补充医疗记入累计,17门诊统筹记入累计,18未报销费用,19医保支付,20个人现金支付,21个人帐户余额

    '获取个人帐户支付和个人现金支付
    '判断第23个参数出参不为空时,适用于涪凌医保 20070929胡华军
    str公务员类别 = gstrOutPara.out22
    If gstrOutPara.out23 <> "" Then
        cur个人帐户 = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
        cur余额 = CCur(gstrOutPara.out21)
        cur全自付 = CCur(gstrOutPara.out20) - cur个人帐户
        cur发生费用 = CCur(gstrOutPara.out1)
        cur先自付 = CCur(gstrOutPara.out10) + CCur(gstrOutPara.out11)
        cur起付段支付 = CCur(gstrOutPara.out7)
        cur统筹支付 = CCur(gstrOutPara.out6)
        cur补充医保 = CCur(gstrOutPara.out14)
        cur公务员补助 = CCur(gstrOutPara.out23)
        
        '校正
        Dim str结算方式 As String
        If cur个人帐户 <> cur帐户支付 Then
            str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户
            str结算方式 = str结算方式 & "||基本医疗保险|" & cur统筹支付
            str结算方式 = str结算方式 & "||补充医疗保险|" & cur补充医保
            str结算方式 = str结算方式 & "||公务员补助|" & cur公务员补助
            
            '试结算与结算不一至,需要修正
            str结算方式 = Mid(str结算方式, 3)
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "试结算与结算不一至,更新预交记录")
        End If
        
        '帐户年度信息
        Call Get帐户信息(TYPE_涪陵, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
        cur统筹报销累计 = cur统筹报销累计 + cur统筹支付 + cur公务员补助
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
                "," & TYPE_涪陵 & "," & Year(datCurr) & "," & cur帐户增加累计 & _
                "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
                cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & _
                cur起付线 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        '保险结算记录
        '性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN, _
         累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,发生费用金额_IN,全自付金额_IN, _
         首先自付金额_IN,进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN, _
         支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_涪陵 & "," & _
                lng病人ID & "," & Year(datCurr) & "," & _
                cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
                cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付段支付 & ",NULL,NULL," & _
                cur发生费用 & "," & cur全自付 & "," & cur先自付 & "," & cur补充医保 & "," & cur统筹支付 & "," & cur公务员补助 & ",NULL," & _
                cur个人帐户 & ",'" & str公务员类别 & "'," & lng主页ID & ",NULL,'" & str就诊编号 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        '---------------------------------------------------------------------------------------------
    Else
        cur个人帐户 = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
        cur余额 = CCur(gstrOutPara.out21)
        cur全自付 = CCur(gstrOutPara.out20) - cur个人帐户
        cur发生费用 = CCur(gstrOutPara.out1)
        cur先自付 = CCur(gstrOutPara.out10) + CCur(gstrOutPara.out11)
        cur起付段支付 = CCur(gstrOutPara.out7)
        cur统筹支付 = CCur(gstrOutPara.out6)
        cur补充医保 = CCur(gstrOutPara.out14)
        'cur公务员补助 = CCur(gstrOutPara.out23)
        
        '校正
       ' Dim str结算方式 As String
        If cur个人帐户 <> cur帐户支付 Then
            str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户
            str结算方式 = str结算方式 & "||基本医疗保险|" & cur统筹支付
            str结算方式 = str结算方式 & "||补充医疗保险|" & cur补充医保
            'str结算方式 = str结算方式 & "||公务员补助|" & cur公务员补助
            
            '试结算与结算不一至,需要修正
            str结算方式 = Mid(str结算方式, 3)
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "试结算与结算不一至,更新预交记录")
        End If
        
        '帐户年度信息
        Call Get帐户信息(TYPE_涪陵, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
        cur统筹报销累计 = cur统筹报销累计 + cur统筹支付
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
                "," & TYPE_涪陵 & "," & Year(datCurr) & "," & cur帐户增加累计 & _
                "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
                cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & _
                cur起付线 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        '保险结算记录
        '性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN, _
         累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,发生费用金额_IN,全自付金额_IN, _
         首先自付金额_IN,进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN, _
         支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_涪陵 & "," & _
                lng病人ID & "," & Year(datCurr) & "," & _
                cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
                cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付段支付 & ",NULL,NULL," & _
                cur发生费用 & "," & cur全自付 & "," & cur先自付 & "," & cur补充医保 & "," & cur统筹支付 & ",NULL,NULL," & _
                cur个人帐户 & ",NULL," & lng主页ID & ",NULL,'" & str就诊编号 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    End If
    住院结算_涪陵 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院虚拟结算_涪陵(rs费用明细 As Recordset, lng病人ID As Long, str医保号 As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
'20070926胡华军增加:cur公务员补助
    
    Dim cur个人帐户支付 As Currency, cur个人现金支付 As Currency, cur帐户余额 As Currency
    Dim cur统筹支付 As Currency, cur医保支付 As Currency, cur补充医保 As Currency, cur公务员补助 As Currency
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str操作员 As String, datCurr As Date, str就诊编号 As String
    Dim curCount As Currency
    
    On Error GoTo errHandle
    '需要先上传费用明细
'    费用明细传递_涪陵 0, rs费用明细
'
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    Set rs明细 = rs费用明细.Clone

    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    curCount = 0
    While Not rs明细.EOF
        curCount = curCount + rs明细!金额
        rs明细.MoveNext
    Wend
    rs明细.MoveFirst
    If curCount = 0 Then
        MsgBox "病人费用为零,不能进行结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str操作员 = UserInfo.姓名
    
    If 记帐传输_涪陵("", 0, "", lng病人ID) = False Then
        Exit Function
    End If
    
    gstrSQL = "Select Nvl(帐户余额,0) AS 帐户余额,nvl(顺序号,0) as 顺序号,中心 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_涪陵)
    str就诊编号 = rsTemp!顺序号
    cur帐户余额 = rsTemp!帐户余额
'    gstr医保机构编码 = rsTemp!中心
    '医保机构编码, 医院编号, 医保就诊编号， 出院日期，操作员，显示标志
    datCurr = zlDatabase.Currentdate
    initType
    mblnReturn = fl_pcalc(gstr医保机构编码, gstr医院编码, str就诊编号, Format(datCurr, "yyyy-MM-dd"), str操作员, "1", "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        住院虚拟结算_涪陵 = ""
        Exit Function
    End If
'间接出口参数:1费用合计,2特殊病种费用,3本次本年帐户支付,4本次历年帐户支付,5累计分段自付,6统筹金支付,7起付段支付,
'             8单位支付,9自费费用,10特检先自付,11特治先自付,12特检费用,13特治费用,14补充医疗保险支付,15本次统筹记入累计,
'             16补充医疗记入累计,17门诊统筹记入累计,18未报销费用,19医保支付,20个人现金支付,21个人帐户余额
    
    '获取个人帐户支付和个人现金支付
    '判断第22个参数出参不为空时,适用于涪凌医保 20070929胡华军
    If gstrOutPara.out22 <> "" Then
            cur个人帐户支付 = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
            cur个人现金支付 = CCur(gstrOutPara.out20)
            cur统筹支付 = CCur(gstrOutPara.out6)
            cur医保支付 = CCur(gstrOutPara.out19)
            cur补充医保 = CCur(gstrOutPara.out14)
            cur公务员补助 = CCur(gstrOutPara.out22)
            If curCount <> CCur(gstrOutPara.out1) Then
                MsgBox "请注意：医保返回结算金额与当前单据金额不符" & vbCrLf & _
                               "医院总额：" & curCount & "    医保返回：" & CCur(gstrOutPara.out1), vbInformation, gstrSysName
            End If
            住院虚拟结算_涪陵 = "个人帐户;" & cur个人帐户支付 & ";0" '不允许修改个人帐户
        '    If cur个人现金支付 <> 0 Then
        '        住院虚拟结算_涪陵 = 住院虚拟结算_涪陵 & "|现金;" & cur个人现金支付 & ";0" '不允许修改现金支付
        '    End If
            If cur统筹支付 <> 0 Then
                住院虚拟结算_涪陵 = 住院虚拟结算_涪陵 & "|基本医疗保险;" & cur统筹支付 & ";0" '不允许修改统筹支付
            End If
            If cur补充医保 <> 0 Then
                住院虚拟结算_涪陵 = 住院虚拟结算_涪陵 & "|补充医疗保险;" & cur补充医保 & ";0"
            End If
            
            If cur公务员补助 <> 0 Then
                住院虚拟结算_涪陵 = 住院虚拟结算_涪陵 & "|公务员补助;" & cur公务员补助 & ";0"
            End If

            If cur帐户余额 < cur个人帐户支付 Then
                gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_涪陵 & ",'帐户余额'," & cur个人帐户支付 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新帐户余额")
            End If
    Else
    
            cur个人帐户支付 = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
            cur个人现金支付 = CCur(gstrOutPara.out20)
            cur统筹支付 = CCur(gstrOutPara.out6)
            cur医保支付 = CCur(gstrOutPara.out19)
            cur补充医保 = CCur(gstrOutPara.out14)
'            cur公务员补助 = CCur(gstrOutPara.out23)
            If curCount <> CCur(gstrOutPara.out1) Then
                MsgBox "请注意：医保返回结算金额与当前单据金额不符" & vbCrLf & _
                               "医院总额：" & curCount & "    医保返回：" & CCur(gstrOutPara.out1), vbInformation, gstrSysName
            End If
            住院虚拟结算_涪陵 = "个人帐户;" & cur个人帐户支付 & ";0" '不允许修改个人帐户
        '    If cur个人现金支付 <> 0 Then
        '        住院虚拟结算_涪陵 = 住院虚拟结算_涪陵 & "|现金;" & cur个人现金支付 & ";0" '不允许修改现金支付
        '    End If
            If cur统筹支付 <> 0 Then
                住院虚拟结算_涪陵 = 住院虚拟结算_涪陵 & "|基本医疗保险;" & cur统筹支付 & ";0" '不允许修改统筹支付
            End If
            If cur补充医保 <> 0 Then
                住院虚拟结算_涪陵 = 住院虚拟结算_涪陵 & "|补充医疗保险;" & cur补充医保 & ";0"
            End If
            
'            If cur公务员补助 <> 0 Then
'                住院虚拟结算_涪陵 = 住院虛拟结算_涪陵 & "|公务员补助:" & cur公务员补助 & ":0"
'            End If
             
            If cur帐户余额 < cur个人帐户支付 Then
                gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_涪陵 & ",'帐户余额'," & cur个人帐户支付 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新帐户余额")
            End If
  End If
            Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    住院虚拟结算_涪陵 = ""
End Function

Public Function 出院登记_涪陵(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    '个人状态的修改
    Dim str就诊编号 As String, rsTemp As New ADODB.Recordset
    Dim bln零费用出院 As Boolean
    
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    
    On Error GoTo errHandle
    '检查该次住院是否没有费用发生
    gstrSQL = "Select nvl(sum(实收金额),0) as 金额  from 住院费用记录 where nvl(附加标志,0)<>9 and 病人ID=[1] and 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病人出院", lng病人ID, lng主页ID)
    If rsTemp.EOF = True Then
        bln零费用出院 = True
    Else
        bln零费用出院 = (rsTemp("金额") = 0)
    End If
    
    If bln零费用出院 = True Then
        gstrSQL = "Select nvl(顺序号,0) as 顺序号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_涪陵)
        str就诊编号 = rsTemp!顺序号
        initType
        mblnReturn = fl_dall(gstr医保机构编码, gstr医院编码, str就诊编号, gstrOutPara)
        If mblnReturn = False Then
            出院登记_涪陵 = False
            Exit Function
        End If
    End If
    '对HIS之中的基础数据进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_涪陵 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_涪陵 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    出院登记_涪陵 = False
End Function

Public Function 医保设置_涪陵() As Boolean
    医保设置_涪陵 = frmSet涪陵.ShowME(TYPE_涪陵)
End Function

Private Function Get病人ID(str医保号 As String, str医保中心编码 As String) As String
'功能：通过医保中心号码和医保号求出病人ID
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 病人ID from 保险帐户 where 险类 = [1] and 医保号 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_涪陵, str医保号)
    If Not rsTmp.BOF Then
        Get病人ID = CStr(rsTmp("病人ID"))
    Else
        Get病人ID = ""
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Get病人ID = ""
End Function

Public Function 记帐传输_涪陵(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0) As Boolean
    Dim lng主页ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim RSPATIENT As New ADODB.Recordset
    
    
    If str单据号 <> "" Then
        '可能存在一张单据内有多个病人的情况
        gstrSQL = "Select Distinct 病人ID From 住院费用记录 Where 记录状态<>0 And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 and 记录性质=[1] and NO=[2]"
        Set RSPATIENT = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人信息", int性质, str单据号)
        With RSPATIENT
            Do While Not .EOF
                '判断是否是医保病人
                gstrSQL = "Select Count(*) Records From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.病人ID=[1] And A.主页ID=B.住院次数 And A.险类=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否医保病人", CLng(!病人ID), TYPE_涪陵)
                If (rsTemp!Records = 1) Then
                    '提取指定病人，指定单据明细
                    gstrSQL = "Select * From 住院费用记录 Where 记录状态<>0 And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 and 记录性质=[1] and NO=[2] And 病人ID=[3] order by 主页ID,序号"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", int性质, str单据号, CLng(!病人ID))
                    记帐传输_涪陵 = 费用明细传递_涪陵(0, rsTemp)
                    If Not 记帐传输_涪陵 Then Exit Function
                Else
                    记帐传输_涪陵 = True
                End If
                .MoveNext
            Loop
        End With
    Else
        If lng病人ID <> 0 Then
            gstrSQL = "Select Max(主页ID) From 病案主页 Where 病人id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
            lng主页ID = rsTemp(0)
        End If
        
        gstrSQL = "Select * From 住院费用记录 Where 记录状态<>0 And Nvl(是否上传,0)=0 And Nvl(实收金额,0)<>0 And nvl(附加标志,0)<>9 and 病人id=[1] And 主页id=[2] order by 主页ID,序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", lng病人ID, lng主页ID)
        记帐传输_涪陵 = 费用明细传递_涪陵(0, rsTemp)
    End If
End Function

Public Function LeftStr(sString As String, iLen As Long) As String
    Dim sStr As String, sTmp As String
    sStr = sString
    If LenB(StrConv(sStr, vbFromUnicode)) > iLen Then
        sTmp = StrConv(sStr, vbFromUnicode)
        sStr = StrConv(LeftB(sTmp, iLen), vbUnicode)
        If asc(Right(sStr, 1)) = 0 Then sStr = StrConv(LeftB(sTmp, iLen - 1), vbUnicode)
    End If
    LeftStr = sStr
End Function

