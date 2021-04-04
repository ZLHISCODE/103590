Attribute VB_Name = "mdlInPatient"
Option Explicit '要求变量声明

Public gobjPatient As Object '病人管理部件
Public gclsInsure As New clsInsure
Public gobjPublicPatient As Object  '病人信息公共部件 zlPublicPatient.clsPublicPatient
Public gobjPlugIn As Object    '插件部件zlPlugIn.clsPlugIn
Public gobjXWHIS As Object     '新网接口部件zl9XWInterface.clsHISInner

'系统参数--------------------------------
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"
Public gbytDec As Byte '费用金额的小数点位数

Public gstrLike As String  '项目匹配方法,%或空
Public gblnMyStyle As Boolean '使用个性化风格
Public gstrIme As String '自动的开启输入法
Public gbytCode As Byte '简码生成方式，0-拼音,1-五笔,2-两者
Public gblnXW As Boolean      '系统参数：“启用医学影像信息系统专业版接口”
Public gobjPublicExpenseBillOperation As Object  '费用公共部件，用户转病区转费用

Public gint门诊诊断输入 As Integer
Public gint住院诊断输入 As Integer
Public gbln入科确定护理等级 As Boolean
Public gbln医生允许才能出院 As Boolean '医生下达出院医嘱才允许病人出院
Public gbln医生允许才能撤销预出院 As Boolean '回退出院医嘱后才允许撤销预出院
Public gbln每次住院新住院号 As Boolean '每次住院使用新的住院号
Public gbln在院病人不准出院结帐 As Boolean '存在未审核的销帐申请是,未病人不允许进行出院

Public gbln入院预交 As Boolean '入院时收预交款
Public gbln入院发卡 As Boolean '入院时办就诊卡
Public gbln入院入科 As Boolean '入院同时入科
Public gbyt入科时间 As Byte '0-入院时间,1-入科时系统时间
Public gbln控制空床 As Boolean

'Public gblnShowCard As Boolean '是否显示明文卡号
Public gblnCheckPass As Boolean '刷卡时要求输入密码
Public gblnMultiBalance As Boolean  '多张单据使用多种结算方式模式
   
'Public gbytCardNOLen As Byte '就诊卡号长度
Public gbytPrepayLen As Byte '票据号码长度
Public gblnPrepayStrict As Boolean '是否严格控制票据
'Public gblnMagcardStrict As Boolean '是否严格控制票据
Public gbyt出院时检查未执行 As Byte        '出院和结帐出院时检查是否有未执行项目:0-不检查,1-检查并提示,2-检查并禁止
Public gbyt转科时检查未执行 As Byte        '转科时检查是否有未执行项目:0-不检查,1-检查并提示,2-检查并禁止
'问题30208 by lesfeng 2010-08-02 撤分参数22及32 新增154、155
Public gbyt出院时检查药品未执行 As Byte    '在出院结帐及病人入出管理中出院时是否检查病人的未发药品项目,0-不检查,1-检查并提示,2-检查并禁止
Public gbyt转科时检查药品未执行 As Byte    '在病人入出管理中转科时是否检查病人的未发药品项目:0-不检查,1-检查并提示,2-检查并禁止
'刘兴洪 问题:????    日期:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '费用小数精度
Public gstrFeePrecisionFmt As String '费用小数格式:0.00000
'61347:刘鹏飞,2013-11-09
Public gbyt病人审核方式 As Byte             '病人审核方式:0-未审核不允许结帐，缺省为0;1-审核时不许调整费用和医嘱（包含医嘱调整和费用调整）
'61492:刘鹏飞,2013-11-11
Public gbyt转科时未审核销帐单据检查 As Byte  '转科时是否检查病人存在未审核的销帐单据:0-不检查,1-检查并提示,2-检查并禁止
'68953:刘鹏飞,2014-08-12
Public gbyt出院时超期护理数据检查 As Byte    '出院时是否检查病人出院时间之后存在护理数据:0-不检查,1-检查并提示,2-检查并禁止
'医嘱相关
Public gbln药疗划价单 As Boolean
Public gbln其他划价单 As Boolean
Public gbln执行后审核 As Boolean
'结构化地址
Public gbln启用结构化地址 As Boolean
Public gbln显示乡镇 As Boolean
Public gblnPatiByID As Boolean   '同一身份证只能对应一个建档病人

'本地参数
Public gbln担保 As Boolean '是否允许输入担保信息
Public gblnSeekName As Boolean '是否通过姓名进行模糊查找
Public gintNameDays As Integer '通过姓名模糊查找天数
Public gstr预交ID As String   '票据领用ID
Public gbln记账 As Long '就诊卡费用以记账方式收取
Public gbln先选病区 As Boolean '入院时先选病区
Public gbln费用计算 As Boolean '计算一次的项目是否在入院是计算(未入住的情况)
Public gbln转病区转费用 As Boolean

Public gbln医疗机构不允许自由录入 As Boolean

Public gbytPrepayPrint As Byte ''0-不打印,1-要打印,2-选择是否打印
Public gbytFPagePrint As Byte ''0-不打印,1-要打印,2-选择是否打印
Public gbytWristletPrint As Byte ''0-不打印,1-要打印,2-选择是否打印(病人入院管理)
Public gbytBabyWristletPrint As Byte ''0-不打印,1-要打印,2-选择是否打印
Public gbytCourseWristletPrint As Byte ''0-不打印,1-要打印,2-选择是否打印(病人入出管理)

'初始化 clsBase
Public gclsBase As New clsBase
Public Const ETO_OPAQUE = 2
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'消息结构类型
Public Enum enumXmlType
    xsString = 1
    xsNumber = 2
    xsDate = 3
    xsTime = 4
    xsDateTime = 5
End Enum


Public Enum EFun
    E入科 = 0
    E转科 = 1
    E换床 = 2
    E包房 = 3
    E出院 = 4
    E转为住院 = 5
    E更改床位等级 = 6
    E调整病人信息 = 7
    E新生儿登记 = 8
    E重算费用 = 9
    E医保病种选择 = 10
    E撤销 = 11
    E修改出院时间 = 12
    E床位对换 = 13
    E转医疗小组 = 14
    E转病区 = 15
    E入病区 = 16
    E病人备注编辑 = 17
End Enum

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
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    
End Enum
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
    g本机公共模块 = 5
    g本机私有模块 = 6
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
    共享系统号 As Long
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

Public Enum 交易Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
    Busi_TransferSwap
    Busi_TransferDelSwap
    Busi_WipeoffMoney
    Busi_SettleSwap
    Busi_SettleDelSwap
    Busi_ComeInSwap
    Busi_LeaveSwap
    Busi_TranChargeDetail
    Busi_LeaveDelSwap
    Busi_RegistSwap
    Busi_RegistDelSwap
    Busi_ComeInDelSwap
    Busi_ModiPatiSwap
    Busi_ChooseDisease
    Busi_IdentifyCancel
End Enum

'内部应用模块号定义
Public Enum ENUM_INSIDE_PROGRAM
    P病区床位管理 = 1130
    P病人入院管理 = 1131
    P病人入出管理 = 1132
End Enum

'结构化地址类型 1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址
Public Enum Enum_IX_ADDRESS
    E_IX_出生地点 = 1
    E_IX_籍贯 = 2
    E_IX_现住址 = 3
    E_IX_户口地址 = 4
    E_IX_联系人地址 = 5
End Enum

Public ParamInfo As SYSPARAM_INFO

Public Function ExecPatiChange(ByVal bytFun As Byte, ByRef frmParent As Form, ByRef strPrivs As String, ParamArray arrPar() As Variant) As Boolean
'功能:执行病人变动相关功能
'参数:bytFun:0-入科,1-转科
'     arrPar:根据不同的功能调用，传入不同的参数
'            入科:病区,床号(入住的目标床位,允许为空),mlng床位科室ID,入科方式(0-入院入科，1-转科入科)
'            转科:病人ID,主页Id
'            换床:病人ID,主页Id,mbytInFun,mstr目标床号
'            出院:病人ID,主页Id
'            转为住院:病人ID,主页Id,住院号,姓名
'            调整床位等级:病人ID,主页Id,mstr床号(当前调整等级的床号)
'            重算费用:lng病人ID, lng主页Id, str姓名
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim strTmp As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error Resume Next    '是因为窗体的Form_Load中执行Unload me会出错
    Select Case bytFun
    Case EFun.E入科
        strTmp = CStr(arrPar(3))
        blnReturn = frmIn.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strTmp, Val(arrPar(4)), CStr(arrPar(5)), strPrivs)
        arrPar(3) = strTmp
    Case EFun.E入病区
        strTmp = CStr(arrPar(3))
        blnReturn = frmCheckIn.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strTmp, Val(arrPar(4)), strPrivs)
        arrPar(3) = strTmp
    Case EFun.E转病区
        blnReturn = frmChangeUnit.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strPrivs)
    Case EFun.E转科
        blnReturn = frmChange.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strPrivs)
    Case EFun.E转医疗小组
        'EFun.E转医疗小组, Me, mstrPrivs, mlngUnit, mrsBeds!病人ID, mrsBeds!主页ID, 3
        blnReturn = frmChangeGroup.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strPrivs)
    Case EFun.E换床
        strTmp = CStr(arrPar(4))
        blnReturn = frmMove.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), Val(arrPar(3)), strTmp, CStr(arrPar(5)), strPrivs)
        arrPar(4) = strTmp
    Case EFun.E床位对换
        strTmp = CStr(arrPar(4))
        blnReturn = frmBedSwap.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), CStr(arrPar(3)), strTmp, strPrivs)
        arrPar(4) = strTmp
    Case EFun.E出院
        blnReturn = frmOut.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), strPrivs)
    '问题27392 by lesfeng 2010-01-14
    Case EFun.E修改出院时间
        ExecPatiChange = frmModifOut.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), strPrivs)
    Case EFun.E转为住院
        Dim strNote  As String, str住院号 As String
        '没有住院号则分配一个
        '问题 26939 by lesfeng 2010-1-4 在执行 ZL_病人变动记录_转住院 时没有将住院号传入
        '77193:刘鹏飞,修改CStr(arrPar(2))=""条件为val(arrPar(2))=0,其它模块可能进行了val导致传入的住院号为0
        If Val(arrPar(2)) = 0 Then
            If gbln每次住院新住院号 = False Then
                strSql = " SELECT Nvl(a.住院号," & vbNewLine & _
                    "            (SELECT 住院号" & vbNewLine & _
                    "             FROM 病案主页" & vbNewLine & _
                    "             WHERE 病人id = a.病人id AND" & vbNewLine & _
                    "                   主页id = (SELECT MAX(主页id) FROM 病案主页 WHERE 病人id = a.病人id AND 住院号 IS NOT NULL))) 住院号" & vbNewLine & _
                    " FROM 病人信息 a" & vbNewLine & _
                    " WHERE 病人id = [1]"
            
                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取住院号", arrPar(0))
                If Not rsTemp.EOF Then
                    str住院号 = Nvl(rsTemp!住院号)
                End If
                If str住院号 = "" Then
                    str住院号 = zlDatabase.GetNextNo(2)
                    strNote = "在留观病人 " & CStr(arrPar(3)) & " 转为住院病人之前，请先为该病人确定一个住院号。"
                    If Not frmInput.InputVal(frmParent, "住院号", strNote, str住院号, 1, 10, False, InStr(strPrivs, ";修改住院号;") <> 0) Then Exit Function
                End If
            Else
                str住院号 = zlDatabase.GetNextNo(2)
                strNote = "在留观病人 " & CStr(arrPar(3)) & " 转为住院病人之前，请先为该病人确定一个住院号。"
                If Not frmInput.InputVal(frmParent, "住院号", strNote, str住院号, 1, 10, False, InStr(strPrivs, ";修改住院号;") <> 0) Then Exit Function
            End If
        Else
            str住院号 = CStr(arrPar(2))
        End If
        
        On Error GoTo errH
        strSql = "ZL_病人变动记录_转住院(" & Val(arrPar(0)) & "," & Val(arrPar(1)) & "," & str住院号 & ")"
        zlDatabase.ExecuteProcedure strSql, App.ProductName
        gblnOK = True
        blnReturn = gblnOK
    Case EFun.E更改床位等级
        blnReturn = frmBedLevel.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), CStr(arrPar(2)))
    Case EFun.E调整病人信息
        blnReturn = frmEditPati.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strPrivs)
    Case EFun.E新生儿登记
        blnReturn = frmBabyReg.ShowMe(Val(arrPar(0)), Val(arrPar(1)), strPrivs, frmParent)
    Case EFun.E病人备注编辑
        blnReturn = frmMemo.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), strPrivs)
    Case EFun.E重算费用
             
        If MsgBox("你确定要将[" & CStr(arrPar(2)) & "]的未结费用按当前费别重算吗?" & vbCrLf & vbCrLf & _
            "本操作将按病人当前费别对应的优惠比率对未结费用重新进行打折计算!", vbInformation + vbYesNo + vbDefaultButton1, App.ProductName) = vbNo Then
            Exit Function
        End If
        
        On Error GoTo errH
        strSql = "Zl_病人未结费用_Recalc(" & Val(arrPar(0)) & "," & Val(arrPar(1)) & ")"
        zlDatabase.ExecuteProcedure strSql, App.ProductName
        gblnOK = True
        blnReturn = gblnOK
    Case EFun.E医保病种选择
        Call gclsInsure.ChooseDisease(Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)))
        blnReturn = True
    Case EFun.E撤销
        blnReturn = ExecUndo(frmParent, strPrivs, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), Val(arrPar(3)), CStr(arrPar(4)))
        
    End Select
    ExecPatiChange = blnReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecUndo(ByRef frmParent As Form, ByVal strPrivs As String, ByVal lngUnit As Long, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int险类 As Integer, ByVal strType As String) As Boolean
    Dim str数据 As String, strBeds As String, str主床位 As String, str等级IDs As String
    Dim str床位串 As String
    Dim strPreBed As String, strBed As String, strInfo As String
    Dim lngBeSwap病人ID As Long, lngBeSwap主页ID As Long
    Dim blnDie As Boolean, blnUndoOut As Boolean, blnUndoPreOut As Boolean
    Dim bln结清 As Boolean, strSql As String, strUndoBeds As String
    Dim rsTmp As ADODB.Recordset, blnTrans As Boolean, blnSwapBed As Boolean
    Dim rsBeSwap As ADODB.Recordset
    Dim arrSQL() As String, intLoop As Integer
    
    Dim clsMipModule As zl9ComLib.clsMipModule
    Dim clsXML As zl9ComLib.clsXML
    Dim rsUndoBegin As ADODB.Recordset, rsUndoEnd As ADODB.Recordset
    Dim rsDeptOper As New ADODB.Recordset '部门人员信息
    Dim rsBedChange As New ADODB.Recordset
    Dim strBeforBed As String, strAfterBed As String
    Dim lngNextPati As Long, lngNextPage As Long '床位对换第二个病人
    Dim bln床位对换 As Boolean
        Dim colSQL As New Collection, i As Long, strSQLTmp As String, rsPati As Recordset
    
    blnUndoOut = strType = "出院"
    blnUndoPreOut = strType = "预出院"
    
    On Error GoTo errH
    
    '撤销病区变更之前检查未执行项目
    If InStr(strType, "病区") > 0 Then
        If gbyt转科时检查未执行 <> 0 Then
    
            strInfo = ExistWaitExe(lng病人ID, lng主页ID)
            If strInfo <> "" Then
                If gbyt转科时检查未执行 = 1 Then
                    If MsgBox("该病人存在尚未执行完成的内容：" & _
                        vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "确定要撤消" & strType & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "该病人存在尚未执行完成的内容：" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许撤消" & strType & "。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    
        If gbyt转科时检查药品未执行 <> 0 Then
            strInfo = ExistWaitDrug(lng病人ID, lng主页ID)
            If strInfo <> "" Then
                If gbyt转科时检查药品未执行 = 1 Then
                    If MsgBox("该病人" & strInfo & vbCrLf & vbCrLf & "确定要撤消" & strType & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "该病人" & strInfo & vbCrLf & vbCrLf & "不允许撤消" & strType & "。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        '61429:刘鹏飞,2013-11-11,转科时销帐未审核单据检查
        If gbyt转科时未审核销帐单据检查 <> 0 Then
            strInfo = ""
            strInfo = ExistWaitQuittance(lng病人ID, lng主页ID)
            If strInfo <> "" Then
                If gbyt转科时未审核销帐单据检查 = 1 Then
                    If MsgBox("该病人" & strInfo & vbCrLf & vbCrLf & "确定要撤消" & strType & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "该病人" & strInfo & vbCrLf & vbCrLf & "不允许撤消" & strType & "。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        '74262:刘鹏飞,2014-06-23,添加条件:终止时间 Is Null
        gstrSQL = "Select 1" & vbNewLine & _
                " From 病人医嘱记录 a" & vbNewLine & _
                " Where 病人id = [1] And 主页id = [2] And 医嘱状态 Not In (4, 8, 9) And Exists" & vbNewLine & _
                " (Select 1" & vbNewLine & _
                "       From 病人变动记录" & vbNewLine & _
                "       Where 病人id = a.病人id And 主页id = a.主页id And 开始原因 = 15 And 开始时间 < a.开嘱时间 And 终止时间 Is Null) And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "撤销转病区时判断是否存在有效医嘱", lng病人ID, lng主页ID)
        
        If Not rsTmp.EOF Then
            MsgBox "该病人在当前病区存在有效医嘱，不允许撤消" & strType & "！"
            Exit Function
        End If
    End If
    
    If InStr(strType, "护理等级变动") > 0 Then
        gstrSQL = "SELECT 1" & vbNewLine & _
                    "FROM 病人变动记录 a, 病人医嘱记录 b, 诊疗收费关系 c" & vbNewLine & _
                    "WHERE a.病人id = b.病人id AND a.主页id = b.主页id AND a.护理等级id = c.收费项目id AND b.诊疗项目id = c.诊疗项目id AND a.病人id = [1] AND" & vbNewLine & _
                    "      a.主页id = [2] AND a.开始原因 = 6 AND a.终止原因 IS NULL AND a.终止时间 IS NULL"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "撤销护理等级判断是否存在护理等级医嘱", lng病人ID, lng主页ID)
        
        If Not rsTmp.EOF Then
            MsgBox "该病人存在护理等级医嘱，不允许撤消" & strType & "！"
            Exit Function
        End If

    End If

    If InStr(strType, "入住") > 0 And InStr(strPrivs, "写病历后撤消入科") = 0 Then '撤消入科/撤消转科入科,检查是否已书写有入院/转科时限要的病历
        gstrSQL = "Select Count(ID) 记录 From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 终止时间 Is Null And 开始原因 = 3"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "有转科入科记录", lng病人ID, lng主页ID)
        
        gstrSQL = "Select Count(B.文件id) 数量" & vbNewLine & _
                    "From 病历时限要求 A, 电子病历记录 B" & vbNewLine & _
                    "Where Instr([3],A.事件)>0 And A.书写时限 > 0 And A.文件id = B.文件id And B.病人id = [1] And B.主页id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否书写过病历", lng病人ID, lng主页ID, IIf(rsTmp!记录 = 0, "入院,首次入院,再次入院", "转科"))
        If rsTmp!数量 > 0 Then
            MsgBox "该病人已书写有入院/转科时限要求的病历，禁止撤消入科！", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    
    '撤消前床位等级检查
    If blnUndoOut Or InStr(strType, "换床") > 0 Or InStr(strType, "入住") > 0 Then '撤销出院或撤销换床或撤销转科入科,撤销入科也会被判断但提取不到数据
        gstrSQL = "Select Distinct a.床号, a.等级id 原住等级, b.等级id 现有等级" & vbNewLine & _
                "From (Select a.床号, a.床位等级id 等级id, a.病区id" & vbNewLine & _
                "       From 病人变动记录 a, 病人变动记录 b" & vbNewLine & _
                "       Where a.病人id = [1] And a.主页id = [2] And a.病人id = b.病人id And a.主页id = b.主页id And" & vbNewLine & _
                "             (b.终止时间 Is Null And b.开始原因 In(3,4) And a.终止时间 = b.开始时间 Or" & vbNewLine & _
                "             b.终止原因 = 1 And a.终止原因 = b.终止原因)) a, 床位状况记录 b" & vbNewLine & _
                "Where Nvl(a.床号, 0) = b.床号 And a.病区id = b.病区id"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查上次入住床位变化", lng病人ID, lng主页ID)
        Do Until rsTmp.EOF
            If Nvl(rsTmp!原住等级, 0) <> Nvl(rsTmp!现有等级, 0) Then
                strUndoBeds = strUndoBeds & Nvl(rsTmp!床号, 0) & " "
            End If
            rsTmp.MoveNext
        Loop
        If Trim(strUndoBeds) <> "" Then
            If MsgBox("床位 " & strUndoBeds & "等级与病人上次入住时的等级不一致" & vbCrLf & "是否继续(继续将自动设置为上次入住时的等级)？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        End If
    End If

    '撤消出院
    '可能存在问题************
    If blnUndoOut Then
        If GetOutState(lng病人ID, lng主页ID) = "死亡" Then blnDie = True
        bln结清 = True
        Set rsTmp = GetMoneyInfo(lng病人ID, , , , , , , lng主页ID)
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then bln结清 = (rsTmp!费用余额 = 0)
        End If
    
        '权限判断
        If HavedInCost(lng病人ID, lng主页ID) Then '零费用不受"结清撤消出院"权限控制
            If bln结清 And InStr(strPrivs, "结清撤消出院") = 0 Then
                MsgBox "该出院病人的费用已结清，你没有权限将该病人撤消出院。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '已编目病案了的不允许撤消出院
        If HaveCatalogue(lng病人ID, lng主页ID) Then
            MsgBox "该病人本次住院的病案已经编目，不允许撤消出院。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '医保病人判断
        If int险类 <> 0 Then
            If Not gclsInsure.GetCapability(support撤消出院, lng病人ID, int险类) Then
                MsgBox "保险病人不能撤消出院！", vbInformation, gstrSysName
                Exit Function
            End If
            '去掉了医保连接匹配检查
        End If
    ElseIf int险类 <> 0 Then
        If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, lng病人ID, int险类) Then
            str数据 = "1"   '检查自动记帐的费用是否已被结帐
        End If
    End If
    
    '费用已审核病人是否允许撤销出院或撤销预出院
    If blnUndoPreOut Or blnUndoOut Then
        If InStr(strPrivs, "已审病人撤消出院") = 0 Then
            If CheckAudited(lng病人ID, lng主页ID) Then
                MsgBox "该病人费用已审核，不允许撤消" & IIf(blnUndoPreOut, "预", "") & "出院。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '61347:刘鹏飞,2013-11-09,如果病人的费用已经完成审核,并且费用审核方式为"审核时不许调整费用和医嘱"，则不允许撤销出院
        If gbyt病人审核方式 = 1 And blnUndoOut = True Then
            If CheckAudited(lng病人ID, lng主页ID, 2) Then
                MsgBox "该病人费用已完成审核,并且费用审核时不许调整费用和医嘱，不允许撤消出院。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If (blnUndoPreOut) Then
            '43579:撤销预出院时检查，当系统参数设置为“下出院医嘱才准出院”时，不允许在界面上撤销预出院，只能通过回退出院医嘱的发送方式来撤销预出院操作
'            If (gbln医生允许才能出院) Then
'                MsgBox "只能通过回退出院医嘱发送的方式来撤销预出院。", vbInformation, gstrSysName
'                Exit Function
'            End If
            '--55791:刘鹏飞,2012-11-13,回退出院医嘱才能撤销出院
             '撤销预出院时，检查如果是由出院医嘱操作的出院，则需要根据参数'回退出院医嘱才能撤销出院'决定是否可以直接取消预出院，如果不是由出院医嘱操作的出院，则可以直接撤销。
            If Check医生下达出院医嘱(lng病人ID, lng主页ID) And gbln医生允许才能撤销预出院 = True Then
                MsgBox "只能通过作废医嘱的方式来撤销预出院。", vbInformation, gstrSysName
                Exit Function
            End If
            
        End If
    End If
        
    If blnUndoOut And blnDie Then
        If MsgBox("该病人出院时已登记为死亡,确实要撤消出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If MsgBox("该操作将撤消病人" & strType & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    If strType Like "*转为住院病人*" Then
        If lng主页ID = 1 Then
            If MsgBox("要同时清除该病人的住院号吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then str数据 = "1"
        Else
            str数据 = ""
        End If
    End If
    
    '出院前入院的床位是否被占用,被占用则返回被占用床和原入住床
    If blnUndoOut Then
        strBeds = GetUsedBeds(lng病人ID, lng主页ID, str等级IDs)
        If strBeds <> "" Then
            If MsgBox("病人原住院床位 " & strBeds & "非空床，是否撤消到其它床位？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        
            Call ExecPatiChange(EFun.E换床, frmParent, strPrivs, lngUnit, lng病人ID, lng主页ID, 2, str等级IDs, strBeds)
            If Not gblnOK Then Exit Function
            'str等级IDs-传回新入住的床号
            strBeds = str等级IDs
            str主床位 = Split(str等级IDs, ",")(0)
        End If
    End If
    '问题28386 by lesfeng 2010-03-06 增加 strType = 入科、入院入科、转科入科（转科）、换床、床位登记变动、护理登记变动、经治医师改动、责任护士改变、转为住院病人、预出院、主治医师变动、主任医师变动、病况变动以及出院
    ReDim Preserve arrSQL(0)
    arrSQL(UBound(arrSQL)) = "zl_病人变动记录_Undo(" & lng病人ID & "," & lng主页ID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & str数据 & "','" & strBeds & "','" & str主床位 & "','" & strType & "')"
    
    '提取撤销前病人相关信息
    Select Case strType
    Case "出院"
        strSql = " Select b.姓名, b.性别, b.住院号, a.Id 变动id,病区Id 撤销前病区ID,科室Id 撤销前科室Id,b.出院病床,A.开始时间" & _
            "   From 病人变动记录 a, 病案主页 b" & _
            "   Where a.病人id = b.病人id And a.主页id = b.主页id And a.终止原因 = 1 And Nvl(a.附加床位, 0) = 0 And b.病人id = [1] And b.主页id = [2]"
        Set rsUndoBegin = zlDatabase.OpenSQLRecord(strSql, "撤销变动前", lng病人ID, lng主页ID)
    Case "转科", "转病区"
        strSql = "Select b.姓名, b.性别, b.住院号, a.Id 变动id,病区Id 撤销前病区ID,科室Id 撤销前科室Id,b.出院病床,A.开始时间" & _
            "   From 病人变动记录 a, 病案主页 b" & _
            "   Where a.病人id = b.病人id And a.主页id = b.主页id And a.开始原因 = [3] And a.开始时间 Is Null And Nvl(a.附加床位, 0) = 0 And 终止时间 Is Null And b.病人id = [1] And b.主页id = [2]"
        Set rsUndoBegin = zlDatabase.OpenSQLRecord(strSql, "撤销变动前", lng病人ID, lng主页ID, IIf(strType = "转科", 3, 15))
    Case Else
        strSql = "Select b.姓名, b.性别, b.住院号, a.Id 变动id,病区Id 撤销前病区ID,科室Id 撤销前科室Id,b.出院病床,A.开始时间" & _
            "   From 病人变动记录 a, 病案主页 b" & _
            "   Where a.病人id = b.病人id And a.主页id = b.主页id And a.开始时间 IS NOT NULL And Nvl(a.附加床位, 0) = 0 And 终止时间 Is Null And b.病人id = [1] And b.主页id = [2]"
        Set rsUndoBegin = zlDatabase.OpenSQLRecord(strSql, "撤销变动前", lng病人ID, lng主页ID)
    End Select
    '检查是否取消床位对换(消息处理需要)
    bln床位对换 = False
    If strType = "换床" And rsUndoBegin.RecordCount > 0 Then
        strSql = "Select a.床号,NVL(a.附加床位,0) 附加床位 , c.出院病床, c.出院科室id, c.当前病区id,b.病人id" & vbNewLine & _
            "    From 病人变动记录 A, 床位状况记录 B, 病案主页 C" & vbNewLine & _
            "    Where a.病人id = [1] And a.主页id = [2] And 终止原因 = 4 And a.病区id = b.病区id And a.病人id = c.病人id And" & vbNewLine & _
            "          a.主页id = c.主页id And a.床号 = b.床号" & vbNewLine & _
            " Order By a.终止时间 Desc,NVL(A.附加床位,0) DESC, a.开始时间 Desc"
         Set rsBedChange = zlDatabase.OpenSQLRecord(strSql, "", lng病人ID, lng主页ID)
         If rsBedChange.RecordCount > 0 Then
            strBeforBed = Nvl(rsBedChange!床号)
            strAfterBed = Nvl(rsBedChange!出院病床)
            If (Not IsNull(rsBedChange!病人ID)) And Val(Nvl(rsBedChange!病人ID, 0)) <> lng病人ID And Nvl(rsBedChange!附加床位, 0) = 0 Then
                strSql = "Select a.病人id, c.主页id, a.床号,NVL(a.附加床位,0) 附加床位, c.出院病床" & vbNewLine & _
                    " From 病人变动记录 a, 病案主页 c" & vbNewLine & _
                    " Where a.病人id = [1] And a.病人id = c.病人id And a.主页id = c.主页id And c.出院日期 Is Null And a.终止原因 = 4" & vbNewLine & _
                    " Order By 终止时间 Desc,NVL(A.附加床位,0) DESC, 开始时间 Desc"
                Set rsBedChange = zlDatabase.OpenSQLRecord(strSql, "", rsBedChange!病人ID, Nvl(rsBedChange!床号))
                If rsBedChange.RecordCount > 0 Then
                    If Nvl(rsBedChange!病人ID, 0) <> 0 And Nvl(rsBedChange!主页ID, 0) <> 0 And strBeforBed = Nvl(rsBedChange!出院病床, 0) And strAfterBed = Nvl(rsBedChange!床号) And Nvl(rsBedChange!附加床位, 0) = 0 Then
                        lngNextPati = Val(rsBedChange!病人ID)
                        lngNextPage = Val(rsBedChange!主页ID)
                        
                        strSql = "Select b.姓名, b.性别, b.住院号, a.Id 变动id,病区Id 撤销前病区ID,科室Id 撤销前科室Id,b.出院病床" & _
                        "   From 病人变动记录 a, 病案主页 b" & _
                        "   Where a.病人id = b.病人id And a.主页id = b.主页id And A.开始原因 = [3] And a.开始时间 IS NOT NULL And Nvl(a.附加床位, 0) = 0 And 终止时间 Is Null And b.病人id = [1] And b.主页id = [2]"
                        Set rsBedChange = zlDatabase.OpenSQLRecord(strSql, "撤销变动前", lngNextPati, lngNextPage, 4)
                        bln床位对换 = rsBedChange.RecordCount > 0
                    End If
                End If
            End If
         End If
    End If
        
    '撤销时未执行费用需要转回原病区
    If strType = "转病区入住" Then
        If CreatePublicExpenseBillOperation() And gbln转病区转费用 Then
            strSQLTmp = "Select ID, 病区id" & vbNewLine & _
                    "From 病人变动记录" & vbNewLine & _
                    "Where 病人id = [1] And 主页id = [2] And Nvl(附加床位, 0) = 0 And 终止原因 = 15 And" & vbNewLine & _
                    "      终止时间 = [3]"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQLTmp, App.ProductName, lng病人ID, lng主页ID, CDate(rsUndoBegin!开始时间))
            If rsPati.RecordCount > 0 Then
                If gobjPublicExpenseBillOperation.zlTurnToWard_Fee_Query(frmParent, 1, lng病人ID, lng主页ID, Val(rsPati!ID & ""), Val(rsUndoBegin!撤销前病区ID & ""), Val(rsPati!病区ID & ""), colSQL) = False Then Exit Function
            End If
        End If
    End If
        
    gcnOracle.BeginTrans: blnTrans = True
    '转病区撤销时费用先执行
    For i = 1 To colSQL.Count
        zlDatabase.ExecuteProcedure colSQL(i), App.ProductName
    Next
    For intLoop = LBound(arrSQL) To UBound(arrSQL)
        zlDatabase.ExecuteProcedure arrSQL(intLoop), App.ProductName
    Next
    
    '撤消出院时医保处理
    If blnUndoOut And int险类 <> 0 Then
        If Not gclsInsure.LeaveDelSwap(lng病人ID, lng主页ID, int险类) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    ExecUndo = True

    '完成病人变动的相关撤销操作触发消息
    On Error Resume Next
    '创建消息对象
    Set clsMipModule = New zl9ComLib.clsMipModule
    Call clsMipModule.InitMessage(glngSys, P病人入出管理, strPrivs, frmParent.hWnd)
    Call AddMipModule(clsMipModule)
    Set clsXML = New zl9ComLib.clsXML
    
    If clsMipModule.IsConnect = True Then
        
        strSql = "Select d.Id 部门id, d.名称 部门名称, p.Id 人员id, p.姓名 人员姓名" & vbNewLine & _
            "    From 人员表 p, 部门人员 r, 部门表 d" & vbNewLine & _
            "    Where p.Id = r.人员id And r.部门id = d.Id"
        Call zlDatabase.OpenRecordset(rsDeptOper, strSql, "获取部门人员信息")
PreNext:
        If strType = "转科入住" Or strType = "转病区入住" Then
            strSql = "Select  a.病区id 撤销后病区id, a.科室id 撤销后科室id, a.经治医师, a.主治医师, a.主任医师, a.病情, c.编码 病情编码, a.责任护士, B.出院病床 床号" & vbNewLine & _
                "   From 病人变动记录 a, 病案主页 b, 病情 c" & vbNewLine & _
                "   Where a.病人id = b.病人id And a.主页id = b.主页id And a.开始原因 = [3] And a.开始时间 Is Null And Nvl(a.附加床位, 0) = 0 And 终止时间 Is Null And" & vbNewLine & _
                "      a.病情 = c.名称(+) And b.病人id = [1] And b.主页id = [2]"
            Set rsUndoEnd = zlDatabase.OpenSQLRecord(strSql, "撤销变动后", lng病人ID, lng主页ID, IIf(strType = "转科入住", 3, 15))
        Else
            strSql = "Select a.病区Id 撤销后病区ID,a.科室Id 撤销后科室Id,a.经治医师,a.主治医师,a.主任医师,a.病情,c.编码 病情编码,a.责任护士,B.出院病床 床号" & _
                "   From 病人变动记录 a, 病案主页 b,病情 C" & _
                "   Where  a.病人id = b.病人id And a.主页id = b.主页id And a.开始时间 IS NOT NULL And Nvl(a.附加床位, 0) = 0 And 终止时间 Is Null And a.病情 = c.名称(+) And b.病人id = [1] And b.主页id = [2]"
            Set rsUndoEnd = zlDatabase.OpenSQLRecord(strSql, "撤销变动后", lng病人ID, lng主页ID)
        End If
        
        clsXML.ClearXmlText '清除缓存中的XML
        '--进行消息组装
        '病人信息
        clsXML.AppendNode "in_patient"
        'patient_id      病人id  1   N
        clsXML.appendData "patient_id", lng病人ID, xsNumber  '病人ID
        'page_id     主页id  1   N
        clsXML.appendData "page_id", lng主页ID, xsNumber '主页ID
        'patient_name        姓名    1   S
        clsXML.appendData "patient_name", Nvl(rsUndoBegin!姓名), xsString '姓名
        'patient_sex     性别    0..1    S
        clsXML.appendData "patient_sex", Nvl(rsUndoBegin!性别), xsString '性别
        'in_number       住院号  1   S
        clsXML.appendData "in_number", Nvl(rsUndoBegin!住院号), xsString '住院号
        clsXML.AppendNode "in_patient", True
        
        'change_cancel 1
        clsXML.AppendNode "change_cancel"
        'change_id       变动id  1   N
        clsXML.appendData "change_id", rsUndoBegin!变动ID, xsNumber
        'cancel_kind     撤消方式    1   S       换床、床位对换、转科、转病区、出院、入院入住、入住、转科入住、床位等级变动、护理等级变动、经治医师改变、责任护士改变、转为住院病人、预出院、主治医师变动、主任医师变动、病况变动、转医疗小组、转病区入住
        clsXML.appendData "cancel_kind", strType, xsString
        'before_area_id      撤销变动前病区id    0..1    N
        If Nvl(rsUndoBegin!撤销前病区ID, 0) <> 0 Then
            clsXML.appendData "before_area_id", Nvl(rsUndoBegin!撤销前病区ID, 0), xsNumber
        End If
        'before_dept_id      撤销变动前科室Id    0..1    N
        clsXML.appendData "before_dept_id", Nvl(rsUndoBegin!撤销前科室Id, 0), xsNumber
        'after_area_id       撤销变动后病区id    0..1    N
        If Nvl(rsUndoEnd!撤销后病区ID, 0) <> 0 Then
            clsXML.appendData "after_area_id", Nvl(rsUndoEnd!撤销后病区ID, 0), xsNumber
        End If
        'after_area_title 撤销变动后病区名称
        rsDeptOper.Filter = "部门ID=" & Val(Nvl(rsUndoEnd!撤销后病区ID, 0))
        If rsDeptOper.RecordCount > 0 Then
            clsXML.appendData "after_area_title", Nvl(rsDeptOper!部门名称), xsString
        End If
        'after_dept_id       撤销变动后科室id    0..1    N
        clsXML.appendData "after_dept_id", Nvl(rsUndoEnd!撤销后科室Id, 0), xsNumber
        'after_dept_title 撤销变动后科室名称
        rsDeptOper.Filter = "部门ID=" & Val(Nvl(rsUndoEnd!撤销后科室Id, 0))
        If rsDeptOper.RecordCount > 0 Then
            clsXML.appendData "after_dept_title", Nvl(rsDeptOper!部门名称), xsString
        Else
            clsXML.appendData "after_dept_title", "", xsString
        End If
        Select Case strType
        Case "换床"
            'after_duty_nurse_id 撤销变动后护士ID
            rsDeptOper.Filter = "人员姓名='" & Nvl(rsUndoEnd!责任护士) & "'"
            If rsDeptOper.RecordCount > 0 Then
                clsXML.appendData "after_duty_nurse_id", Val(Nvl(rsDeptOper!人员ID)), xsNumber
            End If
            'after_duty_nurse 撤销变动后护士姓名
            clsXML.appendData "after_duty_nurse", Nvl(rsUndoEnd!责任护士), xsString
            'after_bed_no 撤销变动后的床号
            clsXML.appendData "after_bed_no", Nvl(rsUndoEnd!床号), xsString
        Case "病况变动"
            'after_situation 撤销变动后的病情
            clsXML.appendData "after_situation", Nvl(rsUndoEnd!病情), xsString
            'after_situation_code 撤销变动后的病情编码
            clsXML.appendData "after_situation_code", Nvl(rsUndoEnd!病情编码), xsString
        Case "转科", "经治医师改变", "责任护士改变", "主治医师变动", "主任医师变动", "转医疗小组"
            'after_in_doctor_id 撤销变动后经治医生ID
            rsDeptOper.Filter = "人员姓名='" & Nvl(rsUndoEnd!经治医师) & "'"
            If rsDeptOper.RecordCount > 0 Then
                clsXML.appendData "after_in_doctor_id", Val(Nvl(rsDeptOper!人员ID)), xsNumber
            End If
            'after_in_doctor 撤销变动后经治医生姓名
            clsXML.appendData "after_in_doctor", Nvl(rsUndoEnd!经治医师), xsString
            'after_treat_doctor_id 撤销变动后主治医生ID
            rsDeptOper.Filter = "人员姓名='" & Nvl(rsUndoEnd!主治医师) & "'"
            If rsDeptOper.RecordCount > 0 Then
                clsXML.appendData "after_treat_doctor_id", Val(Nvl(rsDeptOper!人员ID)), xsNumber
            End If
            'after_treat_doctor 撤销变动后主治医生姓名
            clsXML.appendData "after_treat_doctor", Nvl(rsUndoEnd!主治医师), xsString
            'after_director_doctor_id 撤销变动后主任医生ID
            rsDeptOper.Filter = "人员姓名='" & Nvl(rsUndoEnd!主任医师) & "'"
            If rsDeptOper.RecordCount > 0 Then
                clsXML.appendData "after_director_doctor_id", Val(Nvl(rsDeptOper!人员ID)), xsNumber
            End If
            'after_director_doctor 撤销变动后主任医生姓名
            clsXML.appendData "after_director_doctor", Nvl(rsUndoEnd!主任医师), xsString
            'after_duty_nurse_id 撤销变动后护士ID
            rsDeptOper.Filter = "人员姓名='" & Nvl(rsUndoEnd!责任护士) & "'"
            If rsDeptOper.RecordCount > 0 Then
                clsXML.appendData "after_duty_nurse_id", Val(Nvl(rsDeptOper!人员ID)), xsNumber
            End If
            'after_duty_nurse 撤销变动后护士姓名
            clsXML.appendData "after_duty_nurse", Nvl(rsUndoEnd!责任护士), xsString
            'after_bed_no 撤销变动后的床号
            clsXML.appendData "after_bed_no", Nvl(rsUndoEnd!床号), xsString
        End Select
        clsXML.AppendNode "change_cancel", True
        clsMipModule.CommitMessage "ZLHIS_PATIENT_006", clsXML.XmlText
        If bln床位对换 = True Then
            bln床位对换 = False
            lng病人ID = lngNextPati
            lng主页ID = lngNextPage
            Set rsUndoBegin = rsBedChange
            GoTo PreNext
        End If
    End If
    '卸载消息对象
    If Not (clsMipModule Is Nothing) Then
        Call clsMipModule.CloseMessage
        Call DelMipModule(clsMipModule)
        Set clsMipModule = Nothing
    End If
    If Not (clsXML Is Nothing) Then
        Set clsXML = Nothing
    End If
    
    If Err <> 0 Then Err.Clear
    gblnOK = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOutState(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 出院方式 From 病案主页 Where 病人ID = [1] And 主页ID = [2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        If Not IsNull(rsTmp!出院方式) Then
            GetOutState = rsTmp!出院方式
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetUsedBeds(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef str等级IDs) As String
'功能:检查出院病人以前的床位是否仍是空床,只要其中之一不是,则返回病人的所有床号
'参数：str等级ID：传出，床位对应的等级ID
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strBeds As String
 
    strSql = "Select A.床号,B.状态,A.床位等级ID From 病人变动记录 A,床位状况记录 B  Where A.病人id=[1] And A.主页id=[2] And A.终止原因 = 1" & _
        " And A.病区id = B.病区id And A.床号 = B.床号"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng病人ID, lng主页ID)
    
    rsTmp.Filter = "状态<>'空床'"
    If rsTmp.RecordCount = 0 Then Exit Function
    
    rsTmp.Filter = ""
    Do While Not rsTmp.EOF
        strBeds = strBeds & "," & rsTmp!床号
        str等级IDs = str等级IDs & "," & rsTmp!床位等级id
        rsTmp.MoveNext
    Loop
    If strBeds <> "" Then
        GetUsedBeds = Mid(strBeds, 2)
        str等级IDs = Mid(str等级IDs, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
    Dim strValue As String
    On Error Resume Next
        
    '费用金额小数点位数
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '卡号显示方式
    'gblnShowCard = zldatabase.GetPara(12, glngSys) = "0"
    
    '多张单据使用多种结算方式模式
    gblnMultiBalance = zlDatabase.GetPara(79, glngSys) = "1"
       
    '票据号码长度、就诊卡号长度
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    gbytPrepayLen = Val(Split(strValue, "|")(1))
    'gbytCardNOLen = Val(Split(strValue, "|")(4))
    If gbytPrepayLen = 0 Then gbytPrepayLen = 7
    'If gbytCardNOLen = 0 Then gbytCardNOLen = 7
    
    '票号严格控制
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnPrepayStrict = Mid(strValue, 2, 1) = "1"
    'gblnMagcardStrict = Mid(strValue, 5, 1) = "1"
    
    
    gbln入院预交 = zlDatabase.GetPara(10, glngSys) = "1"
    gbln入院发卡 = zlDatabase.GetPara(11, glngSys) = "1"
    gbln入院入科 = zlDatabase.GetPara(13, glngSys) = "1"
    gbln在院病人不准出院结帐 = Val(zlDatabase.GetPara(31, glngSys))
    gbyt出院时检查未执行 = Val(zlDatabase.GetPara(22, glngSys))
    gbyt转科时检查未执行 = Val(zlDatabase.GetPara(32, glngSys))
    '问题30208 by lesfeng 2010-08-02 撤分参数22及32 新增154、155
    gbyt出院时检查药品未执行 = Val(zlDatabase.GetPara(154, glngSys))
    gbyt转科时检查药品未执行 = Val(zlDatabase.GetPara(155, glngSys))
    
    '61429:刘鹏飞,2013-11-11
    gbyt转科时未审核销帐单据检查 = Val(zlDatabase.GetPara(227, glngSys))
    '68953
    gbyt出院时超期护理数据检查 = Val(zlDatabase.GetPara(235, glngSys))
    
    gbln医生允许才能出院 = zlDatabase.GetPara(43, glngSys) = "1"
    '--55791:刘鹏飞,2012-11-13,回退出院医嘱才能撤销出院
    gbln医生允许才能撤销预出院 = (Val(zlDatabase.GetPara(192, glngSys)) = 1)
    
    '入院登记时刷卡输入密码
    gblnCheckPass = Mid(zlDatabase.GetPara(46, glngSys, , "0000000000"), 5, 1) = "1"
        
    '诊断输入方式
    strValue = zlDatabase.GetPara(65, glngSys, , "11")
    If Len(strValue) = 1 Then strValue = strValue & strValue
    gint门诊诊断输入 = Mid(strValue, 1, 1)
    gint住院诊断输入 = Mid(strValue, 2)
        
    gbln药疗划价单 = zlDatabase.GetPara(79, glngSys) = "1"
    gbln其他划价单 = zlDatabase.GetPara(80, glngSys) = "1"
    gbln执行后审核 = zlDatabase.GetPara(81, glngSys) = "1"
    gbln入科确定护理等级 = zlDatabase.GetPara(99, glngSys) = "1"
    gbln每次住院新住院号 = zlDatabase.GetPara(145, glngSys) = "1"
    '刘兴洪 问题:????    日期:2010-12-06 23:38:53
    '费用单价保留位数
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    '61347:刘鹏飞,2013-11-09
    gbyt病人审核方式 = Val(zlDatabase.GetPara(185, glngSys, , "0"))
    
    gblnXW = Val(zlDatabase.GetPara(255, glngSys)) = 1
    '同一身份证只能对应一个建档病人
    gblnPatiByID = Val(zlDatabase.GetPara(279, glngSys)) = 1
    gbln医疗机构不允许自由录入 = zlDatabase.GetPara(287, glngSys, , "0") = "1"
    
    InitSysPar = True
End Function

Public Sub InitLocPar(ByVal lngModul As Long)
'功能：初始化模块参数
    Dim strValue As String
    On Error Resume Next

    gstrLike = IIf(zlDatabase.GetPara("输入匹配") = 0, "%", "")
    strValue = zlDatabase.GetPara("输入法")
    gstrIme = IIf(strValue = "", "不自动开启", strValue)
    gbytCode = Val(zlDatabase.GetPara("简码方式"))
    gblnMyStyle = zlDatabase.GetPara("使用个性化风格") = "1"
    
    gbln启用结构化地址 = Val(zlDatabase.GetPara("病人地址结构化录入", glngSys)) <> 0
    gbln显示乡镇 = Val(zlDatabase.GetPara("乡镇地址结构化录入", glngSys)) <> 0
          
    If lngModul = P病人入院管理 Then
        gstr预交ID = zlDatabase.GetPara("共用预交票据批次", glngSys, lngModul, "")
        If gstr预交ID <> "" Then Call UpdateShareID(lngModul, gstr预交ID, 2)
        'LED语音报价
        gblnLED = Val(GetSetting("ZLSOFT", "公共全局", "使用", 0)) <> 0
        gblnLedWelcome = Val(zlDatabase.GetPara("LED显示欢迎信息", glngSys, lngModul, 1)) <> 0
        
        gbln担保 = zlDatabase.GetPara("担保信息", glngSys, lngModul) = "1"
        gblnSeekName = zlDatabase.GetPara("姓名模糊查找", glngSys, lngModul) = "1"
        gintNameDays = Val(zlDatabase.GetPara("姓名查找天数", glngSys, lngModul))
        gbln记账 = zlDatabase.GetPara("卡费记帐", glngSys, lngModul) = "1"
        gbln先选病区 = zlDatabase.GetPara("先选病区", glngSys, lngModul) = "1"
        gbln控制空床 = zlDatabase.GetPara("科室下无空床不能登记", glngSys, lngModul) = "1"
        
        gbytPrepayPrint = Val(zlDatabase.GetPara("预交款票据打印", glngSys, lngModul))
        gbytFPagePrint = Val(zlDatabase.GetPara("病案首页打印", glngSys, lngModul))
        gbytWristletPrint = Val(zlDatabase.GetPara("病人腕带打印", glngSys, lngModul))
        
        '36454,刘鹏飞,2012-09-06
        gbln费用计算 = Val(zlDatabase.GetPara("费用计算时机", glngSys, lngModul)) = 1
        
    ElseIf lngModul = P病人入出管理 Then
        gbyt入科时间 = Val(zlDatabase.GetPara("缺省入科时间", glngSys, lngModul, "0"))
        gbytBabyWristletPrint = Val(zlDatabase.GetPara("婴儿腕带打印", glngSys, lngModul))
        '49854:刘鹏飞,2013-10-31,病人腕带打印
        gbytCourseWristletPrint = Val(zlDatabase.GetPara("病人腕带打印", glngSys, lngModul))
        gbln转病区转费用 = zlDatabase.GetPara("转病区转费用", glngSys, lngModul) = "1"
    End If
End Sub

Public Function SaveIDCard(bytStyle As Byte, strNo As String, lng病人ID As Long, lng主页ID As Long, _
        lng病人病区ID As Long, lng病人科室ID As Long, str标识号 As String, str费别 As String, _
        str原卡号 As String, str姓名 As String, str性别 As String, str年龄 As String, str卡号 As String, str密码 As String, _
        cur应收金额 As Currency, cur实收金额 As Currency, str结算方式 As String, Dat发卡时间 As Date, lng领用ID As Long, rsMoney As ADODB.Recordset, ByVal strICCard As String) As String
'功能：产生一条就诊卡费用记录SQL语句
'参数：bytStyle=0-发卡,1-补卡,2-换卡
'      cur金额=就诊卡金额
'      str结算方式=如果为空,表示记帐,不收现金
'      rsMoney:包括就诊卡收费信息的记录集
'      str原卡号=仅换卡时用
'      lng领用ID=当前可用的就诊卡领用ID
'      strICCard=IC卡号,通过读IC卡方式发卡时,同时填写病人信息的IC卡字段
    Dim lngUnitID As Long
    Dim strSql As String
    
'    Select Case rsMoney!科室标志
'        Case 1 '指定科室
'            lngUnitID = GetItemUnitID(rsMoney!科室标志, rsMoney!收费细目ID)
'        Case 2 '病人科室
'            If lng病人科室ID <> 0 Then
'                lngUnitID = lng病人科室ID
'            Else
'                lngUnitID = UserInfo.部门ID
'            End If
'        Case 0, 3, 5, 6
'            lngUnitID = UserInfo.部门ID
'        Case 4 '指定科室
'            lngUnitID = GetItemUnitID(rsMoney!科室标志, rsMoney!收费细目ID)
'    End Select
    
    '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
    Select Case rsMoney!科室标志
        Case 4 '指定科室
            lngUnitID = GetItemUnitID(rsMoney!科室标志, rsMoney!收费细目ID)
        Case 1, 2 '病人科室
            If lng病人科室ID <> 0 Then
                lngUnitID = lng病人科室ID
            Else
                lngUnitID = UserInfo.部门ID
            End If
        Case 0, 3, 5, 6
            lngUnitID = UserInfo.部门ID
    End Select
    
    '调用过程"zl_就诊卡记录_Insert"
    strSql = "zl_就诊卡记录_INSERT(" & bytStyle & ",'" & strNo & "'," & lng病人ID & "," & lng主页ID & "," & _
        IIf(str标识号 = "0", "NULL", str标识号) & ",'" & str费别 & "','" & str原卡号 & "','" & str卡号 & "','" & str密码 & "','" & str姓名 & _
        "','" & str性别 & "','" & str年龄 & "'," & lng病人病区ID & "," & lng病人科室ID & "," & rsMoney!收费细目ID & _
        ",'" & rsMoney!收费类别 & "','" & IIf(IsNull(rsMoney!计算单位), "", rsMoney!计算单位) & "'," & _
        rsMoney!收入项目ID & ",'" & rsMoney!收据费目 & "'," & cur应收金额 & "," & lngUnitID & "," & UserInfo.部门ID & _
        ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(OverTime(Dat发卡时间), "1", "0") & _
        ",To_Date('" & Format(Dat发卡时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
        "'" & str结算方式 & "'," & IIf(lng领用ID = 0, "NULL", lng领用ID) & ",'" & strICCard & "'," & cur应收金额 & "," & cur实收金额 & ")"
    
    SaveIDCard = strSql
End Function

Private Function GetItemUnitID(bytFlag As Byte, lngID As Long) As Long
'功能：返回收费特定项目的执行科室
'参数：bytFlag=执行科室标志,lngID=收费细目ID
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '无明确科室
            GetItemUnitID = UserInfo.部门ID '取操作员所在科室
        Case 4 '指定科室
            strSql = "Select B.执行科室ID From 收费项目目录 A,收费执行科室 B Where B.收费细目ID=A.ID And A.ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lngID)
            
            If Not rsTmp.EOF Then
                GetItemUnitID = rsTmp!执行科室ID '默认取第一个(如有多个)
            Else
                GetItemUnitID = UserInfo.部门ID '如没有指定，则取操作员所在科室
            End If
        Case 1, 2, 3 '病人科室,操作员科室
            GetItemUnitID = UserInfo.部门ID '都取操作员科室
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetSelectPersonal(ByVal strType As String, ByVal strPosts As String, ByRef frmParent As Form) As ADODB.Recordset
'参数:strType=人员性质
'     strPosts=专业技术职务 以逗号分隔
    Dim strSql As String
    
    On Error GoTo errH
    strSql = _
            "Select ID,上级ID,0 as 末级,编码 as 编码,简码,名称" & _
            " From 部门表 Where 编码 Not like '-%' Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
            " And (站点='" & gstrNodeNo & "' Or 站点 is Null)" & _
            " Union All" & _
            " Select Distinct A.ID,C.部门ID as 上级ID,1 as 末级,A.编号,A.简码,A.姓名" & _
            " From 人员表 A,人员性质说明 B,部门人员 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID And B.人员性质='" & strType & _
            "' And C.缺省=1 And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
            
    If strPosts <> "" Then
        strSql = strSql & " And Instr(',' || '" & strPosts & "' || ',', ',' || A.专业技术职务 || ',') > 0"
    End If
    
    Set GetSelectPersonal = frmPubSel.ShowSelect(frmParent, strSql, 2, "人员选择")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDepCharacter(ByVal lngDepID As Long) As String
'功能：获取部门工作性质
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select 工作性质 From 部门性质说明 Where 部门ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lngDepID)
    Do While Not rsTmp.EOF
        If InStr(1, GetDepCharacter & ",", "," & rsTmp!工作性质 & ",") = 0 Then
            GetDepCharacter = GetDepCharacter & "," & rsTmp!工作性质
        End If
        rsTmp.MoveNext
    Loop
    
    If GetDepCharacter <> "" Then GetDepCharacter = Mid(GetDepCharacter, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'问题27392 by lesfeng 2010-01-14
Public Function GetPatiInfoModiOut(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人相关信息
'参数：byt入科方式:0-入院入科，1-转科入科
    Dim strSql As String
    On Error GoTo errH
    '问题28982 by lesfeng 2010-06-09 增加“确诊日期”
    strSql = "Select NVl(B.姓名,A.姓名) 姓名, Nvl(NVL(B.性别,A.性别),'未知') 性别, NVL(B.年龄,A.年龄) 年龄, B.险类, B.病人性质, B.当前病况, B.护理等级id, B.住院医师, B.门诊医师, B.责任护士, B.出院科室id, B.出院科室id 入住科室id," & vbNewLine & _
        "       To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, B.当前病区id, B.住院号, D.名称 as 当前科室, B.出院病床 as 主要床号," & vbNewLine & _
        "     To_Char(B.出院日期, 'YYYY-MM-DD HH24:MI:SS') As 出院日期,是否确诊,确诊日期,出院方式,随诊标志,随诊期限,尸检标志" & vbNewLine & _
        "From 病人信息 A, 病案主页 B, 部门表 D" & vbNewLine & _
        "Where A.病人id = B.病人id And B.病人id = [1] And B.主页id = [2] And B.出院科室id = D.id and B.出院日期 is not null"
    Set GetPatiInfoModiOut = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional byt入住方式 As Byte) As ADODB.Recordset
'功能：获取病人相关信息
'参数：byt入住方式:0-入院入科，1-转科入科，2-入病区
    Dim strSql As String
    '问题31652 by lesfeng 从病案主页直接提取确诊日期，增加是否确诊及确诊日期
    '49163,刘鹏飞,2012-09-07,从病案主页直接提取随诊标志和随诊期限
    On Error GoTo errH
    If byt入住方式 = 0 Then '
        strSql = "Select NVL(B.姓名,A.姓名) 姓名, Nvl(NVL(B.性别,A.性别),'未知') 性别, B.年龄,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI:SS') As 出生日期, B.险类, B.病人性质, B.当前病况, B.护理等级id, B.住院医师, B.门诊医师, B.责任护士, B.出院科室id, B.出院科室id 入住科室id, B.医疗小组id, " & vbNewLine & _
            "       To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, B.当前病区id, B.住院号, D.名称 as 当前科室,E.名称 as 当前病区, B.出院病床 as 主要床号,B.是否确诊,B.确诊日期, B.出院日期, B.出院方式, B.尸检标志,B.随诊标志,B.随诊期限,B.再入院,B.挂号ID " & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 部门表 D,部门表 E" & vbNewLine & _
            "Where A.病人id = B.病人id And B.病人id = [1] And B.主页id = [2] And B.出院科室id = D.id And B.当前病区Id=E.id(+)"
    ElseIf byt入住方式 = 1 Then
        strSql = "Select NVL(B.姓名,A.姓名) 姓名, Nvl(NVL(B.性别,A.性别),'未知') 性别, B.年龄,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI:SS') As 出生日期, B.险类, B.病人性质, B.当前病况, B.护理等级id, B.住院医师, B.门诊医师, B.责任护士, B.出院科室id, C.科室id As 入住科室id, B.医疗小组id, " & vbNewLine & _
            "       To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, B.当前病区id, B.住院号, D.名称 as 当前科室,E.名称 as 当前病区, B.出院病床 as 主要床号,B.是否确诊,B.确诊日期, B.出院日期, B.出院方式, B.尸检标志,B.随诊标志,B.随诊期限,B.再入院,B.挂号ID" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病人变动记录 C, 部门表 D,部门表 E" & vbNewLine & _
            "Where A.病人id = B.病人id And B.病人id = [1] And B.主页id = [2] And C.病人id = B.病人id And C.主页id = B.主页id And" & vbNewLine & _
            "      C.开始原因 = 3 And C.开始时间 Is Null And C.终止时间 Is Null And B.出院科室id = D.id And B.当前病区Id=E.id(+)"
    ElseIf byt入住方式 = 2 Then
        strSql = "Select NVL(B.姓名,A.姓名) 姓名, Nvl(NVL(B.性别,A.性别),'未知') 性别, B.年龄,To_Char(A.出生日期,'YYYY-MM-DD HH24:MI:SS') As 出生日期, B.险类, B.病人性质, B.当前病况, B.护理等级id, B.住院医师, B.门诊医师, B.责任护士, B.出院科室id, C.科室id As 入住科室id, B.医疗小组id, " & vbNewLine & _
            "       To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, B.当前病区id, B.住院号, D.名称 as 当前科室,E.名称 as 当前病区, B.出院病床 as 主要床号,B.是否确诊,B.确诊日期, B.出院日期, B.出院方式, B.尸检标志,B.随诊标志,B.随诊期限,B.再入院,B.挂号ID" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病人变动记录 C, 部门表 D,部门表 E" & vbNewLine & _
            "Where A.病人id = B.病人id And B.病人id = [1] And B.主页id = [2] And C.病人id = B.病人id And C.主页id = B.主页id And" & vbNewLine & _
            "      C.开始原因 = 15 And C.开始时间 Is Null And C.终止时间 Is Null And B.出院科室id = D.id And B.当前病区Id=E.id(+)"
    End If
    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInDept(blnByDept As Boolean, DateBegin As Date, DateEnd As Date, ByVal strNodeNo As String, _
    Optional ByVal strDeptIDs As String) As ADODB.Recordset
'功能：获取指定入院时间范围内的科室或病区
'参数：strDeptIDs-可选科室或病区ID
    Dim strSql As String

    On Error GoTo errH
    If strDeptIDs <> "" Then strSql = strSql & "  And Instr(','||[4]||',',','||B.ID||',')>0 "
    strSql = "Select A.ID, B.编码, B.名称" & vbNewLine & _
            "From (Select Distinct " & IIf(blnByDept, "入院科室id", "入院病区id") & " ID" & vbNewLine & _
            "       From 病案主页" & vbNewLine & _
            "       Where 入院日期 Between [1] And [2] And " & IIf(blnByDept, "入院科室id", "入院病区id") & " Is Not Null) A, 部门表 B" & vbNewLine & _
            "Where A.ID = B.ID  And (B.站点=[3] Or B.站点 is Null) " & strSql & vbNewLine & _
            "Order By B.编码"

    Set GetInDept = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", DateBegin, DateEnd, strNodeNo, strDeptIDs)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SimilarIDs(str国籍 As String, str民族 As String, dat出生日期 As Date, str性别 As String, str姓名 As String, str身份证号 As String) As ADODB.Recordset
'功能：检查病人是否存在相似信息
'返回：相似记录的病人ID串,如"234,235,236"
    On Error GoTo errH
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Integer
    
    strSql = _
        " Select Rownum+1 ID,病人ID,Nvl(身份证号,'未登记') 身份证号,门诊号,住院号,Nvl(家庭地址,'未登记') 地址,To_Char(登记时间,'YYYY-MM-DD') 登记时间 " & _
        " From 病人信息 Where (国籍=[1] And 民族=[2] And 性别=[3] And 姓名=[4] " & _
        " And 出生日期=TO_DATE([5],'YYYY-MM-DD')) Or 身份证号=[6] " & _
        " Order by 病人ID Desc"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", str国籍, str民族, str性别, str姓名, Format(dat出生日期, "YYYY-MM-DD"), str身份证号)
    
    Set SimilarIDs = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMax主页ID(lng病人ID As Long) As Long
'功能：获取病人的最大病案主页ID
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select Nvl(Max(主页ID),0)+1 as 主页ID From 病案主页 Where Nvl(主页ID,0)<>0 And 病人ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID)
    If rsTmp.EOF Then
        GetMax主页ID = 1
    Else
        GetMax主页ID = rsTmp!主页ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function NextBedNo(lngUnitID As Long, str编制类型 As String, str符号 As String) As String
'功能：获取指定病区的下一床位号
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errH
    If str符号 <> "" Then
        gstrSQL = _
                "Select nvl(Max(to_number(Substr(床号, nvl(Length(B.符号),0) + 1))),0) 床号,nvl(max(length(床号)-length(符号)),1) 位长 " & _
                "From 床位状况记录 A, 床位编制分类 B " & _
                "Where A.床位编制 = B.名称 And 床位编制=[1] and  A.床号 like [2] and A.病区id=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlInPatient", str编制类型, str符号 & "%", lngUnitID)
    Else
        gstrSQL = _
                "Select nvl(Max(to_number(Substr(床号, nvl(Length(B.符号),0) + 1))),0) 床号,nvl(max(length(床号)-length(符号)),1) 位长 " & _
                "From 床位状况记录 A, 床位编制分类 B " & _
                "Where A.床位编制 = B.名称 And 床位编制=[1] and  zl_to_number(A.床号)>0 and A.病区id=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlInPatient", str编制类型, lngUnitID)

    End If
    
    If Not rsTmp.EOF Then
        If IsNumeric(rsTmp!床号) Then
            strTmp = rsTmp!床号 + 1
        Else
            strTmp = Val(rsTmp!床号) + 1
        End If
    Else
        NextBedNo = 1
    End If
    
    If Len(strTmp) > rsTmp!位长 Then
        NextBedNo = strTmp
    Else
        NextBedNo = Right("00000" & strTmp, rsTmp!位长)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function isRepeat(lngUnitID As Long, strBeds As String) As String
'功能：判断在指定病区内的一系列床号是否已经存在
'参数：lngUnitID=病区ID,strBeds=床号字符串,如"12,13,15..."
'返回：空=都不存在,否则"12,13..."这些床号重复
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 床号 From 床位状况记录 Where 病区ID=[1] And instr([2],','||床号||',')>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lngUnitID, "," & strBeds & ",")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            isRepeat = isRepeat & rsTmp!床号 & ","
            rsTmp.MoveNext
        Next
        isRepeat = Left(isRepeat, Len(isRepeat) - 1)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistInPatiNO(str住院号 As String, Optional lng病人ID As Long, Optional bln不含本人 As Boolean) As Boolean
'功能：判断指定住院号是否已经存在于数据库中,每次住院新住院号时,预约和修改仍用原来的住院号,排开病人自己
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    If gbln每次住院新住院号 Then
        strSql = "Select 1 From 病案主页 Where 住院号=[1]" & IIf(bln不含本人, " And 病人ID<>[2]", "")
    Else
        strSql = _
            " Select 1 From 病人信息 Where 住院号=[1] And 病人ID<>[2]" & vbNewLine & _
            " UNION ALL" & vbNewLine & _
            " Select 1 From 病案主页 Where 住院号=[1] And 病人ID<>[2]"
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", str住院号, lng病人ID)
    If rsTmp.RecordCount > 0 Then ExistInPatiNO = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistInPatiID(lngPatientID As Long) As Boolean
'功能：判断指定病人ID是否已经存在于数据库中
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errH
    
    strSql = "Select 1 From 病人信息 Where 病人ID=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lngPatientID)
    If rsTmp.RecordCount > 0 Then ExistInPatiID = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check费别适用科室(ByVal str费别 As String, ByVal lng科室ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errH
    
    '适用于所有科室,或当前指定科室
    strSql = "Select 1" & vbNewLine & _
            "From Dual" & vbNewLine & _
            "Where Not Exists (Select 1 From 费别适用科室 Where 费别 = [1]) Or Exists" & vbNewLine & _
            " (Select 1 From 费别适用科室 Where 费别 = [1] And 科室id = [2])"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", str费别, lng科室ID)
    If rsTmp.RecordCount > 0 Then Check费别适用科室 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMaxDate(lng病人ID As Long, lng主页ID As Long, Optional int原因 As Integer) As Date
'功能：获取转科病人最大的上次变动时间
'参数：int原因=返回上次变动的原因
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    GetMaxDate = #1/1/1900#
    int原因 = 0
    
    strSql = "Select 开始时间,开始原因 From 病人变动记录" & _
        " Where 开始时间 is Not NULL And 终止时间 is NULL" & _
        " And 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        GetMaxDate = IIf(IsNull(rsTmp!开始时间), GetMaxDate, rsTmp!开始时间)
        int原因 = Nvl(rsTmp!开始原因, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLastInfo(lngPatientID As Long) As String
'功能：获取病人最后一次预交款单位信息
'返回："缴款单位|单位开户行|单位帐号"
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    'by lesfeng 2010-01-11 性能优化
    '大改：记录性质=1
    strSql = "Select 缴款单位,单位开户行,单位帐号 From 病人预交记录 Where (缴款单位 is Not NULL Or 单位开户行 is Not NUll Or 单位帐号 is Not NULL) And 记录性质=1 And 病人ID=[1] Order by 收款时间 Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lngPatientID)
    
    If Not rsTmp.EOF Then
        GetLastInfo = IIf(IsNull(rsTmp!缴款单位), "", rsTmp!缴款单位) & "|" & IIf(IsNull(rsTmp!单位开户行), "", rsTmp!单位开户行) & "|" & IIf(IsNull(rsTmp!单位帐号), "", rsTmp!单位帐号)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check医生下达出院医嘱(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：判断病人是否处于预出院状态,且存在有效的出院(转院、死亡)医嘱才允许出院(有效的医嘱是指开始执行时间与预出院时间相同，且处于已发送状态[医嘱状态=8])。
'参数：
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    '--55791:刘鹏飞,2012-11-13,作废出院医嘱才能撤销出院
     '如果参数"作废出院医嘱才能撤销出院"为假，医嘱记录就和病人变动记录无法对应。原有SQL(注释内容),就不能允许病人预出院或出院
'    strSQL = "Select a.Id" & vbNewLine & _
'            " From 病人医嘱记录 a, 病人变动记录 b, 病案主页 c, 诊疗项目目录 d" & vbNewLine & _
'            " Where a.病人id = [1] And a.主页id = [2] And a.医嘱状态 = 8 And a.病人id = b.病人id And a.主页id = b.主页id And" & vbNewLine & _
'            "           a.开始执行时间 = b.开始时间+0 And b.开始原因 = 10 And b.病人id = c.病人id And b.主页id = c.主页id And" & vbNewLine & _
'            "           c.状态 = 3 And d.类别='Z' And d.操作类型 In ('5', '6', '11') And a.诊疗项目id = d.Id"
    '102342 启用参数时："下达出院医嘱才能出院",当婴儿下达出院医嘱,病人本身未下达时,不允许病人出院。
    strSql = "Select a.Id" & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
            "Where a.病人id = [1] And a.主页id = [2] And a.医嘱状态 = 8 And Nvl(a.婴儿, 0) = 0 And b.类别 = 'Z' And b.操作类型 In ('5', '6', '11') And" & vbNewLine & _
            "      a.诊疗项目id = b.Id "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID)
    
    Check医生下达出院医嘱 = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check销帐申请(lng病人ID As Long, lng主页ID As Long) As String
'功能：判断病人是存在销帐申请记录
'参数：
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strInfo As String
    '56323:刘鹏飞,2013-02-18,加强销帐为审核单据的提示信息内容
    'strSQL = "Select 1 From 住院费用记录 A, 病人费用销帐 B Where A.病人ID=[1] And A.主页ID=[2] And A.Id = B.费用ID And b.状态=0"
    strSql = "Select distinct A.NO,C.名称 审核科室,D.名称 项目名称 From 住院费用记录 A, 病人费用销帐 B,部门表 C,收费项目目录 D" & vbNewLine & _
        "        Where A.病人ID=[1] And A.主页ID=[2] And A.Id = B.费用ID And b.状态=0 And B.审核部门ID=C.ID And B.收费细目ID=D.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID)
    With rsTmp
        Do While Not .EOF
            If strInfo = "" Then
                strInfo = "单据[" & Nvl(!NO) & "]中的" & Nvl(!项目名称) & "：在" & Nvl(!审核科室, "[未定部门]") & "未审核"
            Else
                strInfo = strInfo & vbCrLf & "单据[" & Nvl(!NO) & "]中的" & Nvl(!项目名称) & "：在" & Nvl(!审核科室, "[未定部门]") & "未审核"
            End If
        rsTmp.MoveNext
        Loop
    End With
    Check销帐申请 = strInfo
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAudited(lng病人ID As Long, lng主页ID As Long, Optional bytAudited As Byte = 1) As Boolean
'功能：判断病人是否已审核
'参数：
'      lng病人ID:病人ID
'      lng主页ID:主页ID
'      bytAudited:0-未审核;1-已审核;2-完成审核
'返回:TRUE OR FALSE
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    If bytAudited = 0 Then
        '未审核
        strSql = "Select 病人id From 病案主页 Where 病人id=[1] And 主页id=[2] And Nvl(审核标志,0)=0"
    ElseIf bytAudited = 1 Then
        '已经审核
        strSql = "Select 病人id From 病案主页 Where 病人id=[1] And 主页id=[2] And Nvl(审核标志,0)>=1"
    Else
        '完成审核
        strSql = "Select 病人id From 病案主页 Where 病人id=[1] And 主页id=[2] And Nvl(审核标志,0)=2"
    End If
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID)
    CheckAudited = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMaxBedLen(Optional lng部门ID As Long, Optional bln占用 As Boolean) As Integer
'功能：获取指定部门的床位号的最大长度
'参数：lng部门ID=病区ID或科室ID,为0表示所有病区或科室
'      bln占用=是否只管被占用的床
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    If Not bln占用 Or lng部门ID = 0 Then
        strSql = "Select Max(Lengthb(床号)) as 长度 From 床位状况记录 Where  病区ID" & IIf(lng部门ID = 0, " is Not NULL", "=[1]")
    Else
        strSql = "Select Max(Lengthb(床号)) as 长度 From 床位状况记录 Where 状态='占用' And 病区ID" & IIf(lng部门ID = 0, " is Not NULL", "=[1]")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInExse", lng部门ID)
    If Not rsTmp.EOF Then GetMaxBedLen = IIf(IsNull(rsTmp!长度), 0, rsTmp!长度)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetPatiLog(lng病人ID As Long, lng主页ID As Long) As ADODB.Recordset
'功能：获取病人变动记录
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    'by lesfeng 2010-01-11 性能优化
    strSql = "Select id,病人ID ,主页ID ,开始时间 ,开始原因,附加床位,病区id,科室id,护理等级id,床位等级id,床号," & _
             "       责任护士,经治医师,主治医师,主任医师,病情,终止人员,终止时间,终止原因,操作员编号,操作员姓名,上次计算时间 " & _
             "  From 病人变动记录" & _
             " Where Nvl(附加床位,0)=0 And 病人ID=[1] And 主页ID=[2] " & _
             " Order by 终止时间 Desc,开始时间 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID)
    
    If Not rsTmp.EOF Then Set GetPatiLog = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetDiagnosticInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
                                  ByVal str诊断类型 As String, ByVal str记录来源 As String, Optional ByVal lngDeptID As Long = 0) As ADODB.Recordset
'功能：获取指定病人的诊断记录'
'参数:
'诊断类型-1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;
'        11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断
'记录来源:1-病历；2-入院登记；3-首页整理(门诊医生站,诊断摘要);
'lngDeptID:入院科室ID（接收预约病人）
    On Error GoTo errH
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intDiagDays As Integer
    
    If lng主页ID = 0 And InStr(1, "," & str记录来源 & ",", ",3,") > 0 And (InStr(1, "," & str诊断类型 & ",", ",1,") > 0 Or InStr(1, "," & str诊断类型 & ",", ",11,") > 0) Then
        '107823接收预约病人时提取门诊诊断规则:
        '1-限制3天内挂号的,避免本次没有挂号的病人把上次的诊断记录读出
        '2-优先取有效天数内，入院科室ID对应入院医嘱所对应的诊断
        '3-未取到入院医嘱对应的诊断,则取有效天数内的最近一次挂号记录对应的第一诊断
        intDiagDays = Val(zlDatabase.GetPara("诊断查找天数", glngSys, glngModul, "3"))
        If lngDeptID = 0 Then
            strSql = "Select 入院科室id　from 病案主页 Where 病人id = [1] And 主页id = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID)
            If Not rsTmp.EOF Then lngDeptID = Val(rsTmp!入院科室ID & "")
        End If
        strSql = "Select a.诊断类型, a.记录来源, a.诊断描述, a.疾病id, a.诊断id, a.出院情况, a.记录日期, a.是否疑诊, a.主页id" & vbNewLine & _
                "From 病人诊断记录 A, 病人诊断医嘱 B, 病人医嘱记录 C, 诊疗项目目录 D, 病人挂号记录 E" & vbNewLine & _
                "Where a.Id = b.诊断id And b.医嘱id = c.Id And c.诊疗项目id + 0 = d.Id And a.主页id = e.Id And a.病人id = [1] And a.记录来源 = 3  And INSTR([2], ',' || A.诊断类型 || ',') > 0 And" & vbNewLine & _
                "      e.记录性质 = 1 And e.记录状态 = 1 And e.登记时间 + 0 > Trunc(Sysdate - [3]) And c.医嘱状态 In (3, 8) And d.类别 = 'Z' And" & vbNewLine & _
                "      Instr(',1,2,', d.操作类型) > 0 And c.执行科室id = [4]" & vbNewLine & _
                "Order By a.记录日期 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, "," & str诊断类型 & ",", intDiagDays, lngDeptID)
        If rsTmp.EOF Then
            strSql = "Select a.诊断类型, a.记录来源, a.诊断描述, a.疾病id, a.诊断id, a.出院情况, a.记录日期, a.是否疑诊, a.主页id" & vbNewLine & _
                    "From 病人诊断记录 A, 病人挂号记录 B" & vbNewLine & _
                    "Where a.主页id = b.Id And a.病人id = [1] And b.记录性质 = 1 And b.记录状态 = 1 And b.登记时间 + 0 > Trunc(Sysdate - [3]) And a.诊断次序 = 1 And" & vbNewLine & _
                    "      Instr([2], ',' || a.诊断类型 || ',') > 0 And a.记录来源 = 3 " & vbNewLine & _
                    "Order By a.记录日期 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, "," & str诊断类型 & ",", intDiagDays)
        End If
    Else
        strSql = " Select 诊断类型,记录来源,诊断描述,疾病ID,诊断ID,出院情况,记录日期,是否疑诊 From 病人诊断记录 " & _
                 " Where 病人ID=[1] And Nvl(主页ID,0)=[2]" & _
                 " And 诊断次序=1  And NVL(编码序号,1) = 1 And instr([3],','||诊断类型||',')>0 And 记录来源 in (" & str记录来源 & ")" & _
                 " Order by 记录日期 Desc"
        '诊断次序-出院时,病案主页整理中可能填写主要诊断,次要诊断等多条记录
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID, "," & str诊断类型 & ",")
    End If
    
    If lng主页ID = 0 Then
        If Not rsTmp.EOF Then
            '103952多次挂号记录,取最近一次挂号记录
            lng主页ID = rsTmp!主页ID
            rsTmp.Filter = "主页ID =" & lng主页ID
        End If
        If Not rsTmp.EOF Then
            Set GetDiagnosticInfo = zlDatabase.CopyNewRec(rsTmp)
        Else
            Set GetDiagnosticInfo = Nothing
        End If
    Else
        If Not rsTmp.EOF Then Set GetDiagnosticInfo = rsTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsureInfo(lng病人ID As Long) As String
'功能：获取住院病人保险帐户信息
'返回："险类名;医保号"
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '连接病案主页,确保本次住院是保险病人,但不一定在院
    strSql = "Select A.名称,B.医保号" & _
        " From 保险类别 A,保险帐户 B,病人信息 C,病案主页 D" & _
        " Where A.序号=B.险类 And B.病人ID=C.病人ID" & _
        " And B.险类=D.险类 And C.病人ID=D.病人ID" & _
        " And D.主页ID=C.主页ID And C.病人ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID)
    
    If Not rsTmp.EOF Then GetInsureInfo = rsTmp!名称 & ";" & rsTmp!医保号
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function



Public Function GetLastAdviceTime(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Date
'功能：获取指定病人最后一条有效的医嘱的时间
'说明：用于病人出院时判断出院时间必须大于该时间
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    GetLastAdviceTime = CDate("1900-01-01")
    
    On Error GoTo errH
    
    '以长嘱最后执行时间为准判断,暂时排开持续性长嘱
    '临嘱有离院带药的情况,如以"出院"医嘱为准,出院时间本来就必须大于该变动时间。
    strSql = "Select Max(Nvl(执行终止时间,Nvl(上次执行时间,开始执行时间))) as 时间" & _
        " From 病人医嘱记录" & _
        " Where Nvl(医嘱期效,0)=0 And 医嘱状态 Not IN(1,2,4)" & _
        " And Not (执行时间方案 is NULL And (Nvl(频率次数, 0) = 0 Or Nvl(频率间隔, 0) = 0 Or 间隔单位 is NULL))" & _
        " And 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID)
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!时间) Then
            GetLastAdviceTime = rsTmp!时间
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveCatalogue(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：判断本次住院病案是否已编目
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select 编目日期 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID)
    
    If Not rsTmp.EOF Then
        HaveCatalogue = Not IsNull(rsTmp!编目日期)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function inBlackList(ByVal lng病人ID As Long) As String
'功能：判断病人是否在黑名单中,并反回加入原因
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select 加入原因 From 特殊病人 Where 撤消时间 is NULL And 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPublic", lng病人ID)
    
    If Not rsTmp.EOF Then inBlackList = rsTmp!加入原因
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitExe(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：检查病人在医技科室是否还有未执行完成(未执行或正在执行)的项目
'返回：医技科室名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select Zl_Pati_Check_Execute(2,[1],[2]) as 内容 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExistWaitExe", lng病人ID, lng主页ID)
    
    If Not rsTmp.EOF Then
        ExistWaitExe = Nvl(rsTmp!内容)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitDrug(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：检查病人在药房是否还有未发药的药品或卫材
'返回：药房和发料部门名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select Zl_Pati_Check_Execute(1,[1],[2]) as 内容 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExistWaitDrug", lng病人ID, lng主页ID)
    
    If Not rsTmp.EOF Then
        ExistWaitDrug = Nvl(rsTmp!内容)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitBool(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：检查病人在血库是否还有未发的血
'返回：血库部门名称
'编制人:刘鹏飞
'问题号：30339,2012-09-14
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select Zl_Pati_Check_Execute(3,[1],[2]) as 内容 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExistWaitBool", lng病人ID, lng主页ID)
    
    If Not rsTmp.EOF Then
        ExistWaitBool = Nvl(rsTmp!内容)
    End If
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Public Function ExistWaitTest(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
''功能：对于已经完成执行的检验和检查，判断检验及检查报告是否填写
''返回：检验和检查项目和执行部门
''编制人:刘鹏飞
''问题号：51613,2012-17-10
''问题号：69009,2013-01-03,取消该检查(已经存在了对未执行项目的检查)
'    Dim strSQL As String
'    Dim rsTmp As New ADODB.Recordset
'    Dim strDrug As String, strStuff As String, strTest As String
'    On Error GoTo errH
'
'    '判断检验及检查报告是否填写
'    strSQL = _
'        " Select Distinct C.类别, C.名称 As 项目, d.名称 As 部门" & vbNewLine & _
'        " From 病人医嘱记录 A,病人医嘱发送 B,病人医嘱报告 E,诊疗项目目录 C,部门表 D" & vbNewLine & _
'        " Where a.病人id = [1] And Nvl(a.主页id, 0) = [2]" & vbNewLine & _
'        "   And a.Id = b.医嘱id And A.ID = E.医嘱ID(+) And B.执行状态 = 1 And a.诊疗项目id = c.Id" & vbNewLine & _
'        "   And b.执行部门id + 0 = d.Id(+) And E.医嘱ID is null" & vbNewLine & _
'        "   And (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null)" & vbNewLine & _
'        "   And exists (select ID from 病人医嘱记录" & vbNewLine & _
'        "         where (诊疗类别 = 'C' And 相关id Is not Null And A.ID = 相关ID)" & vbNewLine & _
'        "            OR (诊疗类别 = 'D' And 相关id Is Null And A.ID = ID))" & vbNewLine & _
'        " order by C.类别, C.名称"
'
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "判断检验及检查报告是否填写", lng病人ID, lng主页ID)
'    strDrug = "": strStuff = ""
'    Do While Not rsTmp.EOF
'            If UCase(rsTmp!类别) = "E" Then '检验项目是以诊疗类别为E的标本医嘱为主记录
'                If strDrug = "" Then
'                    strDrug = Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未定部门]") & "未填写"
'                Else
'                    If InStr(1, vbCrLf & strDrug & vbCrLf, vbCrLf & Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未定部门]") & "未填写" & vbCrLf) = 0 Then
'                        If LenB(StrConv(strDrug & vbCrLf & Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未定部门]") & "未填写", vbFromUnicode)) <= 1000 Then
'                            strDrug = strDrug & vbCrLf & Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未定部门]") & "未填写"
'                        Else
'                            strDrug = strDrug & vbCrLf & "... ..."
'                        End If
'                    End If
'                End If
'            Else
'                If strStuff = "" Then
'                    strStuff = Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未定部门]") & "未填写"
'                Else
'                    If InStr(1, vbCrLf & strStuff & vbCrLf, vbCrLf & Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未定部门]") & "未填写" & vbCrLf) = 0 Then
'                        If LenB(StrConv(strStuff & vbCrLf & Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未定部门]") & "未填写", vbFromUnicode)) <= 1000 Then
'                            strStuff = strStuff & vbCrLf & Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未定部门]") & "未填写"
'                        Else
'                            strStuff = strStuff & vbCrLf & "... ..."
'                        End If
'                    End If
'                End If
'            End If
'
'    rsTmp.MoveNext
'    Loop
'    If strDrug <> "" Then
'        strDrug = "存在未填写的检验报告：" & vbCrLf & vbCrLf & strDrug
'    End If
'    If strStuff <> "" Then
'        strStuff = "存在未填写的检查报告：" & vbCrLf & vbCrLf & strStuff
'    End If
'    strTest = ""
'    If strDrug <> "" And strStuff <> "" Then
'      strTest = strDrug & vbCrLf & vbCrLf & strStuff
'    ElseIf strDrug <> "" Then
'      strTest = strDrug
'    ElseIf strStuff <> "" Then
'      strTest = strStuff
'    End If
'
'    ExistWaitTest = strTest
'
'    Exit Function
'errH:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call SaveErrLog
'End Function

Public Function ExistNurseData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal dOutTime As Date) As String
'功能:检查病人出院时间之后是否存在护理数据
    Dim strSql As String
    Dim strDrug As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    dOutTime = CDate(Format(dOutTime + 1 / 24 / 60, "YYYY-MM-DD HH:MM") & ":00")
    strSql = "Select ID From 病人护理文件 Where 病人ID=[1] and 主页ID=[2] And RowNum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "检查是否存在新版护理文件", lng病人ID, lng主页ID)
    If rsTemp.RecordCount > 0 Then
        '新版规则:病人本人则直接检查出院之后存在护理数据,婴儿只检查母婴同时出院的情况
        strSql = _
            " Select Distinct NVL(婴儿,0) 序号, 文件名称" & vbNewLine & _
            " From 病人护理文件 a, 病人护理数据 b" & vbNewLine & _
            " Where a.Id = b.文件id And b.发生时间 >= [3] And a.病人id = [1] And a.主页id = [2] And" & vbNewLine & _
            "      (Nvl(a.婴儿, 0) = 0 Or" & vbNewLine & _
            "      (Nvl(a.婴儿, 0) <> 0 And Not Exists" & vbNewLine & _
            "       (Select e.婴儿" & vbNewLine & _
            "         From 病人医嘱记录 e, 诊疗项目目录 f" & vbNewLine & _
            "         Where e.诊疗项目id + 0 = f.Id And e.医嘱状态 = 8 And Nvl(e.婴儿, 0) <> 0 And f.类别 = 'Z' And" & vbNewLine & _
            "               Instr([4], ',' || f.操作类型 || ',', 1) > 0 And e.病人id = a.病人id And e.主页id = a.主页id And e.婴儿 = a.婴儿)))"
            
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "护理数据检查", lng病人ID, lng主页ID, dOutTime, ",3,5,11,")
        If rsTemp.EOF Then ExistNurseData = "": Exit Function
        rsTemp.Sort = "序号"
        Do While Not rsTemp.EOF
            If strDrug = "" Then
                strDrug = Nvl(rsTemp!文件名称) & IIf(Val(rsTemp!序号) = 0, "", Space(6) & "婴儿序号:" & Val(rsTemp!序号))
            Else
                strDrug = strDrug & vbCrLf & Nvl(rsTemp!文件名称) & IIf(Val(rsTemp!序号) = 0, "", Space(6) & "婴儿序号:" & Val(rsTemp!序号))
            End If
        rsTemp.MoveNext
        Loop
        
        If strDrug <> "" Then
            strDrug = "存在护理数据的文件名称：" & vbCrLf & vbCrLf & strDrug
        End If
    Else
        '老版:
        strSql = "Select 1 From 病人护理记录 Where 病人id = [1] And 主页id = [2] And 病人来源 = 2 And 发生时间 >= [3] And RowNum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "护理数据检查", lng病人ID, lng主页ID, dOutTime)
        If rsTemp.RecordCount > 0 Then
            strDrug = "OK"
        End If
    End If
    ExistNurseData = strDrug
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ExistWaitQuittance(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：对于检查是否存在未审核销帐的单据
'返回：科室和单据
'编制人:刘鹏飞
'问题号：61429,2013-11-11
    
    Dim strSql As String
    Dim strDrug As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    strSql = _
        " Select Distinct a.No, d.名称 项目, c.名称 As 部门" & vbNewLine & _
        " From 住院费用记录 a, 病人费用销帐 b, 部门表 c, 收费项目目录 d" & vbNewLine & _
        " Where a.Id = b.费用id And a.收费细目id = d.Id And b.审核部门id = c.Id(+) And b.审核时间 Is Null And a.病人id = [1] And" & vbNewLine & _
        "      Nvl(a.主页id, 0) = [2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查是否存在未审核销帐的单据", lng病人ID, lng主页ID)
    strDrug = ""
    Do While Not rsTmp.EOF
        If strDrug = "" Then
            strDrug = "单据[" & Nvl(rsTmp!NO) & "]中的" & Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未知部门]") & "未审核"
        Else
            If InStr(1, vbCrLf & strDrug & vbCrLf, vbCrLf & "单据[" & Nvl(rsTmp!NO) & "]中的" & Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未知部门]") & "未审核" & vbCrLf) = 0 Then
                If LenB(StrConv(strDrug & vbCrLf & "单据[" & Nvl(rsTmp!NO) & "]中的" & Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未知部门]") & "未审核", vbFromUnicode)) <= 1000 Then
                    strDrug = strDrug & vbCrLf & "单据[" & Nvl(rsTmp!NO) & "]中的" & Nvl(rsTmp!项目) & "：在" & Nvl(rsTmp!部门, "[未知部门]") & "未审核"
                Else
                    strDrug = strDrug & vbCrLf & "... ..."
                End If
            End If
        End If
    rsTmp.MoveNext
    Loop
    
    If strDrug <> "" Then
        strDrug = "存在未审核销帐的单据：" & vbCrLf & vbCrLf & strDrug
    End If
    ExistWaitQuittance = strDrug
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExistFeeInsurePatient(lng病人ID As Long) As Boolean
'功能：判断医保病人是否存在未结费用
'返回：
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
        
    strSql = "Select Nvl(sum(B.费用余额),0) 费用余额 From 病人信息 A,病人余额 B Where A.病人ID=B.病人ID And Nvl(A.险类,0)<>0 And A.病人ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPatient", lng病人ID)
    
    If Not rsTmp.EOF Then ExistFeeInsurePatient = (rsTmp!费用余额 <> 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNurseGrade() As ADODB.Recordset
'功能：获取护理等级
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select ID,编码,名称 From 收费项目目录" & _
        " Where 类别='H' And 项目特性>=1 And (撤档时间 is NULL or Trunc(撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by 编码"
    Set GetNurseGrade = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'问题27370 by lesfeng 2010-01-26
Public Function InputDept(ByRef frmParent As Object, ByVal fra入院 As Control, ByVal obj As Control, ByVal str性质 As String, ByVal str服务对象 As String, _
ByVal strInput As String, ByRef blnCancel As Boolean, Optional ByVal intFlag = -1, Optional ByVal lngDeptID = 0, Optional ByVal bln仅操作员部门 As Boolean = False) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示输入指定性质的部门列表
    '入参:frmParent 窗口对象
    '     fra入院 控件 这里可能是 fra入院 主要计算弹出选择器的位置
    '     obj 控件 这里可能是 cbo入院科室 或者 cbo入院病区
    '     str性质='临床','护理','中药房',...,允许为空
    '     str服务对象:以,分离:如1,3
    '     strInput 支持输入编码、简码、名称进行匹配
    '     blnCancel 返回成功与否
    '     intFlag 判断输入方式，先选科室或者先选病区 初始-1不关联病区科室对应
    '     lngDeptId 当intFlag不为-1时，关联病区科室对应的科室或者病区的ID
    '     bln仅操作员部门-操作员的所属部门
    '出参:
    '返回:
    '编制:lesfeng
    '日期:2010-01-25 16:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim lngTxtHeight As Long, vPoint As POINTAPI
    Dim strNo As String, strInputN As String
    Dim strFrom As String, strWhere As String
    
    On Error GoTo errH
    
    vPoint = zlControl.GetCoordPos(fra入院.hWnd, obj.Left, obj.Top)
    lngTxtHeight = obj.Height
    
    strInputN = gstrLike & strInput & "%"
    strNo = strInput & "%"
                    
    str性质 = Replace(str性质, "'", "")
    If str性质 <> "" Then
        If InStr(1, str性质, ",") > 0 Then
            strSql = " And Instr(','||[1]||',',','||B.工作性质||',')>0"
        Else
            strSql = " And B.工作性质 = [1]"
        End If
    End If
    
    If zlCommFun.IsCharChinese(strInput) Or InStr(1, strInput, "-", 0) <> 0 Then
        strSql = strSql & " And (A.名称 Like [4] or A.编码||'-'||A.名称 Like [4])" '输入汉字时只匹配名称
    Else
        strSql = strSql & " And (A.编码 Like [5] Or A.名称 Like [4] Or A.简码 Like [4])"
    End If
    
    If intFlag = -1 Then
        strFrom = ""
        strWhere = ""
    ElseIf intFlag = 1 Then '科室输入 gbln先选病区
        strFrom = ",病区科室对应 D"
        strWhere = " And A.ID = D.科室ID And D.病区ID = [6]"
    Else '病区输入 先选科室
        strFrom = ",病区科室对应 D"
        strWhere = " And A.ID = D.病区ID And D.科室ID = [6]"
    End If

    If bln仅操作员部门 Then strSql = strSql & "  And A.id=C.部门ID and C.人员id =[3]"
    
    strSql = " Select 1 as 排序ID, A.ID,A.编码,A.名称,A.简码,B.工作性质,B.服务对象" & _
        " From 部门表 A,部门性质说明 B " & IIf(bln仅操作员部门, ",部门人员 C", "") & strFrom & _
        " Where B.部门ID=A.ID And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And  Instr(','||[2]|| ',',','|| B.服务对象|| ',')>0 " & strSql & strWhere & _
        " Order by A.编码  Desc"
        '" And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) Order by A.编码  Desc"

    Set InputDept = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "科室", 1, "", "请选择", False, False, True, vPoint.X, vPoint.Y, lngTxtHeight, blnCancel, False, True, str性质, str服务对象, UserInfo.ID, strInputN, strNo, lngDeptID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'问题27370 by lesfeng 2010-02-03
Public Function InputDoctors(ByRef frmParent As Object, ByVal fra入院 As Control, ByVal obj As Control, ByVal bytType As Byte, ByVal str服务对象 As String, _
ByVal strInput As String, ByRef blnCancel As Boolean, Optional ByVal strUnits As String = "") As ADODB.Recordset
'---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医生或护士列表.
    '入参:frmParent 窗口对象
    '     fra入院 控件 这里可能是 fra入院 主要计算弹出选择器的位置
    '     obj 控件 这里可能是 cbo门诊医师
    '     bytType=0-医生，1-护士
    '     str服务对象:以,分离:如1,2,3
    '     strInput 支持输入编号、简码、名称进行匹配
    '     blnCancel 返回成功与否
    '     strUnits=科室或病区ID串,如:18,26,31
    '出参:
    '返回:
    '编制:lesfeng
    '日期:2010-01-25 16:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim lngTxtHeight As Long, vPoint As POINTAPI
    Dim strNo As String, strInputN As String
    Dim strFrom As String, strWhere As String
    
    On Error GoTo errH
    
    vPoint = zlControl.GetCoordPos(fra入院.hWnd, obj.Left, obj.Top)
    lngTxtHeight = obj.Height
    
    strInputN = gstrLike & strInput & "%"
    strNo = strInput & "%"
    
    On Error GoTo errH
    If strUnits <> "" Then
        If InStr(1, strUnits, ",") > 0 Then
            strSql = " And Instr(','|| [3] || ',',',' || B.部门ID || ',')>0"
        Else
            strSql = " And B.部门ID=[3]"
        End If
    End If
    
    If zlCommFun.IsCharChinese(strInput) Or InStr(1, strInput, "-", 0) <> 0 Then
        strSql = strSql & " And (A.姓名 Like [4] or A.简码||'-'||A.姓名 Like [4])" '输入汉字时只匹配名称
    Else
        strSql = strSql & " And (A.编号 Like [5] Or A.姓名 Like [4] Or A.简码 Like [4])"
    End If
    
    strSql = "Select Distinct A.ID,A.编号,A.简码,A.姓名,C.人员性质" & _
             "  From 人员表 A,部门人员 B,人员性质说明 C,部门性质说明 D" & _
             " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质 = [1] And B.部门ID=D.部门ID" & _
             "   And  Instr(','||[2]|| ',',','|| D.服务对象|| ',')>0 " & strSql & _
             "   And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & _
             "   And (A.站点=[6] Or A.站点 is Null)" & _
             " Order by 简码"
    Set InputDoctors = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "医生选择", 1, "", "请选择", False, False, True, vPoint.X, vPoint.Y, lngTxtHeight, blnCancel, False, True, IIf(bytType = 0, "医生", "护士"), str服务对象, strUnits, strInputN, strNo, gstrNodeNo)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDeptDoctors(ByVal lng科室ID As Long) As String
'功能：获取指定科室下面所包含的医生/护士IDs
'返回：医生ID1,医生ID2,...
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long
    
    On Error GoTo errH
    'by lesfeng 2010-01-11 性能优化
    strSql = "Select Distinct A.ID From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质 IN('医生','护士') " & _
                " And B.部门ID=[1] And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & _
                " And (A.站点=[2] Or A.站点 is Null)" & _
                " Order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng科室ID, gstrNodeNo)
    
    strSql = ""
    For i = 1 To rsTmp.RecordCount
        strSql = strSql & "," & rsTmp!ID
        rsTmp.MoveNext
    Next
    GetDeptDoctors = Mid(strSql, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetArea(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'功能：获取地区列表或选择的地区
'参数：
    Dim strSql As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo errH
    vRect = zlControl.GetControlRect(txtInput.hWnd)
    If Not blnShowAll Then
        strSql = " Select 编码 as ID,编码,名称,简码 From 区域" & _
                 " Where (编码 Like [1] Or upper(简码) Like '" & gstrLike & "'||[1]||'%' Or 名称 Like '" & gstrLike & "'||[1]||'%') And  NVL(级数,0)<3 "
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "区域", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSql = "Select 编码 as ID,编码,名称,简码 From 区域 Where NVL(级数,0)<3 "
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "区域", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAddress(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'功能：获取地区列表或选择的地区
'参数：
    Dim strSql As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo errH
    If Not blnShowAll Then
        strSql = " Select 编码 as ID,编码,名称,简码 From 地区" & _
                 " Where 编码 Like [1] Or 简码 Like [1] Or 名称 Like [1]"
        vRect = zlControl.GetControlRect(txtInput.hWnd)
        Set GetAddress = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "地址", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSql = " Select Distinct Substr(名称,1,2) as ID,NULL as 上级ID,0 as 末级,NULL as 编码," & _
                " Substr(名称,1,2) as 名称 From 地区" & _
                " Union All" & _
                " Select 编码 as ID,Substr(名称,1,2) as 上级ID,1 as 末级,编码,名称 " & _
                " From 地区 Order by 编码"
        Set GetAddress = zlDatabase.ShowSQLSelect(frmParent, strSql, 2, "地址", True, txtInput.Text, "", True, True, False, 0, 0, 0, blnCancel, True, True)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get医疗机构(ByVal objText As Object, ByVal objFrom As Form, ByVal bytStyle As Byte, ByVal strCaption As String, ByVal strMsg As String, ByRef vPoint As POINTAPI, ByVal blnCancel As Boolean)
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码,末级 From 医疗机构 Order By 编码"
    Set rsTmp = zlDatabase.ShowSelect(objFrom, strSql, bytStyle, strCaption, , , , , True, True, vPoint.X, vPoint.Y, objText.Height, blnCancel)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有设置""" & strCaption & """数据，请先到" & strMsg & "中设置。", vbInformation, gstrSysName
        End If
        objText.Tag = ""
        zlControl.ControlSetFocus objText
    Else
        objText.Text = rsTmp!名称
        zlControl.ControlSetFocus objText
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetSpc医疗机构(ByVal objText As Object, ByVal objFrom As Form, ByVal strCaption As String, ByVal strSeek As String, ByVal strNote As String, ByVal blnCancel As Boolean, ByVal bln末级 As Boolean, ByRef vPoint As POINTAPI)
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    If zlCommFun.IsCharChinese(objText.Text) Then
        strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 医疗机构 Where 名称 Like [1]"
    Else
        If gbytCode = 1 Then
            strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 医疗机构 Where zlWbCode(名称) Like [1]"
        ElseIf gbytCode = 0 Then
            strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 医疗机构 Where 简码 Like [1]"
        End If
    End If
    Set rsTmp = zlDatabase.ShowSQLSelect(objFrom, strSql, 0, strCaption, bln末级, strSeek, strNote, False, _
        False, True, vPoint.X, vPoint.Y, objFrom.Height, blnCancel, False, False, _
        gstrLike & UCase(objText.Text) & "%")
    If Not rsTmp Is Nothing Then
        objText.Text = rsTmp!名称
    Else
        objText.Tag = ""
        If gbln医疗机构不允许自由录入 Then
            MsgBox "在字典表中未找到该数据,请重新录入！", vbInformation, gstrSysName
            objText.Text = ""
            objText.SetFocus
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetOrgAddress(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'功能：获取合约单位列表
'参数：
    Dim strSql As String, blnCancel As Boolean
    Dim vRect As RECT
    '问题27040 by lesfeng 对合约单位加上撤档时间的处理
    On Error GoTo errH
    If Not blnShowAll Then
        strSql = " Select ID,编码,名称,简码,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
                    " Where 末级=1 And (编码 Like [1] Or 简码 Like [1] Or 名称 Like [1]) " & _
                    " and (撤档时间 IS NULL OR TO_CHAR(撤档时间, 'yyyy-MM-dd') = '3000-01-01') "
        vRect = zlControl.GetControlRect(txtInput.hWnd)
        Set GetOrgAddress = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "单位", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSql = " Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From  合约单位" & _
                 "  Where (撤档时间 IS NULL OR TO_CHAR(撤档时间, 'yyyy-MM-dd') = '3000-01-01') " & _
                 " Start With 上级ID is NULL Connect by Prior ID=上级ID"
        Set GetOrgAddress = zlDatabase.ShowSQLSelect(frmParent, strSql, 2, "单位", True, txtInput.Text, "", True, True, False, 0, 0, 0, blnCancel, True, True)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetPatiBeds(ByVal lng病人ID As Long, Optional ByVal strBed As String) As ADODB.Recordset
'功能：获取病人所占用的床位表信息
'参数：lng病人ID=病人ID
'返回：病人占用的床位记录集
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select A.床号, A.房间号, A.性别分类, A.等级id As 床位等级id, B.名称 As 床位等级, A.科室ID, A.共用, A.状态, C.性别" & vbNewLine & _
        "       From 床位状况记录 A, 收费项目目录 B, 病人信息 C" & vbNewLine & _
        "       Where A.等级id = B.ID And A.病人ID = C.病人ID(+) And A.病人id = [1]" & IIf(strBed = "", "", " And 床号 = [2]")
    '注意:家庭病床病人没有床位
    Set GetPatiBeds = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, strBed)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFreeBeds(ByVal lng病区ID As Long, ByVal lng科室ID As Long, str性别 As String, Optional lng病人ID As Long = 0) As ADODB.Recordset
'功能：获取指定病区和科室的空床
'参数：
    Dim strSql As String, strTmp As String
    
    If InStr(str性别, "男") > 0 Then
        strTmp = "男床,不限床"
    ElseIf InStr(str性别, "女") > 0 Then
        strTmp = "女床,不限床"
    Else
        strTmp = "不限床"
    End If
    
    On Error GoTo errH
    
'    Select 床号, a.房间号, 性别分类, 床位编制, 等级id, 共用, b.性别
'From (Select 床号, 房间号, 性别分类, 床位编制, 等级id, 共用
'       From (Select 床号, 房间号, 性别分类, 床位编制, 等级id, 共用
'              From 床位状况记录
'              Where 病区id = 57 And 病人id Is Null And 状态 = '空床' And Instr('男床,不限床', 性别分类) > 0 And (科室id = 57 Or 科室id Is Null)
'              Union
'              Select 床号, 房间号, 性别分类, 床位编制, 等级id, 共用
'              From 床位状况记录
'              Where 病人id = 701100 And 共用 = 1 And 病区id = 57)
'       Order By LPad(床号, 10, ' ')) A,
'     (Select m.房间号, Wmsys.Wm_Concat(Distinct n.性别) As 性别
'       From 床位状况记录 M, 病人信息 N
'       Where m.病人id = n.病人id(+) And m.病区id = 57 And 房间号 Is Not Null
'       Group By m.房间号) B
'Where a.房间号 = b.房间号(+)

    strSql = "Select 床号, a.房间号, 性别分类, 床位编制, 等级id, 床位等级, 共用, b.性别 From( " & _
        "Select 床号, 房间号, 性别分类, 床位编制, 等级id, 床位等级, 共用" & vbNewLine & _
        "From (" & vbNewLine & _
                "Select 床号, 房间号, 性别分类, 床位编制, 等级id, J.名称 AS 床位等级, 共用" & vbNewLine & _
                    "From 床位状况记录 I, 收费项目目录 J " & vbNewLine & _
                    "Where I.等级ID = J.ID And 病区id = [1] And 病人id Is Null And 状态 = '空床' And Instr([3],性别分类)>0" & _
                    IIf(lng科室ID = 0, "", " And (科室ID = [2] Or 科室ID is Null)")
    If lng病人ID <> 0 Then '-----------------------------------------------------------------病人可以入住共用病区共用病床原住床位
        strSql = strSql & " Union " & _
                        "Select 床号, 房间号, 性别分类, 床位编制, 等级id, Q.名称 AS 床位等级, 共用" & vbNewLine & _
                        "       From 床位状况记录 P, 收费项目目录 Q " & vbNewLine & _
                        "       Where P.等级ID = Q.ID And 病人id = [4] And 共用 = 1 And 病区id = [1]"
    End If
    strSql = strSql & ") ORDER BY LPAD(床号,10,' ')) A ," & _
            "(Select m.房间号, f_List2str(Cast(COLLECT(Distinct n.性别) as t_Strlist)) As 性别" & vbNewLine & _
            "       From 床位状况记录 M, 病人信息 N" & vbNewLine & _
            "       Where m.病人id = n.病人id(+) And m.病区id = [1] And 房间号 Is Not Null" & vbNewLine & _
            "       Group By m.房间号) B Where a.房间号 = b.房间号(+)"

    Set GetFreeBeds = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病区ID, lng科室ID, strTmp, lng病人ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDiagnosticOtherInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
                                  ByVal str诊断类型 As String, ByVal str记录来源 As String) As ADODB.Recordset
'功能：获取指定病人的其它诊断记录'
'参数:
'诊断类型-1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;
'        11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断
'记录来源:1-病历；2-入院登记；3-首页整理(门诊医生站,诊断摘要);

    On Local Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    If InStr(1, "," & str记录来源 & ",", ",3,") > 0 And _
        (InStr(1, "," & str诊断类型 & ",", ",1,") > 0 Or InStr(1, "," & str诊断类型 & ",", ",11,") > 0) Then
        strSql = " Select A.诊断类型,A.记录来源,A.诊断描述,A.疾病ID,A.诊断ID,A.出院情况,A.记录日期,A.是否疑诊,A.诊断次序,C.编码 From 病人诊断记录 A,疾病编码目录 C " & _
                 " Where A.病人ID=[1] And Nvl(A.主页ID,0)=[2] And A.疾病ID = C.ID(+)" & _
                 " And A.诊断次序>1  And NVL(A.编码序号,1) = 1 And instr([3],','||A.诊断类型||',')>0 And A.记录来源 in (" & str记录来源 & ")" & _
                 " Union " & _
                 " Select a.诊断类型,a.记录来源,a.诊断描述,a.疾病ID,a.诊断ID,a.出院情况,a.记录日期,a.是否疑诊,a.诊断次序,C.编码 From 病人诊断记录 a,病人挂号记录 b,疾病编码目录 C  " & _
                 " Where a.病人ID=[1] And A.疾病ID = C.ID(+) " & _
                 " And a.病人ID=b.病人ID And b.登记时间>trunc(sysdate-3) And a.主页ID=b.id" & _
                 " And a.诊断次序>1 And NVL(A.编码序号,1) = 1 And instr([3],','||a.诊断类型||',')>0 And a.记录来源=3 And B.记录性质=1 and B.记录状态=1" & _
                 " Order by 诊断次序 Asc"
    Else
        
        strSql = " Select A.诊断类型,A.记录来源,A.诊断描述,A.疾病ID,A.诊断ID,A.出院情况,A.记录日期,A.是否疑诊,A.诊断次序,C.编码 From 病人诊断记录 A,疾病编码目录 C " & _
                 " Where A.病人ID=[1] And Nvl(A.主页ID,0)=[2] And A.疾病ID = C.ID(+)" & _
                 " And A.诊断次序>1 And NVL(A.编码序号,1) = 1 And instr([3],','||A.诊断类型||',')>0 And A.记录来源 in (" & str记录来源 & ")" & _
                 " Order by 诊断次序 Asc"
                 
        '诊断次序-出院时,病案主页整理中可能填写主要诊断,次要诊断等多条记录
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng病人ID, lng主页ID, "," & str诊断类型 & ",")
    If Not rsTmp.EOF Then Set GetDiagnosticOtherInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SetVsFlexGridChangeHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid, lngNo As Long)
    '功能：初始vsFlexGrid
    '           有一固定行，初始化后，只有一行记录，无固定列。
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'vsGrid:    要初始化的控件

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1
        If lngNo = 0 Then
            .FixedCols = 0
            .Cols = .FixedCols + UBound(arrHead) + 1
            .Rows = .FixedRows + 1
        Else
            .FixedCols = 1
            .Cols = .FixedCols + UBound(arrHead)
            .Rows = .FixedRows + 1
        End If

        For i = 0 To UBound(arrHead)
            If .FixedCols > 0 Then
                .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            Else
                .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            End If
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
               '为了支持zl9PrintMode
                If .FixedCols > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .colAlignment(i) = Val(Split(arrHead(i), ",")(2))
                    .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(.FixedCols + i) = False
                    .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                    .colAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
'                    .ColData
                    '为了支持zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                End If
            Else
                If .FixedCols > 0 Then
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  '为了支持zl9PrintMode
                Else
                    .ColHidden(.FixedCols + i) = True
                    .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
                End If
            End If
            .ColData(i) = Val(Split(arrHead(i), ",")(3)) '将标提作为列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        Next
        
        '固定行文字居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
    End With
End Sub
 
Public Function zlVsInsertIntoRow(ByVal vsGrid As VSFlexGrid, ByVal lngRow As Long, Optional blnBefor As Boolean = False, _
    Optional blnMoveNewRow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '功能:插入行
    '参数:vsGrid-插入行的网格格件
    '     lngRow-当前行
    '     blnBefor-在lngrow之间或之后.true:之间,false-之后
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        If blnBefor Then
            .AddItem "", lngRow
        Else
            .AddItem "", lngRow + 1
        End If
        If blnMoveNewRow = True Then
            If blnBefor Then '
                .Row = lngRow
            Else
                .Row = lngRow + 1
            End If
        End If
    End With
    zlVsInsertIntoRow = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlPvVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng主例 As Long = -1, Optional lng尾列 As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef lngRow As Long = -1, Optional strValue As String)
    ', Optional strHeadMove As String
    '-----------------------------------------------------------------------------------------------------------

    '功能:移动单元格的列
    '入参:blnEdit-当前正处于编辑状态,允许新增行
    '     lng主例-主列,如果<0,则主列为0列,否则为指定的列
    '     lng尾列-尾列,如果<0,则主列为.cols-1,否则为指定的列
    '出参:lngRow-如果存在插入行,则返回被插入的行号,否则返回-1
    '返回:
    '编制:刘兴洪
    '日期:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------

    Dim lngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long
    Dim lngValue As Long
    Dim arrHead As Variant
    Dim j As Long
    Dim lngColValue As Long
    
    Err = 0: On Error GoTo ErrHand:
    'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)

    If lng主例 <> -1 Then
        lngCol = lng主例
    Else
        lngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If lngCol = -1 Then lngCol = 0
    lngLastCol = IIf(lng尾列 < 0, vsGrid.Cols - 1, lng尾列)
    lngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = lngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        lngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                If IsNull(strValue) Or strValue = "" Then
                    arrSplit = Split(.ColData(i) & "||", "||")
                    If IsNull(arrSplit(1)) Or Trim(arrSplit(1)) = "" Then
                        lngValue = 0
                    Else
                        lngValue = Val(arrSplit(1))
                    End If
                Else
                    arrHead = Split(strValue, ";")
                    For j = 0 To UBound(arrHead)
                        lngValue = 1
                        lngColValue = Val(Split(arrHead(j), "||")(0))
                        If i = lngColValue Then
                            lngValue = Val(Split(arrHead(j), "||")(1))
                            Exit For
                        End If
                    Next
                End If
                If .ColHidden(i) Or lngValue >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = lngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    lngRow = .Row
                                End If
                            End If
                            .Col = lngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
ErrHand:
End Sub

Public Function zl_VsGrid_SaveToPara(ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, _
ByVal lngMoudel As Long, ByVal strParaName As String, Optional ByVal bln私有 As Boolean = True, _
    Optional ByVal bln强制恢复保存 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:保存vsFlex的宽度到参数表
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     lngMoudel-模块号
    '返回:保存成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------

    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    If bln强制恢复保存 = False Then
        If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
    End If

    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    zlDatabase.SetPara strParaName, strCol, glngSys, lngMoudel ', bln私有
    zl_VsGrid_SaveToPara = True
End Function

Public Function zl_VsGrid_FromParaRestore(ByVal vsGrid As VSFlexGrid, ByVal strCaption, ByVal lngMoudle As Long, _
    ByVal strParaName As String, Optional bln私有 As Boolean = True, _
    Optional ByVal bln强制恢复保存 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:从参数表中恢复网格的宽度
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     lngMoudle-模块号
    '返回:恢复成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------

    Dim strParaValue As String, intCols As Integer, arrReg As Variant, arrtemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String

    If bln强制恢复保存 = False Then
        If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
    End If

    strParaValue = zlDatabase.GetPara(strParaName, glngSys, lngMoudle, "")
    If strParaValue = "" Then Exit Function
    
    'strParaValue:保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...

    Err = 0: On Error GoTo ErrHand:

    arrReg = Split(strParaValue, "|")
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrtemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrtemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(arrtemp(1))
                If Val(arrtemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_VsGrid_FromParaRestore = True
    Exit Function
ErrHand:
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
    On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
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
    On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub
Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:保存DockPane控件的具体位置
    '入参:frmMain-窗体名
    '     objPance:DockinPane控件
    '      StrKey-键名
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    Err = 0: On Error GoTo ErrHand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlSaveDockPanceToReg = True
ErrHand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:保存DockPane控件的具体位置
    '入参:frmMain-窗体名
    '     objPance:DockinPane控件
    '      StrKey-键名
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(zlDatabase.GetPara("界面区域隐藏", , , True)) = 1
    Err = 0: On Error GoTo ErrHand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlRestoreDockPanceToReg = True
ErrHand:
End Function
Public Function GetBalanceDeposit(ByVal lngBalanceID As Long, ByVal blnNOMoved As Boolean) As ADODB.Recordset
    '功能：获取一张结帐单据的冲预交记录
    Dim strSql As String
    On Error GoTo errH
    strSql = _
        "Select A.ID,A.NO 单据号,A.实际票号 票据号,To_Char(A.收款时间,'YYYY-MM-DD') as 日期,A.结算方式," & _
        " Ltrim(To_Char(A.冲预交,'9999999990.00')) as 金额" & _
        " From " & IIf(blnNOMoved, "H", "") & "病人预交记录 A " & _
        " Where mod(A.记录性质,10)=1 And A.结帐ID = [1]  " & _
        " Order by A.日期,A.结算方式"
    Set GetBalanceDeposit = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lngBalanceID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlIsExistsSquareCard(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查该单据是否为卡结算单据
    '入参:strNos-单据号(可以为多张,用逗号分离)
    '出参:
    '返回:存在,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, strNoIns As String
    strNoIns = Replace(strNos, "'", "")
    On Error GoTo errHandle
    strSql = "Select A.ID As 卡结算id " & _
    "   From 病人卡结算记录 A, 病人预交记录 B, " & _
    "        (Select Column_Value From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist))) J " & _
    "   Where A.结算id = B.ID and ( B.记录性质=2 or B.记录性质=12) And B.NO = J.Column_Value And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "检查结帐单是否存在刷卡记录", strNoIns)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSql As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSql, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '     blnNoBeginTrans:没有事务开始
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSql As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSql = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSql, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub
Public Function StrToNum(ByVal strNumber As String) As Double
    '功能:将字符串转换成数据
    Dim strTmp As String
    strTmp = Replace(strNumber, ",", "")
    StrToNum = Val(strTmp)
End Function
Public Function zl_Get医疗卡类型(lngTypeId As Long) As String()
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据医疗类型ID获取医疗类型
    '入参:lngTypeID-医疗卡类型ID
    '返回:类型对象
    '编制:王吉
    '日期:2012-07-06
    '问题号:51072
    '-----------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As Recordset
    Dim arr(3) As String
    
    strSql = "" & _
    "       Select 密码长度,密码输入限制,是否缺省密码 " & _
    "       From 医疗卡类别 " & _
    "       Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取医疗卡类别", lngTypeId)
    If rsTemp Is Nothing Then zl_Get医疗卡类型 = arr: Exit Function
    If rsTemp.RecordCount <= 0 Then zl_Get医疗卡类型 = arr: Exit Function
    rsTemp.MoveFirst
    arr(0) = Nvl(rsTemp!密码长度, "0")
    arr(1) = Nvl(rsTemp!密码输入限制, "0")
    arr(2) = Nvl(rsTemp!是否缺省密码, "0")
    zl_Get医疗卡类型 = arr
End Function

Public Function 是否已经签约(strCardNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查需要绑定的卡号是否已经签约
    '入参:绑定卡号
    '编制:王吉
    '日期:2012-08-31 11:32:14
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand:
    strSql = "" & _
    "   Select Count(1) as 是否签约 From 病人医疗卡信息 Where 卡号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "医疗卡绑定", strCardNO)
    是否已经签约 = rsTemp!是否签约 > 0
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function


Public Sub AddSQL绑定卡(ByVal lng病人ID As Long, 卡类别ID As Long, strCard As String, strPassWord As String, ByVal dtCurdate As Date, blnICCard As Boolean, ByRef cllPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绑定卡处理
    '入参:lng病人ID;strCard-绑定卡号;strPassWord-加密密码
    '出参:lngCard结帐ID-卡费的结帐ID
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim str变动原因 As String
    Dim strICCard As String
    
    strICCard = IIf(blnICCard, strCard, "")
    str变动原因 = "病人挂号发卡"
          'Zl_医疗卡变动_Insert
          strSql = "Zl_医疗卡变动_Insert("
          '      变动类型_In   Number,
          '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
          strSql = strSql & "" & 11 & ","
          '      病人id_In     住院费用记录.病人id%Type,
          strSql = strSql & "" & lng病人ID & ","
          '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
          strSql = strSql & "" & 卡类别ID & ","
          '      原卡号_In     病人医疗卡信息.卡号%Type,
          strSql = strSql & "'',"
          '      医疗卡号_In   病人医疗卡信息.卡号%Type,
          strSql = strSql & "'" & strCard & "',"
          '      变动原因_In   病人医疗卡变动.变动原因%Type,
          '      --变动原因_In:如果密码调整，变动原因为密码.加密的
          strSql = strSql & "'" & str变动原因 & "',"
          '      密码_In       病人信息.卡验证码%Type,
          strSql = strSql & "'" & strPassWord & "',"
          '      操作员姓名_In 住院费用记录.操作员姓名%Type,
          strSql = strSql & "'" & UserInfo.姓名 & "',"
          '      变动时间_In   住院费用记录.登记时间%Type,
          strSql = strSql & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
          strSql = strSql & "'" & strICCard & "',"
          '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
          strSql = strSql & "NULL)"
     zlAddArray cllPro, strSql
End Sub

Public Function Get医疗卡类别ID(strTypeName As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别ID
    '入参:strTypeName 医疗卡类别名称
    '返回:医疗卡类别ID
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand
    strSql = "" & _
    "   Select ID From 医疗卡类别 Where 名称=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "医疗卡类别", strTypeName)
    If rsTemp Is Nothing Then Get医疗卡类别ID = 0: Exit Function
    If rsTemp.RecordCount <= 0 Then Get医疗卡类别ID = 0: Exit Function
    Get医疗卡类别ID = rsTemp!ID
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zl当前用户身份证是否绑定(lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断当前用户身份证是否已被绑定
    '入参:lng病人ID
    '返回:True 已绑定 false 未绑定
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand
    strSql = "" & _
    " Select count(1) as 是否绑定 From 病人信息 A,病人医疗卡信息 B Where A.身份证号 =B.卡号 And A.病人ID=B.病人ID And A.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "医疗卡绑定", lng病人ID)
    zl当前用户身份证是否绑定 = rsTemp!是否绑定 > 0
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function SetNullValue(varObj As Variant, Optional strDefault As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置空值默认值
    '入参:varObj集合字段对象,strDefault 默认值
    '返回:设置后的值
    '编制:王吉
    '日期:2012-08-31 04:36:33
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If IsNull(varObj) Then
        SetNullValue = strDefault
        Exit Function
    End If
    SetNullValue = CStr(varObj)
End Function


Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         'Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub

Public Function CheckBillExistReplenishData(intType As Integer, _
    Optional lngBalance As Long, Optional strNos As String, _
    Optional ByRef strReplenishNo As String, Optional ByRef blnErrBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否存在二次结算
    '入参:intType:0-收费数据，使用lngBalance为结算序号
    '     intType:1-收费数据，使用strNos为单据号
    '出参：
    '   strReplenishNo 补充结算单据号
    '   blnErrBill 是否异常结算单据
    '返回:True-存在二次结算数据 False-不存在二次结算数据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strReplenishNo = ""
    If intType = 0 Then
        strSql = _
            " Select Max(a.NO) As No,Max(a.费用状态) As 费用状态" & vbNewLine & _
            " From 费用补充记录 A, (Select Distinct 结帐id From 病人预交记录 Where 结算序号 = [1]) B" & vbNewLine & _
            " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 0 And Nvl(a.费用状态,0) <> 2"
        strSql = strSql & _
            " Union All" & _
            " Select Max(a.NO) As No,Max(a.费用状态) As 费用状态 From 费用补充记录 A Where a.结算序号 = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查二次结算", lngBalance)
    Else
        strSql = _
            " Select Max(a.NO) As No,Max(a.费用状态) As 费用状态" & vbNewLine & _
            " From 费用补充记录 A," & vbNewLine & _
            "      (Select Distinct 结帐id" & vbNewLine & _
            "       From 门诊费用记录" & vbNewLine & _
            "       Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(f_Str2list([1])))) B" & vbNewLine & _
            " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 0 And Nvl(a.费用状态,0) <> 2 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查二次结算", strNos)
    End If
    
    strReplenishNo = Nvl(rsTmp!NO)
    blnErrBill = Val(Nvl(rsTmp!费用状态)) = 1
    CheckBillExistReplenishData = strReplenishNo <> ""
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitComponents (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Sub
    End If
End Sub

Public Function CheckAge(ByVal strAge As String, Optional ByVal strBirthDay As String = "", Optional ByVal datCalc As Date) As String
    '功能:年龄合法性检查
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    strBirthDay = Format(strBirthDay, "YYYY-MM-DD HH:mm")
    If IsDate(strBirthDay) Then
        If datCalc = CDate(0) Then
            strSql = "select Zl_Age_Check([1],[2]) From dual"
        Else
            strSql = "select Zl_Age_Check([1],[2],[3]) From dual"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Zl_Age_Check", strAge, CDate(strBirthDay), datCalc)
    Else
        strSql = "select Zl_Age_Check([1]) From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Zl_Age_Check", strAge)
    End If
    CheckAge = Nvl(rsTemp.Fields(0).Value)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CreatePublicPatient() As Boolean
'功能:创建病人信息公共部件对象
    If gobjPublicPatient Is Nothing Then
        On Error Resume Next
        Set gobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        If gobjPublicPatient Is Nothing Then
            MsgBox "创建病人信息公共部件(zlPublicPatient.clsPublicPatient)失败!", vbInformation, gstrSysName
        Else
            Call gobjPublicPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser)
        End If
        Err.Clear: On Error GoTo 0
    End If
    If Not gobjPublicPatient Is Nothing Then CreatePublicPatient = True
End Function

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'功能：外挂创建与检查
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject(, "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
    
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String, Optional ByRef strErr As String = "0")
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    Dim strMsg As String
    
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        strMsg = "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description
        If strErr = "0" Then
            MsgBox strMsg, vbInformation, gstrSysName
        Else
            strErr = strMsg
        End If
    End If
End Sub

Public Function Get病案主页从表(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str信息串 As String) As ADODB.Recordset
'功能：
'    获取病案主页从表项
'参数:
    Dim strSql As String
    Dim intRet As Integer
    
    intRet = UBound(Split(str信息串, ","))
    If intRet = -1 Then '读取病人所有从表信息
        strSql = "Select 信息名,信息值 From 病案主页从表 Where 病人ID =[1] And 主页ID =[2] And 信息值 is Not Null"
    ElseIf intRet = 0 Then '读取指定某个从表信息
        strSql = "Select 信息名,信息值 From 病案主页从表 Where 病人ID =[1] And 主页ID =[2] and 信息名='" & Split(str信息串, ",")(0) & "'" & " And 信息值 is Not Null "
    ElseIf intRet > 0 Then '读取指定的多个从表信息值
        strSql = "Select 信息名, 信息值" & vbNewLine & _
            "From 病案主页从表" & vbNewLine & _
            "Where 病人id = [1] And 主页id = [2] And" & vbNewLine & _
            "      信息名 In (Select * From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) And 信息值 is Not Null "
    End If
    
    On Error GoTo errH
    Set Get病案主页从表 = zlDatabase.OpenSQLRecord(strSql, "读取病案主页从表", lng病人ID, lng主页ID, str信息串)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub LoadStructAddressDef(ByRef strAddress() As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取区域表中的缺省地址
    '入参:PatiAddress-结构化地址控件
    '返回:
    '编制:余伟节
    '日期:2016/1/7
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    On Error GoTo errH
    strSql = "Select 级数,名称,level From 区域 " & _
            " Start With 缺省标志=1 " & _
            " Connect by Prior 上级编码=编码 " & _
            " Order by level Desc "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "缺省区域")
    If rsTmp.RecordCount = 0 Then Exit Sub
    Do While Not rsTmp.EOF
        strAddress(Val(Nvl(rsTmp!级数))) = Nvl(rsTmp!名称)
        rsTmp.MoveNext
    Loop
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ReadStructAddress(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef PatiAddress As Object)
'功能:读取结构化地址
    Dim i As Long
    Dim rsStruct As ADODB.Recordset
    Dim rsAddress As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select a.省, a.市, a.县, a.乡镇, a.其他, a.地址类别 From 病人地址信息 A Where a.病人id = [1] And NVL(a.主页id,0) = [2]"
    Set rsStruct = zlDatabase.OpenSQLRecord(strSql, "读取病人结构化地址", lng病人ID, lng主页ID)
    
    For i = PatiAddress.LBound To PatiAddress.UBound
        rsStruct.Filter = "地址类别=" & i
        If rsStruct.RecordCount > 0 Then
            Call PatiAddress(i).LoadStructAdress(rsStruct!省 & "", rsStruct!市 & "", rsStruct!县 & "", rsStruct!乡镇 & "", rsStruct!其他 & "")
        Else
            If rsAddress Is Nothing Then
                '同一个病人只读取一次
                If lng主页ID <> 0 Then
                    strSql = "Select c.出生地点, c.籍贯, Nvl(b.家庭地址, c.家庭地址) As 现住址, Nvl(b.户口地址, c.户口地址) As 户口地址, Nvl(b.联系人地址, c.联系人地址) As 联系人地址" & vbNewLine & _
                        "From 病案主页 B, 病人信息 C" & vbNewLine & _
                        "Where b.病人id = c.病人id And b.病人id = [1] And b.主页id = [2] "
                Else
                    strSql = "Select c.出生地点, c.籍贯, c.家庭地址 As 现住址,  c.户口地址 As 户口地址,c.联系人地址 As 联系人地址 " & vbNewLine & _
                        "From 病人信息 C" & vbNewLine & _
                        "Where c.病人id = [1] "
                End If
                Set rsAddress = zlDatabase.OpenSQLRecord(strSql, "读取病人结构化地址", lng病人ID, lng主页ID)
            End If
            If rsAddress.RecordCount > 0 Then
                If Nvl(rsAddress.Fields(PatiAddress(i).Tag).Value, "") <> "" Then
                    PatiAddress(i).Value = Nvl(rsAddress.Fields(PatiAddress(i).Tag).Value, "")    '兼容启用结构化地址之前的数据
                End If
            End If
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub CreateStructAddressSQL(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef arrSQL As Variant, ByRef PatiAddress As Object, Optional ByVal bytFunc As Byte = 0)
'功能:创建结构化地址SQL
'参数:
'PatiAddress-结构化地址控件数组名
'arrSQL-返回的SQL数组集合
'bytFunc 可选参数:=1当控件值为空时,代表删除
    Dim i As Long
    
    For i = PatiAddress.LBound To PatiAddress.UBound
        If PatiAddress(i).Value <> "" Then
            '新增\修改
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(1," & lng病人ID & "," & lng主页ID & "," & i & ",'" & PatiAddress(i).value省 & "','" & PatiAddress(i).value市 & "','" & PatiAddress(i).value区县 & "','" & PatiAddress(i).value乡镇 & "','" & PatiAddress(i).value详细地址 & "','" & PatiAddress(i).Code & "')"
        Else
            '删除
            If bytFunc = 1 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_病人地址信息_update(2," & lng病人ID & "," & lng主页ID & "," & i & ")"
            End If
        End If
    Next

End Sub

Public Function CreateXWHIS(Optional ByVal blnMsg As Boolean) As Boolean
'功能：判断 RIS接口部件(zl9XWInterface.clsHISInner) 是否存在，并启用
'参数：blnMsg－创建失败时是否提示

    If Not gblnXW Then Exit Function
    If Not gobjXWHIS Is Nothing Then CreateXWHIS = True: Exit Function
    
    On Error Resume Next
    Set gobjXWHIS = GetObject(, "zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    On Error Resume Next
    If gobjXWHIS Is Nothing Then Set gobjXWHIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    If gobjXWHIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS接口部件(zl9XWInterface)未创建成功！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    CreateXWHIS = True
End Function

Public Function CreatePublicExpenseBillOperation() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共费用部件
    '入参:
    '编制:
    '日期:
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjPublicExpenseBillOperation Is Nothing Then
        Set gobjPublicExpenseBillOperation = CreateObject("zlPublicExpense.clsBillOperation")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
            Exit Function
        End If
    Else
        CreatePublicExpenseBillOperation = True
        Exit Function
    End If
    If gobjPublicExpenseBillOperation Is Nothing Then Exit Function
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If gobjPublicExpenseBillOperation.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)初始化失败，请与系统管理员联系！", vbExclamation, gstrSysName
         Exit Function
    End If
    CreatePublicExpenseBillOperation = True
End Function

Public Function IsAppointPati(ByVal lngRegID As Long, ByRef strBedNO As String) As Boolean
'功能:检查病人是否是预约中心病人
'参数:lngRegID-挂号ID;
'     strPatiInfo-病人信息  格式  床位号,入院科室ID,入院病区ID
'测试地址：http://192.168.32.201:8889/bizdomain/9e404039-9c1a-48a7-b283-093992bffe4a
'测试返回值:[{"RGST_ID":76237,"RGST_NO":"S0000021","PID":84414,"PAT_NAME":"审方测试","IBA_PAT_SEX":"女","PAT_AGE":"30岁","INP_BED_NO":"12","DSTRBT_INP_DEPT_ID":122,"DSTRBT_INP_DEPT":"心血管内科","IBA_WARD":"心血管内科","DSTRBT_INP_WARD_ID":122,"ORDER_ID":1214753,"HOME_PHNO":"123","CONTACTS_PHNO":"234"}]

    Dim strRet As String
    Dim blnRet As Boolean
    
    blnRet = Sys.NewSystemSvr("预约中心", "预约安排查询", "{""rgst_id_in"":""" & lngRegID & """}", strRet)
    If blnRet And strRet <> "" Then
        strRet = Mid(strRet, 2, Len(strRet) - 2)
        If strRet = "" Then Exit Function   '未找到床位
        strBedNO = zlStr.JSONParse("INP_BED_NO", strRet)   '床位号
    End If
    IsAppointPati = blnRet
End Function

