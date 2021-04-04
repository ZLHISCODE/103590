VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicBlood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mfrmBloodPeoPle As frmBloodPeoPle
Attribute mfrmBloodPeoPle.VB_VarHelpID = -1

Public Event PeopleChange(ByVal strRetrun As String)  '要求更新主窗体状态栏文字
Public Event AfterPatiFind(ByVal strIDKindstr As String, ByVal strValue As String, ByVal blnNext As Boolean, blnfind As Boolean) '查找的IDKindStr不存卡片上，则返回事件有调整程序处理
Public Event FilterRefreshByCode(ByVal strCode As String)
'医嘱操作枚举
Public Enum Enum_Advice
    Advice_新开 = 0
    Advice_修改
    Advice_删除
    Advice_校对
    Advice_作废
    Advice_发送
    Advice_回退
    Advice_回退作废 '住院专用
End Enum

Public Enum Enum_CheckType
    发血核对 = 0
    接收核对
    执行核对
End Enum

'该属性用于区分是否是试点用户启用血库系统，如果是则HIS基础程序和临床部件特殊sp版本程序做兼容性处理
'目前改属性支持的试点版本：34.110,35.20   说明：正式版本则不用该属性
Public Property Get IsPilotBlood() As Boolean
    IsPilotBlood = True
End Property

Public Function zlInitCommon(ByVal cnOracle As ADODB.Connection, Optional ByVal strDBUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及数据库连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-当前数据库登录用户名
    '返回:初始化成功,返回true,否则返回False
    '编制:刘鹏飞
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set gcnOracle = cnOracle: gstrDBUser = strDBUser
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = gobjComlib.GetUnitName
    Call InitLocPar
    
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ZlSetFormSize(ByVal frmParent As Object, ByVal lngMinWidth As Long, ByVal lngMinHeight As Long, ByVal lngMaxWidth As Long, ByVal lngMaxHeight As Long, ByRef lngWindowLong As Long, Optional ByVal blnUnLoad As Boolean = False)
'功能：设置窗体大小范围
'说明：窗体加载获取lngWindowLong，窗体卸载传入lngWindowLong
    Call SetFormSize(frmParent, lngMinWidth, lngMinHeight, lngMaxWidth, lngMaxHeight, lngWindowLong, blnUnLoad)
End Sub

Public Function zlBloodUserIdentity(ByVal frmParent As Object, ByVal lngModul As Long, ByVal lngReviewDeptID As Long, ByVal lngCheckDeptId As Long, ByVal str最小时间 As String, ByVal str最大时间 As String, _
    Optional ByVal blnSelectCheckUser As Boolean = True, Optional ByVal intMode As Enum_CheckType = 发血核对, Optional ByVal strIDs As String = "") As Boolean
'功能:血液发血/接收/执行身份验证
'入参：
'       lngModul=调用模块
'       lngReviewDeptID :复查科室ID(确定复查人员) User2
'       lngCheckDeptId：核对科室ID(确定核对人员) User1
'       str最小时间-str最大时间：确定核对时间的范围
'       blnSelectCheckUser：核对者是否需要身份验证(复查者强制需要进行身份验证)
'       strIDs：对于多袋血液进行发血、接收、执行必须传入收发ID串(逗号分割)，便于设置每一代血液的核对结果
    zlBloodUserIdentity = frmUserCheck.ShowMe(frmParent, lngModul, lngReviewDeptID, lngCheckDeptId, str最小时间, str最大时间, blnSelectCheckUser, intMode, strIDs)
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'医嘱调用相关
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'显示配发信息
Public Function zlGetBloodListInfo() As Object
    Set zlGetBloodListInfo = New frmBloodListInfo
End Function

'输血执行
Public Function zlGetBloodExec() As Object
    Set zlGetBloodExec = New frmBloodExec
End Function

Public Sub ViewBloodExecEdit(ByVal frmParent As Object, ByVal lng收发id As Long)
    '功能：输血执行情况查看
    Call frmBloodExecEdit.ViewExecution(frmParent, lng收发id)
End Sub

Public Function ShowBloodInstantRptPrint(ByVal objfrm As Object, ByVal lngActiveID As Long) As Boolean
'功能：供护士站或医技工作站调用(输血执行单打印)
'参数： objFrm--调用主窗体
'           lngActiveID--医嘱ID
    ShowBloodInstantRptPrint = zlBloodInstantRptPrint(objfrm, lngActiveID)
End Function

Public Function AdviceOperation(ByVal lngMoudle As Long, ByVal lng医嘱ID As Long, ByVal intOperation As Enum_Advice, Optional ByVal blnMoved As Boolean = False, _
        Optional ByRef strErrInfo As String = "") As Boolean
'功能：医嘱操作调用接口（新开、删除、发送、回退时此方法的调用请放在医嘱操作事物中调用，修改、校对、作废为操作校验检查，放在事物之前）
'入参:
'       lngMoudle:调用模块号
'       lng医嘱ID:血液医嘱主医嘱ID
'       intOperation:医嘱操作类型(枚举),含：新开、修改、删除、校对、作废、发送、回退
'       blnMoved:病人历史数据是否转出
'出参：
'       strErrInfo：接口返回FALSE时的信息
'返回：成功=TRUE，失败=False
    AdviceOperation = zlAdviceOperation(lngMoudle, lng医嘱ID, intOperation, blnMoved, strErrInfo)
End Function

Public Function GetPrepareBloodRs(ByVal lng医嘱ID As Long, rsBlood As ADODB.Recordset) As Boolean
'功能:获取配血记录相关信息，以记录集返回
'入参：lng医嘱ID：输血医嘱ID
'出参：rsTemp：血液配血记录记录集
'返回：TRUE=成功,False-失败
    Dim strSQL As String
    On Error GoTo ErrHand
    strSQL = "Select Id, 病人来源, 病人id, 主页id, 申请id, 执行部门id, 记录状态, 接收人, 接收时间, 完成人, 完成时间, 记录性质" & vbNewLine & _
        "From 血液配血记录" & vbNewLine & _
        "Where 申请id = [1]"
    Set rsBlood = gobjDatabase.OpenSQLRecord(strSQL, "GetPrepareBloodRs", lng医嘱ID)
    GetPrepareBloodRs = Not rsBlood.EOF
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function BloodPeopleCodeFilter(rs As Recordset)
    Call mfrmBloodPeoPle.FilterRefreshBC(rs)
End Function

Public Function AdviceTermination(ByVal lngMoudle As Long, ByVal lng医嘱ID As Long, Optional ByVal blnCance As Boolean = False, Optional ByVal blnMoved As Boolean = False, _
    Optional ByRef strErrInfo As String = "", Optional ByRef blnFallback As Boolean = False) As Boolean
'功能：医嘱标记未用调用(输血终止功能处理)
'入参:
'       lngMoudle:调用模块号
'       lng医嘱ID:血液医嘱主医嘱ID
'       blnCance:TRUE:取消输血终止,False:输血终止
'       blnMoved:病人历史数据是否转出
'出参：
'       strErrInfo：接口返回FALSE时的信息
'       blnFallback:标记未用的同时是否回退医嘱
'返回：成功=TRUE，失败=False
    AdviceTermination = True
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'病人卡片列表相关
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlGetBloodPeople() As Object
    Set zlGetBloodPeople = mfrmBloodPeoPle
End Function

Public Function zlRefreshBloodPeople(ByVal frmParent As Object, ByVal strRule As String, ByVal rs As ADODB.Recordset, ByVal lngMode As Long, Optional ByVal imgList As Object, Optional ByVal strIDKindstr As String = "") As Boolean
    Call mfrmBloodPeoPle.UserInit(frmParent, strRule, imgList, lngMode, strIDKindstr)
    Call mfrmBloodPeoPle.ShowPeople(rs)
    zlRefreshBloodPeople = True
End Function

Public Sub mfrmBloodPeoPle_CodeFilter(ByVal strCode As String)
    '血袋编号过滤
    RaiseEvent FilterRefreshByCode(strCode)
End Sub

Public Sub BloodPeopleLocation(strTitle As String, strfind As String)
    '定位到指定的人员卡上
    Call mfrmBloodPeoPle.SetCardFocus(strTitle, strfind)
End Sub

Private Sub mfrmBloodPeoPle_AfterPatiFind(ByVal strIDKindstr As String, ByVal strValue As String, ByVal blnNext As Boolean, blnfind As Boolean)
    RaiseEvent AfterPatiFind(strIDKindstr, strValue, blnNext, blnfind)
End Sub

Public Sub FindBloodPeopleByVbKey(Optional ByVal blnNext As Boolean)
    '功能：通过快捷键的方式开始查找，或查找下一个病人
    Call mfrmBloodPeoPle.FindPatiByVbKey(blnNext)
End Sub

Private Sub mfrmBloodPeoPle_CardChanged()
    RaiseEvent PeopleChange(mfrmBloodPeoPle.strReturn)
End Sub

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'输血反应登记
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlShowBloodReaction(ByVal frmParent As Object, ByVal lngSys As Long, ByVal lngMoudle As Long, lng场合 As Long, lng病人ID As Long, lng主页id As Long, _
                    lng病人来源 As Long, Optional ByVal lngisMoudal As Long = 1, Optional ByVal lng收发id As Long = 0)
'-------------------------------------------------------------------------------------------------------------------------
'功能：对指定病人的输血反应进行查看或编辑
'入参：frmParent-调用窗体对象
'          lngMoudle-调用模块号,lngSys--调用系统号
'         lng场合-1:医生站调用;2-输血科调用
'         lng病人id：病人ID
'         lng主页id：主页ID或门诊号(住院病人为主页ID，门诊病人为门诊号)
'         lng病人来源：1-门诊;2-住院
'         lngisMoudal:1-模态，0-非模态,默认为模态窗体
'-------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String
    strPrivs = GetInsidePrivs(2200, p输血反应管理)
    frmBloodReaction.BloodReaction frmParent, lng场合, lng病人ID, lng主页id, lng病人来源, lngSys, lngMoudle, strPrivs, lngisMoudal, lng收发id
End Function

Public Function zlShowBloodReactionRecord(ByVal frmParent As Variant, ByVal lngSys As Long, ByVal lngMoudle As Long, lng场合 As Long, Optional ByVal lngisMoudal As Long = 1)
'-------------------------------------------------------------------------------------------------------------------------
'功能：对操作员所属科室下病人的输血反应进行查看或编辑
'入参：frmParent-调用窗体对象, BH调用穿入父窗体HWND
'          lngMoudle-调用模块号；lngsys-调用系统号
'          lng场合-  0:门诊医生站调用1:住院医生站调用;2-输血科调用
'          lngisMoudal:1-模态，0-非模态
'-------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String
    strPrivs = GetInsidePrivs(2200, p输血反应管理)
    frmBloodReactionRecord.BloodReactionRecord frmParent, lng场合, lngSys, lngMoudle, strPrivs, lngisMoudal
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'血液接收登记
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlShowBloodReceives(ByVal frmParent As Variant, ByVal lngSys As Long, ByVal lngMoudle As Long, Optional ByVal lngisMoudal As Long = 1, _
            Optional int场合 As Integer = 0)
'功能：对操作员所属科室下病人的血液进行接收操作(CodeMan入口调用)
'入参：frmParent-调用窗体对象, BH调用穿入父窗体HWND,lngisMoudal:1-模态，0-非模态
'          lngMoudle-调用模块号  lngSys:调用系统号
'-------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String
    strPrivs = GetInsidePrivs(100, p血液接收登记)
    frmBloodReceivesRecord.BloodReceives frmParent, lngSys, lngMoudle, strPrivs, lngisMoudal, int场合
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'医护站消息状态
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlIsBloodMessageDone(ByVal intMode As Integer, ByVal lng病人ID As Long, ByVal lng就诊id As Long, ByVal int阅读场合 As Integer, ByVal lng阅读部门id As Long) As Boolean
'功能：修正医护站该病人的消息状态
'入参：intMode-模式：1-血库回收消息,2-输血反应消息
'
'-------------------------------------------------------------------------------------------------------------------------
    zlIsBloodMessageDone = IsBloodMessageDone(intMode, lng病人ID, lng就诊id, int阅读场合, lng阅读部门id)
End Function
Private Sub Class_initialize()
    Set mfrmBloodPeoPle = New frmBloodPeoPle
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not mfrmBloodPeoPle Is Nothing Then
        Unload mfrmBloodPeoPle
        Set mfrmBloodPeoPle = Nothing
    End If
    
    If Not gobjPublicExpense Is Nothing Then
        Set gobjPublicExpense = Nothing
    End If
    Call CloseSquareCardObject
    If Err <> 0 Then Err.Clear
End Sub
