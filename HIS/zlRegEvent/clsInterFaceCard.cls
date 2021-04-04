VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterFaceCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mobjOneCard As clsOneCardComLib
Private mlngSys As Long
Private mlngModule As Long

Public Property Get lngModul() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objOneCard
    lngModul = mlngModule
End Property

Public Property Get lngSys() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objOneCard
    lngSys = mlngSys
End Property

Public Property Get objOneCard() As clsOneCardComLib
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objOneCard
    Set objOneCard = mobjOneCard
End Property


Public Function Init(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
                    ByVal cnOracle As ADODB.Connection, ByVal strDbUser As String) As Boolean
    '初始化类模块变量
    On Error GoTo ErrHandler
    
    Set mobjOneCard = New clsOneCardComLib
    If mobjOneCard.zlInitComponents(frmMain, lngModule, lngSys, strDbUser, cnOracle) = False Then Exit Function
    
    mlngSys = lngSys
    mlngModule = lngModule
    Init = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
    
Public Function Is身份证签约(strCardNo As String, Optional ByRef lng身份证类别ID As Long, Optional ByRef str终止使用时间 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查身份证是否可以签约
    '入参:绑定卡号
    '编制:王吉
    '日期:2012-08-31 11:32:14
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCardStatus As Integer
    Dim objCard As Card
    On Error GoTo ErrHand:
    If mobjOneCard.zlGetCardFromTypeName("二代身份证", False, objCard) = False Then Exit Function
    
    If objCard Is Nothing Then Exit Function
    '找到病人说明已签约，不能继续签约
    If mobjOneCard.zlGetPatiID(objCard.接口序号, strCardNo, , , , , , , , , , , , , , intCardStatus) Then Exit Function
    If intCardStatus <> 0 Then Exit Function
    lng身份证类别ID = objCard.接口序号
    str终止使用时间 = objCard.zlGetDefaultDate
    Is身份证签约 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function
Public Function GetPatiID(ByVal frmMain As Object, ByVal objControl As Object, _
                ByVal strCardTypes As String, ByVal strInput As String, ByRef lng病人ID As Long, _
                Optional ByVal strCardPassWord As String, Optional ByVal blnUserCancel As Boolean, _
                Optional ByVal blnNotShowErr As Boolean, Optional ByRef intCardStatus As Integer, _
                Optional ByVal blnShowMergePati As Boolean) As Boolean
    
    If mobjOneCard.zlGetPatiID(strCardTypes, strInput, blnNotShowErr, lng病人ID, strCardPassWord, , , objControl, frmMain, blnShowMergePati, True, , blnUserCancel, , True, intCardStatus) = False Then Exit Function
End Function

Public Function GetPatient(objPati As clsPatientInfo, ByVal frmMain As Object, ByVal objControl As Object, _
                ByVal str查询方式 As String, ByVal lng卡类别id As Long, ByVal strInput As String, _
                Optional ByVal blnCard As Boolean, Optional ByVal blnCont在院 As Boolean = True, _
                Optional ByVal blnSeekName As Boolean, Optional ByVal strName As String, Optional ByVal strSex As String, _
                Optional ByRef strPassWord As String, Optional ByRef blnUserCancel As Boolean, Optional ByRef intCardStatus As Integer, _
                Optional ByRef strValidTime As String, Optional ByVal blnAddNewPati As Boolean, _
                Optional ByVal blnLoadFamily As Boolean, Optional ByVal blnLoadDrug As Boolean, _
                Optional ByVal blnLoadImmune As Boolean, Optional ByVal blnLoadCard As Boolean, _
                Optional cllFamily As Collection, Optional cllDrug As Collection, _
                Optional cllImmune As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人信息
    ' 入参 : int查询方式：0：按病人id查找;-1:卡号模式查找;>0 按卡类别查找
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/6 16:19
    '---------------------------------------------------------------------------------------
    Dim blnValidTime As Boolean '验证卡有效终止时间
    Dim cllPati As Collection, cllOtherFindCons As Collection
    Dim strErrMsg As String
    Dim lngDefaultCardTypeID As Long, lng病人ID As Long
    On Error GoTo ErrHand
    
    Set objPati = New clsPatientInfo
    Set cllFamily = New Collection
    Set cllDrug = New Collection
    Set cllImmune = New Collection
    Set cllOtherFindCons = New Collection
    If str查询方式 = "病人ID" Then
        lng病人ID = Val(strInput)
    ElseIf str查询方式 = "姓名" Or str查询方式 = "姓名或就诊卡" Then
        str查询方式 = "姓名"
        If blnCard Then
            If mobjOneCard.zlGetPatiID(lng卡类别id, strInput, False, lng病人ID, strPassWord, strErrMsg, , objControl, frmMain, , , , blnUserCancel, , , intCardStatus, strValidTime) = False Then lng病人ID = 0
            
        ElseIf blnSeekName Then
            If GetPatiIDByName(frmMain, objControl, strInput, strSex, lng病人ID, blnCont在院, blnAddNewPati) = False Then Exit Function
            If lng病人ID = 0 Then GetPatient = True: Exit Function '新病人
        End If
        If lng病人ID = 0 Then
            If mobjOneCard.zlIsMobileNo(strInput) Then
                lng卡类别id = 0
                str查询方式 = "手机号"
            End If
        End If
    ElseIf str查询方式 = "身份证号" Or str查询方式 = "身份证" Or str查询方式 = "二代身份证" Then
        str查询方式 = "身份证"
        If GetPatiIDByIDCard(frmMain, objControl, strInput, strName, strSex, lng病人ID, blnCont在院, blnUserCancel) = False Then Exit Function
        If blnUserCancel Then Exit Function
        If lng病人ID = 0 Then GetPatient = True: Exit Function '新病人
        
    ElseIf str查询方式 = "手机号" Then
        If mobjOneCard.zlIsMobileNo(strInput) = False Then Exit Function
    ElseIf lng卡类别id > 0 Then
        str查询方式 = lng卡类别id
        blnValidTime = True
    End If
    
    If (str查询方式 = "姓名" Or str查询方式 = "姓名或就诊卡") And Not blnCard And lng病人ID = 0 Then
        GetPatient = True: Exit Function   '新病人
    End If
    If lng病人ID = 0 And str查询方式 <> "病人ID" Then
       If mobjOneCard.zlGetPatiID(str查询方式, strInput, False, lng病人ID, strPassWord, , , objControl, frmMain, , , , blnUserCancel, , Not blnValidTime, intCardStatus, strValidTime) = False Then Exit Function
    End If
    If lng病人ID = 0 Then Exit Function
    If mobjOneCard.zlGetPatiInforFromPatiID(lng病人ID, objPati, strErrMsg, , , blnLoadFamily, blnLoadDrug, blnLoadImmune, blnLoadCard, cllFamily, cllDrug, cllImmune) = False Then lng病人ID = 0: Exit Function
    If Not blnCont在院 And objPati.在院 Then
        If GetInpatiState(objPati.病人ID, objPati.主页ID) = 0 Then lng病人ID = 0: Exit Function
    End If
    If InStr(",姓名,姓名或就诊卡,病人ID,门诊号,手机号,", "," & str查询方式 & ",") > 0 And lng卡类别id = 0 Then
        If strPassWord = "" Then strPassWord = objPati.密码
    End If
    
    GetPatient = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatientOtherInfo(ByVal lng病人ID As Long, ByRef cllCert As Collection, _
                                ByRef cllOther As Collection, ByRef cllContact As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人健康页信息
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/8 16:48
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    If mobjOneCard.zlGetPatiOtherInforFromPatiID(lng病人ID, , , , , , True, True, _
                    , , , cllCert, cllOther, cllContact) = False Then Exit Function
    GetPatientOtherInfo = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatientDrugInfo(ByVal lng病人ID As Long, cllDrug As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人健康页信息
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/8 16:48
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    If mobjOneCard.zlGetPatiOtherInforFromPatiID(lng病人ID, , , , True, , , , , cllDrug) = False Then Exit Function
    GetPatientDrugInfo = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatientExpend(ByVal lng病人ID As Long, ByVal str信息名 As String) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 查询单个病人从表信息值
    ' 入参 : int查询方式：0：按病人id查找;-1:卡号模式查找;>0 按卡类别查找
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/6 16:19
    '---------------------------------------------------------------------------------------
    Dim cllPati As Collection, str信息值 As String
    
    On Error GoTo ErrHand
    If lng病人ID = 0 Or str信息名 = "" Then Exit Function
    If mobjOneCard.zlPatiSvr_GetPatiExtendInfo(lng病人ID, str信息名, cllPati) = False Then Exit Function
    If cllPati Is Nothing Then Exit Function
    If cllPati.Count = 0 Then Exit Function
    str信息值 = cllPati(1)("信息值")
    GetPatientExpend = str信息值
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatientExpends(ByVal lng病人ID As Long, ByVal str信息名 As String, ByRef cllExpends As Collection, Optional ByVal lng就诊ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 查询单个病人从表信息值
    ' 入参 : int查询方式：0：按病人id查找;-1:卡号模式查找;>0 按卡类别查找
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/6 16:19
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    Set cllExpends = New Collection
    If lng病人ID = 0 Or str信息名 = "" Then Exit Function
    If mobjOneCard.zlPatiSvr_GetPatiExtendInfo(lng病人ID, str信息名, cllExpends, , , lng就诊ID) = False Then Exit Function
    If cllExpends Is Nothing Then Exit Function
    GetPatientExpends = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatiVitalSigns(ByVal lng病人ID As Long, ByVal lng挂号ID As String, ByRef cllVital As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人生命体征信息,如果护理记录中没有，则从病人从表中取
    ' 入参 : blnOutPati-门诊病人
    ' 出参 : cllVital:体征信息(Collect)(项目,值,单位)
    ' 返回 : 返回病人的体征信息，包括项目，数值，单位
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 19:48
    '---------------------------------------------------------------------------------------
    Dim cllExpends As Collection, cllExpend As Collection, cllTmp As Collection
    Dim i As Long
    On Error GoTo ErrHand
    If mobjOneCard.zlGetPatiVitalSigns(lng病人ID, lng挂号ID, cllVital) = False Then
        If GetPatientExpends(lng病人ID, "身高,体重,体温,呼吸,脉搏,收缩压,舒张压,血压单位,血糖", cllExpends, lng挂号ID) = False Then Exit Function
        For i = 1 To cllExpends.Count
            Set cllExpend = cllExpends(i): Set cllTmp = New Collection
            '项目单位, b.项目名称, b.记录内容
            cllTmp.Add Nvl(cllExpend("信息名")), "项目名称"
            cllTmp.Add Nvl(cllExpend("信息值")), "项目值"
            cllTmp.Add "", "项目单位"
            cllVital.Add cllTmp, cllTmp("项目名称")
        Next
    End If
    GetPatiVitalSigns = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlCheckMzLgPatiUseDeposit(ByVal lng病人ID As Long, ByVal lng主页Id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查门诊留观病人是否能使用门诊预交
    '入参:lng病人ID-病人ID
    '出参:
    '返回:是门诊留观能够使用门诊预交返回true,否则返回False
    '编制:李南春
    '日期:2019/9/6 15:57:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnLimitDeposit As Boolean, blnMzLgPati As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    blnLimitDeposit = Val(zldatabase.GetPara(Val("323-门诊留观病人预交款使用控制"), glngSys)) <> 0
    If Not blnLimitDeposit Then zlCheckMzLgPatiUseDeposit = True: Exit Function
    
    If mobjOneCard.zlCheckMzLgPati(lng病人ID, lng主页Id, blnMzLgPati) = False Then Exit Function
    zlCheckMzLgPatiUseDeposit = Not blnMzLgPati
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExistOperationData(ByVal lng病人ID As Long, ByVal str挂号单 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否当前病人发生了业务数据
    '入参:lng病人ID-病人ID-指定的病人ID
    '       strNo-挂号单号
    '       lng诊疗ID-挂号ID
    '出参:
    '返回:存在业务数据,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-06 17:21:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavAdvice As Boolean
    If mobjOneCard.zlExistAdvice(lng病人ID, str挂号单, blnHavAdvice) Then
        zlExistOperationData = blnHavAdvice
    Else '服务返回false，不能确定身份有医嘱，则固定返回true
        zlExistOperationData = True
    End If
End Function

Public Sub InitSendCardPreperty(ByVal lngModule As Long, Optional lng卡类别id As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化刷卡属性
    '编制:刘兴洪
    '日期:2011-07-25 11:03:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strSQL As String, blnBoundCard As Boolean
    Dim rsTemp As ADODB.Recordset, str批次 As String, varData As Variant, i As Long
    Dim varTemp  As Variant, ty_Card As Ty_CardProperty
    Dim objCard As Card
    On Error GoTo ErrHand
    If lng卡类别id <> 0 Then
        lngCardTypeID = lng卡类别id
    Else
        lngCardTypeID = Val(zldatabase.GetPara("缺省医疗卡类别", glngSys, lngModule, 0))
    End If
    
    gCurSendCard = ty_Card
    If mobjOneCard.zlGetCard(lngCardTypeID, False, objCard) = False Then Exit Sub
    If objCard Is Nothing Then Exit Sub
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '问题号:57326
    '85565,李南春,2015/7/10:读卡性质
    With ty_Card
        .lng卡类别id = objCard.接口序号
        .str卡名称 = objCard.名称
        .lng卡号长度 = objCard.卡号长度
        .lng结算方式 = objCard.结算方式
        .bln自制卡 = objCard.自制卡
        .bln严格控制 = objCard.是否严格控制
        .str卡号密文 = objCard.卡号密文规则
        .int密码长度 = objCard.密码长度
        .int密码长度限制 = objCard.密码长度限制
        .int密码规则 = objCard.密码规则
        .int密码输入限制 = objCard.密码输入限制
        .bln缺省密码 = objCard.是否缺省密码
        .bln就诊卡 = .str卡名称 = "就诊卡" And objCard.系统
        .str特准项目 = objCard.特定项目
        .bln缺省标志 = objCard.缺省标志
        '问题号:56599
        .bln是否制卡 = objCard.是否制卡
        .bln是否发卡 = objCard.是否发卡
        .bln是否写卡 = objCard.是否写卡
        '问题号:57326
        .lng发卡性质 = objCard.发卡性质
        .bln重复使用 = objCard.卡号重复使用
        .bln刷卡 = objCard.是否刷卡
        .bln扫描卡 = objCard.是否扫描
        .bln接触式读卡 = objCard.是否接触式读卡
        .bln非接触式读卡 = objCard.是否非接触式读卡
        .byt发卡控制 = objCard.发卡控制
        .bln启用回车 = objCard.设备是否启用回车
        .str有效时间 = objCard.缺省有效时间
        .str短名称 = objCard.短名
        If .str特准项目 <> "" Then
            Set .rs卡费 = zlGetSpecialItemFee(.str特准项目)
        Else
            Set .rs卡费 = Nothing
        End If
        str批次 = zldatabase.GetPara("共用医疗卡批次", glngSys, lngModule, "0")
        '领用ID,卡类别ID|...
         .lng共用批次 = 0
        varData = Split(str批次, "|")
        For i = 0 To UBound(varData)
             varTemp = Split(varData(i), ",")
             If Val(varTemp(0)) <> 0 Then
                '140937:李南春,2019/5/21,值传递错误，导致无法取到公用卡号
                If ExistShareBill(Val(varTemp(0)), 5) Then
                    If Val(varTemp(1)) = .lng卡类别id Then
                        .lng共用批次 = Val(varTemp(0)): Exit For
                    End If
                End If
             End If
        Next
    End With
    gCurSendCard = ty_Card
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Function GetPatiIDByComminuty(ByVal int社区 As Integer, ByVal strComminuty As String) As Long
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据社区号获取病人id
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/15 18:35
    '---------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    
    If mobjOneCard.zlGetPatiIDByComminuty(int社区, strComminuty, lng病人ID) = False Then Exit Function
    GetPatiIDByComminuty = lng病人ID
End Function


Public Function GetPatiIDByName(ByVal frmMain As Object, ByVal objControl As Object, ByVal strName As String, _
    ByVal str性别 As String, ByRef lngPatiId As Long, _
    Optional ByVal blnCont住院 As Boolean, Optional ByVal blnAddNewPati As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人姓名，获取病人ID
    '入参:objControl-调用的控件
    '     strName-病人信息
    '     frmMain-调用的主窗体
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-01 11:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnUserCancel As Boolean
    Dim byt场合 As Byte
    Dim rsPati As ADODB.Recordset, rsSel As ADODB.Recordset
    Dim lngDefaultCardTypeID As Long
    
    On Error GoTo ErrHand
    If mlngModule = 1111 Then
        byt场合 = 0
    ElseIf mlngModule = 1113 Then
        byt场合 = 3
    Else
        byt场合 = 4
    End If
    
    If LenB(strName) < 4 Then Exit Function
    lngDefaultCardTypeID = Val(zldatabase.GetPara("缺省医疗卡类别", mlngSys, mlngModule, 0))
    
    If mobjOneCard.zlGetPatiIdFromPatiName(objControl, strName, lngPatiId, frmMain, gintNameDays, lngDefaultCardTypeID, IIf(blnCont住院, 0, 1), byt场合, blnUserCancel, blnAddNewPati) = False Then Exit Function
    If blnUserCancel Then Exit Function
    If blnAddNewPati Then GetPatiIDByName = True: Exit Function

    GetPatiIDByName = lngPatiId <> 0
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPatiIDByIDCard(ByVal frmMain As Object, ByVal objControl As Object, ByVal strIDCard As String, _
    ByVal strName As String, ByVal strSex As String, ByRef lngPatiId As Long, _
    Optional ByVal blnCont住院 As Boolean, Optional ByRef blnUserCancel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人身份证，获取病人ID
    '入参:objControl-调用的控件
    '     strName-病人信息
    '     frmMain-调用的主窗体
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-01 11:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsPati As ADODB.Recordset, rsSel As ADODB.Recordset
    Dim byt场合 As Byte
    
    On Error GoTo ErrHand
    If mlngModule = 1111 Then
        byt场合 = 0
    ElseIf mlngModule = 1113 Then
        byt场合 = 3
    Else
        byt场合 = 4
    End If
    If strIDCard = "" Then Exit Function
    
    If mobjOneCard.zlGetPatiIdFromPatiIDCard(objControl, strIDCard, strName, strSex, lngPatiId, frmMain, IIf(blnCont住院, 0, 1), byt场合, blnUserCancel) = False Then Exit Function
    If blnUserCancel Then Exit Function

    GetPatiIDByIDCard = lngPatiId <> 0
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function


Public Function GetCard(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据卡类别ID获取卡对象
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/25 16:28
    '---------------------------------------------------------------------------------------
    If mobjOneCard.zlGetCard(lngCardTypeID, bln消费卡, objCard) = False Then Exit Function
    If objCard Is Nothing Then Exit Function
    GetCard = True
End Function

    
Public Function GetCardByName(ByVal strCardName As String, ByVal bln消费卡 As Boolean, ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据卡类别名称获取卡对象
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/25 16:28
    '---------------------------------------------------------------------------------------
    If mobjOneCard.zlGetCardFromTypeName(strCardName, bln消费卡, objCard) = False Then Exit Function
    If objCard Is Nothing Then GetCardByName = True
End Function

Public Function GetCards(ByVal bytType As Byte) As Cards
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的卡对象
    '入参:bytType-0-所有医疗卡;
    '             1-启用的医疗卡,
    '             2-所有存在三方账户的三方卡
    '             3-启用的三方账户的医疗卡
    '出参:
    '返回:成功返回本地支持的卡对象
    '编制:刘兴洪
    '日期:2013-10-23 17:43:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set GetCards = mobjOneCard.zlGetCards(bytType)
End Function

Public Function GetPatiCertInfo(ByVal lng病人ID As Long, ByRef cllCert As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据病人ID获取病人有效的证件信息
    ' 入参 :
    ' 出参 : cllCert-证件列表(Collect:卡类别名称,卡号)
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/8 20:00
    '---------------------------------------------------------------------------------------
    If lng病人ID = 0 Then Exit Function
    If mobjOneCard.zlGetPatiCertInfo(lng病人ID, cllCert) = False Then Exit Function
    If cllCert Is Nothing Then Exit Function
    If cllCert.Count = 0 Then Exit Function
    GetPatiCertInfo = True
End Function


Public Function GetPatiIdsByRange(ByVal strCondition As String, ByRef strPatiIds As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据条件值获取符合条件的病人ID
    '入参:strValue :可能是就诊卡号、身份证号、IC卡号、门诊号
    '出参:
    '返回:执行成功返回True，失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetPatiIdsByRange = mobjOneCard.zlGetPatiIdsByRange(strCondition, strPatiIds)
End Function

Public Function CheckPatiInhospital(ByVal lng病人ID As Long, ByRef blnInhospital As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查病人是否在院就诊
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 16:04
    '---------------------------------------------------------------------------------------
    CheckPatiInhospital = mobjOneCard.zlPatiIsInhospital(lng病人ID, blnInhospital)
End Function

Public Function GetPatiInfsByPatiIds(ByVal str病人Ids As String, cllPatiInfos_out As Collection, Optional ByVal blnExpendInfo As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息集
    '入参:str病人Ids-病人ID
    '出参:cllPatiInfos_out-返回的数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetPatiInfsByPatiIds = mobjOneCard.zlGetPatiInfsByPatiIds(2, blnExpendInfo, str病人Ids, cllPatiInfos_out)
End Function

Public Function GetInPatiInfsByPatiId(ByVal lng病人ID As Long, Optional ByVal lng主页Id As Long = -1, Optional cllPatiInfos_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取住院病人病案信息集
    '入参:lng病人ID
    '    :lng主页ID :-1表示返回所有的住院信息
    '出参:cllPatiInfos_out-返回的数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetInPatiInfsByPatiId = mobjOneCard.zlGetPatiBaseInfoByPatiId(lng病人ID, lng主页Id, cllPatiInfos_out)
End Function

Public Function GetPatiInfsByPatiFilter(ByVal blnCont在院 As Boolean, ByVal cllFilter As Collection, cllPatiInfos_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息集
    '入参:blnCont在院-是否查询在院病人
    '    :cllFilter-查询条件
    '出参:cllPatiInfos_out-返回的数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetPatiInfsByPatiFilter = mobjOneCard.zlGetPatiInfsByPatiIds(IIf(blnCont在院, 2, 0), True, "", cllPatiInfos_out, cllFilter, 1)
End Function

Public Function GetAvailabilityCardType() As String
    GetAvailabilityCardType = mobjOneCard.zlGetAvailabilityCardType
End Function

Public Function WriteCard(ByVal frmMain As Object, ByVal lng病人ID As Long, ByVal lng卡类别id As Long) As Boolean
    WriteCard = mobjOneCard.zlBandCardArfter(frmMain, mlngModule, lng卡类别id, lng病人ID)
End Function

Public Function CheckUsed门诊号(ByVal lng病人ID As Long, ByVal str门诊号 As String, _
                ByRef blnUsedByOther As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查门诊号是否被使用
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/12 11:33
    '---------------------------------------------------------------------------------------
    CheckUsed门诊号 = mobjOneCard.zlCheckOutNoIsExist(lng病人ID, str门诊号, blnUsedByOther)
End Function


Public Function zlUpdateOutMedRec(ByVal lng病人ID As Long, Optional ByVal str门诊号 As String, Optional ByVal str建立日期 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 更新门诊病案，没有记录时新增记录,不传str病案号时删除病案记录
    ' 入参 : str病案号-门诊号
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/21 13:53
    '---------------------------------------------------------------------------------------
    zlUpdateOutMedRec = mobjOneCard.zlUpdateOutMedRec(lng病人ID, str门诊号, str建立日期)
End Function


Public Function GetMoneyInfo(lng病人ID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的门诊预交剩余额
    '入参:
    '       curModiMoney=修改时,原单据的当前病人的费用合计
    '       int类型:类型(0-门诊和住院共用;1-门诊;2-住院),-1表示所有
    '       bytModiMoneyType-修改费用的类别(在按类别统计时有效)
    '       blnFamilyMoney-是否读取家属余额
    '出参:
    '返回:病人剩余额
    '编制:刘兴洪
    '日期:2011-07-21 15:33:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, blnFamilyMoney As Boolean
    Dim strSQL As String, strFamilyIds As String
    On Error GoTo errH
    blnFamilyMoney = True
    If blnFamilyMoney Then
        Call mobjOneCard.ZlGetPatiFamilyMember(1, lng病人ID, strFamilyIds)
    End If
    
    strSQL = "Select " & IIf(blnFamilyMoney, "0 As 家属,", "") & _
            "       Nvl(费用余额,0) As 费用余额,Nvl(预交余额,0) As 预交余额" & _
            " From 病人余额" & _
            " Where 性质=1 And 病人ID=[1] And 类型 = 1"
    '79868,读取病人家属余额
    If blnFamilyMoney And strFamilyIds <> "" Then
        strSQL = strSQL & " Union All " & _
                " Select /*+cardinality(B,10) */ " & IIf(blnFamilyMoney, "1 As 家属,", "") & _
                "       Nvl(a.费用余额, 0) As 费用余额, Nvl(a.预交余额, 0) As 预交余额" & _
                " From 病人余额 A, (Select Column_Value as 家属ID from Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) B" & _
                " Where a.病人id = b.家属id And a.性质 = 1 And a.类型 = 1 "
    End If

    strSQL = "Select " & IIf(blnFamilyMoney, "家属,", "") & _
            "       nvl(Sum(费用余额),0) as 费用余额,nvl(Sum(预交余额),0) as 预交余额 " & _
            " From (" & strSQL & ")" & vbCrLf & _
                IIf(blnFamilyMoney, " Group by 家属", "")
    
    Set GetMoneyInfo = zldatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, strFamilyIds)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SimilarIDs(str身份证号 As String) As String
    '功能：检查病人是否存在相似信息
    '返回：相似记录的病人ID串,如"234,235,236"
    Dim i As Integer
    Dim cllPati As Collection
    Dim cllFilter As New Collection
    On Error GoTo errH
    
    cllFilter.Add Array("身份证号", str身份证号)
    If mobjOneCard.zlGetPatiInfsByFilter(2, False, cllFilter, cllPati) = False Then Exit Function
    If cllPati Is Nothing Then Exit Function
    
    For i = 1 To cllPati.Count
        SimilarIDs = SimilarIDs & "|ID:" & cllPati(i)("病人ID")
        SimilarIDs = SimilarIDs & ",姓名:" & cllPati(i)("姓名")
        SimilarIDs = SimilarIDs & ",门诊号:" & IIf(cllPati(i)("门诊号") = "", "无", cllPati(i)("门诊号"))
        SimilarIDs = SimilarIDs & ",身份证号:" & IIf(cllPati(i)("身份证号") = "", cllPati(i)("身份证号"), "未登记")
        SimilarIDs = SimilarIDs & ",地址:" & IIf(cllPati(i)("家庭地址") = "", cllPati(i)("家庭地址"), "未登记")
        SimilarIDs = SimilarIDs & ",登记日期:" & Format(cllPati(i)("登记时间"), "YYYY-MM-DD")
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckMobile(str手机号 As String, Optional ByVal lng病人ID As Long, _
                    Optional ByVal blnShowMsg As Boolean = True, Optional ByRef strErrMsg As String) As Boolean
    '功能：判断指定手机号是否是正确的手机号格式以及是否已经存在于数据库中
    '入参：str手机号-进行检查的手机号
    '      lng病人ID - 检查手机号重复性，不需要检查就传0
    '      blnShowMsg-错误时是否显示提示
    '出参：strErrMsg -错误信息
    '返回：手机号允许使用-true；否则返回False
    Dim blnUsedByOther As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMobileRange As String
    Dim blnQuery As Boolean
    
    On Error GoTo errH
    '127941:李南春,2018/8/10,根据手机号段检查手机号是否合法
    If str手机号 = "" Then CheckMobile = True: Exit Function
    strSQL = "Select Max(检查结果) As 检查结果" & vbNewLine & _
                "From (Select Decode(号码长度, Length([1]), 1, 2) As 检查结果" & vbNewLine & _
                "       From 手机号常用号段表" & vbNewLine & _
                "       Where 号段 = Substr([1], 1, Length(号段)))"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "手机号检查", str手机号)
    
    If rsTmp.RecordCount = 0 Then
        strErrMsg = "未能在【手机号常用号段表】检索到输入的手机号格式，请重新录入！"
    ElseIf Val(Nvl(rsTmp!检查结果)) = 0 Then
        strErrMsg = "未能在【手机号常用号段表】检索到输入的手机号格式，请重新录入！"
    ElseIf Val(Nvl(rsTmp!检查结果)) = 2 Then
        strErrMsg = "输入的手机号位数不正确，请重新录入！"
    End If
    
    If gbln检查手机号重复 Then
        If mobjOneCard.zlCheckPhoneIsExist(lng病人ID, str手机号, blnUsedByOther, Not blnShowMsg, strErrMsg) = False Then Exit Function
        If blnUsedByOther Then
            strErrMsg = "输入的手机号与其他病人重复，是否确定录入？"
            blnQuery = True
        End If
    End If
    
    If strErrMsg <> "" Then
        If Not blnShowMsg Then Exit Function
        If blnQuery Then
            If MsgBox(strErrMsg, vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
                strErrMsg = "": Exit Function
            End If
            strErrMsg = ""
        Else
            MsgBox strErrMsg, vbInformation, gstrSysName
            strErrMsg = "": Exit Function
        End If
    End If
    CheckMobile = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckUsedMCNO(ByVal strMCNO As String) As Boolean
    '功能:检查医保号是否已存在
    Dim blnUsed As Boolean
    If mobjOneCard.zlCheckMCNOIsExist(strMCNO, blnUsed) = False Then Exit Function

    If blnUsed Then
        MsgBox "医保号" & strMCNO & "已存在，请检查。", vbInformation, gstrSysName
    End If
    CheckUsedMCNO = Not blnUsed
End Function

Public Function GetPatiInforMaxLen(ByVal str列名 As String) As Integer
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息中某一个字段的最大长度
    '返回:
    '编制:LNC
    '日期:2019-11-21 21:44:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllcolumn As Collection, cllTmp As Collection, i As Long
    
    On Error GoTo errHandle
    If gPatiMaxLen.intPatiName <> 0 Then
        Select Case str列名
            Case "姓名"
                GetPatiInforMaxLen = gPatiMaxLen.intPatiName
            Case "门诊号"
                GetPatiInforMaxLen = gPatiMaxLen.intPatiMzNo
            Case "性别"
                GetPatiInforMaxLen = gPatiMaxLen.intPatiSex
            Case "年龄"
                GetPatiInforMaxLen = gPatiMaxLen.intPatiAge
            Case "家庭地址"
                GetPatiInforMaxLen = gPatiMaxLen.intFamilyAddr
            Case "户口地址"
                GetPatiInforMaxLen = gPatiMaxLen.intRegistAddr
            Case "出生地点"
                GetPatiInforMaxLen = gPatiMaxLen.intBirthAddr
            Case "联系人地址"
                GetPatiInforMaxLen = gPatiMaxLen.intLinkAddr
        End Select
        Exit Function
    End If
    With gPatiMaxLen
        .intPatiName = 100
        .intPatiMzNo = 18
        .intPatiAge = 20
        .intPatiSex = 4
        .intFamilyAddr = 100
        .intRegistAddr = 100
        .intBirthAddr = 100
        .intLinkAddr = 100
    End With
    '重数据库中读取
    If mobjOneCard.zlGetInputItemLength("病人信息:姓名,门诊号,性别,年龄,家庭地址,户口地址,出生地点,联系人地址", cllcolumn) = False Then Exit Function
    For i = 1 To cllcolumn.Count
        Set cllTmp = cllcolumn(i)
        Select Case cllTmp("列名")
            Case "姓名"
                gPatiMaxLen.intPatiName = Val(cllTmp("列长度"))
                GetPatiInforMaxLen = gPatiMaxLen.intPatiName
            Case "门诊号"
                gPatiMaxLen.intPatiMzNo = Val(cllTmp("列长度"))
                GetPatiInforMaxLen = gPatiMaxLen.intPatiMzNo
            Case "性别"
                gPatiMaxLen.intPatiSex = Val(cllTmp("列长度"))
                GetPatiInforMaxLen = gPatiMaxLen.intPatiSex
            Case "年龄"
                gPatiMaxLen.intPatiAge = Val(cllTmp("列长度"))
                GetPatiInforMaxLen = gPatiMaxLen.intPatiAge
            Case "家庭地址"
                gPatiMaxLen.intFamilyAddr = Val(cllTmp("列长度"))
                GetPatiInforMaxLen = gPatiMaxLen.intFamilyAddr
            Case "户口地址"
                gPatiMaxLen.intRegistAddr = Val(cllTmp("列长度"))
                GetPatiInforMaxLen = gPatiMaxLen.intRegistAddr
            Case "出生地点"
                gPatiMaxLen.intBirthAddr = Val(cllTmp("列长度"))
                GetPatiInforMaxLen = gPatiMaxLen.intBirthAddr
            Case "联系人地址"
                gPatiMaxLen.intLinkAddr = Val(cllTmp("列长度"))
                GetPatiInforMaxLen = gPatiMaxLen.intLinkAddr
        End Select
    Next
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'=====================================================================================================
'结算类
'=====================================================================================================
Public Function zlReadCard(frmMain As Object, _
    ByVal lngCardTypeID As Long, _
    ByVal blnOlnyCardNO As Boolean, _
    ByVal strExpand As String, _
    ByRef strOutCardNO As String, _
    ByRef strOutPatiInforXML As String, _
    Optional ByRef strPhotoFile As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读卡接口
    '入参:frmMain-调用的父窗口
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       blnOlnyCardNO-仅仅读取卡号
    '       strExpand-扩展参数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlReadCard = mobjOneCard.zlReadCard(Me, lngCardTypeID, blnOlnyCardNO, strExpand, strOutCardNO, strOutPatiInforXML, strPhotoFile)
End Function

Public Function zlBandCardArfter(frmMain As Object, _
    ByVal lngCardTypeID As Long, _
    ByVal lng病人ID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绑定接口
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  卡类别ID
    '    lng病人ID   Long    IN   病人ID
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    True:调用成功,False:调用失败
    '调用者:
    '    1.  医疗卡发放管理中绑定卡后调用
    '    2.  病人入院登记;病人挂号管理;病人信息登记中进行卡绑定时调用
    '其他说明:
    '    相关的调用流程见后续各模块的绑定卡的详细说明
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBandCardArfter = mobjOneCard.zlBandCardArfter(frmMain, lngCardTypeID, lng病人ID, strExpend)
End Function

Public Function zlSetBrushCardObject(ByVal lngCardTypeID As Long, _
    ByRef objBrushCard As Object, ByRef strExpend As String, Optional ByVal bln消费卡 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡对象
    '入参:lngCardTypeID-医疗卡类别ID
    '    objBrushCard -指定的刷卡对象(文本框)
    '    strExpend-扩展信息(暂无,待以后扩展用)
    '出参:strExpend-扩展信息(暂无,待以后扩展用)
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-21 14:04:13
    '说明:主要针对刷卡方式的,比如:有些银行卡加密的,需要刷卡后进行解密
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSetBrushCardObject = mobjOneCard.zlSetBrushCardObject(lngCardTypeID, objBrushCard, strExpend, bln消费卡)
End Function

Public Sub zlInitEvents(ByVal Hwnd As Long, ByVal objComEvents As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置公共事件
    '入参:objComEvents-公共事件
    '编制:刘兴洪
    '日期:2012-08-22 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mobjOneCard.zlInitEvents(Hwnd, objComEvents)
End Sub

Public Function zlReturnCashCheck(frmMain As Object, _
                ByVal bytMode As Byte, ByVal objPayInfos As clsPayInfos, _
                ByVal rsBillAdvance As ADODB.Recordset, cllReturnCash As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查所有的三方卡是否支持退现（允许退现、缺省退现、退现方式）
    '入参:
    '出参:cllReturnCash-退现检查
    '返回:
    '编制:李南春
    '日期:2018/3/23 16:37:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim blnCheck As Boolean, bln允许退现 As Boolean, bln缺省退现 As Boolean, str退现方式 As String
    Dim strSQL As String, strBalanceIDs As String, strExpendXML As String, str支付方式 As String
    Dim strSwapNo As String, strSwapMemo As String, str卡号 As String, str三方退现 As String
    Dim dblThirdMoney As Double
    Dim lng结帐id As Long, lngCard结帐ID As Long, i As Long, lng卡类别id As Long, byt消费卡 As Byte
    Dim objCard As Card
    On Error GoTo ErrHand
    Set cllReturnCash = New Collection
    If objPayInfos Is Nothing Then zlReturnCashCheck = True: Exit Function
    If rsBillAdvance Is Nothing Then zlReturnCashCheck = True: Exit Function
    rsBillAdvance.Filter = "卡类别ID <> 0 And 消费卡 = 1 And 记录性质 <> 1 And 记录性质 <> 11"
    If rsBillAdvance.RecordCount > 0 Then
        With rsBillAdvance
            If .RecordCount = 0 Then zlReturnCashCheck = True: Exit Function
            Do While Not .EOF
                i = i + 1
                If lng卡类别id <> Val(Nvl(!卡类别ID)) And lng卡类别id <> 0 Then
                    If mobjOneCard.zlGetCard(lng卡类别id, True, objCard) = False Then
                        MsgBox str支付方式 & "支付接口未启用，不能继续结算。", vbInformation, gstrSysName
                        Exit Function
                    End If
                    str三方退现 = IIf(objCard.是否退现, 1, 0) & "," & IIf(objCard.是否缺省退现, 1, 0) & ","
                    Call cllReturnCash.Add(str三方退现, "X" & lng卡类别id)
                End If
                lng卡类别id = Val(Nvl(!卡类别ID)): str支付方式 = Nvl(!支付方式)
                .MoveNext
                
                If .EOF And lng卡类别id <> 0 Then '最后一列必须执行
                    If mobjOneCard.zlGetCard(lng卡类别id, True, objCard) = False Then
                        MsgBox str支付方式 & "支付接口未启用，不能继续结算。", vbInformation, gstrSysName
                        Exit Function
                    End If
                    str三方退现 = IIf(objCard.是否退现, 1, 0) & "," & IIf(objCard.是否缺省退现, 1, 0) & ","
                    Call cllReturnCash.Add(str三方退现, "X" & lng卡类别id)
                End If
            Loop
        End With
    End If
    
    rsBillAdvance.Filter = "卡类别ID <> 0 And 消费卡 = 0 And 记录性质 <> 1 And 记录性质 <> 11"
    If rsBillAdvance.RecordCount = 0 Then zlReturnCashCheck = True: Exit Function
        
    Do While Not rsBillAdvance.EOF
        If Not ((Val(Nvl(rsBillAdvance!校对标志)) = 2 Or Val(Nvl(rsBillAdvance!固定)) = 1) And bytMode = 6) Then
            blnCheck = True: Exit Do
        End If
        rsBillAdvance.MoveNext
    Loop
    If Not blnCheck Then zlReturnCashCheck = True: Exit Function
    
    rsBillAdvance.MoveFirst
    With rsBillAdvance
        If .RecordCount = 0 Then zlReturnCashCheck = True: Exit Function
        Do While Not .EOF
            '接口退费结算成功了，不再检查退现
            '接口支付结算失败了，不再检查退现
            i = i + 1
            If Not ((Val(Nvl(!校对标志)) = 2 Or Val(Nvl(!固定)) = 1) And bytMode = 6) Or bytMode = 5 Then
                If lng卡类别id <> Val(Nvl(!卡类别ID)) And lng卡类别id <> 0 Then
                    bln缺省退现 = False: str退现方式 = ""
                    strExpendXML = "<INPUT><TKLIST>" & strExpendXML & "</INPUT></TKLIST>"
                    '---------------------------------------------------------------------------------------------------------------------------------------------
                    '功能:退现交易检查
                    '入参:frmMain-调用的主窗体
                    '       lngModule-调用的模块号
                    '       lngCardTypeID-卡类别ID
                    '       strCardNo-卡号
                    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
                    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
                    '       dblMoney-退款金额
                    '       strSwapNo-交易流水号(退款时检查)多种结算方式时，本参数为第一个结算方式的交易流水号
                    '       strSwapMemo-交易说明(退款时传入) 多种结算方式时，本参数为第一个结算方式的交易说明
                    '       strXMLExpend    XML IN  10.35.90后才支持
                    '        <INPUT>
                    '            <TKLIST>    //本次退款列表
                    '                <TK>
                    '                    <TKFS>退款方式</TKFS>
                    '                    <TKJE>退款金额</TKJE>
                    '                    <JYLSH>原交易流水号</JYLSH>
                    '                    <JYSM>原交易说明</JYSM>
                    '                </TK>
                    '                ....
                    '            </TKLIST>
                    '        </INPUT>
                    '出参:
                    '       blnDelDefaultCash_Out-是否缺省退现：接口返回true时有效，true时：表示缺省退成现金（缺省方式为:str缺省退现方式_Out返回值),否则缺省退回原卡，但允许操作员选择退为现金
                    '       strDefaultDelBalance_Out-缺省退现方式,比如：支票，现金等
                    '       strXMLExpend:10.35.90后才支持
                    '        <OUTPUT>
                    '            <SFQSTX>是否缺省退现<SFQSTX>//NUMBER 1 是否缺省退现: 1-缺省;0-不缺省，缺省退回原卡，但以许操作员操作退现
                    '            <QSTKFS>缺省退现退款方式</QSTKFS>//Varchar2 20 缺省退现退款方式即结算方式.名称
                    '                    1.不允许返回三卡方的结算方式
                    '                    2.应避免使用：医保类结算，一卡通本身的结算方式和消费卡的一些特殊结算方式。返回这类方式，将被禁使用这些方式
                    '        </OUTPUT>
                    '返回:退现合法,返回true,否则返回Flase
                    '编制:刘尔旋
                    '日期:2017-7-3
                    '说明:
                    '    1.使用模块：病人结帐管理
                    '    2.触发时机:在结帐列表中点“X"时，调用，如果返回true,则允许退现，否则禁止
                    '---------------------------------------------------------------------------------------------------------------------------------------------
                    bln允许退现 = mobjOneCard.zlReturnCashCheck(frmMain, mlngModule, lng卡类别id, str卡号, strBalanceIDs, dblThirdMoney, strSwapNo, strSwapMemo, strExpendXML, bln缺省退现, str退现方式)
                    str三方退现 = IIf(bln允许退现, 1, 0) & "," & IIf(bln缺省退现, 1, 0) & "," & str退现方式
                    Call cllReturnCash.Add(str三方退现, "K" & lng卡类别id)
                    strExpendXML = ""
                End If
                
                If Val(Nvl(!记录性质)) = 4 And lng结帐id = 0 Then lng结帐id = Val(Nvl(!结帐ID))
                If Val(Nvl(!记录性质)) = 5 And lngCard结帐ID = 0 Then lngCard结帐ID = Val(Nvl(!结帐ID))
                If lng结帐id <> 0 Then strBalanceIDs = "4|" & objPayInfos.Reg_结帐ID
                If lngCard结帐ID <> 0 Then strBalanceIDs = strBalanceIDs & IIf(strBalanceIDs = "", "", "||") & "5|" & objPayInfos.Card_结帐ID
                
                str卡号 = Nvl(!卡号)
                strSwapNo = Nvl(!交易流水号)
                strSwapMemo = Nvl(!交易说明)
                dblThirdMoney = dblThirdMoney + Val(Nvl(!金额))
                lng卡类别id = Val(Nvl(!卡类别ID))
                strExpendXML = strExpendXML & _
                                "<TK>" & vbCrLf & _
                                    "<TKFS>" & Nvl(!结算方式) & "</TKFS>" & vbCrLf & _
                                    "<TKJE>" & Val(Nvl(!金额)) & "</TKJE>" & vbCrLf & _
                                    "<JYLSH>" & Nvl(!交易流水号) & "</JYLSH>" & vbCrLf & _
                                    "<JYSM>" & Nvl(!交易说明) & "</JYSM>" & vbCrLf & _
                                "</TK>" & vbCrLf
            End If
            .MoveNext
            
            If .EOF And lng卡类别id <> 0 Then '最后一列必须执行
                bln缺省退现 = False: str退现方式 = ""
                strExpendXML = "<INPUT><TKLIST>" & strExpendXML & "</INPUT></TKLIST>"
                bln允许退现 = mobjOneCard.zlReturnCashCheck(frmMain, mlngModule, lng卡类别id, str卡号, strBalanceIDs, dblThirdMoney, strSwapNo, strSwapMemo, strExpendXML, bln缺省退现, str退现方式)
                str三方退现 = IIf(bln允许退现, 1, 0) & "," & IIf(bln缺省退现, 1, 0) & "," & str退现方式
                Call cllReturnCash.Add(str三方退现, "K" & lng卡类别id)
                strExpendXML = ""
            End If
        Loop
    End With
    zlReturnCashCheck = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
    
Public Function zlCheckCancelValied(ByVal frmMain As Object, _
                ByVal objRegInfo As clsRegEventInfor, ByVal lng挂号结帐ID As Long, _
                ByVal lng卡费结帐ID As Long, ByVal bytMode As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费时的数据有效性
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng卡类别id As Long, lng消费卡ID As Long
    Dim str验证卡号  As String, strXmlIn As String, str刷卡密码 As String
    Dim str卡号 As String, str交易流水号 As String, str交易说明 As String, str结算信息 As String
    Dim strExpendXML As String, strExpend As String, strNo As String, strErrMsg As String
    Dim cllSquareBalance As Collection, bln检查卡费 As Boolean
    Dim i As Integer, dbl结算金额 As Double
    Dim objThreePay As clsPayInfo, objPay As clsPayInfo, objPayReg As clsSubPayInfo
    Dim objCard As Card
    On Error GoTo errHandle
    If objRegInfo Is Nothing Then zlCheckCancelValied = True: Exit Function
    If objRegInfo.objPayInfos Is Nothing Then zlCheckCancelValied = True: Exit Function
    
    Set cllSquareBalance = New Collection
    strNo = objRegInfo.objPayInfos.Reg_单据号
    '4.3.3.2.6   zlReturnCheck:帐户回退交易前的检查
    'zlPaymentCheck帐户扣款交易检查
    '参数名  参数类型    入/出   备注
    'frmMain Object  In  调用的主窗体
    'lngModule   Long    In  模块号
    'lngCardTypeID   Long    In  卡类别ID:医疗卡类别.ID
    'strCardNo   String  IN  卡号
    'strBalanceIDs:格式:收费类型( 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款)|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    'dblMoney    Double  IN  退款金额
    'strSwapNo   String  In  交易流水号(退款时检查)
    'strSwapMemo String  In  交易说明(退款时传入)
    '    Boolean 函数返回    True:调用成功,False:调用失败
    '说明:
    '在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,以便控制死锁情况。
    With objRegInfo.objPayInfos
        If objRegInfo.Card_变动类型 = CP_退卡 And .Card_卡类别ID > 0 And _
            (.Card_校对标志 <> 2 And bytMode = 6 Or .Card_校对标志 <> 1 And bytMode = 5 Or bytMode = 4) And _
            Not .Card_结算成功 Then bln检查卡费 = True
        If lng挂号结帐ID <> 0 Then
            For Each objPay In objRegInfo.objPayInfos
                If objPay.接口序号 > 0 Then
                    If GetCard(objPay.接口序号, objPay.消费卡, objCard) = False Then
                        MsgBox objPay.名称 & "未启用，不能原样退回！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If objCard.启用 = False Then
                        MsgBox objPay.名称 & "未启用，不能原样退回！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                If objPay.接口序号 > 0 And (objPay.校对标志 <> 2 And bytMode = 6 Or objPay.校对标志 <> 1 And bytMode = 5 Or bytMode = 4) And Not objPay.结算成功 Then
                    str结算信息 = "4|" & lng挂号结帐ID
                    strExpend = GetRegThreeBalance(strNo, objPay)
                    dbl结算金额 = objPay.支付金额
                    If bln检查卡费 And objPay.接口序号 = .Card_卡类别ID And objPay.消费卡 = .Card_消费卡 Then
                        str结算信息 = str结算信息 & "||5|" & lng卡费结帐ID
                        strExpend = strExpend & GetCardThreeBalance(objRegInfo.objPayInfos)
                        dbl结算金额 = dbl结算金额 + objRegInfo.objPayInfos.Card_结算金额
                        bln检查卡费 = False
                    End If
                    If strExpend <> "" Then
                        strExpendXML = "<TFDATA>" & vbCrLf & _
                                            "<YCTF>" & IIf(bytMode = 6, 1, 0) & "</YCTF>" & vbCrLf & _
                                            "<TKLIST>" & strExpend & "</TKLIST>" & vbCrLf & _
                                        "</TFDATA>"
                    End If
                    If mobjOneCard.zlReturnCheck(frmMain, mlngModule, _
                        objPay.接口序号, objPay.消费卡, objPay.卡号, str结算信息, dbl结算金额, _
                        objPay.交易流水号, objPay.交易说明, strExpendXML) = False Then
                        Exit Function
                    End If
                    If objPay.消费卡 And gbln消费卡退费验卡 Or objPay.消费卡 = False And _
                        objCard.是否退款验卡 Then
                        If objPay.消费卡 Then
                            'Array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
                            cllSquareBalance.Add Array(objPay.接口序号, objPay.消费卡ID, 0, _
                                                objPay.卡号, "", "", False, dbl结算金额)
                        End If
                        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
                        If mobjOneCard.zlBrushCard(frmMain, _
                            mlngModule, Nothing, objPay.接口序号, objPay.消费卡, _
                            objRegInfo.objPati.姓名, objRegInfo.objPati.性别, objRegInfo.objPati.年龄, dbl结算金额, _
                            str卡号, str刷卡密码, True, True, False, False, _
                            cllSquareBalance, False, True, strXmlIn) = False Then _
                            Exit Function
                    End If
                End If
            Next
        End If
        
        If bln检查卡费 Then
            If .Card_卡类别ID > 0 Then
                If GetCard(.Card_卡类别ID, .Card_消费卡, objCard) = False Then
                    MsgBox objPay.名称 & "未启用，不能原样退回！", vbInformation, gstrSysName
                    Exit Function
                End If
                If objCard.启用 = False Then
                    MsgBox objPay.名称 & "未启用，不能原样退回！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            str结算信息 = "5|" & lng卡费结帐ID
            strExpend = GetCardThreeBalance(objRegInfo.objPayInfos)
            strExpendXML = "<TFDATA>" & vbCrLf & _
                                "<YCTF>" & IIf(bytMode = 6, 1, 0) & "</YCTF>" & vbCrLf & _
                                "<TKLIST>" & strExpend & "</TKLIST>" & vbCrLf & _
                           "</TFDATA>"
            If mobjOneCard.zlReturnCheck(frmMain, mlngModule, _
                .Card_卡类别ID, .Card_消费卡, .Card_支付卡号, str结算信息, .Card_结算金额, .Card_交易流水号, _
                .Card_交易说明, strExpendXML) = False Then
                Exit Function
            End If
            If .Card_消费卡 And gbln消费卡退费验卡 Or .Card_消费卡 = False And _
                .Card_CardPro.是否退款验卡 Then
                If .Card_消费卡 Then
                    Set cllSquareBalance = New Collection
                    'Array(卡类别ID,消费卡ID,刷卡金额, 卡号,密码,限制类别,是否密文,剩余未退金额)
                    cllSquareBalance.Add Array(.Card_卡类别ID, .Card_消费卡ID, 0, _
                                        .Card_支付卡号, "", "", False, .Card_结算金额)
                End If
                strXmlIn = "<IN><CZLX>2</CZLX></IN>"
                If mobjOneCard.zlBrushCard(frmMain, mlngModule, _
                    Nothing, .Card_卡类别ID, .Card_消费卡, objRegInfo.objPati.姓名, objRegInfo.objPati.性别, _
                    objRegInfo.objPati.年龄, .Card_结算金额, str卡号, str刷卡密码, True, True, False, _
                    False, cllSquareBalance, False, True, strXmlIn) = False Then _
                    Exit Function
            End If
        End If
    End With
    
    zlCheckCancelValied = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlPaymentMoney(ByVal frmMain As Object, ByVal lngModul As Long, _
                    ByVal objPayInfos As clsPayInfos, objThreePay As clsPayInfo, ByVal dbl卡费 As Double, _
                    ByRef cllPro As Collection, ByRef cllThreeSwap As Collection, _
                    ByRef intSwapStyle As Integer, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllPro-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String, strValue As String
    Dim lngCardFeeID As Long, strCardNo As String, dblCardMoney As Double, dbl挂号费 As Double
    Dim strBalanceIDs As String
    Dim strSQL As String, rsCard As ADODB.Recordset
    Dim i As Integer, lng预交ID As Long
    Dim objThreePayInfo As clsPayInfo, objPay As clsPayInfo
    Dim lng卡类别id  As Long, bln消费卡 As Boolean, str支付卡号 As String
    
    On Error GoTo errHandle
    If objPayInfos Is Nothing Then zlPaymentMoney = True: Exit Function
    If objThreePay Is Nothing Then zlPaymentMoney = True: Exit Function
    If (objThreePay.支付金额 = 0 And dbl卡费 = 0) Or objThreePay.校对标志 = 2 Then zlPaymentMoney = True: Exit Function
    
    dbl挂号费 = objThreePay.支付金额 - dbl卡费
    If RoundEx(dbl挂号费, 6) <> 0 Then strBalanceIDs = objPayInfos.Reg_结帐ID
    If RoundEx(dbl卡费, 6) <> 0 Then strBalanceIDs = strBalanceIDs & IIf(strBalanceIDs = "", "", ",") & objPayInfos.Card_结帐ID
    
'    If cbo结算方式.ItemData(cbo结算方式.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If RoundEx(dbl卡费, 6) <> 0 Then
        strSQL = "Select 关联交易ID,No,冲预交 From 病人预交记录 where 结帐ID = [1]"
        Set rsCard = zldatabase.OpenSQLRecord(strSQL, "卡费信息", objPayInfos.Card_结帐ID)
        If Not rsCard.EOF Then
            lngCardFeeID = Val(Nvl(rsCard!关联交易ID))
            dblCardMoney = Val(Nvl(rsCard!冲预交))
            strCardNo = Nvl(rsCard!NO)
        End If
    End If

    strSwapExtendInfor = "<INPUT>" & vbCrLf & _
                            "<QRCODE>" & objThreePay.QRCode & "</QRCODE>" & vbCrLf & _
                         "</INPUT>"
    If mobjOneCard.zlPaymentMoney(frmMain, mlngModule, objThreePay.接口序号, objThreePay.消费卡, objThreePay.卡号, strBalanceIDs, "", objThreePay.支付金额, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        If Not objThreePay.消费卡 Then
            If zlXML_Init("OUTPUT") Then
                If zlXML_LoadXMLToDOMDocument(strSwapExtendInfor, False) Then
                    Call zlXML_GetNodeValue("JYZT", , strValue)
                    intSwapStyle = Val(strValue)
                    Call zlXML_GetNodeValue("ERRMSG", , strValue)
                    strErrMsg = strValue
                End If
            End If
        End If
        Exit Function
    End If
    '更新三交交易数据
    '必须先更新挂号费，再更新卡费，将对支付的金额累计到挂号费上进行找补
    Call objThreePay.ReMoveAll
    If RoundEx(dbl挂号费, 6) <> 0 Then
        '问题:58322
        str支付卡号 = objThreePay.卡号
        If zlAddUpdateSwapSQL(False, objPayInfos.Reg_结帐ID, objThreePay, str支付卡号, strSwapGlideNO, strSwapMemo, cllPro, 2, 4, dbl挂号费, objThreePay.关联交易ID, strSwapExtendInfor, objPayInfos.Reg_单据号, , strErrMsg, dbl卡费) = False Then
            Call objThreePay.ReMoveAll: Exit Function
        End If
        Call zlAddThreeSwapSQLToCollection(False, objPayInfos.Reg_结帐ID, objThreePay.接口序号, objThreePay.消费卡, str支付卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    If RoundEx(dbl卡费, 6) <> 0 Then
        str支付卡号 = objThreePay.卡号
        If zlAddUpdateSwapSQL(False, objPayInfos.Card_结帐ID, objThreePay, str支付卡号, strSwapGlideNO, strSwapMemo, cllPro, 2, 5, dbl卡费, lngCardFeeID, strSwapExtendInfor, objPayInfos.Card_单据号, , strErrMsg, dbl挂号费) = False Then
            Call objThreePay.ReMoveAll: Exit Function
        End If
        Call zlAddThreeSwapSQLToCollection(False, objPayInfos.Card_结帐ID, objThreePay.接口序号, objThreePay.消费卡, str支付卡号, strSwapExtendInfor, cllThreeSwap)
        '更新发卡结算信息,当strAdvance = "" 时只更新校对标志 = 2
    End If
    
    If RoundEx(objThreePay.支付金额, 6) <> RoundEx(dbl挂号费 + dbl卡费, 6) Then
        objThreePay.支付金额 = dbl挂号费 + dbl卡费
    End If
    zlPaymentMoney = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function

Public Function zlReturnMoney(ByVal frmMain As Object, _
                            ByVal bytMode As Byte, ByRef objRegInfor As clsRegEventInfor, _
                            ByVal lng挂号结帐ID As Long, ByVal lng卡费结帐ID As Long, _
                            ByVal lng挂号冲销ID As Long, ByVal lng退卡冲销ID As Long, _
                            ByVal blnYBCommit As Boolean, ByRef blnTrans As Boolean, ByRef blnAbNormal As Boolean, _
                            ByRef intSwapStyle As Integer, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用回退接口,三方卡调用成功后立即执行并提交
    '入参:
    '返回:成功,返回true,否则返回False
    '编制:李南春
    '日期:2019/2/1 11:26:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, i As Long
    Dim bln结算卡费 As Boolean, blnFirst As Boolean '第一笔三方卡交易
    Dim bln消费卡 As Boolean, lng卡类别id As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim str卡号 As String, strSwapGlideNO As String, strSwapMemo As String, str结算信息 As String, strSwapExtendInfor As String
    Dim strTemp As String, strExpend As String, strExpendXML As String, strValue As String
    Dim lng关联ID As Long, lngCard关联ID As Long
    Dim dblCardMoney As Double, dbl结算金额 As Double
    Dim objPayInfos As clsPayInfos, objPay As clsPayInfo, objSubPay As clsSubPayInfo
    Dim cllUpdate As Collection, cllThreeSwap As Collection, cllBack As Collection, cllPro As Collection
    
    '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO
    'cllBalance.Add Array(Val(Nvl(rsTmp!卡类别ID)), Trim(Nvl(rsTmp!卡号)), IIf(Val(Nvl(rsTmp!结算卡序号)) <> 0, 1, 0), Trim(Nvl(rsTmp!交易流水号)), Trim(Nvl(rsTmp!交易说明))), strNO
    Set objPayInfos = objRegInfor.objPayInfos
    If objPayInfos Is Nothing Then zlReturnMoney = True: Exit Function
    
    On Error GoTo ErrHand
    If objRegInfor.Card_变动类型 = CP_退卡 And objPayInfos.Card_卡类别ID > 0 And _
        (objPayInfos.Card_校对标志 <> 2 And bytMode = 6 Or objPayInfos.Card_校对标志 <> 1 And bytMode = 5 Or bytMode = 4) Then bln结算卡费 = True
        
    If (bln结算卡费 And Not objPayInfos.Card_消费卡 Or objPayInfos.Reg_Hav三方卡) And gblnAsyncCharge Then
        gcnOracle.CommitTrans
        blnAbNormal = True
        gcnOracle.BeginTrans
    End If
    blnFirst = True
    '可能会退多笔三方卡交易，因此成功一笔提交一笔
    For Each objPay In objPayInfos
        If objPay.接口序号 > 0 And Not objPay.消费卡 And _
            (objPay.校对标志 <> 2 And bytMode = 6 Or objPay.校对标志 <> 1 And bytMode = 5 Or bytMode = 4) Then
            str结算信息 = "4|" & lng挂号结帐ID
            strSwapExtendInfor = "4|" & lng挂号冲销ID: strTemp = ""
            strExpend = GetRegThreeBalance(objPayInfos.Reg_单据号, objPay)
            dbl结算金额 = objPay.支付金额
            '三方卡同时支付卡费
            If bln结算卡费 And objPay.接口序号 = objPayInfos.Card_卡类别ID And objPay.消费卡 = objPayInfos.Card_消费卡 Then
                str结算信息 = str结算信息 & "||5|" & lng卡费结帐ID
                strSwapExtendInfor = strSwapExtendInfor & "||5|" & lng退卡冲销ID
                strExpend = strExpend & GetCardThreeBalance(objPayInfos)
                dbl结算金额 = dbl结算金额 + objPayInfos.Card_结算金额
            End If
            strTemp = strSwapExtendInfor
            If strExpend <> "" Then
                strExpendXML = "<INPUT><TKLIST>" & strExpend & "</TKLIST></INPUT>"
            End If
            strSwapGlideNO = objPay.交易流水号
            strSwapMemo = objPay.交易说明
            str卡号 = objPay.卡号

            'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, ByVal dblMoney As Double, ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, ByRef strSwapExtendInfor As String, ByRef strExpendXML as string) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:帐户扣款回退交易
            '入参:frmMain-调用的主窗体
            '       lngModule-调用的模块号
            '       lngCardTypeID-卡类别ID:医疗卡类别.ID
            '       strCardNo-卡号
            '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
            '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
            '       dblMoney-退款金额
            '       strSwapNo-交易流水号(扣款时的交易流水号)
            '       strSwapMemo-交易说明(扣款时的交易说明)
            '       strSwapExtendInfor-传入，本次退费的冲销ID：
            '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
            '       strSwapExtendInfor-传出，交易的扩展信息
            '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
            If mobjOneCard.zlReturnMoney(frmMain, mlngModule, objPay.接口序号, objPay.消费卡, str卡号, _
                str结算信息, dbl结算金额, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strExpendXML) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                If Not objPay.消费卡 And gblnAsyncCharge Then
                    If zlXML_Init("OUTPUT") Then
                        If zlXML_LoadXMLToDOMDocument(strExpendXML, False) Then
                            Call zlXML_GetNodeValue("JYZT", , strValue)
                            intSwapStyle = Val(strValue)
                            Call zlXML_GetNodeValue("ERRMSG", , strValue)
                            strErrMsg = strValue
                        End If
                    End If
                    If bytMode <> 6 And intSwapStyle <> 1 Then GoTo ErrBack
                End If
                Exit Function
            End If
            
            On Error GoTo ErrExcPro:
            Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
            If zlAddUpdateSwapSQL(False, lng挂号冲销ID, objPay, _
                str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, IIf(gblnAsyncCharge, 2, 0), _
                4, objPay.支付金额, objPay.关联交易ID, strExpendXML, objPayInfos.Reg_单据号, True, _
                strErrMsg) = False Then Exit Function
            If strTemp <> strSwapExtendInfor Then
                Call zlAddThreeSwapSQLToCollection(False, lng挂号冲销ID, objPay.接口序号, _
                    objPay.消费卡, str卡号, strSwapExtendInfor, cllThreeSwap)
            End If
                
            If bln结算卡费 And objPay.接口序号 = objPayInfos.Card_卡类别ID And objPay.消费卡 = objPayInfos.Card_消费卡 Then '三方支付中有卡费
                If zlAddUpdateSwapSQL(False, lng退卡冲销ID, objPay, _
                    str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, IIf(gblnAsyncCharge, 2, 0), _
                    5, objPayInfos.Card_结算金额, objPayInfos.Card_关联交易ID, _
                    strExpendXML, objPayInfos.Card_单据号, True, strErrMsg) = False Then
                    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
                    blnFirst = False: intSwapStyle = 1  '结算数据异常，需要强制退出
                    GoTo ErrBack
                End If
                If strTemp <> strSwapExtendInfor Then
                    Call zlAddThreeSwapSQLToCollection(False, lng退卡冲销ID, _
                         objPay.接口序号, objPay.消费卡, str卡号, strSwapExtendInfor, cllThreeSwap)
                End If
                bln结算卡费 = False
            End If
            If Not blnTrans Then gcnOracle.BeginTrans: blnTrans = True
            zlExecuteProcedureArrAy cllUpdate, "zlReturnMoney", True, True
            On Error Resume Next
            zlExecuteProcedureArrAy cllThreeSwap, "zlReturnMoney", True, True
            On Error GoTo ErrHand
            gcnOracle.CommitTrans: blnTrans = False
            blnFirst = False: blnAbNormal = True
            objPay.校对标志 = IIf(bytMode = 5, 1, 2)
        End If
    Next
    
    If bln结算卡费 Then
        str结算信息 = "5|" & lng卡费结帐ID
        strSwapExtendInfor = "5|" & lng退卡冲销ID
        strTemp = strSwapExtendInfor
        strExpend = GetCardThreeBalance(objPayInfos)
        If strExpend <> "" Then
            strExpendXML = "<INPUT><TKLIST>" & strExpend & "</TKLIST></INPUT>"
        End If
        strSwapGlideNO = objPayInfos.Card_交易流水号
        strSwapMemo = objPayInfos.Card_交易说明
        str卡号 = objPayInfos.Card_支付卡号

        'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, ByVal dblMoney As Double, ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, ByRef strSwapExtendInfor As String, ByRef strExpendXML as string) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:帐户扣款回退交易
        '入参:frmMain-调用的主窗体
        '       lngModule-调用的模块号
        '       lngCardTypeID-卡类别ID:医疗卡类别.ID
        '       strCardNo-卡号
        '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
        '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        '       dblMoney-退款金额
        '       strSwapNo-交易流水号(扣款时的交易流水号)
        '       strSwapMemo-交易说明(扣款时的交易说明)
        '       strSwapExtendInfor-传入，本次退费的冲销ID：
        '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
        '       strSwapExtendInfor-传出，交易的扩展信息
        '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
        If mobjOneCard.zlReturnMoney(frmMain, mlngModule, objPayInfos.Card_卡类别ID, _
            objPayInfos.Card_消费卡, str卡号, str结算信息, objPayInfos.Card_结算金额, _
            strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strExpendXML) = False Then
            If Not objPay.消费卡 And gblnAsyncCharge Then
                If zlXML_Init("OUTPUT") Then
                    If zlXML_LoadXMLToDOMDocument(strExpendXML, False) Then
                        Call zlXML_GetNodeValue("JYZT", , strValue)
                        intSwapStyle = Val(strValue)
                        Call zlXML_GetNodeValue("ERRMSG", , strValue)
                        strErrMsg = strValue
                    End If
                End If
                If bytMode <> 6 And intSwapStyle <> 1 Then
                    Call zlGetCancelSql(lng挂号冲销ID, IIf(blnYBCommit, 0, lng退卡冲销ID), cllBack)
                    zlExecuteProcedureArrAy cllBack, "zlReturnMoney", False, False
                    If blnFirst Then blnAbNormal = False
                End If
            End If
            Exit Function
        End If
        
        On Error GoTo ErrExcPro:
        Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
        Set objPay = New clsPayInfo
        objPay.接口序号 = objPayInfos.Card_卡类别ID
        objPay.消费卡 = objPayInfos.Card_消费卡
        If zlAddUpdateSwapSQL(False, lng退卡冲销ID, objPay, _
            str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, IIf(gblnAsyncCharge, 2, 0), _
            5, objPayInfos.Card_结算金额, Val(objPayInfos.Card_关联交易ID), _
            strExpendXML, objPayInfos.Card_单据号, True, strErrMsg) = False Then
            If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
            blnFirst = False: intSwapStyle = 1  '结算数据异常，需要强制退出
            GoTo ErrBack
        End If
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng退卡冲销ID, _
                 objPayInfos.Card_卡类别ID, objPayInfos.Card_消费卡, str卡号, strSwapExtendInfor, cllThreeSwap)
        End If
        If Not blnTrans Then gcnOracle.BeginTrans: blnTrans = True
        zlExecuteProcedureArrAy cllUpdate, "zlReturnMoney", True, True
        On Error Resume Next
        zlExecuteProcedureArrAy cllThreeSwap, "zlReturnMoney", True, True
        On Error GoTo ErrHand
        gcnOracle.CommitTrans: blnTrans = False
        objPayInfos.Card_校对标志 = IIf(bytMode = 5, 1, 2)
    End If
    zlReturnMoney = True
    Exit Function
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
ErrExcPro:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    strErrMsg = "接口支付成功但HIS保存数据失败，原因：" & Err.Description
    Exit Function
ErrBack:
    If Not objPay.消费卡 And gblnAsyncCharge Then
        Call zlGetCancelSql(lng挂号冲销ID, lng退卡冲销ID, cllBack, , , blnYBCommit Or Not blnFirst Or bytMode = 6)
        zlExecuteProcedureArrAy cllBack, "zlReturnMoney", False, False
        If blnFirst Then blnAbNormal = False
    End If
End Function

Public Function zlCheckBackCard(ByVal frmMain As Form, ByVal objService As zlPublicExpense.clsService, _
                                ByVal bytCheck As Byte, objRegInfor As clsRegEventInfor, _
                                ByVal bln允许退现 As Boolean, ByVal bln强制退现 As Boolean, _
                                ByVal strCardPrivs As String, ByRef str强制退现操作员 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退号时的退卡操作是否合法
    '入参:bytCheck:1-读取费用时检查是否允许同时退卡费；2-确认退款方式后，检查是否同时退卡费
    '     objRegInfor - 挂号数据
    '     bln允许退现 - 支付卡费的三方卡是否支持退现
    '     bln强制退现 - 不支持退现时是否允许强制退现，允许强制时要检查权限
    '     strCardPrivs - 消费卡操作权限
    '出参:str强制退现操作员 - 强制退现的操作员
    '返回:
    '编制:李南春
    '日期:2019/2/1 17:15:36
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objPay As clsPayInfo, objCard As Card
    Dim strErrMsg As String
    On Error GoTo ErrHand
    If objRegInfor Is Nothing Then Exit Function
    If objService Is Nothing Then Exit Function
    With objRegInfor.objPayInfos
        Select Case bytCheck
            Case 1
                '需要卡费和挂号费同时是用三方卡支付金进行此检查
                If .Card_单据号 <> "" Then
                    '113613：李南春，2018/1/18，退卡时检查当前卡是否允许退卡
                    If objService.zlPatiSvr_delcard_check(2, objRegInfor.objPati.病人ID, objRegInfor.Card_卡类别ID, objRegInfor.Card_卡号, True, strErrMsg) = False Then
                        If Not .Card_CardPro Is Nothing And .Card_卡类别ID <> 0 Then
                            If .Card_CardPro.是否全退 And bln允许退现 = False Then
                                str强制退现操作员 = ""
                                If strErrMsg = "" Then strErrMsg = "卡号" & objRegInfor.Card_卡号 & "不允许退卡"
                                If bln强制退现 Then '允许强制退现
                                    If InStr(strCardPrivs, ";三方退款强制退现;") = 0 Then
                                        str强制退现操作员 = zldatabase.UserIdentifyByUser(frmMain, "强制退现验证", glngSys, 1151, "三方退款强制退现")
                                    Else
                                        str强制退现操作员 = UserInfo.姓名
                                    End If
                                End If
                                
                                If str强制退现操作员 <> "" Then
                                    If MsgBox(strErrMsg & "，只能以强制退现方式退挂号费，是否退号？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                        Exit Function
                                    End If
                                Else
                                    If MsgBox(strErrMsg & "，不能单独退挂号费！", vbInformation + vbOK, gstrSysName) = vbNo Then
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                        
                        If str强制退现操作员 = "" Then
                            If MsgBox(strErrMsg & "，是否单独退挂号费？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                Exit Function
                            End If
                        End If
                        objRegInfor.Card_变动类型 = CP_仅退号
                    Else
                        If MsgBox("该病人挂号时发过卡,退号同时退卡吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            objRegInfor.Card_变动类型 = CP_仅退号
                        Else
                            objRegInfor.Card_变动类型 = CP_退卡
                        End If
                    End If
                End If
            Case 2
                '目前卡费只能原样退，三方卡不支持部分退的情况下，挂号又强制退现，会导致后续不能退卡
                '三方卡同时支付的卡费和挂号费,但又不支持部分退，只能退现
                If objRegInfor.Card_变动类型 = CP_仅退号 And .Card_卡类别ID > 0 Then
                    On Error Resume Next
                    Set objPay = objRegInfor.objPayInfos(IIf(.Card_消费卡, "X", "K") & .Card_卡类别ID)
                    On Error GoTo ErrHand
                    If .Card_CardPro.是否全退 And Not objPay Is Nothing Then
                        MsgBox .Card_CardPro.名称 & "不支持部分退，不能单独退挂号费！", vbInformation + vbOK, gstrSysName
                        Exit Function
                    End If
                ElseIf objRegInfor.Card_变动类型 = CP_退卡 Then
                    If mobjOneCard.zlGetCard(objRegInfor.Card_卡类别ID, False, objCard) Then
                        If Not objCard.自制卡 Then
                            If zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "卡号:" & .Card_单据号 & "卡为院外卡发,请选择退卡或取消绑定操作", "退卡,取消绑定", frmMain, vbQuestion) = "退卡" Then
                                objRegInfor.Card_变动类型 = CP_退卡
                            Else
                                objRegInfor.Card_变动类型 = CP_解绑
                            End If
                        End If
                    End If
                End If
        End Select
    End With
    zlCheckBackCard = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCardThreeBalance(ByVal objPayInfos As clsPayInfos) As String
    Dim strExpend As String
    If objPayInfos Is Nothing Then Exit Function
    With objPayInfos
        strExpend = "<TK>" & vbCrLf & _
                        "<TKFS>" & .Card_结算方式 & "</TKFS>" & vbCrLf & _
                        "<TKJE>" & .Card_结算金额 & "</TKJE>" & vbCrLf & _
                        "<JYLSH>" & .Card_交易流水号 & "</JYLSH>" & vbCrLf & _
                        "<JYSM>" & .Card_交易说明 & "</JYSM>" & vbCrLf & _
                        "<DJH>" & .Card_单据号 & "</DJH>" & vbCrLf & _
                        "<DJXZ>" & 5 & "</DJXZ>" & vbCrLf & _
                    "</TK>"
    End With
    GetCardThreeBalance = strExpend
End Function

Private Function GetRegThreeBalance(ByVal strNo As String, ByVal objPay As clsPayInfo) As String
    Dim strExpend As String
    Dim objPayReg As clsSubPayInfo
    If objPay Is Nothing Then Exit Function
    
    If objPay.Count > 1 Then
        For Each objPayReg In objPay
            With objPayReg
                strExpend = strExpend & _
                            "<TK>" & vbCrLf & _
                                "<TKFS>" & .结算方式 & "</TKFS>" & vbCrLf & _
                                "<TKJE>" & .结算金额 & "</TKJE>" & vbCrLf & _
                                "<JYLSH>" & .交易流水号 & "</JYLSH>" & vbCrLf & _
                                "<JYSM>" & .交易说明 & "</JYSM>" & vbCrLf & _
                                "<DJH>" & strNo & "</DJH>" & vbCrLf & _
                                "<DJXZ>" & 4 & "</DJXZ>" & vbCrLf & _
                            "</TK>"
            End With
        Next
    Else
        strExpend = "<TK>" & vbCrLf & _
                        "<TKFS>" & objPay.结算方式 & "</TKFS>" & vbCrLf & _
                        "<TKJE>" & objPay.支付金额 & "</TKJE>" & vbCrLf & _
                        "<JYLSH>" & objPay.交易流水号 & "</JYLSH>" & vbCrLf & _
                        "<JYSM>" & objPay.交易说明 & "</JYSM>" & vbCrLf & _
                        "<DJH>" & strNo & "</DJH>" & vbCrLf & _
                        "<DJXZ>" & 4 & "</DJXZ>" & vbCrLf & _
                    "</TK>"
    End If
    GetRegThreeBalance = strExpend
End Function

Public Function zlBrushCard(ByVal frmMain As Object, ByVal lngModul As Long, _
                    ByVal rsMoney As ADODB.Recordset, objRegInfo As clsRegEventInfor, _
                    objThreePay As clsPayInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '参数：dblMoney：IN/OUT,传入支付录入的金额，返回接口实际支付的金额
    '     str结算方式：OUT,接口实际选择的支付方式
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strXmlIn As String, str卡号 As String, str密码 As String
    Dim dblMoney As Double, dbl结算金额_OUT As Double, dbl挂号费 As Double, dbl卡费 As Double
    Dim bln消费卡 As Boolean
    Dim str结算方式_OUT As String
    
    On Error GoTo errHandle
    '68991
    If mobjOneCard Is Nothing Then Exit Function
    If objRegInfo Is Nothing Then Exit Function
    If objThreePay Is Nothing Then zlBrushCard = True: Exit Function
    If objThreePay.支付金额 = 0 Then zlBrushCard = True: Exit Function
    If objThreePay.结算成功 Then zlBrushCard = True: Exit Function
    dblMoney = objThreePay.支付金额
    
    If objThreePay.接口序号 = 0 Then
        MsgBox objThreePay.名称 & "异常,请检查!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln消费卡 As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl金额 As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln退现 As Boolean = False, _
    Optional ByVal bln余额不足禁止 As Boolean = True, _
    Optional ByRef varSquareBalance As Variant, _
    Optional ByVal bln转预交 As Boolean = False, _
    Optional ByVal blnAllPay As Boolean = False, _
    Optional ByVal strXmlIn As String = "", _
    Optional ByVal str费用来源 As String, _
    Optional ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定支付类别,弹出刷卡窗口
    '入参:rsClassMoney:收费类别,金额
    '        lngCardTypeID-为零时,为老一卡通刷卡
    '       bln余额不足禁止-目前只针对消费卡,表示余额不足时,禁止继续操作,否则用余额进行支付
    '       str费用来源 - 当前支付费用的费用来源，多种用逗号分隔(使用消费卡支付时传入)
    '       lng病人ID - 病人ID(使用消费卡支付时传入)
   '58322
   'NEXTTODO:消费卡支持多张
    If objThreePay.校对标志 <> 2 Then
         strXmlIn = "<IN>" & vbCrLf & _
                        "<CZLX>0</CZLX>" & vbCrLf & _
                        "<QRCODE>" & objThreePay.QRCode & "</QRCODE>" & vbCrLf & _
                    "</IN>"
        If mobjOneCard.zlBrushCard(frmMain, lngModul, rsMoney, objThreePay.接口序号, objThreePay.消费卡, _
            objRegInfo.objPati.姓名, objRegInfo.objPati.性别, objRegInfo.objPati.年龄, dblMoney, str卡号, str密码, _
            False, True, False, True, Nothing, False, True, strXmlIn, "1", objRegInfo.objPati.病人ID, str结算方式_OUT, dbl结算金额_OUT) = False Then Exit Function
        
        '如果接口返回的金额小于支付，不允许继续挂号
        If RoundEx(dblMoney, 6) <> RoundEx(dbl结算金额_OUT, 6) And Not objThreePay.消费卡 Then
            MsgBox objThreePay.名称 & "支付金额为" & Format(dbl结算金额_OUT, "0.00") & "元，与本次缴款金额 " & _
                    Format(dblMoney, "0.00") & "元不一致，不能继续支付。", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        str卡号 = objThreePay.卡号
        str密码 = objThreePay.支付密码
    End If
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
     strXmlIn = "<IN>" & vbCrLf & _
                    "<QRCODE>" & objThreePay.QRCode & "</QRCODE>" & vbCrLf & _
                "</IN>"
    If mobjOneCard.zlPaymentCheck(frmMain, lngModul, objThreePay.接口序号, objThreePay.消费卡, _
        str卡号, dblMoney, "", strXmlIn) = False Then Exit Function
    
    objThreePay.卡号 = str卡号
    objThreePay.支付密码 = str密码
    If str结算方式_OUT <> "" Then objThreePay.结算方式 = str结算方式_OUT

    zlBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlSwapIsSucces(ByVal frmMain As Object, ByVal lngModule As Long, _
                        ByVal bytMode As Byte, ByVal bytInState As Byte, _
                        ByVal objRegInfor As clsRegEventInfor, _
                        ByVal rsBillAdvance As ADODB.Recordset, _
                        ByVal vsfPay As VSFlexGrid) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查结算信息中所有三方卡的实际结算情况
    '入参:intSwapType-是否重退检查
    '     objPayInfos - 结算汇总信息
    '     rsBillAdvance - 交易明细数据
    '     intBusiness
    '编制:李南春
    '日期:2019/1/30 15:19:37
    '说明:退号时不退卡费，需要将卡费结算剔除，前提是允许部分退
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllThreeUp As Collection, cllThreeSwap As Collection
    Dim objThreePay As clsPayInfo, objPayInfos As clsPayInfos
    Dim blnSwapSuc As Boolean, blnHavNoSuccess As Boolean, bln重退 As Boolean, bln检查卡费 As Boolean
    Dim blnInVsf As Boolean
    Dim intBusiness As Integer, intSwapType As Integer, i As Integer
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strBalanceIDs As String, strXMLExpend As String, strFilter As String, strErrMsg As String
    
    If mobjOneCard Is Nothing Then Exit Function
    If objRegInfor Is Nothing Then Exit Function
    Set objPayInfos = objRegInfor.objPayInfos
    If objPayInfos Is Nothing Then Exit Function
    If rsBillAdvance Is Nothing Then Exit Function

    intSwapType = IIf(bytMode = 6, 1, 0)
    rsBillAdvance.Filter = "卡类别ID > 0 And 消费卡 = 0 And 固定 = 0 And 校对标志 = 1"
    rsBillAdvance.Sort = "卡类别ID"
    If rsBillAdvance.RecordCount = 0 Then Exit Function
    
    strBalanceIDs = "4|" & objPayInfos.Reg_结帐ID
    If objPayInfos.Card_结帐ID <> 0 Then strBalanceIDs = strBalanceIDs & "||5|" & objPayInfos.Card_结帐ID
    
    On Error GoTo ErrHand
    '更新三方交易信息
    '已rsBillAdvance进行检查，不区分卡费和挂号费
    Set objThreePay = New clsPayInfo
    rsBillAdvance.MoveFirst
    Do While Not rsBillAdvance.EOF
        If objThreePay.接口序号 <> Val(Nvl(rsBillAdvance!卡类别ID)) And objThreePay.接口序号 <> 0 Then
            blnSwapSuc = False: blnInVsf = False: intBusiness = 0
            With vsfPay
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("卡类别ID"))) = objThreePay.接口序号 And Val(.TextMatrix(i, .ColIndex("消费卡"))) = 0 Then
                        If Val(.Cell(flexcpData, i, .ColIndex("校对标志"))) = 1 Then
                            blnSwapSuc = True
                        End If
                        blnInVsf = True: Exit For
                    End If
                Next
            End With
            If blnInVsf = True And blnSwapSuc = False Then
                '    zlSwapIsSucces(ByVal frmMain As Object, ByVal lngModule As Long, ByVal intSwapType As Integer, _
                '    ByVal lngCardTypeID As Long, ByVal strBalanceIDs As String, ByVal strCardNo As String, ByVal dblSwapMoney As Double, _
                '    ByRef intSwapStatus_Out As Integer, Optional strErrMsg As String, Optional strXMLExpend As String) As Boolean
                blnSwapSuc = mobjOneCard.zlSwapIsSucces(frmMain, lngModule, intSwapType, objThreePay.接口序号, _
                    strBalanceIDs, objThreePay.卡号, objThreePay.支付金额, intBusiness, strErrMsg, strXMLExpend)
            End If
            If blnSwapSuc Then
                strFilter = strFilter & IIf(strFilter <> "", " Or ", "") & "卡类别ID = " & objThreePay.接口序号 & " And 消费卡 = 0 And 校对标志 = 1)"
                If blnHavNoSuccess = False Then blnHavNoSuccess = True
            Else
                If intBusiness = 1 Then
                    If strErrMsg = "" Then strErrMsg = "交易正在处理中, 请稍后再试！"
                    MsgBox strErrMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            Set objThreePay = New clsPayInfo
        End If
        
        objThreePay.支付类型 = Pay_ThreePay
        objThreePay.名称 = Nvl(rsBillAdvance!支付方式)
        objThreePay.支付金额 = Val(Nvl(rsBillAdvance!金额)) + objThreePay.支付金额
        objThreePay.结算方式 = Nvl(rsBillAdvance!结算方式)
        objThreePay.结算性质 = Val(Nvl(rsBillAdvance!结算性质))
        objThreePay.接口序号 = Val(Nvl(rsBillAdvance!卡类别ID))
        objThreePay.消费卡 = False
        If Val(Nvl(rsBillAdvance!记录性质)) = 4 Or objThreePay.交易流水号 = "" Then '优先以挂号的为准
            objThreePay.卡号 = Nvl(rsBillAdvance!卡号)
            objThreePay.校对标志 = Val(Nvl(rsBillAdvance!校对标志))
            objThreePay.交易流水号 = Nvl(rsBillAdvance!交易流水号)
            objThreePay.交易说明 = Nvl(rsBillAdvance!交易说明)
            objThreePay.关联交易ID = Val(Nvl(rsBillAdvance!关联交易ID))
        End If
        rsBillAdvance.MoveNext
        If rsBillAdvance.EOF Then
            blnSwapSuc = False: blnInVsf = False: intBusiness = 0
            With vsfPay
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("卡类别ID"))) = objThreePay.接口序号 And Val(.TextMatrix(i, .ColIndex("消费卡"))) = 0 Then
                        If Val(.Cell(flexcpData, i, .ColIndex("校对标志"))) = 1 Then
                            blnSwapSuc = True
                        End If
                        blnInVsf = True: Exit For
                    End If
                Next
            End With
            If blnInVsf = True And blnSwapSuc = False Then
                '    zlSwapIsSucces(ByVal frmMain As Object, ByVal lngModule As Long, ByVal intSwapType As Integer, _
                '    ByVal lngCardTypeID As Long, ByVal strBalanceIDs As String, ByVal strCardNo As String, ByVal dblSwapMoney As Double, _
                '    ByRef intSwapStatus_Out As Integer, Optional strErrMsg As String, Optional strXMLExpend As String) As Boolean
                blnSwapSuc = mobjOneCard.zlSwapIsSucces(frmMain, lngModule, intSwapType, objThreePay.接口序号, _
                    strBalanceIDs, objThreePay.卡号, objThreePay.支付金额, intBusiness, strErrMsg, strXMLExpend)
            End If
            If blnSwapSuc Then
                strFilter = strFilter & IIf(strFilter <> "", " Or ", "") & "(卡类别ID = " & objThreePay.接口序号 & " And 消费卡 = 0 And 校对标志 = 1)"
                If blnHavNoSuccess = False Then blnHavNoSuccess = True
            Else
                If intBusiness = 1 Then
                    If strErrMsg = "" Then strErrMsg = "交易正在处理中, 请稍后再试！"
                    MsgBox strErrMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            Set objThreePay = New clsPayInfo
        End If
    Loop
    
    If bytMode = 5 And bytInState = 1 Then
        rsBillAdvance.Filter = "校对标志 = 2 And 结算性质 <> 3 And 结算性质 <> 4"
        If rsBillAdvance.RecordCount > 0 Or blnHavNoSuccess Then
            MsgBox "挂号单中的部分结算方式已支付成功，不允许作废，请继续重收完成挂号。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If strFilter <> "" Then
        rsBillAdvance.Filter = strFilter
        Do While Not rsBillAdvance.EOF
            rsBillAdvance!固定 = 1
            rsBillAdvance.Update
            rsBillAdvance.MoveNext
        Loop
    End If
    zlSwapIsSucces = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSwapIsSuccesWithSave(ByVal frmMain As Object, ByVal lngModule As Long, _
                        ByVal objRegInfor As clsRegEventInfor, _
                        ByRef objPay As clsPayInfo, ByRef blnSwapIsSuc As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查结算信息中所有三方卡的实际结算情况
    '入参:intSwapType-是否重退检查
    '     objPayInfos - 结算汇总信息
    '     rsBillAdvance - 交易明细数据
    '出参:blnSwapIsSuc - 交易是否成功
    '编制:李南春
    '日期:2019/1/30 15:19:37
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSwapSuc As Boolean, blnHavNoSuccess As Boolean, bln重退 As Boolean, bln检查卡费 As Boolean
    Dim intBusiness As Integer, intSwapType As Integer
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strBalanceIDs As String, strXMLExpend As String, strFilter As String, strErrMsg As String
    Dim dbl卡费 As Double, dbl挂号费 As Double
    
    If mobjOneCard Is Nothing Then Exit Function
    If objRegInfor Is Nothing Then Exit Function
    If objPay Is Nothing Then Exit Function

    If (objRegInfor.Card_变动类型 = CP_发卡 Or objRegInfor.Card_变动类型 = CP_退卡) And objRegInfor.objPayInfos.Card_卡类别ID = objPay.接口序号 And objRegInfor.objPayInfos.Card_消费卡 = objPay.消费卡 Then
        dbl卡费 = objRegInfor.objPayInfos.Card_结算金额
        strBalanceIDs = "5|" & objRegInfor.objPayInfos.Card_结帐ID
    End If
    dbl挂号费 = objPay.支付金额 - dbl卡费
    If RoundEx(dbl挂号费, 6) > 0 Then strBalanceIDs = strBalanceIDs & IIf(strBalanceIDs = "", "", "||") & "4|" & objRegInfor.objPayInfos.Reg_结帐ID
    
    On Error GoTo ErrHand
    '更新三方交易信息
    '    zlSwapIsSucces(ByVal frmMain As Object, ByVal lngModule As Long, ByVal intSwapType As Integer, _
    '    ByVal lngCardTypeID As Long, ByVal strBalanceIDs As String, ByVal strCardNo As String, ByVal dblSwapMoney As Double, _
    '    ByRef intSwapStatus_Out As Integer, Optional strErrMsg As String, Optional strXMLExpend As String) As Boolean
    blnSwapSuc = mobjOneCard.zlSwapIsSucces(frmMain, lngModule, 0, objPay.接口序号, _
        strBalanceIDs, objPay.卡号, objPay.支付金额, intBusiness, strErrMsg, strXMLExpend)
    If blnSwapSuc Then
        blnSwapIsSuc = True
    Else
        If intBusiness = 1 Then
            If strErrMsg = "" Then strErrMsg = "交易正在处理中, 请稍后再试！"
            MsgBox strErrMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
'    If RoundEx(dbl挂号费, 6) > 0 Then
'        strSQL = "Select 关联交易ID from 病人预交记录 where 记录性质 = 4 And 结帐ID = [1] And 卡类别ID = [2]"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlSwapIsSuccesWithSave", objRegInfor.objPayInfos.Reg_结帐ID, objPay.接口序号)
'        If rsTmp.EOF Then
'            MsgBox "未或取到待支付数据，不能继续！", vbInformation, gstrSysName
'            Exit Function
'        End If
'        objPay.关联交易ID = Val(Nvl(rsTmp!关联交易ID))
'    End If
    zlSwapIsSuccesWithSave = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetIDKindStr(ByVal strIDKindStr As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取读卡的种类
    '入参:strIDKindStr    String  IN
    '有两种格式:
    '一种是缺省的:短名1|全名1|读卡标志1;…. ;短名n|全名n|读卡标志n
    '另一种是该接口返回的格式:短名|全名|读卡标志|卡类别ID|卡号长度;…
    '出参:
    '返回: 短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '           是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密);…
    '        其中:卡类别ID|长度是本次增加的,由调用者根据情况来确认.
    '       比如:身|身份证号|0|0|18;IC|IC卡号|1|0|8;门|门诊号|0|0|0;就|就诊卡|0|0|0;建|建行卡|0|0|10
    '      出现错误时,返回空
    '编制:刘兴洪
    '日期:2011-06-14 14:43:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetIDKindStr = mobjOneCard.zlGetIDKindStr(strIDKindStr)
End Function

Public Function ReCalcOld(ByVal DateBir As Date, ByRef cbo年龄单位 As ComboBox, Optional ByVal lng病人ID As Long, Optional ByVal blnSetControl As Boolean = True) As String
'功能:根据出生日期重新计算病人的年龄,重设年龄单位
'参数:blnSetControl是否设置年龄单位控件
'返回:年龄,年龄单位
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String, str年龄 As String
 
    On Error GoTo errH
    str年龄 = mobjOneCard.Zl_Calc_Age(lng病人ID, Format(DateBir, "YYYY-MM-DD hh:mm:ss"))
    If blnSetControl = False Then
        ReCalcOld = str年龄
        Exit Function
    End If
    
    If str年龄 <> "" Then
        If str年龄 Like "*岁" Or str年龄 Like "*月" Or str年龄 Like "*天" Then
            strTmp = Mid(str年龄, 1, Len(str年龄) - 1)
            If IsNumeric(strTmp) Then
                Call cbo.Locate(cbo年龄单位, Mid(str年龄, Len(str年龄), 1))
            Else
                strTmp = str年龄
                cbo年龄单位.ListIndex = -1
            End If
        Else
            strTmp = str年龄
            If IsNumeric(strTmp) Then
                cbo年龄单位.ListIndex = 0
            Else
                cbo年龄单位.ListIndex = -1
            End If
        End If
    End If
    If cbo年龄单位.ListIndex = -1 Then
        cbo年龄单位.Visible = False
    Else
        If cbo年龄单位.Visible = False Then cbo年龄单位.Visible = True
    End If
    
    ReCalcOld = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiAddrssInfo(ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
                ByVal str地址类别 As String, ByRef cllAddrList As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人结构化地址信息
    ' 入参 : int地址类别:查询的地址类别：1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址；为0时表示查询所有类型的地址信息
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    GetPatiAddrssInfo = mobjOneCard.zlGetPatiAddrssInfo(lng病人ID, lng主页Id, str地址类别, cllAddrList)
End Function


Public Function ReadPatPricture(ByVal objControl As Object, ByVal lng病人ID As Long, Optional ByRef str采集图片 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人照片
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-13 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    On Error GoTo ErrHand

    If lng病人ID = 0 Then ReadPatPricture = True: Exit Function
    '67776:刘尔旋,2013-11-20,提取无照片的病人信息，照片没有清除
    Set objControl.Picture = Nothing
    'strTmp = Sys.ReadLob(glngSys, 27, lng病人ID)
    strTmp = mobjOneCard.zlGetPatiPhoto(lng病人ID)
    str采集图片 = strTmp
    objControl.Picture = LoadPicture(strTmp)
    If strTmp <> "" Then Kill strTmp
    ReadPatPricture = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SavePatPicture(ByVal lng病人ID As Long, ByVal strPhotoFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存病人照片
    '入参:
    '出参:
    '返回:返回病人照片Base64
    '编制:李南春
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPatiPhoto As String
    
    If lng病人ID = 0 Then Exit Function
    
    strPatiPhoto = zlStr.EncodeBase64_File(strPhotoFile)
    SavePatPicture = mobjOneCard.zlSavePatiPhoto(lng病人ID, strPatiPhoto)
End Function

Public Function DeletePatPicture(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除病人照片
    '入参:
    '出参:
    '返回:
    '编制:李南春
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If lng病人ID = 0 Then Exit Function
    DeletePatPicture = mobjOneCard.zlDeletePatiPhoto(lng病人ID)
End Function

Public Function GetInpatiState(ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
                Optional ByVal intPatiType As Integer) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人住院状态
    '入参:lng病人ID；lng主页ID
    '     intPatiType:病人性质 0-普通住院病人 1-门诊留观病人 2-住院留观病人
    '返回:cllPati_Out 病人信息
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As New Collection
    GetInpatiState = -1
    If lng病人ID = 0 Or lng主页Id = 0 Then Exit Function
    Call mobjOneCard.zlGetInpatiState(lng病人ID, lng主页Id, intPatiType, cllPati)
    If cllPati.Count > 0 Then
        GetInpatiState = cllPati("病人状态")
    End If
End Function

Private Sub Class_Terminate()
    Set mobjOneCard = Nothing
End Sub
