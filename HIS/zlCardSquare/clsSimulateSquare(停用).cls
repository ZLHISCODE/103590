VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSimulateSquareCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarObjectEvts As Object
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
'Private WithEvents mfrmTimer As frmTimer    '接收读卡接口返回的读卡信息
Public Function zlCardDevSet(ByVal frmMain As Object, Optional lngCardCode As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlCardDevSet(设备参数设置接口):主要是配置相关的设备的相关参数
    '入参:frmMain Form    HIS传入 调用的主窗体
    '出参:
    '返回:true:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:18:38
    'HIS调用说明.
    '    1.  在门诊收费管理的参数设置'设备配置' zlCardDevSet
    '    2.  在住院结帐界面的参数设置'设备配置' zlCardDevSet
    '---------------------------------------------------------------------------------------------------------------------------------------------
   ' frmSetDrvParaSet.ShowMe frmMain, lngCardCode
End Function

Public Function zlBrushCardSquare(ByVal intCallType As Integer, frmMain As Object, ByVal lng接口编号 As Long, ByVal rsFeeList As ADODB.Recordset, ByVal dbl最大消费 As Double, ByRef rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: zlBrushCardSquare (刷卡结算接口)
    '入参:frmMain:HIS传入 调用的主窗体
    '     intCallType : HIS传入 0-  门诊费用调用 1-  住院结帐调用
    '     rsFeeList: HIS传入 如果是门诊多单据,则所有单据的明细,如果是住院结帐 , 则是本次结帐的所有明细
    '     dbl最大消费 :  HIS传入 表示刷卡不能超过此金额
    '
    '出参:rsSquare : 接口返回    本地记录集:接口传入空结构(接口返回相关的数据) , 结构如下:
    '                接口编号 , 消费卡ID, 结算方式, 结算金额, 卡号卡名称, 交易流水号, 交易时间, 备注
    '     rsSquare说明:主要是解决同一单据,刷多张卡消费的情况.,如果本次刷多张卡 , 则传入接口中已经刷过的卡信息
    '返回:true:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:18:38
    '说明:
    '    1.  在门诊收费界面时,HIS在点"结算卡"时,调用本接口
    '    2.  在住院结帐界面时,HIS在点"结算卡"时,调用本接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
   ' zlBrushCardSquare = frmSquareBrushCard.zlShowBrushCard(frmMain, lng接口编号, intCallType, rsFeeList, dbl最大消费, rsSquare)
End Function
Public Function zlCheckDelSquareValied(frmMain As Object, ByVal intCallType As Integer, ByVal str结帐ID_IN As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:在执行退费时,检查相关的接口部件是否正常
    '入参:
    '出参:
    '返回:正常,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-31 16:39:47
    '说明;
    '     在退费时，需要进行相关的检查
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlCheckDelSquareValied = True
End Function
Public Function zlDelSquareFee(ByVal intCallType As Integer, frmMain As Object, ByVal str结帐ID_IN As String) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: zlSquareFee (结算接口)
    '入参:frmMain:HIS传入 调用的主窗体
    '     intCallType : HIS传入 0-  门诊费用调用 1-  住院结帐调用
    '     str结帐ID_IN: HIS传入 本次结帐的结帐ID集
    '出参:
    '返回:true:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:18:38
    '说明:
    '    1. "门诊收费管理"和"住院结帐管理"中作废时,调用此接口
    '注:
    '  此接口由于是在HIS事务中 , 因此不能在此接口存在与用户交互的操作
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlDelSquareFee = True
    
End Function
 
Public Function zlBrushCardToDeposit(ByVal lng预交ID As Long, ByVal lng接口编号 As Long, ByRef strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：刷卡存预交交易
    '入参： lng预交ID-预交ID
    '           lng接口编号-结算卡编号
    '出参：strBlanceInfor-返回刷卡信息:
    '         用||分隔: 接口编号||消费卡ID(可传'')||结算方式||结算金额||卡号||交易流水号||交易时间(yyyy-mm-dd hh24:mi:ss)||备注
    '返回：成功返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-18 11:33:22
    '说明：在预交款管理中增加预交时，点击确定按钮时调用(事务中调用)
    '          出参一定要传入正确,否则会出现程序错误
    '------------------------------------------------------------------------------------------------------------------------
    zlBrushCardToDeposit = True
End Function
Public Function zlDepositDel(ByVal lng预交ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：退预交交易
    '入参： lng预交ID-预交ID
    '返回：成功返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-18 11:33:22
    '说明：1. 在预交款管理中的退款, 点确定时(事务中调用)，调用该交易
     '------------------------------------------------------------------------------------------------------------------------
     zlDepositDel = True
End Function


''Public Function zlReadCard(frmMain As Object, ByRef strCardNo As String) As Boolean
''    '---------------------------------------------------------------------------------------------------------------------------------------------
''    '功能:读卡接口:暂不考虑病人相关接口
''    '入参:frmMain-调用的主窗体
''    '出参:strCardNo-返回的卡号
''    '返回:函数返回    True:调用成功,False:调用失败
''    '编制:刘兴洪
''    '日期:2009-12-15 15:44:09
''    '---------------------------------------------------------------------------------------------------------------------------------------------
''    Dim objFile As New FileSystemObject, strPath As String, strFile As String, strCurCardNo As String
''    Dim objText As TextStream
''
''    Err = 0: On Error GoTo ErrHand:
''    strPath = App.Path & "\调试"
''    If objFile.FolderExists(strPath) = False Then objFile.CreateFolder (strPath)
''    strFile = strPath & "\CardInfor.txt"
''    If objFile.FileExists(strFile) = False Then objFile.CreateTextFile strFile, True
''
''    With gTy_TestBug
''        '不存在
''        If .objSquareCard Is Nothing Then Exit Function
''        If .bytType = 1 Then
''            If .strStartNo <> "" Then
''                '求出的是结束号
''                '不能太大,随机产生在10以内的单据
''                strCurCardNo = Val(Replace(.strStartNo, .objSquareCard.前缀文本, "")) + Int((Rnd * 100))
''                strCurCardNo = String(.objSquareCard.卡号长度 - Len(strCurCardNo), "0") & strCurCardNo
''                strCardNo = .objSquareCard.前缀文本 & strCurCardNo
''            Else
''                '随机产生一个号码
''                strCurCardNo = Int(Rnd * IIf(.objSquareCard.卡号长度 > 5, 10000, .objSquareCard.卡号长度))
''                strCurCardNo = String(.objSquareCard.卡号长度 - Len(strCurCardNo), "0") & strCurCardNo
''                strCardNo = .objSquareCard.前缀文本 & strCurCardNo
''            End If
''            '向调试文本中加入随机号
''            Set objText = objFile.OpenTextFile(strFile, ForAppending)
''            objText.WriteLine "1" & vbTab & strCardNo: objText.Close
''        Else
''            '读取已经刷个卡的数据
''            strCardNo = Read模拟卡号(strFile)
''        End If
''    End With
''    zlReadCard = True
''    Exit Function
''ErrHand:
''    If ErrCenter = 1 Then Resume
''End Function

Public Function zlCloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    '函数模板，不实现具体功能
End Function
Private Sub Class_Terminate()
    Call zlCloseCard
End Sub
'******************************************************************************
'新接口处理:
'   主要是适应新的收费流程,目前增加如下接口
'   zlInitComponents:初始化部件,参数有所变化
'   zlPaymentCheck:扣款前检查
'   zlPaymentMoney:扣钱
'******************************************************************************
Public Function zlInitComponents(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    ByVal cnOracle As ADODB.Connection, _
    Optional blnDeviceSet As Boolean = False, _
    Optional strExpand As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlInitComponents (初始化接口部件)
    '入参: frmMain-调用的主窗体
    '        lngModule-HIS调用模块号
    '       lngSys-传入的系统号
    '       strDBUser-数据库用户名
    '       cnOracle -HIS/三方机构
    '       blnDeviceSet-设备设置调用初始化
    '       strExpand-扩展信息(暂留,现为空)
    '返回:函数返回True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 11:49:39
    '说明:如果是第三方接口主动读取数据,则在渠道更改的部件中传入连接.
    '    1.如果第三接口存在签到的情况，可以在此接口中进行签道.
    '    2.如果第三方接口读取数据,需要先初始化.
    '调用者:
    '    1.  进入门诊收费、挂号时调用本接口
    '    2.  进入医生站、护士站、医技站等时调用本接口
    '    3.  第三方接口需要主动读取数据时,必须先初始化.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlInitComponents = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlBrushCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal strPatiName As String, ByVal strSex As String, _
    ByVal strOld As String, ByVal dbl金额 As Double, _
    Optional ByRef strCardNo As String, _
    Optional ByRef strPassWord As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定支付类别,弹出
    '入参:rsClassMoney:收费类别,金额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:18:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBrushCard = True
 End Function
Public Function zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, _
    ByVal strCardNo As String, strExpand As String, dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取帐户余额
    '入参:frmMain-调用的主窗体
    '        lngModule-模块号
    '        strCardNo-卡号
    '        strExpand-预留，为空,以后扩展
    '出参:dblMoney-返回帐户余额
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 16:29:48
    '说明:
    '       在所有需要扣款的地方，都要检查帐户余额是否充足,帐户不充足时不允许扣款.
    '       如果某些第三方接口不存在余额接口，可以固定返回一定的金额。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    '获取帐户余额
    dblMoney = 0
    strSQL = _
        "Select a.余额" & vbNewLine & _
        "From 消费卡信息 A" & vbNewLine & _
        "Where a.卡号 = [1] And a.接口编号 = [2]" & vbNewLine & _
        "      And 序号 = (Select Max(序号) From 消费卡信息 B Where 卡号 = [1] And 接口编号 = [2])" & vbNewLine & _
        "Order By a.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取帐户余额", strCardNo, strCardTypeID)
    If Not rsTemp.EOF Then
        dblMoney = Round(Val(NVL(rsTemp!余额)), 2)
    End If
    zlGetAccountMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       strCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       dblMoney-支付金额(退款时为负数)
    '       strNos-本次支付所涉及的单据
    '       strXMLExpend-(XML串:验证密码:自助机用)
    '出参:
    '   strXMLExpend-(XML串:错误信息)
    '返回:扣款合法,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 16:42:43
    '说明:
    '   在调用扣款前，由于存在Oracle事务问题， 所以再调用扣款交易前， _
    '   先进行数据的合法性检查,以便控制死锁情况。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlPaymentCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlPaymentMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, ByVal strPrepayNos As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
    ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '       strSwapExtendInfor -(自助机时，传入自助机输入的密码)
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 17:13:48
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
   On Error GoTo errHandle
    zlPaymentMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReadCard(frmMain As Object, _
    ByVal lngModule As Long, _
     ByVal blnOlnyCardNO As Boolean, _
    ByVal strExpand As String, _
    ByRef strOutCardNo As String, _
    ByRef strOutclsPatientInfoXml As String, _
    Optional ByRef strPhotoFile As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读卡接口
    '入参:frmMain-调用的父窗口
    '       lngModule-调用的模块号
    '       strExpand-扩展参数,暂无用
    '       blnOlnyCardNO-仅仅读取卡号
    '出参:strOutCardNO-返回的卡号
    '       strOutclsPatientInfoXML-(病人信息返回.XML串)
    '       strPhotoFile-当从卡中读取取像片后,传给HIS的像片文件(含完整路径)
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 11:56:41
    '说明:
    '   １. 在所有绑定卡时,需要建立病人信息档案时，读取该接口
    '   ２. 在所有需要刷卡进行就诊的地方，都需要调用该接口
    'strOutclsPatientInfoXML参数说明
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    '    姓名    Varchar2    100
    '    性别    Varchar2    4
    '    年龄    Varchar2    10
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    '    出生地点    Varchar2    50
    '    身份证号    VARCHAR2    18
    '    其他证件    Varchar2    20
    '    职业    Varchar2    80
    '    民族    Varchar2    20
    '    国籍    Varchar2    30
    '    学历    Varchar2    10
    '    婚姻状况    Varchar2    4
    '    区域    Varchar2    30
    '    家庭地址    Varchar2    50
    '    家庭电话    Varchar2    20
    '    家庭地址邮编    Varchar2    6
    '    监护人  Varchar2    64
    '    联系人姓名  Varchar2    64
    '    联系人关系  Varchar2    30
    '    联系人地址  Varchar2    50
    '    联系人电话  Varchar2    20
    '    工作单位    Varchar2    100
    '    单位电话    Varchar2    20
    '    单位邮编    Varchar2    6
    '    单位开户行  Varchar2    50
    '    单位帐号    Varchar2    20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Errhand:
      ' 处理IC卡号
    strOutCardNo = ""
    If mobjICCard Is Nothing Then
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If Not mobjICCard Is Nothing Then
        strOutCardNo = mobjICCard.Read_Card()
    End If
    zlReadCard = strOutCardNo <> ""
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlMakeCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal stInputclsPatientInfoXml As String, _
    ByRef strPhotoFile As String, _
    ByRef strOutclsPatientInfoXml As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:制卡接口
    '入参:
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  卡类别ID
    '    stInputclsPatientInfoXml String(XML) In  HIS传入已经健档的病人信息
    '    strPhotoFile    String  In  表示像片文件(含完整路径)
    '    strExpend   String  In   XML,暂无,待以后扩展
    '出参:
    '    strOutclsPatientInfoXml  String(XML) Out 调用接口后,传给HIS系统已经制卡后的病人信息,以便建立病人档案
    '    strExpend   String  Out  XML,暂无,待以后扩展
    '    strPhotoFile    String  In  表示像片文件(含完整路径)
    '返回 :True:调用成功,False:调用失败
    '调用者:
    '    目前只在医疗卡发放管理中调用此接口(主要是考虑输入的信息较多,其他窗口不适应)，相关的调用流程医疗卡发放管理的制卡
    '  strInputclsPatientInfoXML及strOutclsPatientInfoXML格式
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    '    姓名    Varchar2    100
    '    性别    Varchar2    4
    '    年龄    Varchar2    10
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    '    出生地点    Varchar2    50
    '    身份证号    VARCHAR2    18
    '    其他证件    Varchar2    20
    '    职业    Varchar2    80
    '    民族    Varchar2    20
    '    国籍    Varchar2    30
    '    学历    Varchar2    10
    '    婚姻状况    Varchar2    4
    '    区域    Varchar2    30
    '    家庭地址    Varchar2    50
    '    家庭电话    Varchar2    20
    '    户口邮编    Varchar2    6
    '    监护人  Varchar2    64
    '    联系人姓名  Varchar2    64
    '    联系人关系  Varchar2    30
    '    联系人地址  Varchar2    50
    '    联系人电话  Varchar2    20
    '    工作单位    Varchar2    100
    '    单位电话    Varchar2    20
    '    单位邮编    Varchar2    6
    '    单位开户行  Varchar2    50
    '    单位帐号    Varchar2    20
    '    病人ID  Number  18      可以为空
    '    ABO血型 Varchar2    10      信息名='ABO'
    '    信息值=(A型;B型;O型;AB型;不详)
    '    RH  Varchar2    10      信息名='RH'
    '    信息值=(阴;阳;不详;未查)
    '    哮喘标志    Number  2       信息名='医学警示'
    '    信息值=' 哮喘;心脏病; 心脑血管病;…'
    '    心脏病标志  Number  2
    '    心脑血管病标志  Number  2
    '    癫痫病标志  Number  2
    '    凝血紊乱标志    Number  2
    '    糖尿病标志  Number  2
    '    青光眼标志  Number  2
    '    透析标志    Number  2
    '    器官移植标志    Number  2
    '    器官缺失标志    Number  2
    '    可装卸义肢标志  Number  2
    '    心脏起搏器标志  Number  2
    '    其他医学警示    Varchar2    100
    '    联系信息
    '        姓名    Varchar2    30
    '        关系    Varchar2    30
    '        电话    Varchar2    30
    '        身份证号    Varchar2    30
    '    健康档案编号    Varchar2    18
    '    新农合证号  Varchar2    18
    '    其他证件
    '        信息名  Varchar2    15
    '        信息值  Varchar2    100
    '    其他信息 可以返回上面接点不存在的数据
    '        信息名  Varchar2    20      例: 工作单位, 血压等
    '        信息值  Varchar2    50      比如:信息名=血压;信息值='XXmmHg'
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlMakeCard = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlBandCardArfter(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng病人ID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绑定卡接口
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
    '56615
    zlBandCardArfter = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlMzInforWriteToCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng病人ID As Long, _
    ByVal lngBalanceID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写门诊信息接口
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  传入写卡类别ID:
    '           1)传入刷卡的类别ID
    '           2)传入零时,需要选择某个卡类别ID
    '    lng病人ID   Long    In  病人ID
    '    lngBalanceID    Long    In  结算序号(某次结算的序号)
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    True:调用成功,False:调用失败
    '调用时机:
    '         医疗卡类别.是否写卡=1才调用
    '调用者:
    '    1.  目前收费管理完成后,调用该接口
    '    2.  退费完成后,调用该接口
    '说明:
    '        门诊相关信息的写入 , 主要是渠道编制的接口内部写入
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '56615
    zlMzInforWriteToCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlZyInforWriteToCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng病人ID As Long, _
    ByVal lngBalanceID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写住院信息接口
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lngCardTypeID   Long    In  传入写卡类别ID:
    '           1)传入刷卡的类别ID
    '           2)传入零时,需要选择某个卡类别ID
    '    lng病人ID   Long    In  病人ID
    '    lngBalanceID    Long    In  结帐ID(可以不传入)
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    True:调用成功,False:调用失败
    '调用时机:
    '        医疗卡类别.是否写卡=1才调用
    '调用者:
    '    1.  出院时
    '    2.  结帐时
    '    3.  结帐作废时
    '说明:
    '       住院相关信息的写入 , 主要是渠道编制的接口内部写入
    '---------------------------------------------------------------------------------------------------------------------------------------------
      '56615
    zlZyInforWriteToCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReturncheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, ByVal dblMoney As Double, _
    ByVal strSwapNo As String, ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户回退交易前的检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(退款时检查)
    '       strSwapMemo-交易说明(退款时传入)
    '       strXMLExpend    XML IN  可选参数(扩展用).暂未传入
    '返回:退款合法,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-05-26 17:24:55
    '说明:
    '    在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,
    '    以便控制死锁情况。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlReturncheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
    ByRef strSwapExtendInfor As String) As Boolean
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
    '       strSwapExtendInfor-本次退费的冲销ID：
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '出参: strSwapNo-交易流水号(退款交易流水号)
    '         strSwapMemo-交易说明(退款交易说明)
    '       strSwapExtendInfor-交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-26 17:31:32
    '说明:
    '       在所有需要对已经消费的地方进行退款的情况下，调用该接口,目前规划在:收费室；挂号等。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlReturnMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlThreeInterFaceMgrTool(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方接口管理工具
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       卡类别IID
    '编制:刘兴洪
    '日期:2012-02-13 10:17:26
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
     zlThreeInterFaceMgrTool = True
 End Function

Public Property Get zlInitEvents() As Object
    '初始化事件(公共事件)
    Set zlInitEvents = mvarObjectEvts
End Property
Public Property Set zlInitEvents(ByVal vNewValue As Object)
    '初始化事件(公共事件)
    Set mvarObjectEvts = vNewValue
End Property

Public Sub SetEnable(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置是否处于读卡状态
    '入参:blnEnabled-False表示该设备不处于读卡状态,true表示处于读卡状态
    '编制:刘兴洪
    '日期:2012-03-24 15:12:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
   ' mfrmTimer.SetEnable blnEnabled
End Sub

Private Sub mfrmTimer_ShowCardNO(ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接收扫描出来的结果
    '入参:strCardNo-扫描出来的卡号
    '编制:刘兴洪
    '日期:2012-03-24 15:21:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '触发公共事件部件,主要是调用触发方法
    Call mvarObjectEvts.zlShowCardInfor(0, strCardNo, strXmlCardInfor, strExpended, blnCancel)
End Sub
 
Public Function zlHealthArchiveIsSHow(frmMain As Object, _
    ByVal lngModule As Long, _
    ByRef strFunName As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:居民健档档案
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    strFunName  String  Out 返回功能名, 即在医生工作站的菜单中显示的名称(菜单位置放在"工具"菜单下)
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    '        Boolean 函数返回    True:表示应该加入档案显示菜单,False:不加入档案菜单
    '调用者:
    '    1.  门诊(住院)医生工作站调用,医生站
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlHealthArchiveIsSHow = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlHealthArchivesShow(frmMain As Object, _
    ByVal lngModule As Long, ByVal lng病人ID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:居民健档档案
    '    frmMain Object  In  调用的主窗体
    '    lngModule   Long    In  调用的模块号
    '    lng病人ID-病人ID
    '    strExpend   String  In/Out  XML,暂无,待以后扩展
    ' 函数返回    医生工作站不做返回判断,渠道编制的接口部件中可以返回值
    ' 调用者:门诊(住院)医生工作站后调用
    '说明:     该接口分两层:
    '    1.  一层是在zl9CardSquare部件中该接口不返回值,由医生站直接调用,该接口调用渠道编制的接口界面,否则按标准的健康卡档案界面显示(见下图)
    '    2.  二层是渠道编制的接口部件:该接口允许返回值为TRUE时,表示利用渠道自己编制的界面展示,否则将按HIS固定的健康档案信息显示(见下图)
    '调用者:
    '    1.  门诊(住院)医生工作站调用,医生站
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlHealthArchivesShow = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReleaseComponent() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:释放资源接口
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-06-24 09:44:54
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlReleaseComponent = True
End Function

Public Function ZlGetParaConfig(ByVal frmMain As Object, _
    ByVal intPara As Integer, _
    Optional strErrMsg As String, Optional strExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取接口参数
    '入参: frmMain-调用的主窗体
    '       intPara: 包含如下值
    '                1-刷卡和支付在同一页面:true-新模式；False-旧模式
    '       strExpend-扩展参数，暂留，现传为空
    '出参:strErrMsg-返回的错误信息
    '       strExpend-扩展参数，暂留，现传为空
    '返回:函数返回True:调用成功,False:调用失败
   '日期:2013-06-15 20:22:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ZlGetParaConfig = False
End Function

Public Function ZlGetCardNo(ByRef strCardNo As String, Optional strExpend As String) As Boolean
    '获取刷卡卡号
    '出参:strCardNo-刷卡卡号
    '     strExpend-扩展参数，暂留，现传为空
    ZlGetCardNo = True
End Function

