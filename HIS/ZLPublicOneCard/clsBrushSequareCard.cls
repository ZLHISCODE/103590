VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrushSequareCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Private mobjCard  As Object
Private mstrInterFace As String '读卡时选择的接口名称
Private mlng接口序号 As Long
Private mobjCurCard As clsCard
Private mobjDataBase As clsDataBase
Private mcnOracle As ADODB.Connection
Private mobjOneDataObject As clsOneCardDataObject
Public Event zlBrushCarding(ByVal strCardNo As String)

Public Function zlInitInterFacel(ByVal lng接口序号 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的接口部件
    '编制:刘兴洪
    '日期:2009-12-15 23:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As clsCard
    Dim objCardObjs As Cards, objCards As Cards
    
    Dim i As Long
    mlng接口序号 = lng接口序号:
    
    Err = 0: On Error Resume Next
    If mobjOneDataObject.zlGetCardObjs_XF(objCardObjs) = False Then Exit Function
    
    '检查设备是否启用
    Set mobjCard = objCardObjs("K" & lng接口序号)
    
    If Err <> 0 Then
            Err = 0: On Error Resume Next
            If mobjOneDataObject.zlGetCards_XF(objCards) = False Then Exit Function
            
            Set objCard = objCards("K" & lng接口序号)
            If Err <> 0 Then
                ShowMsgbox "部件:" & mlng接口序号 & "未找到!"
                Call mobjDataBase.WritLog("clsBrushSequareCard.zlInitInterFacel", "", "部件:" & mlng接口序号 & "未找到!")
                Exit Function
            End If
            If Err <> 0 Then
                ShowMsgbox "部件:" & mlng接口序号 & "-" & objCard.名称 & "没有正常启用!" & vbCrLf & "详细的信息为:" & Err.Description
                Call mobjDataBase.WritLog("clsBrushSequareCard.zlInitInterFacel", "", "部件:" & mlng接口序号 & "-" & objCard.名称 & "没有正常启用!" & vbCrLf & "详细的信息为:" & Err.Description)
            End If
          Exit Function
    End If
    
    Err = 0: On Error Resume Next
    If mobjOneDataObject.zlGetCards_XF(objCards) = False Then Exit Function
    
    Set objCard = objCards("K" & lng接口序号)
    If objCard.是否自动读取 = 1 Then gbln自动读取 = True
    Set mobjCurCard = objCard
    If Err <> 0 Then
        ShowMsgbox "部件:" & mlng接口序号 & "未找到!"
        Call mobjDataBase.WritLog("clsBrushSequareCard.zlInitInterFacel", "", "部件:" & mlng接口序号 & "未找到!")
        Exit Function
    End If
    
    If objCardObjs("K" & lng接口序号).自制卡 Then
        Call frmCardBrush.Init(Me)
        frmCardBrush.tmrMain.Interval = Val(GetSetting("ZLSOFT", "公共全局\SquareCard\" & lng接口序号, "自动读取间隔", 300))
        With gTy_TestBug
            Set .objSquareCard = objCard
        End With
    End If
    zlInitInterFacel = True
End Function

'------------------------------------------------------------------
'相关的刷卡操作
Friend Sub zlBrushCarding(ByVal strCardNo As String)
    '刷卡时
    RaiseEvent zlBrushCarding(strCardNo)
End Sub
Public Sub zlSetAutoBrush(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置自动刷卡
    '编制:刘兴洪
    '日期:2009-12-15 17:38:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gbln自动读取 = True Then
        frmCardBrush.tmrMain.Enabled = blnEnabled
    End If
End Sub

'------------------------------------------------------------------
'相关的接口操作
'保持属性值的局部变量
'接口类
Public Function zlInitComponents(ByVal intCallType As Integer, cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlInitComponents (初始化接口部件)
    '入参:intCallType Byte    HIS传入 0-  门诊费用调用  1-  住院结帐调用
    '     cnOracle    ADODB.Connection    HIS传入 传入连接对象
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.消费卡时.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mcnOracle = cnOracle
    If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    If zlGetOneCardDataObject(mcnOracle, mobjOneDataObject) = False Then Exit Function
    zlInitComponents = True
End Function

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
    zlCardDevSet = mobjCard.zlCardDevSet(frmMain, lngCardCode)
    
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
   zlBrushCardSquare = mobjCard.zlBrushCardSquare(intCallType, Me, lng接口编号, rsFeeList, dbl最大消费, rsSquare)
End Function
Public Function zlSquareFee(ByVal intCallType As Integer, frmMain As Object, ByVal str结帐ID_IN As String, ByVal rsSquare As ADODB.Recordset) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: zlSquareFee (结算接口)
    '入参:frmMain:HIS传入 调用的主窗体
    '     intCallType : HIS传入 0-  门诊费用调用 1-  住院结帐调用
    '     str结帐ID_IN: HIS传入 本次结帐的结帐ID集
    '     rsSquare :  本次应刷卡的交易
    '出参:
    '返回:true:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:18:38
    '说明:
    '    1. 在"门诊收费"界面点"确定"时,调用本接口
    '    2. 在"住院结帐"界面点"确定"时,调用本接口
    '注:
    '  此接口由于是在HIS事务中 , 因此不能在此接口存在与用户交互的操作
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    
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
End Function
Public Function zlSimulationBrushCard(ByVal frmMain As Object, ByVal Dbl刷卡金额 As Double, _
    ByRef lng接口编号 As Long, ByRef strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：选择指定卡类型
    '入参：frmMain HIS传入 调用的主窗体
    '         Dbl刷卡金额 HIS传入 传入预交界面中的金额
    '         Lng接口编号          HIS不传入
    '出参：Lng接口编号 传出    以何种结算卡结算
    '         strBlanceInfor  传出    用||分隔: 接口编号||消费卡ID(可传'')||结算方式||结算金额||卡号||交易流水号||交易时间(yyyy-mm-dd hh24:mi:ss)||备注
    '返回：
    '编制：刘兴洪
    '日期：2010-06-18 11:33:22
    '说明：在预交款管理中增加预交时，点击确定按钮时调用(事务前调用)
    '------------------------------------------------------------------------------------------------------------------------
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
End Function
Public Function zlReadCard(frmMain As Object, ByRef strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读卡接口:暂不考虑病人相关接口
    '入参:frmMain-调用的主窗体
    '出参:strCardNo-返回的卡号
    '返回:函数返回    True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:44:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjCard Is Nothing Then Exit Function
    If mobjCurCard Is Nothing Then Exit Function
    zlReadCard = mobjCard.zlReadCard(Me, strCardNo)
End Function

Public Function zlCloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    '函数模板，不实现具体功能
End Function
Private Sub Class_Terminate()
    Call zlCloseCard
    Err = 0: On Error Resume Next
    If Not mobjCard Is Nothing Then Set mobjCard = Nothing
    If Not mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub

'----------------------
'以下是内部过程
'----------------------
Private Sub Class_Initialize()
    On Error GoTo errHandle
    Set mobjCard = Nothing
    mstrInterFace = ""
    glngInstanceCount = glngInstanceCount + 1
    Exit Sub
errHandle:
End Sub

Public Property Get zlobjCard() As Object
    Set zlobjCard = mobjCard
End Property

Public Property Let zlobjCard(ByVal vNewValue As Object)
    Set mobjCard = vNewValue
End Property

Public Property Get zlSetSquareNo() As Long
    zlSetSquareNo = mlng接口序号
End Property

 


