VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZHPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String
Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    '初始化公共部件
    If gobjComLib Is Nothing Then
        Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
        Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
        Set gobjControl = GetObject("", "zl9Comlib.clsControl")
        Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    End If
End Sub
Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set gobjComLib = Nothing
    Set gobjCommFun = Nothing
    Set gobjControl = Nothing
    Set gobjDatabase = Nothing
    Err = 0
End Sub
'*************************************************************************************************************************************
'与HIS的相关接口
'保持属性值的局部变量
'接口类
Public Function zlInitCompoent(ByVal intCallType As Integer, cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:zlInitCompoent (初始化接口部件)
    '入参:intCallType Byte    HIS传入 0-  门诊费用调用  1-  住院结帐调用
    '     cnOracle    ADODB.Connection    HIS传入 传入连接对象
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlInitCompoent = True
    Exit Function
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
    frmSetDrvParaSet.ShowMe frmMain, lngCardCode
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
        zlBrushCardSquare = True
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
    '只填标志
    zlSquareFee = True
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
Public Function zlSimulationBrushCard(ByVal frmMain As Object, ByVal Dbl刷卡金额 As Double, _
    ByRef lng接口编号 As Long, ByRef strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：选择指定卡类型
    '入参：frmMain HIS传入 调用的主窗体
    '         Dbl刷卡金额 HIS传入 传入预交界面中的金额
    '         Lng接口编号          HIS不传入
    '出参：Lng接口编号 传出    以何种结算卡结算
    '         strBlanceInfor  传出    用||分隔: 接口编号||消费卡ID(可传'')||结算方式||结算金额||卡号||交易流水号||交易时间(yyyy-mm-dd hh24:mi:ss)||备注
    '          strBlanceInfor-必须返回值,否则会报错.
    '返回：
    '编制：刘兴洪
    '日期：2010-06-18 11:33:22
    '说明：在预交款管理中增加预交时，点击确定按钮时调用(事务前调用)
    '------------------------------------------------------------------------------------------------------------------------
    strBlanceInfor = lng接口编号 & "|| ||POS结算||100||432143||222||" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "||备注"
    zlSimulationBrushCard = True
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
Public Function zlCheckDepositDelValied(ByVal lng预交ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查退费交易接口
    '入参： lng预交ID-预交ID
    '返回：成功返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-18 11:33:22
    '说明：1. 在预交款管理中的退款,点确定时(事务之前调用)，调用该交易
     '------------------------------------------------------------------------------------------------------------------------
    zlCheckDepositDelValied = True
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






