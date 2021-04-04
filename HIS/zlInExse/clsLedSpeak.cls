VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLedSpeak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'LED控制
Private mblnLED As Boolean        '结帐时是否启用LED设备报价
Private mblnLedWelcome As Boolean '是否在输完病人后提示欢迎信息
Private mlngModul As Long
Private mlngSys As Long
Private mobjCom As MSComm
Private mfrmMain As Object
Private mdbl缴款 As Double, mdbl找零 As String, mdbl结算总额 As Double  '主要记录需要播放的数据，以例直接使用zlSpeak_Money


Public Property Let blnLED(ByVal vData As Boolean)
    mblnLED = vData
End Property
Public Property Get blnLED() As Boolean
    blnLED = mblnLED
End Property

Public Property Let 缴款(ByVal vData As Double)
    mdbl缴款 = vData
End Property
Public Property Get 缴款() As Double
    缴款 = mdbl缴款
End Property
Public Property Let 找零(ByVal vData As Double)
    mdbl找零 = vData
End Property
Public Property Get 找零() As Double
    找零 = mdbl找零
End Property
Public Property Let 结算总额(ByVal vData As Double)
    mdbl结算总额 = vData
End Property
Public Property Get 结算总额() As Double
    结算总额 = mdbl结算总额
End Property

Public Sub zlInit(ByVal mfrmMain As Object, ByVal lngModul As Long, ByVal lngSys As Long, ByVal objCom As Object, Optional blnNotStartLed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化部件
    '入参:mfrmMain-调用的主窗口
    '     lngModul-模块号
    '     objCom-MsComm控件
    '     lngSys-系统号
    '     blnNotStartLed-true-表示不启用Led,False-表示根据设置来决定是否启用一卡通
    '编制:刘兴洪
    '日期:2018-09-26 14:24:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlngModul = lngModul: mlngSys = lngSys: Set mobjCom = objCom
    mblnLED = False: mblnLedWelcome = False
    If Not blnNotStartLed Then  '表示根据设置来决定是否启用一卡通
        If lngModul = 1133 Or lngModul = 1134 Or lngModul = 1135 Or lngModul = 1137 Then
            mblnLED = Val(GetSetting("ZLSOFT", "公共全局", "使用", 0)) <> 0
        End If
        If lngModul = 1133 Or lngModul = 1134 Or lngModul = 1135 Or lngModul = 1137 Then
            mblnLedWelcome = zlDatabase.GetPara("LED显示欢迎信息", mlngSys, mlngModul, "1") = "1"
        End If
    End If
    Call zlInitLed(objCom)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlInitLed(ByVal objCom As Object, Optional blnReSetCom As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化Led
    '入参:objCom-MsComm控件
    '     blnReSetCom-是否重新设置Com端口
    '编制:刘兴洪
    '日期:2018-09-26 14:55:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mblnLED Then Exit Sub
    
    If blnReSetCom Then zl9LedVoice.Reset objCom
    zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
End Sub
Public Sub zlLedShowWelcome(Optional objPati As clsPatientInfo, Optional blnReSetCom As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示欢迎信息
    '入参:blnReSetCom-是否重新设置Com端口
    '编制:刘兴洪
    '日期:2015-01-08 10:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mblnLED = False Then Exit Sub
    
    If mblnLedWelcome Then
        Call zlInitLed(mobjCom, blnReSetCom)
        zl9LedVoice.Speak "#1"
    End If
    Call zlDisplayPatient(objPati)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub zlSpeak_请付款(ByVal dblMoney As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:播放请你付款
    '入参:dblMoney-播放当前你需要支付的金额
    '编制:刘兴洪
    '日期:2018-09-26 16:15:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Then Exit Sub
    zl9LedVoice.Speak "#21 " & Format(dblMoney, "0.00")  '请您付款XXXX元
End Sub

Public Sub zlDispCharge(ByVal dbl总额 As Double, ByVal dbl缴款 As Double, ByVal dbl找零 As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示应收,实收,找零,并发声
    '编制:刘兴洪
    '日期:2018-09-26 16:36:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If mblnLED = False Then Exit Sub
    '显示应收,实收,找零,并发声
    zl9LedVoice.DispCharge Format(dbl总额, "0.00"), Format(dbl缴款, "0.00"), Format(dbl找零, "0.00")
    
End Sub
Public Sub zlSpeak_PayMoney(ByVal dbl结算总额 As Double, ByVal dbl缴款 As Double, ByVal dbl找零 As Double, Optional blnPaySpeakMoney As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:播放请你付款(此过程会改变缴款，找零、结算总额的属性值 )
    '入参:dbl缴款-本次支付缴款金额
    '     dbl找零-找零
    '     blnPaySpeakMoney-true:播放预收、找零及当面点清;False-不播放
    '编制:刘兴洪
    '日期:2018-09-26 16:15:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Then Exit Sub
    '显示应收,实收,找零,并发声
    mdbl缴款 = dbl缴款: mdbl结算总额 = dbl结算总额: mdbl找零 = dbl找零
    Call zlDispCharge(dbl结算总额, dbl缴款, dbl找零)
    
    If Not blnPaySpeakMoney Then Exit Sub
    '播放预收、找零及当面点清
    Call zlSpeak_Money
End Sub
Public Sub zlSpeak_Money()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:播放预收、找零及当面点清（需事先对缴款，找零属性赋值）
    '编制:刘兴洪
    '日期:2018-09-26 16:43:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Then Exit Sub
    
    zl9LedVoice.Speak "#22 " & Format(mdbl缴款, "0.00")  '预收XXXX元
    zl9LedVoice.Speak "#23 " & Format(mdbl找零, "0.00")   '找零XXXX元
    zl9LedVoice.Speak "#3"   '#3  --请当面点清, 谢谢!
End Sub
Public Sub zlSpeak_找零(ByVal dbl找零 As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:播放预收、找零及当面点清（需事先对缴款，找零属性赋值）
    '编制:刘兴洪
    '日期:2018-09-26 16:43:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Then Exit Sub
    zl9LedVoice.Speak "#23 " & Format(dbl找零, "0.00")  '找零XXXX元
End Sub



Public Sub zlDisplayPatient(objPati As clsPatientInfo)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示病人信息
    '入参:objPati-病人信息集
    '编制:刘兴洪
    '日期:2018-09-26 14:50:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Or objPati Is Nothing Then Exit Sub
    zl9LedVoice.DisplayPatient objPati.姓名 & " " & objPati.性别 & " " & objPati.年龄, objPati.病人ID
End Sub
Public Sub zlClearDisplayPatient()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除显示病人信息
    '入参:objPati-病人信息集
    '编制:刘兴洪
    '日期:2018-09-26 14:50:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Then Exit Sub
    zl9LedVoice.DisplayPatient ""
End Sub

Public Sub zlClearDisplayBank(Optional blnSpace As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除显示信息
    '入参:blnSpace 是否传入空格
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-09-26 17:17:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mblnLED Then Exit Sub
    zl9LedVoice.DisplayBank IIf(blnSpace, " ", "")
End Sub

Public Sub zlDisplayBank(ByVal cllBanks As Collection, Optional blnIsWait4 As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:Led信息显示（在任意行上显示任意字符）
    '入参:cllBanks-显示的个数集,格式为:array(项目,金额)
    '     blnIsWait4-是否等待4秒
    '编制:刘兴洪
    '日期:2015-01-07 13:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtTmpDate As Date, varBank(1 To 20) As String
    Dim i As Long
    If Not mblnLED Or cllBanks Is Nothing Then Exit Sub
    
    On Error GoTo errHandle
    
    For i = 1 To cllBanks.Count
        If i <= 20 Then
            varBank(i) = cllBanks(i)(0) & cllBanks(i)(1) '格式如:"帐户余额" & Format(mobjInsureInfor.帐户余额, "0.00")
        End If
    Next
    
    Select Case cllBanks.Count
    Case 1
        zl9LedVoice.DisplayBank varBank(1)
    Case 2
        zl9LedVoice.DisplayBank varBank(1), varBank(2)
    Case 3
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3)
    Case 4
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4)
    Case 5
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5)
    Case 6
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6)
    Case 7
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7)
    Case 8
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8)
    Case 9
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9)
    Case 10
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10)
    Case 11
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11)
    Case 12
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12)
    Case 13
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13)
    Case 14
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14)
    Case 15
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15)
    Case 16
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15), varBank(16)
    Case 17
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15), varBank(16), varBank(17)
    Case 18
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15), varBank(16), varBank(17), varBank(18)
    Case 19
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15), varBank(16), varBank(17), varBank(18), varBank(19)
    Case Else   '最多20个
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15), varBank(16), varBank(17), varBank(18), varBank(19), varBank(20)
    End Select
    If blnIsWait4 Then
        '延迟时间
        dtTmpDate = Time
        Do While Time < DateAdd("s", 4, dtTmpDate)  '这个在10.34.0中只有医保才加4秒，在10.35后中总费用也要加4秒，造成用户等待时间过长，因此还是恢复只有医保才增加4秒（为什么要增加,以前不详)
        Loop
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



Private Sub Class_Initialize()
    Err = 0: On Error Resume Next

End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    
    If mblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset mobjCom
    End If
    Set mobjCom = Nothing
    Set mfrmMain = Nothing
End Sub
