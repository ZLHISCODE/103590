VERSION 5.00
Begin VB.UserControl ucQRCodePayButton 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1530
   ScaleHeight     =   570
   ScaleWidth      =   1530
   ToolboxBitmap   =   "ucQRCodePayButton.ctx":0000
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
   Begin VB.Shape shapBack 
      BorderColor     =   &H8000000A&
      Height          =   465
      Left            =   30
      Top             =   0
      Width           =   945
   End
   Begin VB.Image imgHot 
      Height          =   240
      Left            =   690
      Picture         =   "ucQRCodePayButton.ctx":0312
      Top             =   135
      Width           =   240
   End
   Begin VB.Image ImgNormal 
      Height          =   240
      Left            =   1155
      Picture         =   "ucQRCodePayButton.ctx":089C
      Top             =   165
      Width           =   240
   End
End
Attribute VB_Name = "ucQRCodePayButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'功能:扫码付控件
'函数及方法:
'  一、开放方法及函数
'       1.zlInit-初始化方法(主要是给Oracle联接及创建扫码付进行设置)
'       2.zlReReadQRCode:重新读取二维码代码
'  二、内部方法及函数
'       1.GetQRCodePayment-获取扫码付对象
'       2.ReadQRCode-读取扫码付需要支付的二维码
'       3.SetBorderVisible:设置边框线的显示
'       4.DrawButtonStyle：绘制按钮的样式
'       5.ClearButtonTag:清除按钮内置信息
'       6.Refresh:重新刷新界面
'属性:
'   1.BorderStyle-控件是否带边框线
'   2.Appearance-控件的显示样式
'   3.CardTypeIDs-本机支持的卡类别IDs,多个用逗号分隔,只读属性（由方法：zlInit设置)
'   4.PicAlignment-图标对齐方式(相对于文本对齐)
'   5.CaptionAlignMent-文本对齐方式
'   6.Caption-按钮文本
'事件:
'   1.zlErrShow-错误显示（当发生错误时，触发该事件)
'   2.zlQRCodePayment-获取二维码成功或失败后，发生支付事件
'   3.zlGetPayMoney-获取本次支付的金额(在点击按钮时触发该事件)
'说明:
'  1.本部件需要使用时，必须要结合“zlQRCodePayMent.clsQRCodePayment”部件一起才有效
'  2.调用顺序:
'      1)首先：调用“zlInit”进行初始化，未初始化成功，则不允许使用该控件，界面层可以不显示该控件
'      2)其次，通过事件“GetQrCodePayment”返回本次要支付的金额
'      2)最后,通过事件"zlQRCodePayment"事件进行支付
'      3)如果发生错误，则通过“zlErrShow”事件显示相关的错误信息.
'      4)如果想通过快键实现，则调用"zlReReadQRCode"方法
'编制:刘兴洪
'日期:2019-03-04 19:19:10
'*********************************************************************************************************************************************

Public Enum PayButton_Appearance
    ShowFlat = 0
    Show3D = 1 '较浅的按钮
    ShowEdge3D = 2 '较深的按钮
End Enum
Public Enum PayButton_BorderStyle
    ShowNone = 0
    ShowFixed_Single = 1
End Enum

Public Enum PayButton_Alignment
    LeftAgnmt = 0
    CenterAgnmt
    RightAgnmt
End Enum

Public Enum PayButton_PicAlignment
    TxtLeftAgnmt = 0
    TxtDownCenterAgnmt = 1
    TxtTopCenterAgnmt = 2
    TxtRightAgnmt = 3
End Enum

'Const m_def_Enabled = 0
Const m_def_BorderStyle = 0
Const m_def_Appearance = 0

'属性变量:
Dim m_Caption As String
Dim m_CaptionAlignMent As PayButton_Alignment
Dim m_PicAlignMent As PayButton_PicAlignment

Dim m_CardTypeIDs As String '本机支持的卡类别ID
'Dim m_Enabled As Boolean
Dim m_BorderStyle As PayButton_BorderStyle
Dim m_Appearance As PayButton_Appearance

'事件声明:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Event zlErrShow(ByVal strErrMsg As String, ByVal lngErrNum As Long)
Event zlGetPayMoney(ByRef dblMoney As Double, ByRef strExpend As String, ByRef blnCancel As Boolean)
'blnCancel:入参时:true:表示读取二维码失败，否则读取二维码成功;出参时:表示终止本次支付
'lngCardTypeID:读取二维码成功时，为二维码支付的卡类别ID,否则为0
Event zlQRCodePayment(ByVal lngCardTypeID As Long, ByVal strPayMentQRCode As String, ByVal strExpendXML As String, ByRef blnCancel As Boolean)


'模块级变量
Private mcnOracle As ADODB.Connection
Private mstrDBUser As String
Private mlngSys As Long, mlngModul As Long
Private mobjQRCodePayment As Object '二维码支付或扫码支付对象
Private mfrmMain As Object
'缺省属性值:
Const m_def_Caption = ""
Const m_def_CaptionAlignMent = 0
Const m_def_PicAlignMent = 0
Const m_def_CardTypeIDs = ""

Public Function zlInit(ByVal frmMain As Object, ByVal strCardTypeIDs As String, Optional ByVal lngSys As Long, Optional ByVal lngModul As Long, _
    Optional cnOracle As ADODB.Connection, Optional strDBUser As String, Optional ByRef strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化
    '入参:frmMain-调用的主窗口
    '     lngSys : 系统编号
    '     lngModul:需要执行的功能序号
    '     cnOracle:主程序的数据库连接
    '     strCardTypeIDs-本机支持的扫码付的类别IDs(多个用逗号分隔),如:1,2,3...
    '出参:strErrMsg_out-返回的错误信息
    '编制:刘兴洪
    '日期:2019-03-04 19:36:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objQRCode As Object
    
    Set mcnOracle = cnOracle: mstrDBUser = strDBUser: mlngModul = lngModul: mlngSys = lngSys
    Set mfrmMain = frmMain: m_CardTypeIDs = strCardTypeIDs
    
    If GetQrCodePayment(objQRCode, strErrMsg_out) = False Then Exit Function
    zlInit = True
    Exit Function
errHandle:
    strErrMsg_out = Err.Description
    RaiseEvent zlErrShow(strErrMsg_out, Err.Number)
End Function
Public Function zlReReadQRCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新读取二维码代码
    '返回:读取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-03-11 20:14:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlReReadQRCode = ReadQRCode
End Function
Private Function GetQrCodePayment(ByRef objQRCode_Out As Object, Optional ByRef strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取扫码付对象
    '出参:objQRCode_Out-扫码付对象
    '     ErrMsg_out-获取扫码付对象失败
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-03-04 19:36:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnInitComponent As Boolean, strExpandXML As String
    If Not mobjQRCodePayment Is Nothing Then
        Set objQRCode_Out = mobjQRCodePayment
        GetQrCodePayment = True: Exit Function
    End If
    
    Err = 0: On Error Resume Next
    Set mobjQRCodePayment = CreateObject("zlReadQRCode.clsReadQRCode")    '固定名称
    If Err.Number <> 0 Then
        'strErrMsg_out = "创建扫码付对象(zlQRCodePayMent.clsQRCodePayment)失败,请检查该对象是否正确！"
        'RaiseEvent zlErrShow(strErrMsg_out, 0)
        Exit Function
    End If
    Err = 0: On Error GoTo ErrHand:
    '初始化接口部件
    '功能:zlInitComponents (初始化接口部件)
    '入参:   lngSys-传入的系统号
    '       strDBUser-数据库用户名
    '       cnOracle -HIS/三方机构
    '       strCardTypeIDs-本机支持的扫码付的类别IDs(多个用逗号分隔),如:1,2,3...
    '       strExpandXML-扩展信息,暂无,待以后扩展
    blnInitComponent = mobjQRCodePayment.zlInitComponents(mlngSys, mstrDBUser, mcnOracle, m_CardTypeIDs, strExpandXML)
    If Not blnInitComponent Then
        Exit Function
    End If
    Set objQRCode_Out = mobjQRCodePayment
    GetQrCodePayment = blnInitComponent
    Exit Function
ErrHand:
    strErrMsg_out = Err.Description
    RaiseEvent zlErrShow(strErrMsg_out, Err.Number)
End Function

Private Function ReadQRCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取扫码付需要支付的二维码
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-03-04 20:03:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objQRCode As Object, strErrMsg_out As String, lngCardTypeID_Out As Long, strQRCode_Out As String, strExpendXML As String
    Dim blnCancel As Boolean, dblMoney As Double, strExpend As String
    On Error GoTo errHandle
    
    '1.获取当前扫码付的支付金额
    
    blnCancel = False: strExpendXML = "": dblMoney = 0
    RaiseEvent zlGetPayMoney(dblMoney, strExpend, blnCancel) 'strExpend:暂无用处
    If blnCancel Or dblMoney = 0 Then Exit Function
    
    If dblMoney = 0 Then
        strErrMsg_out = "付款金额为零，不需要进行扫码付款!"
        RaiseEvent zlErrShow(strErrMsg_out, 0)
        Exit Function
    End If
    
    If dblMoney < 0 Then
        strErrMsg_out = "不支持负数的扫码付!"
        RaiseEvent zlErrShow(strErrMsg_out, 0)
        Exit Function
    End If
    
    strErrMsg_out = ""
    If GetQrCodePayment(objQRCode, strErrMsg_out) = False Then Exit Function
    '调用读取支付码
    '    zlReadQRCode(frmMain As Object, _
    '    ByVal lngModule As Long,
    '    ByVal dblMoney As Double,
    '    ByRef lngCardTypeID_Out As  Long, _
    '    ByRef strQRCode As String,byref strExpand As String, _) As Boolean
    '    '----------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读取支付的二维码代码接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpendXML-扩展参数,暂无用
    '    '出参:lngCardTypeID_Out-返回的支付卡类别ID
    '    '       strQRCode_Out-返回的支付码
    '    '       strExpendXML-待以后扩展
    '    '返回:函数返回    True:调用成功,False:调用失败
    If objQRCode.ReadQRCode(mfrmMain, mlngModul, dblMoney, lngCardTypeID_Out, strQRCode_Out, strExpendXML) = False Then
        blnCancel = True
        RaiseEvent zlQRCodePayment(0, "", "", blnCancel)
        Exit Function
    End If
    RaiseEvent zlQRCodePayment(lngCardTypeID_Out, strQRCode_Out, strExpendXML, blnCancel)
    ReadQRCode = Not blnCancel
    Exit Function
errHandle:
    strErrMsg_out = Err.Description
    RaiseEvent zlErrShow(strErrMsg_out, Err.Number)
End Function

 
Private Sub DrawButtonStyle(ByVal intAppearance As PayButton_Appearance, Optional blnHotImg As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绘制按钮的样式
    '入参:0=平面,-1=凹下,1=凸起(较浅的按钮),-2=深凹下,2=深凸起(较深的按钮)
    '编制:刘兴洪
    '日期:2019-03-04 17:25:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intAlignMent As gAlignment
    Dim imgPicture As IPictureDisp
    Dim intStyle As Integer
    
    intAlignMent = m_CaptionAlignMent
    
    If blnHotImg Or m_Appearance <> ShowFlat And intAppearance <> -1 And UserControl.Enabled Then
        Set imgPicture = imgHot.Picture
    Else
        Set imgPicture = ImgNormal.Picture
    End If
    
    'intStyle=0=平面,-1=凹下,1=凸起,-2=深凹下,2=深凸起
    'intType:0-所有;1-OnlyDown;2-OnlyKind
    Select Case intAppearance
    Case ShowFlat '平面
        intStyle = 0
    Case ShowEdge3D '较深的按钮
        intStyle = 2
    Case Show3D '较浅的按钮
        intStyle = 1
    Case Else
        intStyle = intAppearance
    End Select
    zlRaisEffectEx picButton, intStyle, m_Caption, m_CaptionAlignMent, imgPicture, m_PicAlignMent
End Sub
Private Sub ClearButtonTag()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除按钮内置信息
    '编制:刘兴洪
    '日期:2019-03-04 17:40:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
     shapBack.BorderColor = &H8000000A
     picButton.Tag = ""
End Sub

Private Sub SetBorderVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置边框线的显示
    '入参:
    '编制:刘兴洪
    '日期:2019-03-04 18:51:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    shapBack.Visible = BorderStyle = ShowFixed_Single And m_Appearance = ShowFlat
    Call UserControl_Resize
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=5
Public Sub Refresh()
     Call SetBorderVisible
     Call DrawButtonStyle(m_Appearance)
End Sub


Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    shapBack.Visible = False
    Call DrawButtonStyle(-1)  '绘制按钮:0=平面,-1=凹下,1=凸起(较浅的按钮),-2=深凹下,2=深凸起(较深的按钮)
    Call ClearButtonTag   '清除按钮内置信息
End Sub
Private Sub picButton_Resize()
    Err = 0: On Error Resume Next
    Call DrawButtonStyle(Appearance)    '绘制按钮
    Call ClearButtonTag   '清除按钮内置信息
End Sub


Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Appearance = ShowEdge3D Or Appearance = Show3D Then Exit Sub
    '只有平面的，才有会有
    If picButton.Tag = "In" Then
         If X < 0 Or Y < 0 Or X > picButton.Width Or Y > picButton.Height Then
             picButton.Tag = ""
             ReleaseCapture
             shapBack.BorderColor = &H8000000A
             Call DrawButtonStyle(m_Appearance)  '绘制按钮
             Call SetBorderVisible
         End If
     Else
         picButton.Tag = "In"
         SetCapture picButton.hWnd
         shapBack.BorderColor = vbBlue
         Call DrawButtonStyle(IIf(shapBack.Visible, ShowFlat, Show3D), True)    '绘制按钮
     End If
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    If Button <> 1 Then Exit Sub
    Call DrawButtonStyle(m_Appearance)  '绘制按钮
    Call ClearButtonTag
    Call SetBorderVisible   '显示边框线
    Call ReadQRCode
   ' RaiseEvent Click
End Sub
Private Sub UserControl_ExitFocus()
    '平面,恢复平面
    Call DrawButtonStyle(m_Appearance)  '绘制按钮
End Sub
'
'
''注意！不要删除或修改下列被注释的行！
''MemberInfo=0,0,0,0
'Public Property Get Enabled() As Boolean
'    Enabled = m_Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    m_Enabled = New_Enabled
'    PropertyChanged "Enabled"
'End Property
 
'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As PayButton_BorderStyle
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As PayButton_BorderStyle)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Call SetBorderVisible
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=1,0,0,0
Public Property Get Appearance() As PayButton_Appearance
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As PayButton_Appearance)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    
    Call DrawButtonStyle(m_Appearance)  '绘制按钮
    Call ClearButtonTag
End Property
Private Sub UserControl_Resize()
    Dim lngX As Long, lngY As Long
    Err = 0: On Error Resume Next
    
    With shapBack
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    
    With UserControl
        lngX = 0: lngY = 0
        If shapBack.Visible Then
            lngX = shapBack.BorderWidth * Screen.TwipsPerPixelX
            lngY = shapBack.BorderWidth * Screen.TwipsPerPixelY
        End If
        picButton.Move .ScaleLeft + lngX, .ScaleTop + lngY, .ScaleWidth - (.ScaleLeft + lngX * 2), .ScaleHeight - (.ScaleTop + lngX * 2)
    End With
End Sub

Private Sub UserControl_Initialize()
    Err = 0: On Error Resume Next
    Call DrawButtonStyle(m_Appearance)  '绘制按钮
    Call ClearButtonTag
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
'    m_Enabled = m_def_Enabled
    m_BorderStyle = m_def_BorderStyle
    m_Appearance = m_def_Appearance
    Call SetBorderVisible
    m_CardTypeIDs = m_def_CardTypeIDs
    m_CaptionAlignMent = m_def_CaptionAlignMent
    m_PicAlignMent = m_def_PicAlignMent
    m_Caption = m_def_Caption
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_CardTypeIDs = PropBag.ReadProperty("CardTypeIDs", m_def_CardTypeIDs)
    
    Call SetBorderVisible
    m_CaptionAlignMent = PropBag.ReadProperty("CaptionAlignMent", m_def_CaptionAlignMent)
    m_PicAlignMent = PropBag.ReadProperty("PicAlignMent", m_def_PicAlignMent)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Call DrawButtonStyle(m_Appearance)  '绘制按钮
    Call ClearButtonTag
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'    picButton.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    picButton.ToolTipText = PropBag.ReadProperty("ToolTipString", "")
End Sub



Private Sub UserControl_Terminate()
    Err = 0: On Error Resume Next
    
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If Not mobjQRCodePayment Is Nothing Then Set mobjQRCodePayment = Nothing
    If Not mfrmMain Is Nothing Then Set mfrmMain = Nothing
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("CardTypeIDs", m_CardTypeIDs, m_def_CardTypeIDs)
    Call PropBag.WriteProperty("CaptionAlignMent", m_CaptionAlignMent, m_def_CaptionAlignMent)
    Call PropBag.WriteProperty("PicAlignMent", m_PicAlignMent, m_def_PicAlignMent)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'    Call PropBag.WriteProperty("ToolTipText", picButton.ToolTipText, "")
    Call PropBag.WriteProperty("ToolTipString", picButton.ToolTipText, "")
End Sub
'
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
 
'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,1,1,
Public Property Get CardTypeIDs() As String
    CardTypeIDs = m_CardTypeIDs
End Property

Public Property Let CardTypeIDs(ByVal New_CardTypeIDs As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_CardTypeIDs = New_CardTypeIDs
    PropertyChanged "CardTypeIDs"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=1,0,0,0
Public Property Get CaptionAlignment() As PayButton_Alignment
    CaptionAlignment = m_CaptionAlignMent
End Property

Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As PayButton_Alignment)
    m_CaptionAlignMent = New_CaptionAlignment
    PropertyChanged "CaptionAlignMent"
    Call DrawButtonStyle(m_Appearance)  '绘制按钮
    Call ClearButtonTag
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=1,0,0,0
Public Property Get PicAlignMent() As PayButton_PicAlignment
    PicAlignMent = m_PicAlignMent
End Property

Public Property Let PicAlignMent(ByVal New_PicAlignMent As PayButton_PicAlignment)
    m_PicAlignMent = New_PicAlignMent
    PropertyChanged "PicAlignMent"
    
    Call DrawButtonStyle(m_Appearance)  '绘制按钮
    Call ClearButtonTag
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Call DrawButtonStyle(m_Appearance)  '绘制按钮
    Call ClearButtonTag
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
    Call DrawButtonStyle(m_Appearance)  '绘制按钮
    Call ClearButtonTag
End Property
 
'注意！不要删除或修改下列被注释的行！
'MappingInfo=picButton,picButton,-1,ToolTipText
Public Property Get ToolTipString() As String
Attribute ToolTipString.VB_Description = "返回/设置当鼠标在控件上暂停时显示的文本。"
    ToolTipString = picButton.ToolTipText
End Property
Public Property Let ToolTipString(ByVal New_ToolTipString As String)
    picButton.ToolTipText() = New_ToolTipString
    PropertyChanged "ToolTipString"
End Property

