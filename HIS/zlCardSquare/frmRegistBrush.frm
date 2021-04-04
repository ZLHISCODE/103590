VERSION 5.00
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#3.4#0"; "zlIDKind.ocx"
Begin VB.Form frmRegistBrush 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "挂号刷卡验证"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   375
      Left            =   800
      TabIndex        =   9
      Top             =   1515
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   661
      Appearance      =   2
      IDKindStr       =   $"frmRegistBrush.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "宋体"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   5805
      TabIndex        =   5
      Top             =   0
      Width           =   5805
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   6000
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "剩余款额：1000.00，本次金额：1000.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   465
         Width           =   4320
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4845
         Picture         =   "frmRegistBrush.frx":009D
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人：张永康，男，30岁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   255
         TabIndex        =   6
         Top             =   105
         Width           =   2640
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   2700
      Width           =   6900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4485
      TabIndex        =   3
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3270
      TabIndex        =   2
      Top             =   2865
      Width           =   1100
   End
   Begin VB.TextBox txtCard 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1455
      TabIndex        =   1
      Top             =   1500
      Width           =   3015
   End
   Begin VB.CommandButton cmdReadIC 
      Caption         =   "读卡"
      Height          =   405
      Left            =   4500
      TabIndex        =   0
      Top             =   1500
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卡号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   1590
      Width           =   570
   End
End
Attribute VB_Name = "frmRegistBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mintCount As Integer
'Private mobjICCard As Object 'IC卡对象
Private mlng病人ID As String
Private mlngSys As Long
Private mblnTest As Boolean
Private mblnPreCard As Boolean
Private mblnUnload As Boolean

'--------------------------------------------------
'卡相关:
Private mobjKeyboard As Object
Private mblnPassInputCardNo As Boolean  '是否密文输入卡号
Private mobjSquareCard As Object
Private mlng医疗卡长度 As Long
Private mlngModul As Long
Private mstrPassWord As String
Private mlngDefaultCardTypeID As Long '缺省的刷卡类别ID
Private mblnBrushCard As Boolean
Private mlngCardTypeID As Long  '卡类别ID
Private mstrCardNo As String    '卡号
Private mobjPatiCardObject As clsCardObject
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'--------------------------------------------------
Public Function ShowMe(frmParent As Object, ByVal lngSys As Long, ByVal lng病人ID As Long, _
    ByVal cur金额 As Currency, Optional lngModul As Long = 0, _
    Optional lngCardTypeID As Long, Optional ByRef strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证窗体入口
    '入参:frmParent-调用的主窗体
    '       lngSys-系统号
    '       lng病人ID-指定的病人ID
    '       lngModul-模块号
    '       lngCardTypeID-缺省卡类别ID
    '出参:strCardNo-返回卡号
    '       lngCardTypeID-卡类别ID
    '返回:验证成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-10 16:35:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, rsTemp As ADODB.Recordset
     
    Dim strSQL As String, intMouse As Integer
    mlngSys = lngSys: mlngModul = lngModul: mlngDefaultCardTypeID = lngCardTypeID
    mblnOK = False: mintCount = 3: mlng病人ID = lng病人ID
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0
    mblnTest = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard", "TestCardNO", 0)) = 1
    mblnTest = IsDesinMode Or mblnTest
 
    '读取就诊卡信息
    On Error GoTo errH
    strSQL = "" & _
    "   Select A.姓名,A.性别,A.年龄,A.就诊卡号,A.卡验证码, " & _
    "              nvl(B.余额,0) as 余额" & _
    "   From 病人信息 A, " & _
    "       (   Select 病人ID,nvl(Sum(预交余额),0)-nvl(sum(费用余额),0) as 余额 " & _
    "           From  病人余额 " & _
    "           Where 病人ID=[1] and 性质=1 and decode([2],0,0,类型)=[2]  Group by 病人ID) B " & _
    "   Where A.病人ID=[1] And A.病人ID=B.病人ID(+) "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng病人ID, 1)
    If rsTmp.EOF Then
        MsgBox "病人信息不存在,请检查!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If IIf(IsNull(rsTmp!就诊卡号), "", rsTmp!就诊卡号 & "") = "" Then
        '问题:43449
        strSQL = "Select Count(Distinct 卡类别ID) as 类别数 From 病人医疗卡信息 Where  病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng病人ID)
        If IIf(IsNull(rsTemp!类别数), 0, Val(rsTemp!类别数 & "")) = 0 Then
            '--未发卡,直接返回true,不用验卡
            ShowMe = False: Exit Function
        End If
    End If
    
    Me.lblPati.Caption = "病人：" & zlCommFun.Nvl(rsTmp!姓名) & _
        IIf(Not IsNull(rsTmp!性别), "，" & rsTmp!性别, "") & _
        IIf(Not IsNull(rsTmp!年龄), "，" & rsTmp!年龄, "")
    Me.lblMoney.Caption = "剩余款额：" & Format(rsTmp!余额, "0.00") & "，本次金额：" & Format(cur金额, "0.00")
    Me.txtCard.Tag = zlCommFun.Nvl(rsTmp!就诊卡号)
    mstrCardNo = "": lngCardTypeID = 0
    On Error GoTo 0
    'IC卡对象
    On Error Resume Next
    'Set mobjICCard = CreateObject("zlICCard.clsICCard")
    On Error GoTo 0
    Me.Show 1, frmParent
    ShowMe = mblnOK
    lngCardTypeID = mlngCardTypeID
    strCardNo = mstrCardNo
    Screen.MousePointer = intMouse
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub CmdCancel_Click()
    Unload Me
End Sub
Private Sub CmdOK_Click()
    Dim strPassWord As String
    If UCase(txtCard.Text) <> UCase(txtCard.Tag) Then
        MsgBox "当前卡号与病人的卡号不相符！", vbExclamation, gstrSysName
        Unload Me: Exit Sub '卡号不匹配，不准重试
    End If
    If Val(cmdReadIC.Tag) <> mlng病人ID Or Val(cmdReadIC.Tag) = 0 Then
        MsgBox "当前卡号与病人的卡号不相符！", vbExclamation, gstrSysName
        Unload Me: Exit Sub '卡号不匹配，不准重试
    End If
    mstrCardNo = txtCard.Text
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdReadIC_Click()
    Call IDKind_Click(IDKind.GetCurCard)
End Sub

Private Sub Form_Activate()
    If IDKind.ListCount = 0 Then Unload Me: Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
       Select Case KeyCode
        Case vbKeyF4
            If IDKind.Enabled Then
                If Shift = vbShiftMask Then
                    IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDKindStr, ";")), IDKind.IDKind - 1)
                Else
                    IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDKindStr, ";")), 0, IDKind.IDKind + 1)
                End If
            End If
        End Select
End Sub
Private Sub Form_Load()
    Call CreateObjectKeyboard
    Call zlInitData
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Set mobjICCard = Nothing
    Set mobjKeyboard = Nothing
End Sub

 

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNo As String, strExpand
    Dim strOutPatiInforXml As String
     If objCard Is Nothing Then Exit Sub
    If IsCardType(IDKind, "IC卡号") Then
        Exit Sub
    End If
    lng卡类别ID = objCard.接口序号
    mlngCardTypeID = lng卡类别ID
    Call CreatePayObject
    If lng卡类别ID = 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    strExpand = lng卡类别ID
    If mobjSquareCard.zlReadCard(Me, mlngModul, True, strExpand, strOutCardNo, strOutPatiInforXml) = False Then Exit Sub
    txtCard.Text = strOutCardNo
    '问题号:42948
    If txtCard.Text <> "" Then
        If GetPatient(Trim(txtCard.Text)) = False Then
                txtCard.Text = ""
                If txtCard.Enabled Then txtCard.SetFocus
                zlControl.TxtSelAll txtCard
                Exit Sub
        End If
     End If
     If txtCard.Text <> "" Then
        Call CmdOK_Click
     Else
         txtCard.SetFocus
     End If
End Sub

'获取idkind的默认kind值
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind的默认Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.名称)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

'控件名称是否匹配
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
     End Select
End Function
                

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
    mlng医疗卡长度 = objCard.卡号长度
    '第7位后,就只能用索引,不然取不到数
    mblnPassInputCardNo = IDKind.ShowPassText
    txtCard.MaxLength = mlng医疗卡长度
    txtCard.PasswordChar = IIf(mblnPassInputCardNo, "*", "")
    '85565,李南春,2015/7/19:读卡性质
'    mblnBrushCard = Mid(objCard.读卡性质, 1, 1) = 0 And Mid(objCard.读卡性质, 2, 1) = 0
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtCard.Text <> "" Then txtCard.Text = ""
    txtCard.Locked = Not (objCard.是否刷卡 Or objCard.是否扫描)
    cmdReadIC.Visible = objCard.是否接触式读卡
    If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtCard.Locked Or txtCard.Text <> "" Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
     

    intIndex = IDKind.GetKindIndex(objCard.名称)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex

    txtCard.Text = objPatiInfor.卡号
    Call txtCard_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
   
End Sub

'
'Private Sub txtPass_LostFocus()
'    ClosePassKeyboard txtPass
'End Sub
Private Sub txtCard_Change()
    txtCard.Tag = "": cmdReadIC.Tag = ""
    'lblPass.Tag = "":
    'txtPass.Enabled = txtCard.Text <> ""
    'If Not txtPass.Enabled Then txtPass.Text = ""
End Sub

Private Sub txtCard_GotFocus()
    Call zlControl.TxtSelAll(txtCard)
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    mblnPreCard = False

    '是否刷卡完成
    blnCard = KeyAscii <> 8 And Len(txtCard.Text) = mlng医疗卡长度 - 1 And txtCard.SelLength <> Len(txtCard.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txtCard.Text = txtCard.Text & Chr(KeyAscii)
            txtCard.SelStart = Len(txtCard.Text)
        End If
        KeyAscii = 0
        If GetPatient(Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            zlControl.TxtSelAll txtCard
            Exit Sub
       End If
       mblnPreCard = blnCard
       If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
       If blnCard Then Call CmdOK_Click
       Exit Sub
    Else
        If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        
        If mblnTest Then Exit Sub
        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = timer
            If txtCard.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(txtCard.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                txtCard.Text = Chr(KeyAscii)
                txtCard.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
End Sub

Private Sub txtCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtCard.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtCard.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtCard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtCard.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
'Private Sub txtPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button <> 2 Or mblnTest Then Exit Sub
'    glngTXTProc = GetWindowLong(txtPass.hWnd, GWL_WNDPROC)
'    Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, AddressOf WndMessage)
'End Sub
'
'Private Sub txtPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'     If Button <> 2 Or mblnTest Then Exit Sub
'    Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, glngTXTProc)
'End Sub
'
'Private Sub txtPass_GotFocus()
'    Call zlControl.TxtSelAll(txtPass)
'    OpenPassKeyboard txtPass
'End Sub
'
'Private Sub txtPass_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        If mblnPreCard Then
'             If (GetAsyncKeyState(VK_RETURN) And &H1) <> 0 Then
'                txtPass.Text = ""
'                Exit Sub
'             End If
'        End If
'        mblnPreCard = False
'        Call cmdOK_Click
'    ElseIf KeyAscii = 22 Then
'        KeyAscii = 0 '不允许粘贴
'    Else
'        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
'                KeyAscii = 0 '去除特殊符号，并且不允许粘贴
'        End If
'    End If
'
'End Sub

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function OpenPassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Sub CreatePayObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建支付对象接口
    '编制:刘兴洪
    '日期:2011-06-22 13:15:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng卡类别ID As Long, bln消费卡 As Boolean, int自动读取 As Integer
    Dim strKey As String
    Dim i As Long
    Set mobjSquareCard = Nothing:
    Err = 0: On Error Resume Next
    If zlGetCardObj(Me, mlngCardTypeID, False, mobjPatiCardObject) = False Then
        Set mobjPatiCardObject = Nothing
        Set mobjSquareCard = Nothing
        Exit Sub
    End If
    Set mobjSquareCard = mobjPatiCardObject.CardObject
    If Err <> 0 Then
        MsgBox "未找到" & IDKind.GetCurCard.名称 & "所对应的部件,请检查", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If mobjSquareCard Is Nothing Then Exit Sub
End Sub
Private Sub zlInitData()
    Dim strExpend As String, i As Integer
    Dim strKey As String
    Dim lngCardID As Long
    strKey = GetIDKindStr("", True)
    If strKey = "" Then
        mblnUnload = True
        Exit Sub
    End If
    IDKind.IDKindStr = strKey
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, "", txtCard)
    IDKind.ShowPropertySet = InStr(";" & gstrPrivs & ";", "参数设置") > 0
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
End Sub
Private Function GetPatient(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-26 00:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    On Error GoTo errH
    mstrPassWord = ""
    lng卡类别ID = IDKind.GetCurCard.接口序号
    If lng卡类别ID = 0 Then
      If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
    Else
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If GetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
    End If
    If lng病人ID <= 0 Then GoTo NotFoundPati:
    If mlng病人ID <> lng病人ID Then
       MsgBox "当前卡号与病人所持有的卡号不相符,请检查！", vbExclamation, gstrSysName
       txtCard.Text = ""
       Exit Function '卡号不匹配，不准重试
    End If
    txtCard.Tag = strInput
    cmdReadIC.Tag = lng病人ID
    mstrPassWord = strPassWord
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
NotFoundPati:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
    Else
        MsgBox "未找到当前卡的持有病人,请检查!", vbOKOnly + vbInformation, gstrSysName
    End If
    txtCard.Tag = "": cmdReadIC.Tag = ""
End Function
Private Function IsDesinMode() As Boolean
      '刘兴洪 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function

