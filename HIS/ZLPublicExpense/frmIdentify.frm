VERSION 5.00
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmIdentify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病人身份验证"
   ClientHeight    =   4008
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   6540
   Icon            =   "frmIdentify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4008
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2325
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1350
      Width           =   3015
   End
   Begin VB.TextBox txtCard 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2325
      TabIndex        =   1
      Top             =   1912
      Width           =   3015
   End
   Begin VB.TextBox txtPass 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2340
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2475
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3420
      TabIndex        =   3
      Top             =   3450
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3450
      Width           =   1100
   End
   Begin VB.Frame fraDown 
      Height          =   30
      Left            =   -30
      TabIndex        =   9
      Top             =   3225
      Width           =   7290
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      ScaleHeight     =   1056
      ScaleWidth      =   6540
      TabIndex        =   10
      Top             =   0
      Width           =   6540
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   1140
         X2              =   8715
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   1140
         X2              =   8715
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Label lblFamilyRest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家属余额:9999999.00"
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
         Left            =   4260
         TabIndex        =   17
         Tag             =   "家属余额:"
         Top             =   750
         Width           =   2280
      End
      Begin VB.Label lblRest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人余额:9999999.00"
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
         Left            =   1140
         TabIndex        =   16
         Tag             =   "病人余额:"
         Top             =   750
         Width           =   2280
      End
      Begin VB.Label lblPatiType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人类型:普通患者"
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
         Left            =   1140
         TabIndex        =   15
         Tag             =   "病人类型:"
         Top             =   420
         Width           =   2040
      End
      Begin VB.Label lblFeeType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别:普通"
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
         Left            =   4740
         TabIndex        =   14
         Tag             =   "费别:"
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄:30岁"
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
         Left            =   4740
         TabIndex        =   13
         Tag             =   "年龄:"
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别:未知"
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
         Left            =   3330
         TabIndex        =   12
         Tag             =   "性别:"
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:琪玛多吉"
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
         Left            =   1140
         TabIndex        =   11
         Tag             =   "姓名:"
         Top             =   90
         Width           =   1560
      End
      Begin VB.Image Image1 
         Height          =   576
         Left            =   240
         Picture         =   "frmIdentify.frx":058A
         Top             =   132
         Width           =   576
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -105
         X2              =   7470
         Y1              =   1035
         Y2              =   1035
      End
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   420
      Left            =   1665
      TabIndex        =   7
      Top             =   1905
      Width           =   630
      _ExtentX        =   1101
      _ExtentY        =   741
      Appearance      =   2
      IDKindStr       =   "就|就诊卡|0|0|0|0|0|;IC|IC卡号|1|0|0|0|0|"
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
      ShowPropertySet =   -1  'True
      NotContainFastKey=   ""
      AllowAutoICCard =   -1  'True
      AllowAutoIDCard =   -1  'True
      BackColor       =   -2147483633
      SaveRegType     =   4
      ProductName     =   "一卡通消费支付"
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "刷卡金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1155
      TabIndex        =   5
      Top             =   1425
      Width           =   1140
   End
   Begin VB.Label lblCardNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卡号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   6
      Top             =   1980
      Width           =   570
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密  码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1425
      TabIndex        =   8
      Top             =   2580
      Width           =   870
   End
End
Attribute VB_Name = "frmIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mintCount As Integer
Private mstr病人IDs As String
Private mlngSys As Long
Private mblnPreCard As Boolean
Private mobjCard As Card '当前处理的卡
'--------------------------------------------------
'卡相关:
Private mobjKeyboard As Object
Private mobjOneCardComLib As zlOneCardComLib.clsOneCardComLib
Private mlngModul As Long
Private mstrPassWord As String
Private mlngDefaultCardTypeID As Long '缺省的刷卡类别ID
Private mblnBrushCard As Boolean
Private Const VK_RETURN = &HD
Private mblnCheckPassWord As Boolean
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mstrRegSection As String
Private mlngPreBrushCardTypeID As Long '上次刷卡类别

Public Function ShowMe(frmParent As Object, ByVal lngSys As Long, ByVal lng病人ID As Long, _
    ByVal cur金额 As Currency, Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0, _
    Optional lngDefaultCardTypeID As Long = 0, _
    Optional blnCheckPassWord As Boolean = True, _
    Optional blnFamilyMoney As Boolean, _
    Optional strFamilyPatiIDs As String = "", _
    Optional bln刷卡验证 As Boolean = True, _
    Optional bln无密码不验卡 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证窗体入口
    '入参:frmParent-调用的主窗体
    '       lngSys-系统号
    '       lng病人ID-指定的病人ID
    '       lngModul-模块号
    '       bytOperationType-业务类型(0-不区分;1-门诊;2-住院)
    '       mlngDefaultCardTypeID-缺省的刷卡类别ID
    '       blnCheckPassWord-验证密码(true-验证密码,false-只刷卡,不输入密码)
    '       blnFamilyMoney-是否读取家属预交余额
    '       strFamilyPatiIDs-病人家属的病人ID
    '       bln刷卡验证-是否进行刷卡验证，主要用于不刷卡验证时读取家属IDs
    '       bln无密码不验卡-病人的所有医疗卡都没有设置密码时是否验卡，当为True时，只要有一张卡设置了密码都要进行验卡,112418
    '出参:
    '返回:验证成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-10 16:35:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strSQL As String, intMouse As Integer
    Dim cur家属余额 As Currency, cur病人余额 As Currency
    Dim objPati As clsPatientInfo, objExpenceSvr As clsExpenceSvr
    Dim cllPatiFee As Collection, objPatiFee As clsPatiFeeinfor, i As Long
    
    mblnCheckPassWord = blnCheckPassWord
    mlngSys = lngSys: mlngModul = lngModul: mlngDefaultCardTypeID = lngDefaultCardTypeID
    mblnOk = False: mintCount = 3: mstr病人IDs = lng病人ID
    
    intMouse = Screen.MousePointer
    strFamilyPatiIDs = ""
    Screen.MousePointer = 0
    
    '读取就诊卡信息
    On Error GoTo errH
    If zlGetOneCardComLibObject(Me, mlngModul, mobjOneCardComLib) = False Then Exit Function
    
    If blnFamilyMoney Then '获取家属信息
        If mobjOneCardComLib.ZlGetPatiFamilyMember(0, lng病人ID, strFamilyPatiIDs) = False Then Exit Function
        If strFamilyPatiIDs <> "" Then mstr病人IDs = mstr病人IDs & "," & strFamilyPatiIDs
    End If
    
    '不用刷卡验证直接返回
    If Not bln刷卡验证 Then ShowMe = True: Exit Function
    
    '检查病人及家属是否有卡，只要其中任何一人有卡都需要刷卡，79868
    '问题:43449，如果病人没有发卡的,则允许不输入密码及刷卡操作,直接进行扣款
    If mobjOneCardComLib.ZlGetPatiCardInfo(mstr病人IDs, rsTemp) = False Then Exit Function
    If rsTemp.EOF Then
        '无记录,直接返回true,不用验卡
        ShowMe = True: Exit Function
    Else
        rsTemp.Filter = "密码<>'' And 密码<>null"
        If rsTemp.RecordCount = 0 And bln无密码不验卡 Then
            '所有卡都无密码,直接返回true,不用验卡
            ShowMe = True: Exit Function
        End If
    End If
    
    '获取病人信息
    If mobjOneCardComLib.zlGetPatiInforFromPatiID(lng病人ID, objPati) = False Then
        MsgBox "获取病人信息失败，请检查!", vbOKOnly, gstrSysName
        Exit Function
    End If
    lblName.Caption = lblName.Tag & objPati.姓名
    lblSex.Caption = lblSex.Tag & objPati.性别
    lblAge.Caption = lblAge.Tag & objPati.年龄
    lblPatiType.Caption = lblPatiType.Tag & objPati.病人类型
    lblFeeType.Caption = lblFeeType.Tag & objPati.费别
    
    '获取病人及家属余额
    cur病人余额 = 0: cur家属余额 = 0
    Set objExpenceSvr = New clsExpenceSvr
    If objExpenceSvr.zlInitCommon(glngSys, mlngModul, gcnOracle, gstrDBUser) = False Then Exit Function
    If objExpenceSvr.zlExseSvr_GetRemainMoneyByBatch(mstr病人IDs, bytOperationType, cllPatiFee) = False Then Exit Function
    For i = 1 To cllPatiFee.Count
        Set objPatiFee = cllPatiFee(i)
        If objPatiFee.病人ID = lng病人ID Then
            cur病人余额 = cur病人余额 + objPatiFee.剩余款
        Else
            cur家属余额 = cur家属余额 + objPatiFee.剩余款
        End If
    Next
    lblRest.Caption = lblRest.Tag & Format(cur病人余额, "0.00")
    lblFamilyRest.Caption = lblFamilyRest.Tag & Format(cur家属余额, "0.00")
    
    txtMoney.Text = Format(cur金额, "0.00")
    Me.Show 1, frmParent
    ShowMe = mblnOk
    
    Screen.MousePointer = intMouse
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡的有效性
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-19 17:04:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPassWord As String
    Dim str名称 As String
    On Error GoTo errHandle
    
    If mobjCard Is Nothing Then Exit Function
    If mobjCard.名称 Like "*卡号" Then
        str名称 = mobjCard.名称
    ElseIf mobjCard.名称 Like "*身份证" Then
        str名称 = "身份证号"
    ElseIf mobjCard.名称 Like "*卡" Then
        str名称 = mobjCard.名称 & "卡号"
    Else
        str名称 = mobjCard.名称 & "卡卡号"
    End If

    If UCase(Trim(txtCard.Text)) = "" Then Exit Function
    If Not InStr("," & mstr病人IDs & ",", "," & Val(lblPass.Tag) & ",") > 0 Or Val(lblPass.Tag) = 0 Then
        MsgBox "当前" & str名称 & "与病人的" & str名称 & "不相符！", vbExclamation, gstrSysName
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Function
    End If
    
    If Not mblnCheckPassWord Then isValied = True: Exit Function
    strPassWord = gobjComlib.zlCommFun.zlStringEncode(txtPass.Text)
    If strPassWord <> mstrPassWord Then
        If mintCount = 1 Then
            MsgBox "三次密码输入错误,不能再输入！", vbExclamation, gstrSysName
        Else
            MsgBox "密码输入错误！", vbExclamation, gstrSysName
        End If
        txtPass.Text = "": mintCount = mintCount - 1
        If mintCount = 0 Then
            Unload Me '密码错误，可输入2次
        ElseIf txtPass.Enabled Then
            txtPass.SetFocus
        End If
        Exit Function
    End If
    isValied = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IDKind.ActiveFastKey
End Sub
Private Sub Form_Load()
    Dim intIdKind As Integer
    
    mstrRegSection = "私有模块\" & gstrDBUser & "\界面设置\" & Me.Name & Me.Name
    mlngPreBrushCardTypeID = GetSetting("ZLSOFT", mstrRegSection, "缺省卡类别ID", 0)
    
    Call CreateObjectKeyboard
    Call IDKind.zlInit(Me, mlngSys, mlngModul, gcnOracle, gstrDBUser, mobjOneCardComLib, "", txtCard)
    If mlngPreBrushCardTypeID <> 0 Then
       intIdKind = IDKind.GetKindIndex(mlngPreBrushCardTypeID)
       If intIdKind <> 0 Then
           IDKind.IDKind = intIdKind
       End If
    End If
    
    Call SetCtrlVisible
    HookDefend txtPass.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not IDKind.GetCurCard Is Nothing Then
         SaveSetting "ZLSOFT", mstrRegSection, "缺省卡类别ID", IDKind.GetCurCard.接口序号
    End If
    
    On Error Resume Next
    Set mobjKeyboard = Nothing
    Set mobjCard = Nothing
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As Card)
    txtCard.PasswordChar = ""
    '85565,李南春,2015/7/10:读卡性质
    mblnBrushCard = objCard.是否刷卡 Or objCard.是否扫描
    If txtCard.Text <> "" Then txtCard.Text = ""
    txtCard.Locked = Not mblnBrushCard
    If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As Card, objPatiInfor As clsPatientInfo, blnCancel As Boolean)
    txtCard.Text = objPatiInfor.卡号
    
    If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComlib.zlControl.TxtSelAll txtCard
            Exit Sub
    End If
    
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
    If txtCard.Text = "" Then
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Sub
    End If
    If mblnCheckPassWord Then txtPass.SetFocus: Exit Sub
    Call cmdOK_Click
End Sub

Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub

Private Sub txtCard_Change()
    lblPass.Tag = "": txtCard.Tag = ""
    txtPass.Enabled = txtCard.Text <> ""
    If Not txtPass.Enabled Then txtPass.Text = ""
End Sub

Private Sub txtCard_GotFocus()
    Call gobjComlib.zlControl.TxtSelAll(txtCard)
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    mblnPreCard = False

    '是否刷卡完成
    blnCard = KeyAscii <> 8 And Len(txtCard.Text) = IDKind.GetCurCard.卡号长度 - 1 And txtCard.SelLength <> Len(txtCard.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txtCard.Text = txtCard.Text & Chr(KeyAscii)
            txtCard.SelStart = Len(txtCard.Text)
        End If
        KeyAscii = 0
        If GetPatient(IDKind.GetCurCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComlib.zlControl.TxtSelAll txtCard
            Exit Sub
        End If
        mblnPreCard = blnCard
        If mblnCheckPassWord Then
            If txtPass.Enabled Then txtPass.SetFocus
        Else
            Call cmdOK_Click: Exit Sub
        End If
    Else
        If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If

        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 And IDKind.GetCurCard.是否持卡消费 = True Then
            sngNow = Timer
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

Private Sub txtPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    glngTXTProc = GetWindowLong(txtPass.hWnd, GWL_WNDPROC)
    Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, AddressOf WndMessage)
End Sub

Private Sub txtPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button <> 2 Then Exit Sub
    Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, glngTXTProc)
End Sub

Private Sub txtPass_GotFocus()
    If txtCard.Text <> "" And mstrPassWord = "" Then Call cmdOK_Click: Exit Sub
    Call gobjComlib.zlControl.TxtSelAll(txtPass)
    OpenPassKeyboard txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If mblnPreCard Then
            '60580
            mblnPreCard = False
             If (GetAsyncKeyState(VK_RETURN) And &H1) <> 0 Then
                txtPass.Text = ""
                Exit Sub
             End If
        End If
        mblnPreCard = False
        Call cmdOK_Click
    ElseIf KeyAscii = 22 Then
        KeyAscii = 0 '不允许粘贴
    Else
        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        End If
    End If
    '60580
    mblnPreCard = False
End Sub

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
    If gobjComlib.ErrCenter() = 1 Then
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
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    Optional blnIDCard As Boolean = False, Optional blnICCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人信息
    '入参:objCard-按指定的卡类别进行读卡
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-26 00:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long
    
    On Error GoTo errH
    
    mstrPassWord = ""
    Set mobjCard = Nothing
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Function
    If mobjOneCardComLib.zlGetPatiID(lng卡类别ID, strInput, True, lng病人ID, strPassWord, strErrMsg, lng卡类别ID, Nothing, Me, False, True) = False Then
        '进行模糊查找:-1:医疗卡类别(但是如果当前的卡号长度不够的话,会存在问题)
        If mobjOneCardComLib.zlGetPatiID(-1, strInput, True, lng病人ID, strPassWord, strErrMsg, lng卡类别ID, Nothing, Me, False, True) = False Then
            GoTo NotFoundPati:
        End If
    End If
    If lng病人ID <= 0 Then GoTo NotFoundPati:
    If Not InStr("," & mstr病人IDs & ",", "," & lng病人ID & ",") > 0 Then
        If objCard.名称 Like "*卡号" Then
            MsgBox "当前" & objCard.名称 & "与病人所持有的" & objCard.名称 & "不相符,请检查！", vbExclamation, gstrSysName
        ElseIf objCard.名称 Like "*身份证" Then
            MsgBox "当前身份证号与病人所持有的身份证号不相符,请检查！", vbExclamation, gstrSysName
        ElseIf objCard.名称 Like "*卡" Then
            MsgBox "当前" & objCard.名称 & "卡号与病人所持有的" & objCard.名称 & "卡号不相符,请检查！", vbExclamation, gstrSysName
        Else
            MsgBox "当前" & objCard.名称 & "卡卡号与病人所持有的" & objCard.名称 & "卡卡号不相符,请检查！", vbExclamation, gstrSysName
        End If
        txtCard.Text = ""
        Exit Function '卡号不匹配，不准重试
    End If
    txtCard.Tag = strInput
    lblPass.Tag = lng病人ID
    mstrPassWord = strPassWord
    Set mobjCard = objCard
    GetPatient = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Exit Function
NotFoundPati:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
    Else
        MsgBox "未找到当前卡的持有病人,请检查!", vbOKOnly + vbInformation, gstrSysName
        txtCard.Text = ""
    End If
    txtCard.Tag = "": lblPass.Tag = ""
End Function

Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的visible属性
    '编制:刘兴洪
    '日期:2012-03-13 11:28:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    lblFamilyRest.Visible = InStr(mstr病人IDs, ",") > 0 '没有家属则隐藏家属余额的显示，79868
    lblPass.Visible = mblnCheckPassWord
    txtPass.Visible = mblnCheckPassWord
    If mblnCheckPassWord Then Exit Sub
    With txtCard
        .Top = picTop.Top + picTop.Height + (fraDown.Top - (picTop.Top + picTop.Height) - .Height) \ 2
        IDKind.Top = .Top
        lblCardNO.Top = .Top + (.Height - lblCardNO.Height) \ 2
    End With
    If Err <> 0 Then Err.Clear
End Sub


