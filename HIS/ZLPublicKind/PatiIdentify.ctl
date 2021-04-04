VERSION 5.00
Begin VB.UserControl PatiIdentify 
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   990
   ScaleWidth      =   4800
   ToolboxBitmap   =   "PatiIdentify.ctx":0000
   Begin ZLPublicKind.IDKindNew IDKind 
      Height          =   330
      Left            =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   135
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "宋体"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox picTxtBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1110
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   495
      ScaleWidth      =   3600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   3600
      Begin VB.TextBox txtPatient 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   105
         Width           =   3990
      End
      Begin VB.Shape shpLine 
         BorderColor     =   &H80000003&
         Height          =   435
         Left            =   0
         Top             =   15
         Width           =   2685
      End
   End
   Begin VB.Label lbl卡号 
      AutoSize        =   -1  'True
      Caption         =   "卡号:"
      Height          =   180
      Left            =   105
      TabIndex        =   2
      Top             =   675
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "PatiIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit
Public Enum Pati_ShowCardNo
    ShowNone = 0
    ShowOnlyCardNo = 1
    ShowShortNameAndCardNo = 2
    ShowFullNameAndCardNo = 3
End Enum
Public Enum Input_Appearance
    ShowFlat = 0
    Show3D = 1
    ShowDeepen3D = 2
End Enum
Public Enum Pati_InputBoxAlignment
    Input_Top_Justify = 0
    Input_Down_Justify = 1
    Input_Center = 2
End Enum
Public Enum TextAlignment
    Text_Left_Justify = 0
    Text_Right_Justify = 1
    Text_Center = 2
End Enum
'事件声明:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtPatient,txtPatient,-1,KeyDown
Attribute KeyDown.VB_Description = "当用户在拥有焦点的对象上按下任意键时发生。"
'读卡之前
Event FindPatiBefore(ByVal objCard As Object, ByRef blnCard As Boolean, ByRef strShowText As String, ByRef objPatiInfor As Object, ByRef blnFindPatied As Boolean, ByRef blnCancel As Boolean)
Event FindPatiArfter(ByVal objCard As Object, ByVal blnCard As Boolean, ByRef ShowName As String, ByRef objHisPatiInfor As Object, _
        ByRef objPatiInfor As Object, ByRef strErrMsg As String, ByRef blnCancel As Boolean)

Event ItemClick(Index As Integer, objCard As Object)   'MappingInfo=IDKInd,IDKInd,-1,ItemClick
Event Click(objCard As Object)   'MappingInfo=IDKInd,IDKInd,-1,Click
Event KeyPress(KeyAscii As Integer)
Event Change()

'缺省属性值:
Const m_def_MustBrushCard = False
Const m_def_InputBoxAlignment = Pati_InputBoxAlignment.Input_Center
Const m_def_HiddenMoseRightKey = True
Const m_def_ShowCardNO = Pati_ShowCardNo.ShowNone
Const m_def_FindPatiShowName = True
Const m_def_病人病区ID = 0
Const m_def_IDKindWidth = 0
Const m_def_OnlyReadDefaultCard = True
Const m_def_ReadThreePatiInfor = True
Const m_def_InputAppearance = 1
Const m_def_OnlyThreeCard = False
'属性变量:
Private m_MustBrushCard As Boolean
Private m_InputBoxAlignment As Pati_InputBoxAlignment
Private m_HiddenMoseRightKey As Boolean
Private m_ShowCardNO As Pati_ShowCardNo
Private m_FindPatiShowName As Boolean
Private m_病人病区ID As Long
Private m_IDKindWidth As Long
Private m_OnlyReadDefaultCard As Boolean
Private m_ReadThreePatiInfor As Boolean
Private mblnNotAutoSel As Boolean
Private m_InputAppearance As Input_Appearance
Private m_OnlyThreeCard As Boolean
'----------------------------------------------------------------------
'控件引用
Private mobjParent As Object
Private mobjCurPati As clsPatiInfor
Private mobjCurCardData As clsPatiInfor   '当前三方卡病人数据或IC卡类存储的病人信息
Private mstrInput As String
Private mblnInternal As Boolean
Private mcnOracle As ADODB.Connection
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private Sub IDKind_Click(objCard As Object)
    RaiseEvent Click(objCard)
End Sub

Private Sub IDKind_LostFocus()
    Call IDKind.Refrash
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As Object, objPatiInfor As Object, blnCancel As Boolean)
    Dim lng病人ID As Long, strErrMsg As String
    If txtPatient.Enabled = False Or txtPatient.Locked Then Exit Sub
    If FindPatiBefor(objPatiInfor.卡号, True, objCard, objPatiInfor) = True Then
        Exit Sub
    End If
    Call GetPatient(True, objPatiInfor.卡号, objCard, lng病人ID, strErrMsg)
End Sub

Private Function FindPatiBefor(ByVal strInput As String, ByVal blnCard As Boolean, _
    objCard As Object, objPati As clsPatiInfor) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人之前
    '入参:strInput-输入的值或刷卡或读卡的值
    '        blnCard-是否刷卡或读卡(True,表示刷卡或读卡)否则为手工输入
    '        objCard-当前的卡类别
    '        objPati-当前读卡的病人信息
    '出参:
    '返回:其他应用程序返回查找到病人或Cancel返回true,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-09-24 15:13:43
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnFindPatied As Boolean, blnCancel As Boolean, strShowName As String
    Dim lng病人ID As Long, strErrMsg As String
    
    If txtPatient.Enabled = False Or txtPatient.Locked Then FindPatiBefor = True: Exit Function
    
    blnFindPatied = False: blnCancel = False
    If Not blnCard Then
        '不是刷卡或者
        Set objPati = Nothing
    End If
    mblnInternal = True
    txtPatient.Text = strInput: strShowName = strInput
    mblnInternal = False
    RaiseEvent FindPatiBefore(objCard, blnCard, strShowName, objPati, blnFindPatied, blnCancel)
    If objCard.接口序号 = 0 And Not blnCard Then
        txtPatient.PasswordChar = ""
    End If
    If Trim(txtPatient.Text) <> strShowName Or strShowName = "" Then txtPatient.PasswordChar = ""
    mblnInternal = True
    txtPatient.Text = strShowName
    mblnInternal = False
    Set mobjCurCardData = objPati
    '如果先找到,直接就不处理后面的数据了
    If blnFindPatied Then
        If FindPatiShowName And Not objPati Is Nothing Then
             mblnInternal = True
             txtPatient.Text = objPati.姓名
             mblnInternal = False
             txtPatient.PasswordChar = ""
        End If
        txtPatient.SelStart = Len(txtPatient.Text)
        Call SelPati
        Call SetShowCardNo(objCard, objPati)
        FindPatiBefor = True
        Exit Function
    End If
    txtPatient.SelStart = Len(txtPatient.Text)
    Call SelPati
    If blnCancel Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        TxtSelAll txtPatient
        Call SetShowCardNo(objCard, objPati)
        FindPatiBefor = True
        Exit Function
    End If
End Function


 Private Sub FindPatiedArfter(ByVal strInput As String, ByVal blnCard As Boolean, _
    objCard As Object, objPati As Object)
    Dim blnCancel As Boolean, strShowName As String, strErrMsg As String
    Dim lng病人ID As Long
    
    If Not objPati Is Nothing Then
        lng病人ID = objPati.病人ID
    End If
    If lng病人ID = 0 Then Set objPati = Nothing
    blnCancel = False: strShowName = strInput
    RaiseEvent FindPatiArfter(objCard, blnCard, strShowName, objPati, mobjCurCardData, strErrMsg, blnCancel)
    Set mobjCurPati = objPati
    mstrInput = strInput
    If Trim(txtPatient.Text) <> strShowName Or strShowName = "" Then txtPatient.PasswordChar = ""
    mblnInternal = True
    txtPatient.Text = strShowName
    mblnInternal = False
    If blnCancel Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        TxtSelAll txtPatient
        Call SetShowCardNo(objCard, objPati)
        Exit Sub
    End If
    If FindPatiShowName And Not objPati Is Nothing Then
         mblnInternal = True
         txtPatient.Text = objPati.姓名
         mblnInternal = False
         txtPatient.PasswordChar = ""
    End If
    txtPatient.SelStart = Len(txtPatient.Text)
    Call SelPati
    If objPati Is Nothing Then
        Set objPati = New clsPatiInfor
        If blnCard Then objPati.卡号 = strInput
    End If
    
    Call SetShowCardNo(objCard, objPati)
End Sub


Private Sub SetShowCardNo(ByVal objCard As Card, ByVal objPati As clsPatiInfor)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置显示的卡号
    '入参:objCard-当前的卡类别
    '        objPati-当前的卡信息
    '编制:刘兴洪
    '日期:2012-09-24 11:35:52
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, strCardNo As String
    lbl卡号.Caption = ""
    If objCard Is Nothing Or objPati Is Nothing Or ShowCardNo = ShowNone Then Exit Sub
    If objPati.卡号 = "" Then Exit Sub
    
    strCardNo = objCard.zlCardNOEncrypt(objPati.卡号)
    If ShowCardNo = ShowShortNameAndCardNo Then
        If objCard.接口序号 > 0 Then
             lbl卡号.Caption = "卡号(" & objCard.短名 & "):" & strCardNo
        ElseIf Not (objCard.名称 Like "姓名*" Or objCard.是否模糊查找) Then
            lbl卡号.Caption = objCard.短名 & ":" & strCardNo
        End If
        Exit Sub
    End If
    If ShowCardNo = ShowFullNameAndCardNo Then
        If objCard.接口序号 > 0 Then
             lbl卡号.Caption = objCard.名称 & "卡号:" & strCardNo
        ElseIf Not (objCard.名称 Like "姓名*" Or objCard.是否模糊查找) Then
            lbl卡号.Caption = objCard.名称 & ":" & strCardNo
        End If
        Exit Sub
    End If
    If ShowCardNo = ShowOnlyCardNo Then
        If objCard.接口序号 > 0 Then
             lbl卡号.Caption = strCardNo
        ElseIf Not (objCard.名称 Like "姓名*" Or objCard.是否模糊查找) Then
            lbl卡号.Caption = strCardNo
        End If
        Exit Sub
    End If
    
End Sub
Private Sub picTxtBack_Click()
   If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
End Sub

Private Sub picTxtBack_GotFocus()
  If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub
Private Sub picTxtBack_Resize()
    Call TxtMove
End Sub
Private Sub TxtMove()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:文本框控件移动
    '编制:刘兴洪
    '日期:2012-09-25 16:11:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngStep As Long
    Err = 0: On Error Resume Next
    If InputAppearance = ShowFlat Or InputAppearance = Show3D Then lngStep = 20
    If InputAppearance = ShowDeepen3D Then lngStep = 40
    'lngStep = IIf(InputAppearance = ShowFlat, 20, 40)
    With picTxtBack
        shpLine.Left = .ScaleLeft
        shpLine.Width = .ScaleWidth
        shpLine.Height = .ScaleHeight
        shpLine.Top = .ScaleTop
    End With
    With picTxtBack
        Select Case InputBoxAlignment
        Case Input_Top_Justify
            With picTxtBack
                txtPatient.Left = .ScaleLeft + lngStep
                txtPatient.Top = .ScaleTop + lngStep
                txtPatient.Width = .ScaleWidth - txtPatient.Left * 2
                txtPatient.Height = .ScaleHeight - txtPatient.Top * 2
            End With
        Case Input_Down_Justify
            txtPatient.Height = .TextHeight("刘兴洪")
            txtPatient.Top = .ScaleHeight - txtPatient.Height - lngStep
            txtPatient.Left = .ScaleLeft + lngStep
            txtPatient.Width = .ScaleWidth - txtPatient.Left * 2
        Case Else
            txtPatient.Height = .TextHeight("刘兴洪")
            txtPatient.Top = (.ScaleHeight - txtPatient.Height - 2 * lngStep) \ 2
            txtPatient.Left = .ScaleLeft + lngStep
            txtPatient.Width = .ScaleWidth - txtPatient.Left * 2
        End Select
    End With
End Sub

Private Sub txtPatient_Change()
    If txtPatient.Visible = False Or txtPatient.Locked = True Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
    If mblnInternal = True Then Exit Sub
    RaiseEvent Change
End Sub

Private Sub txtPatient_GotFocus()
    If txtPatient.Visible = False Or txtPatient.Locked = True Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.SetAutoReadCard(Trim(txtPatient.Text) = "")
    Call SetBrushCardObject(True)
    If mblnNotAutoSel Then
        mblnNotAutoSel = False
        txtPatient.SelStart = 1
        Exit Sub
    End If
    TxtSelAll txtPatient
End Sub

Private Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡接口
    '返回: true-成功，false-失败
    '编制:李南春
    '日期:2015/7/23 10:37:39
    '问题:85565
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String, objCurCard As Card
    Dim objPubOneCard As clsPublicOneCard
    
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    
    Set objCurCard = IDKind.GetCurCard
    If objCurCard Is Nothing Then
        Set objCurCard = IDKind.GetfaultCard
    Else
        If objCurCard.接口序号 = 0 And objCurCard.名称 Like "*姓名*" Then Set objCurCard = IDKind.GetfaultCard
    End If
    
    If objCurCard Is Nothing Then Exit Function
    If objCurCard.接口序号 = 0 Or objCurCard.接口程序名 = "" Or Not (objCurCard.是否刷卡 Or objCurCard.是否扫描) Then Exit Function
    
    If zlGetPubOneCard(mcnOracle, objPubOneCard) = False Then Exit Function
    
    If objPubOneCard.objThirdSwap.zlSetBrushCardObject(objCurCard.接口序号, IIf(blnComm, txtPatient, Nothing), strExpend, objCurCard.消费卡) Then
        If mobjCommEvents Is Nothing And AllowAutoCommCard = True Then
            Set mobjCommEvents = New clsCommEvents
            Call objPubOneCard.objThirdSwap.zlInitEvents(UserControl.hWnd, mobjCommEvents)
        End If
    End If
    Set objPubOneCard = Nothing
End Function


Private Sub txtPatient_LostFocus()
    If txtPatient.Visible = False Or txtPatient.Locked = True Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.SetAutoReadCard(False)
    Call SetBrushCardObject(False)
End Sub

 Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Or HiddenMoseRightKey = False Then Exit Sub
    glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
    Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
End Sub
Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Or HiddenMoseRightKey = False Then Exit Sub
    Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    '需要检查是否刷卡
    Dim blnCard As Boolean  '是否刷卡
    Dim objDefault As Card
    Dim blnCardPass As Boolean  ' 卡号加密显示
    Dim lngDefaultLen As Long, strInputText As String, blnCancel As Boolean, strErrMsg As String
    Dim lng病人ID As Long, objPati As New clsPatiInfor, objThreePati As New clsPatiInfor
    Dim objCard As Card, blnFindPatied As Boolean
    Dim blnPass As Boolean
    
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IDKind.GetCurCard Is Nothing Then Exit Sub
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = 0 Then Exit Sub
    
    If IDKind.GetCurCard.名称 Like "*姓*名*" Or IDKind.GetCurCard.模糊查找项 Then
        '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnPass = txtPatient.PasswordChar <> ""
            blnCard = InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
            txtPatient.IMEMode = 0
            blnPass = txtPatient.PasswordChar = "" And blnPass
            If blnPass Then
                If txtPatient.SelLength = Len(txtPatient.Text) Then
                    mblnInternal = True
                    txtPatient.Text = ""
                    mblnInternal = False
                End If
                SendKeys Chr(KeyAscii): KeyAscii = 0: Exit Sub
            End If
        End If
        
        '76196
        If FindPatiShowName = False And KeyAscii = 13 And Not mobjCurPati Is Nothing And mstrInput = txtPatient.Text Then
            Call FindPatiedArfter(mstrInput, blnCard, IDKind.GetCurCard, mobjCurPati)
            Exit Sub
        End If
        If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
            Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then Exit Sub
        '刷卡
        If KeyAscii <> 13 Then
             mblnInternal = True
             txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
             mblnInternal = False
        End If
        KeyAscii = 0
        strInputText = txtPatient.Text
        Set objCard = Nothing
        If blnCard Then
            objPati.卡号 = strInputText
            Set objCard = IDKind.GetfaultCard
        End If
        If objCard Is Nothing Then
            Set objCard = IDKind.GetCurCard
        End If
        
        Set objPati = Nothing
        If FindPatiBefor(strInputText, blnCard, objCard, objPati) = True Then
            Exit Sub
        End If
        If Not blnCard Then
            '不是刷卡时
            If Left(strInputText, 1) = "-" And IsNumeric(Mid(strInputText, 2)) Then
                '病人ID
                '78219:李南春,2014/9/23,限制病人ID的查询范围在long数据类型之内
                If Val(Mid(strInputText, 2)) > 2147483647 Then
                    MsgBox "输入的ID过长，请检查输入或使用其它查询方式！", vbInformation, gstrSysName
                    Exit Sub
                End If
                lng病人ID = Val(Mid(strInputText, 2))
                Call GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg)
            ElseIf Left(strInputText, 1) = "*" And IsNumeric(Mid(strInputText, 2)) Then '门诊号
                Call GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "门诊号", Mid(strInputText, 2))
            ElseIf Left(strInputText, 1) = "+" And IsNumeric(Mid(strInputText, 2)) Then '住院号
                Call GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "住院号", Mid(strInputText, 2))
            ElseIf Left(strInputText, 1) = "/" Then
                '床号需特殊处理:无病区,不能根据床号来查找
                If 病人病区ID = 0 Then Exit Sub
                lng病人ID = zlGetPatiIDFromBedNumber(mcnOracle, 病人病区ID, Mid(strInputText, 2))
                If lng病人ID <> 0 Then Call GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg)
            ElseIf IsMobileNo(strInputText) Then
                '103000：李南春，2017/2/7，按手机号查找
                If GetPatiIDFromCardType(mcnOracle, "手机号", strInputText, False, lng病人ID, , strErrMsg) = False Then lng病人ID = 0
                If lng病人ID = 0 Then GoTo NotFound
                Call GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg)
            End If
            Call FindPatiedArfter(strInputText, blnCard, objCard, objPati)
            Exit Sub
        End If
        Set objCard = IDKind.GetCurCard
        Call GetPatient(blnCard, strInputText, objCard, lng病人ID, strErrMsg, IDKind.GetDefaultCardTypeID, IDKind.Cards.按缺省卡查找)
        Exit Sub
    End If
    If InStr(1, ",门诊号,住院号,手机号,", "," & IDKind.GetCurCard.名称 & ",") > 0 Then
        '只能为数字
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack And InStr(",3,22,24,", "," & KeyAscii & ",") = 0 Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    End If
    
    blnCard = zlIsBrushCard(txtPatient, KeyAscii)
    If IDKind.GetCurCard.接口序号 > 0 And MustBrushCard Then
        '必须刷卡
        If Not blnCard And KeyAscii <> 13 And KeyAscii <> 8 Then
            If KeyAscii > 32 Then
                mblnInternal = True
                txtPatient.Text = Chr(KeyAscii): txtPatient.SelStart = 1
                mblnInternal = False
            End If
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then Exit Sub

    If KeyAscii <> 13 Then
        mblnInternal = True
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
        mblnInternal = False
    End If
    KeyAscii = 0
    strInputText = txtPatient.Text
    
    Set objPati = Nothing
    Set objCard = IDKind.GetCurCard
    If FindPatiBefor(strInputText, blnCard, objCard, objPati) = True Then Exit Sub
    If blnCard Then '刷卡
        Call GetPatient(blnCard, strInputText, objCard, lng病人ID, strErrMsg): Exit Sub
    End If
    '不是刷卡时
    If Left(strInputText, 1) = "-" And IsNumeric(Mid(strInputText, 2)) Then
        '病人ID
        lng病人ID = Val(Mid(strInputText, 2))
        Call GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg)
    ElseIf Left(strInputText, 1) = "*" And IsNumeric(Mid(strInputText, 2)) Then '门诊号
        Call GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "门诊号", Mid(strInputText, 2))
    ElseIf Left(strInputText, 1) = "+" And IsNumeric(Mid(strInputText, 2)) Then '住院号
        Call GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "住院号", Mid(strInputText, 2))
    ElseIf Left(strInputText, 1) = "/" Then
        '床号需特殊处理:无病区,不能根据床号来查找
        If 病人病区ID = 0 Then GoTo NotFound:
        lng病人ID = zlGetPatiIDFromBedNumber(mcnOracle, 病人病区ID, Mid(strInputText, 2))
        If lng病人ID <> 0 Then Call GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg)
    ElseIf IsMobileNo(strInputText) Then
        '103000：李南春，2017/2/7，按手机号查找
        If GetPatiIDFromCardType(mcnOracle, "手机号", strInputText, False, lng病人ID, , strErrMsg, 0) = False Then lng病人ID = 0
        If lng病人ID = 0 Then GoTo NotFound
        Call GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg)
    Else
        Call GetPatient(blnCard, strInputText, objCard, lng病人ID, strErrMsg)
        Exit Sub
    End If
NotFound:
    Call FindPatiedArfter(strInputText, blnCard, objCard, objPati)
End Sub
Private Sub SelPati()
    '选中病人文本框信息
    If ActiveControl Is txtPatient Then
        TxtSelAll txtPatient
    End If
End Sub
Private Sub UserControl_GotFocus()
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub
Private Sub UserControl_Initialize()
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub UserControl_Terminate()
    Err = 0: On Error Resume Next
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If Not mobjCurCardData Is Nothing Then Set mobjCurCardData = Nothing
    If Not mobjParent Is Nothing Then Set mobjParent = Nothing
    If Not mobjCurPati Is Nothing Then Set mobjCurPati = Nothing
End Sub
Private Sub UserControl_Resize()
    Dim lngHeight As Long
    Err = 0: On Error Resume Next
    lngHeight = IIf(lbl卡号.Visible = False, 0, lbl卡号.Height + 10)
    lbl卡号.Left = ScaleLeft
    lbl卡号.Top = ScaleHeight - lbl卡号.Height
    
    With IDKind
        .Left = ScaleLeft
        .Top = ScaleTop
        .Height = ScaleHeight - lngHeight
    End With
    With picTxtBack
        .Left = IDKind.Left + IDKind.Width + 20
        .Top = IDKind.Top
        .Height = IDKind.Height
        .Width = ScaleWidth - .Left
    End With
    Call SetInputAppearance
End Sub
Private Sub IDKInd_ItemClick(Index As Integer, objCard As Object)
    RaiseEvent ItemClick(Index, objCard)
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Call UserControl_Resize
End Sub
 

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKInd,IDKInd,-1,Cards
Public Property Get Cards() As Object
    Set Cards = IDKind.Cards
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,Font
Public Property Get Font() As Font
    Set Font = txtPatient.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtPatient.Font = New_Font
    Set picTxtBack.Font = New_Font
    PropertyChanged "Font"
    Call UserControl_Resize
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKInd,IDKInd,-1,IDKindStr
Public Property Get IDKindStr() As String
    IDKindStr = IDKind.IDKindStr
End Property
Public Property Let IDKindStr(ByVal New_IDKindStr As String)
    IDKind.IDKindStr() = New_IDKindStr
    PropertyChanged "IDKindStr"
End Property
'
'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKInd,IDKInd,-1,Font
Public Property Get IDKindFont() As Font
    Set IDKindFont = IDKind.Font
End Property
Public Property Set IDKindFont(ByVal New_IDKindFont As Font)
    Set IDKind.Font = New_IDKindFont
    PropertyChanged "IDKindFont"
End Property
'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set txtPatient.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set picTxtBack.Font = picTxtBack.Font
    IDKind.IDKindStr = PropBag.ReadProperty("IDKindStr", "")
    
    Set IDKind.Font = PropBag.ReadProperty("IDKindFont", Ambient.Font)
    IDKind.AutoSize = PropBag.ReadProperty("AutoSize", False)
    IDKind.Appearance = PropBag.ReadProperty("IDKindAppearance", 1)
    m_InputAppearance = PropBag.ReadProperty("InputAppearance", m_def_InputAppearance)
    IDKind.ShowSortName = PropBag.ReadProperty("ShowSortName", False)
    IDKind.CaptionAlignment = PropBag.ReadProperty("CaptionAlignment", 2)
    m_ReadThreePatiInfor = PropBag.ReadProperty("ReadThreePatiInfor", m_def_ReadThreePatiInfor)
    IDKind.ShowPropertySet = PropBag.ReadProperty("ShowPropertySet", False)
    IDKind.DefaultCardType = PropBag.ReadProperty("DefaultCardType", "")
    m_OnlyReadDefaultCard = PropBag.ReadProperty("OnlyReadDefaultCard", m_def_OnlyReadDefaultCard)
    IDKind.BorderStyle = PropBag.ReadProperty("IDkindBorderStyle", 0)

    Call UserControl_Resize: Call SetInputAppearance
 
    m_IDKindWidth = PropBag.ReadProperty("IDKindWidth", IDKind.Width)
    IDKind.Width = m_IDKindWidth
    Dim blnEnable As Boolean
    blnEnable = PropBag.ReadProperty("Enabled", True)
    UserControl.Enabled = blnEnable
    IDKind.Enabled = blnEnable
    m_病人病区ID = PropBag.ReadProperty("病人病区ID", m_def_病人病区ID)
    m_OnlyThreeCard = PropBag.ReadProperty("OnlyThreeCard", m_def_OnlyThreeCard)
    m_FindPatiShowName = PropBag.ReadProperty("FindPatiShowName", m_def_FindPatiShowName)
    mblnInternal = True
    txtPatient.Text = PropBag.ReadProperty("Text", "")
    mblnInternal = False
    txtPatient.Locked = PropBag.ReadProperty("Locked", False)
    m_ShowCardNO = PropBag.ReadProperty("ShowCardNO", m_def_ShowCardNO)
    Call InitFace
    m_HiddenMoseRightKey = PropBag.ReadProperty("HiddenMoseRightKey", m_def_HiddenMoseRightKey)
    Set lbl卡号.Font = PropBag.ReadProperty("CardNoShowFont", Ambient.Font)
    lbl卡号.ForeColor = PropBag.ReadProperty("CardNOForColor", &H80000012)
    m_InputBoxAlignment = PropBag.ReadProperty("InputBoxAlignment", m_def_InputBoxAlignment)
    
    Call TxtMove
    txtPatient.Alignment = PropBag.ReadProperty("TextAlignment", 0)
    m_MustBrushCard = PropBag.ReadProperty("MustBrushCard", m_def_MustBrushCard)
    IDKind.AllowAutoCommCard = PropBag.ReadProperty("AllowAutoCommCard", False)
    IDKind.AllowAutoICCard = PropBag.ReadProperty("AllowAutoICCard", False)
    IDKind.AllowAutoIDCard = PropBag.ReadProperty("AllowAutoIDCard", False)
    IDKind.IDKind = PropBag.ReadProperty("IDKindIDX", 1)
    IDKind.NotContainFastKey = PropBag.ReadProperty("NotContainFastKey", "")
    IDKind.NotAutoAppendKind = PropBag.ReadProperty("NotAutoAppendKind", False)
    txtPatient.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtPatient.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    picTxtBack.BackColor = txtPatient.BackColor
    txtPatient.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtPatient.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtPatient.SelText = PropBag.ReadProperty("SelText", "")
    txtPatient.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    txtPatient.IMEMode = PropBag.ReadProperty("IMEMode", 0)
    mblnNotAutoSel = PropBag.ReadProperty("NotAutoSel", False)
    Call SetInputAppearance
End Sub


'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", txtPatient.Font, Ambient.Font)
    Call PropBag.WriteProperty("IDKindStr", IDKind.IDKindStr, "")
    Call PropBag.WriteProperty("IDKindFont", IDKindFont, Ambient.Font)
    Call PropBag.WriteProperty("AutoSize", IDKind.AutoSize, False)
    Call PropBag.WriteProperty("IDKindAppearance", IDKind.Appearance, 1)
    Call PropBag.WriteProperty("InputAppearance", m_InputAppearance, m_def_InputAppearance)
    Call PropBag.WriteProperty("ShowSortName", IDKind.ShowSortName, False)
    Call PropBag.WriteProperty("CaptionAlignment", IDKind.CaptionAlignment, 2)
    Call PropBag.WriteProperty("ReadThreePatiInfor", m_ReadThreePatiInfor, m_def_ReadThreePatiInfor)
    Call PropBag.WriteProperty("ShowPropertySet", IDKind.ShowPropertySet, False)
    Call PropBag.WriteProperty("DefaultCardType", IDKind.DefaultCardType, "")
    Call PropBag.WriteProperty("OnlyReadDefaultCard", m_OnlyReadDefaultCard, m_def_OnlyReadDefaultCard)
    Call PropBag.WriteProperty("IDkindBorderStyle", IDKind.BorderStyle, 0)
    Call PropBag.WriteProperty("IDKindWidth", m_IDKindWidth, m_def_IDKindWidth)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("病人病区ID", m_病人病区ID, m_def_病人病区ID)
    Call PropBag.WriteProperty("OnlyThreeCard", m_OnlyThreeCard, m_def_OnlyThreeCard)
    Call PropBag.WriteProperty("FindPatiShowName", m_FindPatiShowName, m_def_FindPatiShowName)
    Call PropBag.WriteProperty("Text", txtPatient.Text, "")
    Call PropBag.WriteProperty("Locked", txtPatient.Locked, False)
    Call PropBag.WriteProperty("ShowCardNO", m_ShowCardNO, m_def_ShowCardNO)
    Call PropBag.WriteProperty("HiddenMoseRightKey", m_HiddenMoseRightKey, m_def_HiddenMoseRightKey)
    Call PropBag.WriteProperty("CardNoShowFont", lbl卡号.Font, Ambient.Font)
    Call PropBag.WriteProperty("CardNOForColor", lbl卡号.ForeColor, &H80000012)
    Call PropBag.WriteProperty("InputBoxAlignment", m_InputBoxAlignment, m_def_InputBoxAlignment)
    Call TxtMove
    Call PropBag.WriteProperty("TextAlignment", txtPatient.Alignment, 0)
    Call PropBag.WriteProperty("MustBrushCard", m_MustBrushCard, m_def_MustBrushCard)
    Call PropBag.WriteProperty("AllowAutoCommCard", IDKind.AllowAutoCommCard, False)
    Call PropBag.WriteProperty("AllowAutoICCard", IDKind.AllowAutoICCard, False)
    Call PropBag.WriteProperty("AllowAutoIDCard", IDKind.AllowAutoIDCard, False)
    Call PropBag.WriteProperty("IDKindIDX", IDKind.IDKind, 1)
    Call PropBag.WriteProperty("NotContainFastKey", IDKind.NotContainFastKey, "")
    Call PropBag.WriteProperty("NotAutoAppendKind", IDKind.NotAutoAppendKind, False)
    Call PropBag.WriteProperty("ForeColor", txtPatient.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BackColor", txtPatient.BackColor, &H80000005)
    Call PropBag.WriteProperty("SelStart", txtPatient.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtPatient.SelLength, 0)
    Call PropBag.WriteProperty("SelText", txtPatient.SelText, "")
    Call PropBag.WriteProperty("PasswordChar", txtPatient.PasswordChar, "")
    Call PropBag.WriteProperty("IMEMode", txtPatient.IMEMode, 0)
    Call PropBag.WriteProperty("NotAutoSel", mblnNotAutoSel, False)
End Sub

Public Function IsMobileNo(ByVal strInput As String) As Boolean
    IsMobileNo = IDKind.IsMobileNo(strInput)
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKInd,IDKInd,-1,AutoSize
Public Property Get AutoSize() As Boolean
    AutoSize = IDKind.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    IDKind.AutoSize() = New_AutoSize
    PropertyChanged "AutoSize"
    Call UserControl_Resize
End Property
'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKInd,IDKInd,-1,Appearance
Public Property Get IDKindAppearance() As IDKind_Appearance
    IDKindAppearance = IDKind.Appearance
End Property

Public Property Let IDKindAppearance(ByVal New_IDKindAppearance As IDKind_Appearance)
    IDKind.Appearance() = New_IDKindAppearance
    PropertyChanged "IDKindAppearance"
End Property
Private Sub SetInputAppearance()
    Call picTxtBack_Resize
    shpLine.Visible = InputAppearance = 0
    If InputAppearance = ShowDeepen3D Then
        Call zlRaisEffect(picTxtBack, -2, " ")
    ElseIf InputAppearance = Show3D Then
        Call zlRaisEffect(picTxtBack, -1, " ")
    Else
         Call zlRaisEffect(picTxtBack, 0, " ")
    End If
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get InputAppearance() As Input_Appearance
    InputAppearance = m_InputAppearance
End Property

Public Property Let InputAppearance(ByVal New_InputAppearance As Input_Appearance)
    m_InputAppearance = New_InputAppearance
    PropertyChanged "InputAppearance"
    Call SetInputAppearance
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_InputAppearance = m_def_InputAppearance
    m_ReadThreePatiInfor = m_def_ReadThreePatiInfor
    Call SetInputAppearance
    m_OnlyReadDefaultCard = m_def_OnlyReadDefaultCard
    m_IDKindWidth = m_def_IDKindWidth
    m_病人病区ID = m_def_病人病区ID
    m_OnlyThreeCard = m_def_OnlyThreeCard
    m_FindPatiShowName = m_def_FindPatiShowName
    m_ShowCardNO = m_def_ShowCardNO
    Call InitFace
    m_HiddenMoseRightKey = m_def_HiddenMoseRightKey
    m_InputBoxAlignment = m_def_InputBoxAlignment
    Call TxtMove
    m_MustBrushCard = m_def_MustBrushCard
    mblnNotAutoSel = False
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKInd,IDKInd,-1,ShowSortName
Public Property Get ShowSortName() As Boolean
    ShowSortName = IDKind.ShowSortName
End Property

Public Property Let ShowSortName(ByVal New_ShowSortName As Boolean)
    IDKind.ShowSortName() = New_ShowSortName
    PropertyChanged "ShowSortName"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKInd,IDKInd,-1,CaptionAlignment
Public Property Get CaptionAlignment() As IDKind_CaptionAlignment
    CaptionAlignment = IDKind.CaptionAlignment
End Property
Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As IDKind_CaptionAlignment)
    IDKind.CaptionAlignment() = New_CaptionAlignment
    PropertyChanged "CaptionAlignment"
End Property

Public Sub zlInit(ByVal frmMain As Object, Optional ByVal lngSys As Long, Optional ByVal lngModul As Long, _
    Optional cnOracle As ADODB.Connection, Optional strDBUser As String, _
    Optional objPublicOneCard As Object, _
    Optional strIDKindStr As String = "", _
    Optional strProductName As String = "", Optional ByVal blnIsObjRegisterAlone As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化
    '入参:frmMain-调用的主窗口
    '     lngSys : 系统编号
    '     lngModul:需要执行的功能序号
    '     objPublicOneCard-卡结算公共部件
    '     cnOracle:主程序的数据库连接
    '     strIDKindStr-身份识别的类别项,有两种格式:
    '               一种是缺省的:短名1|全名1|读卡标志1;…. ;短名n|全名n|读卡标志n
    '               一种是扩展格式:短名1|全名1|读卡标志1|卡类别ID1|卡号长度1|缺省标志1(1-当前缺省;0-非缺省)|是否存在帐户1(1-存在帐户;0-不存在帐户)|卡号密文1(第几位至第几位加密,空为不加密);…
    '     strProductName-产品模块名称(主要用于相关模块参数的保存)
    '     是否使用独立的注册部件(True:使用:zlRegisterAlone.DLL,否则使用zlRegister.dll)
    '编制:刘兴洪
    '日期:2012-08-16 10:45:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjParent = frmMain
    Set mcnOracle = cnOracle
    gblnIsObjRegisterAlone = blnIsObjRegisterAlone
    '118959:李南春，2018/1/3,绑定刷卡文本框
    Call IDKind.zlInit(frmMain, lngSys, lngModul, cnOracle, strDBUser, objPublicOneCard, strIDKindStr, , strProductName, OnlyThreeCard, blnIsObjRegisterAlone)
    Call InitFace: lbl卡号.Caption = ""
    Call UserControl_Resize
End Sub
'------------------------------------------------------------------------
'业务操作
Private Function GetPatient(ByVal blnBrushCard As Boolean, ByVal strInput As String, ByRef objCard As Card, _
      ByRef lng病人ID As Long, ByRef strErrMsg As String, Optional ByVal lngDefaultCardTypeID As Long, Optional ByVal bln缺省卡查找 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人ID
    '入参:strInput-输入的相关值
    '        objCard-当前读取的卡类别(可以为出参)
    '出参:lng病人ID-返回:病人ID
    '        objCard-返回读取的卡类别
    '        strErrMsg-返回的错语信息
    '返回: 成功,返回指定的病人ID
    '编制:刘兴洪
    '日期:2012-08-20 17:34:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardPassWord As String, lngCardTypeID As Long
    Dim objPati As New clsPatiInfor
    
    On Error GoTo errHandle
    
    '指定卡类别
    If objCard.接口序号 > 0 Then
       If GetPatiIDFromCardType(mcnOracle, objCard.接口序号, strInput, False, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID) = False Then lng病人ID = 0
       If lng病人ID = 0 Then GoTo NotFindPati:
       
       If lngCardTypeID <> objCard.接口序号 And lngCardTypeID <> 0 Then
            Set objCard = IDKind.GetIDKindCard(lngCardTypeID, CardTypeID)
       End If
        Call GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg)
        objPati.卡号 = strInput: objPati.密码 = strCardPassWord
       Call FindPatiedArfter(strInput, blnBrushCard, objCard, objPati)
       GetPatient = True: Exit Function
    End If
    If objCard.名称 Like "姓名*" Or objCard.模糊查找项 Then
        If bln缺省卡查找 Then
            lngCardTypeID = lngDefaultCardTypeID
        Else
            lngCardTypeID = -1
        End If
        If blnBrushCard Then
            If GetPatiIDFromCardType(mcnOracle, lngCardTypeID, strInput, True, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID) Then
                If lngCardTypeID <> objCard.接口序号 And lngCardTypeID <> 0 Then
                     Set objCard = IDKind.GetIDKindCard(lngCardTypeID, CardTypeID)
                ElseIf Not IDKind.GetfaultCard Is Nothing Then
                    Set objCard = IDKind.GetfaultCard
                End If
            ElseIf IsMobileNo(strInput) Then
                If GetPatiIDFromCardType(mcnOracle, "手机号", strInput, True, lng病人ID, strCardPassWord) Then
                    Set objCard = IDKind.GetIDKindCard("手机号", CardTypeName)
                Else
                    lng病人ID = 0
                    If strErrMsg <> "" Then MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
                End If
               
            Else
                lng病人ID = 0
                If strErrMsg <> "" Then MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
            End If
            
            If lng病人ID = 0 Then GoTo NotFindPati:
            Call GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg)
            objPati.卡号 = strInput: objPati.密码 = strCardPassWord
            Call FindPatiedArfter(strInput, blnBrushCard, objCard, objPati)
            GetPatient = True: Exit Function
        End If
    End If
    Select Case objCard.名称
    Case "姓名"
    Case "医保号", "门诊号", "住院号", "就诊卡", "IC卡号", "手机号"
        If GetPatiIDFromCardType(mcnOracle, objCard.名称, strInput, False, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID) = False Then lng病人ID = 0
        If lng病人ID = 0 Then GoTo NotFindPati:
        If lngCardTypeID <> objCard.接口序号 And lngCardTypeID <> 0 Then
             Set objCard = IDKind.GetIDKindCard(lngCardTypeID, CardTypeID)
        End If
        Call GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg)
        objPati.卡号 = strInput: objPati.密码 = strCardPassWord
        Call FindPatiedArfter(strInput, blnBrushCard, objCard, objPati)
        GetPatient = True: Exit Function
    Case Else
            GoTo NotFindPati:
    End Select
    GoTo NotFindPati:
     Exit Function
errHandle:
    strErrMsg = Err.Description
NotFindPati:
    objPati.卡号 = strInput
    Call FindPatiedArfter(strInput, blnBrushCard, objCard, objPati)
End Function

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    Dim objPati As clsPatiInfor, strErrMsg As String
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard(strCardType, CardTypeName)
    If objCard Is Nothing Then Exit Sub
    
    If strXmlCardInfor <> "" Then
       Call zlGetPatiInforFromXML(mcnOracle, strXmlCardInfor, objPati, strErrMsg)
    End If
    If objPati Is Nothing Then Set objPati = New clsPatiInfor
    objPati.卡号 = strCardNo
    Call IDKind_ReadCard(objCard, objPati, blnCancel)
End Sub


'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get ReadThreePatiInfor() As Boolean
    ReadThreePatiInfor = m_ReadThreePatiInfor
End Property

Public Property Let ReadThreePatiInfor(ByVal New_ReadThreePatiInfor As Boolean)
    m_ReadThreePatiInfor = New_ReadThreePatiInfor
    PropertyChanged "ReadThreePatiInfor"
End Property
'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKInd,IDKInd,-1,ShowPropertySet
Public Property Get ShowPropertySet() As Boolean
    ShowPropertySet = IDKind.ShowPropertySet
End Property

Public Property Let ShowPropertySet(ByVal New_ShowPropertySet As Boolean)
    IDKind.ShowPropertySet() = New_ShowPropertySet
    PropertyChanged "ShowPropertySet"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKInd,IDKInd,-1,DefaultCardType
Public Property Get DefaultCardType() As String
    DefaultCardType = IDKind.DefaultCardType
End Property

Public Property Let DefaultCardType(ByVal New_DefaultCardType As String)
    IDKind.DefaultCardType() = New_DefaultCardType
    PropertyChanged "DefaultCardType"

End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get OnlyReadDefaultCard() As Boolean
    OnlyReadDefaultCard = m_OnlyReadDefaultCard
End Property

Public Property Let OnlyReadDefaultCard(ByVal New_OnlyReadDefaultCard As Boolean)
    m_OnlyReadDefaultCard = New_OnlyReadDefaultCard
    PropertyChanged "OnlyReadDefaultCard"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,BorderStyle
Public Property Get IDkindBorderStyle() As IDKind_BorderStyle
    IDkindBorderStyle = IDKind.BorderStyle
End Property

Public Property Let IDkindBorderStyle(ByVal New_IDkindBorderStyle As IDKind_BorderStyle)
    IDKind.BorderStyle() = New_IDkindBorderStyle
    PropertyChanged "IDkindBorderStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get IDKindWidth() As Long
    IDKindWidth = IDKind.Width
End Property

Public Property Let IDKindWidth(ByVal New_IDKindWidth As Long)
    m_IDKindWidth = New_IDKindWidth
    IDKind.Width = m_IDKindWidth
    PropertyChanged "IDKindWidth"
    Call UserControl_Resize
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    IDKind.Enabled = New_Enabled
    txtPatient.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get 病人病区ID() As Long
    病人病区ID = m_病人病区ID
End Property

Public Property Let 病人病区ID(ByVal New_病人病区ID As Long)
    m_病人病区ID = New_病人病区ID
    PropertyChanged "病人病区ID"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get OnlyThreeCard() As Boolean
    OnlyThreeCard = m_OnlyThreeCard
End Property

Public Property Let OnlyThreeCard(ByVal New_OnlyThreeCard As Boolean)
    m_OnlyThreeCard = New_OnlyThreeCard
    PropertyChanged "OnlyThreeCard"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,true
Public Property Get FindPatiShowName() As Boolean
    FindPatiShowName = m_FindPatiShowName
End Property

Public Property Let FindPatiShowName(ByVal New_FindPatiShowName As Boolean)
    m_FindPatiShowName = New_FindPatiShowName
    PropertyChanged "FindPatiShowName"
End Property
'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "返回/设置控件中包含的文本。"
    Text = txtPatient.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtPatient.Text() = New_Text
    PropertyChanged "Text"
End Property
'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "决定控件是否可编辑。"
    Locked = txtPatient.Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
    txtPatient.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Public Property Get objTxtInput() As Object
    Set objTxtInput = txtPatient
End Property
Public Property Get objIDKind() As Object
    Set objIDKind = IDKind
End Property
'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get ShowCardNo() As Pati_ShowCardNo
    ShowCardNo = m_ShowCardNO
End Property

Public Property Let ShowCardNo(ByVal New_ShowCardNO As Pati_ShowCardNo)
    m_ShowCardNO = New_ShowCardNO
    Call InitFace
    Call UserControl_Resize
    PropertyChanged "ShowCardNO"
End Property
Private Sub InitFace()
    lbl卡号.Visible = ShowCardNo <> ShowNone
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get HiddenMoseRightKey() As Boolean
Attribute HiddenMoseRightKey.VB_Description = "屏蔽鼠标右键的快捷菜单"
    HiddenMoseRightKey = m_HiddenMoseRightKey
End Property

Public Property Let HiddenMoseRightKey(ByVal New_HiddenMoseRightKey As Boolean)
    m_HiddenMoseRightKey = New_HiddenMoseRightKey
    PropertyChanged "HiddenMoseRightKey"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=lbl卡号,lbl卡号,-1,Font
Public Property Get CardNoShowFont() As Font
Attribute CardNoShowFont.VB_Description = "返回一个 Font 对象。"
    Set CardNoShowFont = lbl卡号.Font
End Property

Public Property Set CardNoShowFont(ByVal New_CardNoShowFont As Font)
    Set lbl卡号.Font = New_CardNoShowFont
    PropertyChanged "CardNoShowFont"
    Call UserControl_Resize
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=lbl卡号,lbl卡号,-1,ForeColor
Public Property Get CardNOForColor() As OLE_COLOR
Attribute CardNOForColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    CardNOForColor = lbl卡号.ForeColor
End Property

Public Property Let CardNOForColor(ByVal New_CardNOForColor As OLE_COLOR)
    lbl卡号.ForeColor() = New_CardNOForColor
    PropertyChanged "CardNOForColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get InputBoxAlignment() As Pati_InputBoxAlignment
Attribute InputBoxAlignment.VB_Description = "输入时,光标对齐方式"
    InputBoxAlignment = m_InputBoxAlignment
End Property
Public Property Let InputBoxAlignment(ByVal New_InputBoxAlignment As Pati_InputBoxAlignment)
    m_InputBoxAlignment = New_InputBoxAlignment
    PropertyChanged "InputBoxAlignment"
    Call TxtMove
End Property
'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,Alignment
Public Property Get TextAlignment() As TextAlignment
    TextAlignment = txtPatient.Alignment
End Property

Public Property Let TextAlignment(ByVal New_TextAlignment As TextAlignment)
    txtPatient.Alignment() = New_TextAlignment
    PropertyChanged "TextAlignment"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get MustBrushCard() As Boolean
    MustBrushCard = m_MustBrushCard
End Property

Public Property Let MustBrushCard(ByVal New_MustBrushCard As Boolean)
    m_MustBrushCard = New_MustBrushCard
    PropertyChanged "MustBrushCard"
End Property

Public Function zlIsBrushCard(ByVal txtInput As Object, KeyAscii As Integer) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否刷卡
    '入参:txtInput-输入文本框
    '       KeyAscii
    '出参:
    '返回:是刷卡返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-09-26 11:05:43
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, strText As String
     '刷卡时含有特殊符号的由调用方取消输入
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then Exit Function
    blnCard = False
    '处理当前键入后显示的内容(还未显示出来)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    '判断是否在刷卡
     If KeyAscii > 32 Then
        sngNow = Timer
        If txtInput.Text = "" Or strText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True '用一台笔记本测试，一般在0.014左右
        End If
    End If
    zlIsBrushCard = blnCard
End Function

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,ActiveFastKey
Public Function ActiveFastKey() As Boolean
    If Trim(txtPatient.Text) <> "" Then Exit Function
    ActiveFastKey = IDKind.ActiveFastKey()
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,AllowAutoCommCard
Public Property Get AllowAutoCommCard() As Boolean
    AllowAutoCommCard = IDKind.AllowAutoCommCard
End Property

Public Property Let AllowAutoCommCard(ByVal New_AllowAutoCommCard As Boolean)
    IDKind.AllowAutoCommCard() = New_AllowAutoCommCard
    PropertyChanged "AllowAutoCommCard"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,AllowAutoICCard
Public Property Get AllowAutoICCard() As Boolean
    AllowAutoICCard = IDKind.AllowAutoICCard
End Property

Public Property Let AllowAutoICCard(ByVal New_AllowAutoICCard As Boolean)
    IDKind.AllowAutoICCard() = New_AllowAutoICCard
    PropertyChanged "AllowAutoICCard"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,AllowAutoIDCard
Public Property Get AllowAutoIDCard() As Boolean
    AllowAutoIDCard = IDKind.AllowAutoIDCard
End Property

Public Property Let AllowAutoIDCard(ByVal New_AllowAutoIDCard As Boolean)
    IDKind.AllowAutoIDCard() = New_AllowAutoIDCard
    PropertyChanged "AllowAutoIDCard"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,GetCardNoLen
Public Property Get GetCardNoLen() As Integer
    GetCardNoLen = IDKind.GetCardNoLen
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,GetCurCard
Public Property Get GetCurCard() As Object
    Set GetCurCard = IDKind.GetCurCard
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,GetDefaultCardNoLen
Public Property Get GetDefaultCardNoLen() As Integer
    GetDefaultCardNoLen = IDKind.GetDefaultCardNoLen
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,GetDefaultCardTypeID
Public Function GetDefaultCardTypeID() As Long
    GetDefaultCardTypeID = IDKind.GetDefaultCardTypeID()
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,GetIDKindCard
Public Function GetIDKindCard(ByVal strCardType As String, Optional MachMode As Mach_Mode) As Object
    Set GetIDKindCard = IDKind.GetIDKindCard(strCardType, MachMode)
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,GetKindIndex
Public Function GetKindIndex(ByVal strCardType As String) As Integer
    GetKindIndex = IDKind.GetKindIndex(strCardType)
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,GetfaultCard
Public Property Get GetfaultCard() As Object
    Set GetfaultCard = IDKind.GetfaultCard
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,IDKind
Public Property Get IDKindIDX() As Integer
    IDKindIDX = IDKind.IDKind
End Property

Public Property Let IDKindIDX(ByVal New_IDKind As Integer)
    IDKind.IDKind() = New_IDKind
    PropertyChanged "IDKindIDX"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,NotContainFastKey
Public Property Get NotContainFastKey() As String
    NotContainFastKey = IDKind.NotContainFastKey
End Property

Public Property Let NotContainFastKey(ByVal New_NotContainFastKey As String)
    IDKind.NotContainFastKey() = New_NotContainFastKey
    PropertyChanged "NotContainFastKey"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,Refrash
Public Sub Refrash()
    Call IDKind.Refrash
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,NotAutoAppendKind
Public Property Get NotAutoAppendKind() As Boolean
    NotAutoAppendKind = IDKind.NotAutoAppendKind
End Property

Public Property Let NotAutoAppendKind(ByVal New_NotAutoAppendKind As Boolean)
    IDKind.NotAutoAppendKind() = New_NotAutoAppendKind
    PropertyChanged "NotAutoAppendKind"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = txtPatient.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtPatient.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = txtPatient.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtPatient.BackColor() = New_BackColor
    picTxtBack.BackColor() = New_BackColor
    Call SetInputAppearance
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "返回/设置选定文本的起始点。"
    SelStart = txtPatient.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtPatient.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "返回/设置选定的字符数。"
    SelLength = txtPatient.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtPatient.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "返回/设置包含当前选定文本的字符串。"
    SelText = txtPatient.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtPatient.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "返回/设置一个值，决定是否在控件中显示用户键入字符或保留区字符。"
    PasswordChar = txtPatient.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtPatient.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtPatient,txtPatient,-1,IMEMode
Public Property Get IMEMode() As Integer
Attribute IMEMode.VB_Description = "返回/设置输入方法编辑器的当前操作模式。"
    IMEMode = txtPatient.IMEMode
End Property

Public Property Let IMEMode(ByVal New_IMEMode As Integer)
    txtPatient.IMEMode() = New_IMEMode
    PropertyChanged "IMEMode"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=IDKind,IDKind,-1,NotAutoSel
Public Property Get NotAutoSel() As Boolean
    NotAutoSel = mblnNotAutoSel
End Property

Public Property Let NotAutoSel(ByVal New_NotAutoSel As Boolean)
    mblnNotAutoSel = New_NotAutoSel
    PropertyChanged "NotAutoSel"
End Property
Public Function zlGetPatiInforFromPatiID(ByVal lng病人ID As Long, ByRef objPatiInfor As Object, ByRef strErrMsg As String, Optional strOtherName As String = "", _
    Optional strOtherValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID,获取病信息并将病人信息加载给病人对象
    '入参:lng病人ID-病人ID
    '     strOtherName-其他条件名称:如门诊号,住院号，医保号等
    '     strOtherValue-其他条件值
    '出参:objPati-返回病人信息对象
    '     strErrMsg-发生错误时，返回的错误信息
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-10 15:17:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiInforFromPatiID = IDKind.zlGetPatiInforFromPatiID(lng病人ID, objPatiInfor, strErrMsg, strOtherName, strOtherValue)
End Function

Public Function zlGetPatiIDFromCardType(ByVal strCardType As String, ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef lng病人ID As Long, _
    Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, Optional objCtl As Object = Nothing, Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0, Optional ByVal blnNotCheckValidDate As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的医疗类别和卡号,获取对应的病人ID
    '入参:strCardType-卡类别,如果为数字,这为卡类别ID,如果为字符,则为类别名称
    '       strCardNo-卡号
    '       blnNotShowErrMsg-不显示错误的提示信息
    '       frmMain-调用的主窗体
    '       objCtl-调用的控件
    '       blnShowMergePati-当出现多个满足条件的病人时,是否显示合并功能按钮
    '       blnOnlyContractPati-签约病人
    '       blnUserCancel-选择器中，用户选择了取消
    '       lngShowCardNoTypeID-过滤出多条病信息时，弹出选择器中显示的卡号的卡类别ID,0-表示不显示卡号；>0表示显示指定卡号类别的ID
    '       blnNotCheckValidDate-是否对卡终止使用时间进行检查,true-不检查终止使用时间,false-检查
    '出参:strErrMsg-返回的错误信息
    '       lng病人ID-返回的病人ID
    '       strCardPass-返回卡号的密码
    '       lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '返回:获取病人ID成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-14 17:07:51
    '说明:只有存在医疗类别的才调用此函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiIDFromCardType = IDKind.zlGetPatiIDFromCardType(strCardType, strCardNo, blnNotShowErrMsg, lng病人ID, _
        strCardPassWord, strErrMsg, lngCardTypeID, objCtl, frmMain, blnShowMergePati, blnOnlyContractPati, _
        blnCertificate, blnUserCancel, lngShowCardNoTypeID, blnNotCheckValidDate)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Property Get GetPubOneCardObject() As clsPublicOneCard
    Set GetPubOneCardObject = IDKind.GetPubOneCardObject
End Property
 
