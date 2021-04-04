VERSION 5.00
Begin VB.Form frmSquareBrushCardSimple 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结算卡刷卡"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSquareBrushCardSimple.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   465
      Left            =   5010
      TabIndex        =   12
      Top             =   3330
      Width           =   1335
   End
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消(&C)"
      Height          =   465
      Left            =   6360
      TabIndex        =   14
      Top             =   3330
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "刷卡信息"
      Height          =   2880
      Left            =   180
      TabIndex        =   13
      Top             =   255
      Width           =   7605
      Begin VB.TextBox txtEdit 
         Height          =   405
         Index           =   3
         Left            =   1695
         MaxLength       =   100
         TabIndex        =   11
         Top             =   2115
         Width           =   5745
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1695
         TabIndex        =   9
         Top             =   1575
         Width           =   2535
      End
      Begin VB.TextBox txtEdit 
         Height          =   405
         Index           =   0
         Left            =   1695
         TabIndex        =   1
         Top             =   495
         Width           =   2550
      End
      Begin VB.TextBox txtEdit 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1695
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1020
         Width           =   2550
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "备注(&S)"
         Height          =   285
         Index           =   4
         Left            =   615
         TabIndex        =   10
         Top             =   2175
         Width           =   1020
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "本次刷卡(&X)"
         Height          =   285
         Index           =   3
         Left            =   45
         TabIndex        =   8
         Top             =   1635
         Width           =   1590
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Index           =   1
         Left            =   5445
         TabIndex        =   7
         Top             =   1020
         Width           =   1935
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Index           =   0
         Left            =   5430
         TabIndex        =   3
         Top             =   495
         Width           =   1965
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "卡类型"
         Height          =   285
         Index           =   0
         Left            =   4530
         TabIndex        =   2
         Top             =   555
         Width           =   855
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "卡号(&N)"
         Height          =   285
         Index           =   1
         Left            =   615
         TabIndex        =   0
         Top             =   555
         Width           =   1020
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "密码(&W)"
         Height          =   285
         Index           =   2
         Left            =   615
         TabIndex        =   4
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "当前余额"
         Height          =   285
         Index           =   10
         Left            =   4260
         TabIndex        =   6
         Top             =   1080
         Width           =   1140
      End
   End
   Begin VB.Label lbl失效额 
      Height          =   240
      Left            =   210
      TabIndex        =   15
      Top             =   3420
      Width           =   4455
   End
End
Attribute VB_Name = "frmSquareBrushCardSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng接口编号 As Long, mdbl本次刷卡 As Double, mstrBlanceInfor As String, mintSucces As Integer
Private Type CardInfor
    lng消费卡ID As Long
    str卡号 As String
    dbl余额 As Double
    dbl最大消费额 As Double
    dbl失效面额 As Double '采取的原则是,先进先出的法则:先消费卡面额,再消费允值额:此金额为,到期后未消费的金额
    str限制类别 As String
    str接口名称 As String
    str结算方式 As String
End Type
Private mTyCurCardInfor As CardInfor

Private Enum mtxtIdx
    idx_txt卡号 = 0
    idx_txt密码 = 1
    idx_txt本次刷卡 = 2
    idx_txt备注 = 3
End Enum
Private Enum mlblIdx
    idx_lbl卡类型 = 0
    idx_lbl余额 = 1
    idx_lbl本次刷卡 = 3
    idx_lbl备注 = 4
End Enum
Private WithEvents mobjBrushCard As clsBrushSequareCard
Attribute mobjBrushCard.VB_VarHelpID = -1
Private mblnChange As Boolean
Private mblnCardNoSHowPW As Boolean
Private mobjKeyboard As Object

Private Function CheckDepended() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的关联性
    '编制:刘兴洪
    '日期:2009-12-24 12:13:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Dim rsTemp As New ADODB.Recordset
    Set rsTemp = zlGet消费卡接口
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    rsTemp.Find "编号=" & mlng接口编号, , , 1
    If rsTemp.EOF Then
        ShowMsgbox "接口未找到(编号为" & mlng接口编号 & "),请检查!"
        Exit Function
    End If
    With mTyCurCardInfor
        .str接口名称 = Nvl(rsTemp!名称)
        .str结算方式 = Nvl(rsTemp!结算方式)
        txtEdit(mtxtIdx.idx_txt卡号).MaxLength = Len(Nvl(rsTemp!前缀文本)) + Val(Nvl(rsTemp!卡号长度))
    End With
    CheckDepended = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlShowBrushCard(ByVal frmMain As Object, ByVal lng接口编号 As Long, dbl本次刷卡 As Double, _
    strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：刷卡接口
    '入参：frmMain-调用的主窗体
    '       dbl本次刷卡-本次刷卡额
    '出参:strBlanceInfor-返回结算信息( 用||分隔: 接口编号||消费卡ID(可传'')||结算方式||结算金额||卡号||交易流水号||交易时间(yyyy-mm-dd hh24:mi:ss)||备注)
    '返回:调用成功,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-18 15:27:01
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    mdbl本次刷卡 = dbl本次刷卡: mlng接口编号 = lng接口编号: mintSucces = 0
    If CheckDepended = False Then Exit Function
    
    txtEdit(mtxtIdx.idx_txt本次刷卡).Text = Format(dbl本次刷卡, "###0.00;-###0.00;0.00;0.00")
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlShowBrushCard = mintSucces > 0
    strBlanceInfor = mstrBlanceInfor
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmd取消_Click()
    mintSucces = 0: Unload Me
End Sub
Private Sub cmd确定_Click()
    Dim dt交易时间 As Date

    '不存在,表示需要检查是否合法
    If CheckInput = False Then Exit Sub
    dt交易时间 = zlDatabase.Currentdate
    
    ' 接口编号||消费卡ID(可传'')||结算方式||结算金额||卡号||交易流水号||交易时间(yyyy-mm-dd hh24:mi:ss)||备注
    mstrBlanceInfor = mlng接口编号
    mstrBlanceInfor = mstrBlanceInfor & "||" & mTyCurCardInfor.lng消费卡ID
    mstrBlanceInfor = mstrBlanceInfor & "||" & mTyCurCardInfor.str结算方式
    mstrBlanceInfor = mstrBlanceInfor & "||" & Val(txtEdit(mtxtIdx.idx_txt本次刷卡).Text)
    mstrBlanceInfor = mstrBlanceInfor & "||" & mTyCurCardInfor.str卡号
    mstrBlanceInfor = mstrBlanceInfor & "||" & ""
    mstrBlanceInfor = mstrBlanceInfor & "||" & Format(dt交易时间, "yyyy-mm-dd HH:MM:SS")
    mstrBlanceInfor = mstrBlanceInfor & "||" & Replace(Trim(txtEdit(mtxtIdx.idx_txt本次刷卡).Text), "|", "")
    mintSucces = mintSucces + 1
    mblnChange = False
    Unload Me
End Sub
Private Function CheckInput() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入是否合法
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 17:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    If txtEdit(mtxtIdx.idx_txt卡号).Text <> Trim(txtEdit(mtxtIdx.idx_txt卡号).Tag) Or Trim(txtEdit(mtxtIdx.idx_txt卡号).Text) = "" Then
        ShowMsgbox "未刷卡或刷卡不正确,请检查!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt卡号)
        Exit Function
    End If
    If CheckInputPassWord = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt密码)
        Exit Function
    End If
    If CheckInputSquareMoney = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt本次刷卡)
        Exit Function
    End If
    CheckInput = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckInputPassWord() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查密码输入是否正确
    '返回:正确,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 14:27:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Trim(txtEdit(mtxtIdx.idx_txt密码).Tag) <> "" And Trim(txtEdit(mtxtIdx.idx_txt密码).Text) = "" Then
        ShowMsgbox "密码未输入,请检查!"
        Exit Function
    End If
    
    If Trim(txtEdit(mtxtIdx.idx_txt密码).Tag) <> zlCommFun.zlStringEncode(Trim(txtEdit(mtxtIdx.idx_txt密码).Text)) Then
        ShowMsgbox "密码输入错误,请检查!"
        Exit Function
    End If
    CheckInputPassWord = True
End Function

Private Function CheckInputSquareMoney() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的本次消费金额是否正确
    '返回:正确,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 14:27:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt本次刷卡).Text), 16, True, True, 0, "卡面额") = False Then
        Exit Function
    End If
    If Val(lblInfor(mlblIdx.idx_lbl余额).Caption) < Val(Trim(txtEdit(mtxtIdx.idx_txt本次刷卡).Text)) Then
        ShowMsgbox "卡余额不足(" & Format(Val(lblInfor(mlblIdx.idx_lbl余额).Caption), gVbFmtString.FM_金额) & "元),请检查!"
        Exit Function
    End If
    CheckInputSquareMoney = True
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call CreateObjectKeyboard
    '检查是否启用了相关的刷卡程序
    Set mobjBrushCard = New clsBrushSequareCard
    Call mobjBrushCard.zlInitInterFacel(mlng接口编号)
    mblnCardNoSHowPW = zlIsCardNoShowPW(mlng接口编号)
    If mblnCardNoSHowPW Then
        txtEdit(mtxtIdx.idx_txt卡号).PasswordChar = "*"
    Else
        txtEdit(mtxtIdx.idx_txt卡号).PasswordChar = ""
    End If
    
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index <> mtxtIdx.idx_txt密码 Then txtEdit(Index).Tag = ""
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Select Case Index
    Case mtxtIdx.idx_txt卡号
        zlCommFun.OpenIme False
        gTy_TestBug.BytType = 2
        If Not mobjBrushCard Is Nothing Then Call mobjBrushCard.zlSetAutoBrush(Trim(txtEdit(Index).Text) = "")
    Case mtxtIdx.idx_txt备注
        zlCommFun.OpenIme True
    Case Else
        zlCommFun.OpenIme False
        If Index = mtxtIdx.idx_txt密码 Then
            Call OpenPassKeyboard(txtEdit(Index))
        End If
    End Select
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim str编码 As String, str名称 As String, lngID As Long
    Dim strCardNo As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case mtxtIdx.idx_txt卡号
        If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab
        '考虑可能存在操作员乱刷卡的情况,因此暂不开放如下功能:
        If IsDesinMode = False Then Exit Sub
        
        If txtEdit(Index).Text = "" Then
            '直接调读卡
            If mobjBrushCard.zlReadCard(Me, strCardNo) = False Then
                Exit Sub
            End If
            txtEdit(Index).Text = strCardNo
            txtEdit(Index).Tag = strCardNo
        End If
        
        If zlBrusCard(Trim(txtEdit(Index))) = False Then
            zlCtlSetFocus txtEdit(Index)
        Else
            If txtEdit(mtxtIdx.idx_txt密码).Tag = "" Then
                If txtEdit(mtxtIdx.idx_txt备注).Enabled And txtEdit(mtxtIdx.idx_txt备注).Visible Then txtEdit(mtxtIdx.idx_txt备注).SetFocus
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        End If
        
    Case mtxtIdx.idx_txt备注
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    Case mtxtIdx.idx_txt密码
        If CheckInputPassWord = False Then
            zlControl.ControlSetFocus txtEdit(Index): Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt本次刷卡
        If CheckInputSquareMoney = False Then
            zlControl.ControlSetFocus txtEdit(Index): Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim blnCard As Boolean
    
    Select Case Index
    Case mtxtIdx.idx_txt卡号
        If InStr(1, "'~～|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If IsDesinMode Then Exit Sub
        Call BrushCard(txtEdit(Index), KeyAscii)
    Case mtxtIdx.idx_txt备注
        If InStr(1, "'|'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        blnCard = zlInputIsCard(txtEdit(Index), KeyAscii, glngSys, mblnCardNoSHowPW)
        If blnCard = True Then KeyAscii = 0
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
    Case mtxtIdx.idx_txt密码
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
    Case mtxtIdx.idx_txt本次刷卡
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m金额式
    Case Else
    End Select
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = mtxtIdx.idx_txt密码 Then
        Call ClosePassKeyboard(txtEdit(Index))
    End If
End Sub

Private Sub txtEdit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> mtxtIdx.idx_txt卡号 Then Exit Sub
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtEdit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> mtxtIdx.idx_txt卡号 Then Exit Sub
    If Button = 2 Then
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case mtxtIdx.idx_txt卡号
    Case mtxtIdx.idx_txt备注
    Case mtxtIdx.idx_txt密码
        If CheckInputPassWord = False Then
        End If
    Case mtxtIdx.idx_txt本次刷卡
        If CheckInputSquareMoney = False Then
           'Cancel = 1
        End If
    Case Else
    End Select
End Sub

Private Function zlBrusCard(ByVal strCardNo As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：刷卡操作
    '编制：刘兴洪
    '日期：2010-06-18 15:12:22
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim rsTemp As ADODB.Recordset, i As Long, strTemp As String, blnFind As Boolean
    
    With mTyCurCardInfor
        .dbl失效面额 = 0
        .dbl余额 = 0
        .dbl最大消费额 = 0
        .str卡号 = ""
        .lng消费卡ID = 0
    End With
    
    gstrSQL = "" & _
    "   Select a.Id,a.卡类型,a.卡号,a.序号,a.可否充值,to_char(a.有效期,'yyyy-mm-dd hh24:mi:ss') as 有效期,  a.密码," & _
    "          to_char(a.回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间 , " & _
    "          decode(a.当前状态,2,'回收',3,'退卡','回收') as 当前状态, " & _
    "          to_char(a.卡面金额," & gOraFmtString.FM_金额 & ") as 卡面金额 ," & _
    "          to_char(a.销售金额," & gOraFmtString.FM_金额 & ") as 销售金额 ," & _
    "          to_char(a.充值折扣率," & gOraFmtString.FM_折扣率 & ") as 充值折扣率 ," & _
    "          to_char(a.余额," & gOraFmtString.FM_金额 & ") as 余额 ," & _
    "          to_char(a.停用日期,'yyyy-mm-dd hh24:mi:ss') as 停用日期," & _
    "          a.限制类别 " & _
    "   From 消费卡目录 A  " & _
    "   Where A.卡号 = [1] and A.接口编号=[2] And 序号 = (Select Max(序号) From 消费卡目录 B Where 卡号 = A.卡号 and 接口编号=A.接口编号)  " & _
    "   Order by a.序号"
    Err = 0: On Error GoTo ErrHand:
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strCardNo, mlng接口编号)
    If rsTemp.EOF Then
       ShowMsgbox "未找到相关的消费卡记录,请检查!"
        Exit Function
    End If
    '检查:
    '是否回收
    If Nvl(rsTemp!回收时间, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox IIf(mblnCardNoSHowPW, "", "卡号为" & strCardNo & "的") & "消费卡已经被" & Nvl(rsTemp!当前状态) & ",不能再刷卡"
        Exit Function
    End If
    '是否停用
    If Nvl(rsTemp!停用日期, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox IIf(mblnCardNoSHowPW, "", "卡号为" & strCardNo & "的") & "消费卡已经被停止使用,不能再刷卡"
        Exit Function
    End If
    '是否停用
    If Nvl(rsTemp!停用日期, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox IIf(mblnCardNoSHowPW, "", "卡号为" & strCardNo & "的") & "消费卡已经被停止使用,不能再刷卡"
        Exit Function
    End If
    
    '检查效期
    mTyCurCardInfor.dbl余额 = Val(Nvl(rsTemp!余额))
    lbl失效额.Visible = False
    If Nvl(rsTemp!有效期, "3000-01-01 00:00:00") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
       '到了有效期
       If Val(Nvl(rsTemp!可否充值)) = 1 Then
          '允许允值的,到期的,不能消费卡面金额,只能消费允值部分
          mTyCurCardInfor.dbl失效面额 = zlGet失效面额(Val(Nvl(rsTemp!ID)), mlng接口编号)
          mTyCurCardInfor.dbl余额 = IIf(mTyCurCardInfor.dbl余额 - mTyCurCardInfor.dbl失效面额 < 0, 0, mTyCurCardInfor.dbl余额 - mTyCurCardInfor.dbl失效面额)
          If mTyCurCardInfor.dbl失效面额 <> 0 Then
            lbl失效额.Caption = "当前卡号失效金额(卡面额)为：" & Format(mTyCurCardInfor.dbl失效面额, gVbFmtString.FM_金额) & "元"
            lbl失效额.Visible = True
            lbl失效额.ForeColor = vbRed
          End If
       Else
            '不允许允值的,不能再进行消费
            ShowMsgbox IIf(mblnCardNoSHowPW, "", "卡号为" & strCardNo & "的") & "消费卡已经失效,不能再刷卡"
            Exit Function
       End If
    End If
    If mTyCurCardInfor.dbl余额 <= 0 Then
        ShowMsgbox IIf(mblnCardNoSHowPW, "", "卡号为" & strCardNo & "的") & "消费卡已经没有余额,不能再刷卡"
        Exit Function
    End If
    
    With mTyCurCardInfor
        .lng消费卡ID = Val(Nvl(rsTemp!ID))
        .str卡号 = Nvl(rsTemp!卡号)
        .str限制类别 = Nvl(rsTemp!限制类别)
    End With
    txtEdit(mtxtIdx.idx_txt卡号).Text = Nvl(rsTemp!卡号)
    txtEdit(mtxtIdx.idx_txt卡号).Tag = Nvl(rsTemp!卡号)
    lblInfor(mlblIdx.idx_lbl余额).Caption = Format(Val(Nvl(rsTemp!余额)), gVbFmtString.FM_金额)
    lblInfor(mlblIdx.idx_lbl卡类型).Caption = Nvl(rsTemp!卡类型)
    txtEdit(mtxtIdx.idx_txt密码).Tag = Nvl(rsTemp!密码)
    '缺省值:余额不足,缺省余额,否则为最大消费额
    If mTyCurCardInfor.dbl余额 < mdbl本次刷卡 Then
        txtEdit(mtxtIdx.idx_txt本次刷卡).Text = Format(mTyCurCardInfor.dbl余额, gVbFmtString.FM_金额)
    Else
        txtEdit(mtxtIdx.idx_txt本次刷卡).Text = Format(mdbl本次刷卡, gVbFmtString.FM_金额)
    End If
    zlBrusCard = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub ClearCtlData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除控件数据
    '编制:刘兴洪
    '日期:2009-12-24 11:11:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    txtEdit(mtxtIdx.idx_txt本次刷卡) = "0.00"
    txtEdit(mtxtIdx.idx_txt卡号) = ""
    txtEdit(mtxtIdx.idx_txt密码) = ""
    txtEdit(mtxtIdx.idx_txt密码).Tag = ""
    lblInfor(mlblIdx.idx_lbl卡类型).Caption = ""
    lblInfor(mlblIdx.idx_lbl余额).Caption = "0.00"
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub BrushCard(ByVal objEdit As Object, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡操作(目前只支持有卡进行刷卡)
    '编制:刘兴洪
    '日期:2010-02-09 14:07:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    
    '是否刷卡完成
    blnCard = KeyAscii <> 8 And Len(objEdit.Text) = objEdit.MaxLength - 1 And objEdit.SelLength <> Len(objEdit.Text)
    
    If blnCard Then
        If KeyAscii <> 13 Then
            objEdit.Text = objEdit.Text & Chr(KeyAscii)
            objEdit.SelStart = Len(objEdit.Text)
        End If
        KeyAscii = 0
        '刷卡处理:
        If zlBrusCard(Trim(objEdit)) = False Then
            zlCtlSetFocus objEdit
        Else
            If txtEdit(mtxtIdx.idx_txt密码).Tag = "" Then
                If txtEdit(mtxtIdx.idx_txt备注).Enabled And txtEdit(mtxtIdx.idx_txt备注).Visible Then txtEdit(mtxtIdx.idx_txt备注).SetFocus
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        End If
    Else
        If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = Timer
            If objEdit.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(objEdit.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                objEdit.Text = Chr(KeyAscii)
                objEdit.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
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


