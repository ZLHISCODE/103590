VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmModiCardPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "密码修改"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   Icon            =   "frmModiCardPass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
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
      Left            =   6120
      TabIndex        =   5
      Top             =   360
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
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
      Left            =   6120
      TabIndex        =   6
      Top             =   870
      Width           =   1200
   End
   Begin VB.PictureBox picPass 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   285
      ScaleHeight     =   3735
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   285
      Width           =   5625
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1095
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   3030
         Width           =   4245
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1095
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2520
         Width           =   4245
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2010
         Width           =   4245
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   5520
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请将[XX]从刷卡器上划过，  然后连续两次输入相同的密码！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   180
         Width           =   8550
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   120
         Picture         =   "frmModiCardPass.frx":058A
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验证"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   390
         TabIndex        =   10
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   9
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   390
         TabIndex        =   7
         Top             =   1500
         Width           =   630
      End
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   4065
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5955
      _Version        =   589884
      _ExtentX        =   10504
      _ExtentY        =   7170
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmModiCardPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mlngCardTypeID As Long
Private mblnCheckOldPass As Boolean

Private mobjKeyboard As Object, mblnTest As Boolean
Private mblnFirst As Boolean
Private mblnOK As Boolean

Private Enum mTextIndex
    txt_卡号 = 0
    txt_原密码 = 1
    txt_新密码 = 2
    txt_验证密码 = 3
End Enum

Private Type Ty_CardType '卡类别信息
    str名称 As String
    lng卡号长度 As Long
    lng密码长度 As Long
    int密码长度限制 As Integer
    byt密码规则 As Byte
End Type
Private mTy_CardType As Ty_CardType
Private mlngCardID  As Long
Private mstrOldPassWord As String

Public Function zlModifyPass(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, Optional blnCheckOldPass As Boolean = True) As Boolean
    '功能:调整密码入口参数
    '入参:frmMain-调用的主窗体
    '     lngModule -模块号
    '     lngCardTypeID-消费卡接口编号
    '返回:修改成功,返回true,否则返回false
    mlngModule = lngModule: mlngCardTypeID = lngCardTypeID
    mblnCheckOldPass = blnCheckOldPass
    
    mblnOK = False
    On Error Resume Next
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlModifyPass = mblnOK
End Function

Private Sub Form_Load()
    mblnFirst = True
    
    mblnTest = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard", "TestCardNO", 0)) = 1
    mblnTest = IsDesinMode Or mblnTest
    
    If mblnCheckOldPass = False Then
        lblNotes.Top = 180
        lblNotes.Caption = "请将[XX]从刷卡器上划过后，" & vbCrLf & "连续两次输入相同的新密码！"
        txtEdit(mTextIndex.txt_原密码).Enabled = False
        txtEdit(mTextIndex.txt_原密码).BackColor = &H8000000F
    Else
        lblNotes.Top = 180
        lblNotes.Caption = "请将[XX]从刷卡器上划过后，" & vbCrLf & "输入旧密码与两次相同的新密码！"
    End If
    
    If InitCardInfor() = False Then
        ShowMsgbox "当前卡类别未启用或已被删除，您不能进行修改密码操作，请到【参数设置>设备配置】或者 【消费卡管理】中启用！"
        Unload Me: Exit Sub
    End If
    Call ClearFace
    
    Call CreateObjectKeyboard
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    zlControl.ControlSetFocus txtEdit(mTextIndex.txt_卡号)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
        '去除特殊符号，并且不允许粘贴
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Function InitCardInfor() As Boolean
    '功能:初始化卡类别信息
    '返回:初始化成功,返回true,否则返回False
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    Set rsTemp = zlGet消费卡接口
    rsTemp.Filter = "编号=" & mlngCardTypeID
    If rsTemp.EOF Then Exit Function
    
    With mTy_CardType
        .str名称 = Nvl(rsTemp!名称)
        .lng卡号长度 = Val(Nvl(rsTemp!卡号长度))
        .lng密码长度 = Val(Nvl(rsTemp!密码长度))
        .int密码长度限制 = Val(Nvl(rsTemp!密码长度限制))
        .byt密码规则 = Nvl(rsTemp!密码规则)
    End With
    
    lblNotes.Caption = Replace(lblNotes.Caption, "[XX]", "[" & mTy_CardType.str名称 & "]")
    InitCardInfor = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码键盘
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
    If ErrCenter() = 1 Then Resume
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否则False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, bln确认密码) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否则False
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function isValied() As Boolean
    '功能:检查输入的数据是否有效
    '返回:数据有效,返回true,否则返回False
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, strPassWord As String
    
    On Error GoTo ErrHandler
    strCardNo = Trim(txtEdit(mTextIndex.txt_卡号).Text)
    If CheckCardNo(mlngCardTypeID, strCardNo) = False Then Exit Function
    
    strPassWord = zlCommFun.zlStringEncode(txtEdit(mTextIndex.txt_原密码).Text) '密码加密
    If strPassWord <> mstrOldPassWord And mblnCheckOldPass Then
        ShowMsgbox "卡片原密码输入错误，请重新输入密码！"
        zlControl.ControlSetFocus txtEdit(mTextIndex.txt_原密码)
        Exit Function
    End If
    
    If txtEdit(mTextIndex.txt_新密码).Text = "" Then
        If MsgBox("当前设置的密码为空，确定要这样设置吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            zlControl.ControlSetFocus txtEdit(mTextIndex.txt_新密码)
            Exit Function
        End If
    Else
        Select Case mTy_CardType.int密码长度限制
        Case 0
        Case 1
            If Len(txtEdit(mTextIndex.txt_新密码).Text) <> mTy_CardType.lng密码长度 Then
                ShowMsgbox "密码必须输入" & mTy_CardType.lng密码长度 & "位！"
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_新密码)
                zlControl.TxtSelAll txtEdit(mTextIndex.txt_新密码)
                Exit Function
             End If
        Case Else
            If Len(txtEdit(mTextIndex.txt_新密码).Text) <= Abs(mTy_CardType.int密码长度限制) Then
                ShowMsgbox "密码必须输入" & Abs(mTy_CardType.int密码长度限制) & "位以上！"
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_新密码)
                zlControl.TxtSelAll txtEdit(mTextIndex.txt_新密码)
                Exit Function
             End If
        End Select
        If mTy_CardType.byt密码规则 = 1 Then '密码只允许为数字
            If IsNumeric(txtEdit(mTextIndex.txt_新密码).Text) = False Then
                ShowMsgbox "密码只能包含数字，请重新输入！"
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_新密码)
                zlControl.TxtSelAll txtEdit(mTextIndex.txt_新密码)
                Exit Function
            End If
        End If
    End If
    
    If txtEdit(mTextIndex.txt_新密码).Text <> txtEdit(mTextIndex.txt_验证密码).Text Then
        ShowMsgbox "两次输入的密码不一致，请重新输入！"
        zlControl.ControlSetFocus txtEdit(mTextIndex.txt_新密码)
        zlControl.TxtSelAll txtEdit(mTextIndex.txt_新密码)
        Exit Function
    End If
    
    isValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    If isValied = False Then Exit Sub
    
    'Zl_消费卡密码_Update
    strSQL = "Zl_消费卡密码_Update("
    '  消费卡id_In   In 消费卡信息.Id%Type,
    strSQL = strSQL & "" & mlngCardID & ","
    '  密码_In       In 消费卡信息.密码%Type,
    strSQL = strSQL & "'" & mstrOldPassWord & "',"
    '  修改密码_In   In 消费卡信息.密码%Type,
    strSQL = strSQL & "'" & zlCommFun.zlStringEncode(txtEdit(mTextIndex.txt_新密码).Text) & "',"
    '  操作员姓名_In In 消费卡变动记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  强制修改_In   In Number := 0 --是否强制修改密码,0-否,1-是
    strSQL = strSQL & "" & IIf(mblnCheckOldPass, 0, 1) & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    MsgBox "密码修改成功！", vbOKOnly + vbInformation, gstrSysName
    mblnOK = True
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngCardID = 0
    mstrOldPassWord = ""
End Sub

Private Sub ClearFace()
    txtEdit(mTextIndex.txt_卡号).PasswordChar = IIf(mTy_CardType.byt密码规则 <> 0, "*", "")
    txtEdit(mTextIndex.txt_卡号).Text = ""
    txtEdit(mTextIndex.txt_新密码).Text = "": txtEdit(mTextIndex.txt_验证密码).Text = ""
End Sub

Private Function CheckCardNo(ByVal lngCardTypeID As Long, ByVal strCardNo As String) As Boolean
    '检查卡号的合法性
    '入参:lngCardTypeID-消费卡接口编号
    '     strCardNO-卡号
    '返回:成功返回True,失败返回False
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = _
        "Select a.Id, a.卡类型, a.卡号, a.序号, a.可否充值, a.接口编号, a.密码, a.限制类别," & vbNewLine & _
        "       To_Char(a.有效期, 'yyyy-mm-dd hh24:mi:ss') As 有效期," & vbNewLine & _
        "       To_Char(a.回收时间, 'yyyy-mm-dd hh24:mi:ss') As 回收时间," & vbNewLine & _
        "       To_Char(a.停用日期, 'yyyy-mm-dd hh24:mi:ss') As 停用日期," & vbNewLine & _
        "       Decode(a.当前状态, 2, '回收', 3, '退卡', '回收') As 当前状态" & vbNewLine & _
        "From 消费卡信息 A" & vbNewLine & _
        "Where a.卡号 = [1] And a.接口编号 = [2]" & vbNewLine & _
        "      And 序号 = (Select Max(序号) From 消费卡信息 B Where 卡号 = a.卡号 And 接口编号 = a.接口编号)" & vbNewLine & _
        "Order By a.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCardNo, mlngCardTypeID)
    If rsTemp.EOF Then
        ShowMsgbox "未找到相关的" & mTy_CardType.str名称 & "信息，请检查！"
        Exit Function
    End If
    mlngCardID = Val(Nvl(rsTemp!id))
    mstrOldPassWord = Nvl(rsTemp!密码)
    
    '检查当前刷卡的合法性
    '是否回收
    If Nvl(rsTemp!回收时间, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的" & mTy_CardType.str名称 & "已经被" & Nvl(rsTemp!当前状态) & "，不能再刷卡！"
        Exit Function
    End If
    
    '是否停用
    If Nvl(rsTemp!停用日期, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的" & mTy_CardType.str名称 & "已经被停止使用，不能再刷卡！"
        Exit Function
    End If
    
    CheckCardNo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtEdit_Change(Index As Integer)
    Dim blnEnabled As Boolean
    
    On Error GoTo ErrHandler
    Select Case Index
    Case mTextIndex.txt_卡号
        blnEnabled = Trim(txtEdit(mTextIndex.txt_卡号).Text) <> ""
        If mblnCheckOldPass = True Then txtEdit(mTextIndex.txt_原密码).Enabled = blnEnabled
        txtEdit(mTextIndex.txt_新密码).Enabled = blnEnabled
        txtEdit(mTextIndex.txt_验证密码).Enabled = blnEnabled
            
        txtEdit(mTextIndex.txt_新密码).Text = ""
        txtEdit(mTextIndex.txt_验证密码).Text = ""
        txtEdit(mTextIndex.txt_原密码).Text = ""
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    On Error GoTo ErrHandler
    Select Case Index
    Case mTextIndex.txt_卡号
        txtEdit(mTextIndex.txt_卡号).PasswordChar = IIf(mTy_CardType.byt密码规则 <> 0, "*", "")
    Case mTextIndex.txt_原密码, mTextIndex.txt_新密码, mTextIndex.txt_验证密码
        OpenPassKeyboard txtEdit(Index), True
    End Select
    zlControl.TxtSelAll txtEdit(Index)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    
    On Error GoTo ErrHandler
    Select Case Index
    Case mTextIndex.txt_卡号
        '是否刷卡完成
        blnCard = KeyAscii <> 8 And Len(txtEdit(mTextIndex.txt_卡号).Text) = mTy_CardType.lng卡号长度 - 1 _
            And txtEdit(mTextIndex.txt_卡号).SelLength <> Len(txtEdit(mTextIndex.txt_卡号).Text)
        If blnCard Or KeyAscii = 13 Then
            If KeyAscii <> 13 Then
                txtEdit(mTextIndex.txt_卡号).Text = txtEdit(mTextIndex.txt_卡号).Text & Chr(KeyAscii)
                txtEdit(mTextIndex.txt_卡号).SelStart = Len(txtEdit(mTextIndex.txt_卡号).Text)
            End If
            KeyAscii = 0
    
            If CheckCardNo(mlngCardTypeID, Trim(txtEdit(mTextIndex.txt_卡号).Text)) = False Then
                If txtEdit(mTextIndex.txt_卡号).Enabled Then txtEdit(mTextIndex.txt_卡号).SetFocus
                zlControl.TxtSelAll txtEdit(mTextIndex.txt_卡号)
                Exit Sub
            End If
            If mblnCheckOldPass Then
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_原密码)
            Else
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_新密码): Exit Sub
            End If
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
                If txtEdit(mTextIndex.txt_卡号).Text = "" Then
                    sngBegin = sngNow
                ElseIf Format((sngNow - sngBegin) / (Len(txtEdit(mTextIndex.txt_卡号).Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                    txtEdit(mTextIndex.txt_卡号).Text = Chr(KeyAscii)
                    txtEdit(mTextIndex.txt_卡号).SelStart = 1
                    KeyAscii = 0
                    sngBegin = sngNow
                End If
            End If
        End If
    Case mTextIndex.txt_原密码
        If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
    Case mTextIndex.txt_新密码
        Call CheckInputPassWord(KeyAscii, mTy_CardType.byt密码规则 = 1)
        If KeyAscii = 13 Then
            KeyAscii = 0
            If txtEdit(mTextIndex.txt_新密码).Text = "" And txtEdit(mTextIndex.txt_验证密码).Text = "" Then
                zlControl.ControlSetFocus cmdOK
            Else
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_验证密码)
            End If
        End If
    Case mTextIndex.txt_验证密码
        Call CheckInputPassWord(KeyAscii, mTy_CardType.byt密码规则 = 1)
        If KeyAscii = 13 Then
            KeyAscii = 0: Call cmdOK_Click
        End If
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    On Error GoTo ErrHandler
    Select Case Index
    Case mTextIndex.txt_原密码, mTextIndex.txt_新密码, mTextIndex.txt_验证密码
        OpenPassKeyboard txtEdit(Index), False
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
