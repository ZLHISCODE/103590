VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmModiPatiPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "密码修改"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9135
   Icon            =   "frmModiPatiPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPass 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   180
      ScaleHeight     =   3495
      ScaleWidth      =   6720
      TabIndex        =   13
      Top             =   540
      Width           =   6720
      Begin VB.TextBox txt卡号 
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
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   1110
         Width           =   4845
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   0
         TabIndex        =   14
         Top             =   750
         Width           =   6555
      End
      Begin VB.TextBox txtPati 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   1095
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1650
         Width           =   4845
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1815
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Left            =   4125
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1815
      End
      Begin VB.TextBox txtPass 
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
         Left            =   1095
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2700
         Width           =   1815
      End
      Begin VB.TextBox txtAudi 
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
         Left            =   4125
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   2700
         Width           =   1815
      End
      Begin VB.Label lbl卡号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   330
         TabIndex        =   17
         Top             =   1140
         Width           =   660
      End
      Begin VB.Label lblNotes 
         BackStyle       =   0  'Transparent
         Caption         =   "请将[XX]从刷卡器上轻轻划过，  然后连续两次输入相同的密码！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1110
         TabIndex        =   15
         Top             =   120
         Width           =   5325
      End
      Begin VB.Label lbl病人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
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
         Left            =   390
         TabIndex        =   1
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
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
         Left            =   390
         TabIndex        =   3
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
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
         Left            =   3360
         TabIndex        =   5
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label Label5 
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
         Left            =   105
         TabIndex        =   7
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label Label6 
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
         Left            =   3390
         TabIndex        =   9
         Top             =   2760
         Width           =   630
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   120
         Picture         =   "frmModiPatiPass.frx":06EA
         Top             =   0
         Width           =   720
      End
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
      Left            =   7410
      TabIndex        =   12
      Top             =   825
      Width           =   1500
   End
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
      Left            =   7410
      TabIndex        =   11
      Top             =   315
      Width           =   1500
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   4545
      Left            =   60
      TabIndex        =   16
      Top             =   150
      Width           =   7275
      _Version        =   589884
      _ExtentX        =   12832
      _ExtentY        =   8017
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmModiPatiPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'------------------------------------------------------
'入参
Private mlngModule As Long, mlngCardTypeID As Long
Private mstrCardNo As String, mlng病人ID As Long
'-------------------------------------------------------
Private mblnDO As Boolean
Private mobjKeyboard As Object
Private mblnOk As Boolean
Private mrsInfo As ADODB.Recordset
Private mobjCardObject As clsCardObject
Private mobjICCard As Object
Private mblnFirst As Boolean
Private mblnCheckOldPass As Boolean
Private WithEvents mobjIDCard As zlIDCard.clsIDCard '问题号:54278
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents '问题号:56597
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mobjSquare As Object '问题号:56597

Public Function zlModifyPass(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional lng病人ID As Long, Optional strCardNo As String, _
    Optional blnCheckOldPass As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:入口
    '入参:frmMain-调用的主窗体
    '       lngModule -模块号
    '       lngCardTypeId-卡类别ID
    '       lng病人ID-病人ID
    '       strCardNo-卡号
    '返回:修改成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-07-29 11:08:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngCardTypeID = lngCardTypeID: mlngModule = lngModule: mlng病人ID = lng病人ID
    mstrCardNo = strCardNo: mblnOk = False
    mblnCheckOldPass = blnCheckOldPass
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlModifyPass = mblnOk
End Function
Private Sub InitTaskPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化InitTaskPancel
    '编制:刘兴洪
    '日期:2011-06-30 18:20:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    wndTaskPanel.HotTrackStyle = xtpTaskPanelHighlightItem
    Set tkpGroup = wndTaskPanel.Groups.Add(1, "请刷卡后输入修改密码")
    Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
   Set Item.Control = picPass
    tkpGroup.CaptionVisible = False
   ' Call Item.SetMargins(0, -19, 0, -4)
    picPass.BackColor = Item.BackColor
    Me.BackColor = Item.BackColor
    cmdOK.BackColor = Item.BackColor
    cmdCancel.BackColor = Item.BackColor
    tkpGroup.Expandable = False
    wndTaskPanel.Reposition
    wndTaskPanel.DrawFocusRect = True
End Sub
Private Function InitCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡片信息
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-29 14:25:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
   Set mobjCardObject = zlGetClsCardObject(mlngCardTypeID, False)
   If Err <> 0 Then Err = 0: Exit Function
   If mobjCardObject Is Nothing Then Exit Function
   If mobjCardObject.CardPreporty.名称 = "就诊卡" And mobjCardObject.CardPreporty.系统 Then
            lbl卡号.BorderStyle = 1: lbl卡号.Tag = "1"
   Else
        If mobjCardObject.CardPreporty.是否接触式读卡 Then
            lbl卡号.BorderStyle = 1: lbl卡号.Tag = "1"
        Else
            lbl卡号.BorderStyle = 0: lbl卡号.Tag = "0"
        End If
    End If
    '108779：李南春,2017/5/8,密码不固定就不应该限制10位
    If mobjCardObject.CardPreporty.密码长度限制 <> 0 Then
        txtPass.MaxLength = mobjCardObject.CardPreporty.密码长度
        txtAudi.MaxLength = mobjCardObject.CardPreporty.密码长度
    End If
    lblNotes.Caption = Replace(lblNotes.Caption, "[XX]", "[" & mobjCardObject.CardPreporty.名称 & "]")
    InitCardInfor = True
End Function

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

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
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

Private Sub cmdCancel_Click()
    If txtPati.Text <> "" And Val(txtPati.Tag) <> 0 Then
        Call ClearFace:
        If txt卡号.Enabled Then txt卡号.SetFocus
        Exit Sub
    End If
    mstrCardNo = ""
    Unload Me
End Sub
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的数据是否有效
    '返回:数据有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-29 11:15:42
    '---------------------------------------------------------------------------------------------------------------------------------------------\
    Dim str称呼 As String
    str称呼 = IIf(glngSys Like "8??", "客户", "病人")
    
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then
        MsgBox "不能读取" & str称呼 & "信息，请确定是否正确刷卡！", vbInformation, gstrSysName
        Call ClearFace: txt卡号.SetFocus: Exit Function
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "不能读取" & str称呼 & "信息，请确定是否正确刷卡！", vbInformation, gstrSysName
        Call ClearFace: txt卡号.SetFocus: Exit Function
    End If
    If txtPass.Text <> txtAudi.Text Then
        MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
        Exit Function
    End If
    If txtPass.Text = "" Then
        Select Case mobjCardObject.CardPreporty.密码输入限制
            Case 0 '无限制
            Case 1 '未输入提醒
                If MsgBox("未输入密码将会影响帐户的使用安全,是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
                    Exit Function
                End If
            Case 2 '为输入禁止
                MsgBox "未输入卡密码,不能进行发卡！", vbExclamation, gstrSysName
                If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
                Exit Function
        End Select
    Else
        '108779:李南春,2017/5/8,检查密码长度
        If txtPass.Visible Then
            Select Case mobjCardObject.CardPreporty.密码长度限制
            Case 0
            Case 1
                If Len(txtPass.Text) <> mobjCardObject.CardPreporty.密码长度 Then
                    MsgBox "注意:" & vbCrLf & "密码必须输入" & mobjCardObject.CardPreporty.密码长度 & "位", vbOKOnly + vbInformation
                    txtPass.Text = "": txtAudi.Text = ""
                    If txtPass.Enabled Then txtPass.SetFocus
                    Exit Function
                 End If
            Case Else
                If Len(txtPass.Text) < Abs(mobjCardObject.CardPreporty.密码长度限制) Then
                    MsgBox "注意:" & vbCrLf & "密码必须输入" & Abs(mobjCardObject.CardPreporty.密码长度限制) & "位以上.", vbOKOnly + vbInformation
                    txtPass.Text = "": txtAudi.Text = ""
                    If txtPass.Enabled Then txtPass.SetFocus
                    Exit Function
                 End If
            End Select
        End If
    End If
        
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ModifPatiPass() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改病人的密码
    '返回:修改成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-29 11:18:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim lng病人ID As Long, Curdate As Date, cllPro As Collection
   Dim strSQL As String, strPassWord As String
   
    On Error GoTo errHandle
    strPassWord = zlCommFun.zlStringEncode(txtPass.Text)     '密码加密
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    Set cllPro = New Collection
    Curdate = zlDatabase.Currentdate
      'Zl_医疗卡变动_Insert
       strSQL = "Zl_医疗卡变动_Insert("
      '      变动类型_In   Number,
      '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
      strSQL = strSQL & "" & 5 & ","
      '      病人id_In     住院费用记录.病人id%Type,
      strSQL = strSQL & "" & lng病人ID & ","
      '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
      strSQL = strSQL & "" & mlngCardTypeID & ","
      '      原卡号_In     病人医疗卡信息.卡号%Type,
      strSQL = strSQL & "'" & mstrCardNo & "',"
      '      医疗卡号_In   病人医疗卡信息.卡号%Type,
      strSQL = strSQL & "'" & mstrCardNo & "',"
      '      变动原因_In   病人医疗卡变动.变动原因%Type,
      '      --变动原因_In:如果密码调整，变动原因为密码.加密的
      strSQL = strSQL & "'" & "密码调整" & "',"
      '      密码_In       病人信息.卡验证码%Type,
      strSQL = strSQL & "'" & strPassWord & "',"
      '      操作员姓名_In 住院费用记录.操作员姓名%Type,
      strSQL = strSQL & "'" & UserInfo.姓名 & "',"
      '      变动时间_In   住院费用记录.登记时间%Type,
      strSQL = strSQL & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
      '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
      strSQL = strSQL & "NULL,"
      '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
      strSQL = strSQL & "NULL)"
     Call zlAddArray(cllPro, strSQL)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    ModifPatiPass = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
ErrSaveRollTo:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    If ModifPatiPass = False Then Exit Sub
    MsgBox "密码修改成功!", vbOKOnly + vbInformation, gstrSysName
    mblnOk = True
    mstrCardNo = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitCardInfor = False Then Unload Me: Exit Sub
    Call ClearFace
    If mstrCardNo <> "" Then
        If GetPatient(mstrCardNo) = False Then
            Call ClearFace: If txt卡号.Enabled Then txt卡号.SetFocus
            Exit Sub
        End If
        If txtPass.Enabled Then txtPass.SetFocus
    Else
        If txt卡号.Enabled Then txt卡号.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    If glngSys Like "8??" Then lbl病人.Caption = "客户"
    
    Call CreateObjectKeyboard
    Call InitTaskPancel
    '问题号:56597
    Set mobjCommEvents = New zl9CommEvents.clsCommEvents
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Set mobjCommEvents = Nothing
End Sub

Private Sub lbl卡号_Click()
    Dim strCardNo As String, strOutXml As String, strExpand As String
  
    If mlngCardTypeID = 0 Then Exit Sub
    If mobjCardObject.CardObject Is Nothing Then Exit Sub
    If Not mobjCardObject.CardPreporty.是否接触式读卡 Then Exit Sub
    
    If mobjCardObject.CardPreporty.名称 Like "IC卡*" And mobjCardObject.CardPreporty.系统 = True Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txt卡号.Text = mobjICCard.Read_Card()
            If txt卡号.Text <> "" Then
                If Not GetPatient(txt卡号.Text) Then
                    Call ClearFace
                    txt卡号.SetFocus: Exit Sub
                End If
            End If
        End If
        Exit Sub
    End If
    
    '问题号:54278
    If mobjCardObject.CardPreporty.名称 Like "*身份证*" And mobjCardObject.CardPreporty.接口程序名 = "" Then
        If mobjIDCard Is Nothing Then
            Set mobjIDCard = CreateObject("zlIDCard.clsIDCard")
        End If
        mobjIDCard.SetEnabled True
        Exit Sub
    End If
    If mobjCardObject.CardObject.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    
    txt卡号.Text = Trim(strCardNo)
    If Trim(txt卡号.Text) = "" Then Exit Sub
    If Not GetPatient(txt卡号.Text) Then
        Call ClearFace: If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
        Exit Sub
    End If
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
'问题号:56597
    If strCardType <> "" Then mlngCardTypeID = Val(strCardType)
    If strCardNo = "" Or strCardType = "" Then Exit Sub
    If Not GetPatient(strCardNo) Then
        Call ClearFace: If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
        Exit Sub
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
        '问题号:54278
        txt卡号.Text = Trim(strID)
        If Trim(txt卡号.Text) = "" Then Exit Sub
        If Not GetPatient(txt卡号.Text) Then
            Call ClearFace: If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
            Exit Sub
        End If
End Sub

Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    OpenPassKeyboard txtAudi, True
End Sub

Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    '108779：李南春,2017/5/8,密码不固定就不应该限制10位
    Call CheckInputPassWord(KeyAscii, mobjCardObject.CardPreporty.密码规则 = 1)
    If KeyAscii = 13 Then
        KeyAscii = 0: cmdOK_Click
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAudi_LostFocus()
   ClosePassKeyboard txtPass
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    OpenPassKeyboard txtPass, False
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    '108779：李南春,2017/5/8,密码不固定就不应该限制10位
    Call CheckInputPassWord(KeyAscii, mobjCardObject.CardPreporty.密码规则 = 1)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            cmdOK.SetFocus
        Else
            txtAudi.SetFocus
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub

Private Sub txtPati_GotFocus()
    zlControl.TxtSelAll txtPati
End Sub
Private Function GetPatient(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '编制:刘兴洪
    '日期:2011-07-29 11:34:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, strWhere As String
    
    Set mrsInfo = Nothing
    On Error GoTo errH
    '其他类别的号码
    If GetPatiID(mlngCardTypeID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
    If lng病人ID = 0 Then GoTo NotFoundPati:
    mstrCardNo = strInput
    
    If lng病人ID <= 0 Then GoTo NotFoundPati:
    strSQL = "" & _
    "   Select 病人ID,门诊号,住院号,就诊卡号,姓名,性别,年龄" & _
    "   From 病人信息 " & _
    "   Where 病人ID=[1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    txtPati.Text = Nvl(mrsInfo!姓名)
    txtPati.Tag = Val(mrsInfo!病人ID)
    txtSex.Text = Nvl(mrsInfo!性别)
    txtAge.Text = Nvl(mrsInfo!年龄)
    txtPass.Text = "": txtAudi.Text = ""
    txtPass.Tag = strPassWord
    If mblnCheckOldPass Then
        If zlCommFun.VerifyPassWord(Me, strPassWord, txtPati.Text, txtSex.Text, txtAge.Text, True) = False Then
            Call ClearFace
            Exit Function
        End If
    End If
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = Nothing
    Exit Function
NotFoundPati:
    If strErrMsg = "" Then
        MsgBox "不能读取" & IIf(glngSys Like "8??", "客户", "病人") & "信息，请确定是否正确刷卡！", vbInformation, gstrSysName
    End If
    Set mrsInfo = Nothing
End Function
Private Sub ClearFace()
    txt卡号.PasswordChar = IIf(mobjCardObject.CardPreporty.卡号密文规则 <> "", "*", "")
    txt卡号.Text = ""
    txtPass.Text = "": txtPati.Text = ""
    txtSex.Text = "": txtAge.Text = ""
    txtPass.Text = "": txtAudi.Text = ""
End Sub

Private Sub txt卡号_Change()
    txtPass.Enabled = Trim(txt卡号.Text) <> ""
    txtPass.BackColor = IIf(txtPass.Enabled = False, txtPati.BackColor, txt卡号.BackColor)
    txtAudi.Enabled = Trim(txt卡号.Text) <> ""
    txtAudi.BackColor = IIf(txtAudi.Enabled = False, txtPati.BackColor, txt卡号.BackColor)
End Sub

Private Sub txt卡号_GotFocus()
    Dim strExpend As String
    
    On Error GoTo Errhand
    zlControl.TxtSelAll txt卡号
    txt卡号.PasswordChar = IIf(mobjCardObject.CardPreporty.卡号密文规则 <> "", "*", "")
    '问题号:56597
    '初始化IC卡
    If mobjCardObject.CardPreporty.名称 Like "IC卡*" And mobjCardObject.CardPreporty.系统 = True Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        Exit Sub
    End If
    '初始化二代身份证
    If mobjCardObject.CardPreporty.名称 Like "*身份证*" And mobjCardObject.CardPreporty.接口程序名 = "" Then
        If mobjIDCard Is Nothing Then
            Set mobjIDCard = CreateObject("zlIDCard.clsIDCard")
        End If
        mobjIDCard.SetEnabled True
        Exit Sub
    End If
    If mobjSquare Is Nothing Then Set mobjSquare = CreateObject("zl9CardSquare.clsCardSquare")
    '初始化射频卡对象
    '86152:李南春,2015/7/6,初始化对象
    If Err <> 0 Then Exit Sub
    mobjSquare.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
    mobjSquare.zlInitEvents Me.hWnd, mobjCommEvents
    mobjSquare.SetEnabled True
    
    '85565:李南春,2015/7/21,调用刷卡接口
    Err = 0: On Error Resume Next
    If mobjCardObject.CardPreporty.接口序号 = 0 Or mobjCardObject.CardPreporty.接口程序名 = "" Then Exit Sub
    If Not (mobjCardObject.CardPreporty.是否刷卡 Or mobjCardObject.CardPreporty.是否扫描) Then Exit Sub
    
    Call mobjSquare.zlSetBrushCardObject(mobjCardObject.CardPreporty.接口序号, txt卡号, strExpend, _
                                        mobjCardObject.CardPreporty.消费卡)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub txt卡号_KeyPress(KeyAscii As Integer)
     '问题号:58066
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
     
     If (Len(txt卡号.Text) = mobjCardObject.CardPreporty.卡号长度 - 1 And KeyAscii <> 8) Or (KeyAscii = 13 And Trim(txt卡号.Text) <> "") Then
            If KeyAscii <> 13 Then
                txt卡号.Text = txt卡号.Text & Chr(KeyAscii)
                txt卡号.SelStart = Len(txt卡号.Text)
            End If
            KeyAscii = 0
            If Not GetPatient(txt卡号.Text) Then
                Call ClearFace
                txt卡号.SetFocus: Exit Sub
            End If
            txtPass.SetFocus
        End If
End Sub

Private Sub txt卡号_LostFocus()
    '问题号:56597
   If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled False
   If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
   If Not mobjSquare Is Nothing Then mobjSquare.SetEnabled False
   
End Sub
