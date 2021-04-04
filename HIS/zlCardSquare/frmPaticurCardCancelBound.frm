VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPaticurCardCancelBound 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "取消卡号绑定"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8700
   Icon            =   "frmPaticurCardCancelBound.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7155
      TabIndex        =   15
      Top             =   405
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7155
      TabIndex        =   11
      Top             =   1005
      Width           =   1395
   End
   Begin VB.PictureBox picPass 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   195
      ScaleHeight     =   3135
      ScaleWidth      =   6420
      TabIndex        =   0
      Top             =   495
      Width           =   6420
      Begin VB.CommandButton cmdALL 
         Caption         =   "取消所有绑定的医疗卡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1080
         TabIndex        =   14
         Top             =   3420
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.CommandButton cmdAllType 
         Caption         =   "取消所有绑定的[就诊卡]"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1080
         TabIndex        =   13
         Top             =   3255
         Visible         =   0   'False
         Width           =   4815
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1815
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
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1815
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
         Left            =   1110
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1215
         Width           =   4845
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   0
         TabIndex        =   2
         Top             =   900
         Width           =   6555
      End
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
         Left            =   1110
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2475
         Width           =   4845
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   135
         Picture         =   "frmPaticurCardCancelBound.frx":0ECA
         Top             =   90
         Width           =   720
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
         TabIndex        =   10
         Top             =   1920
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
         Left            =   375
         TabIndex        =   9
         Top             =   1920
         Width           =   630
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
         Left            =   375
         TabIndex        =   8
         Top             =   1275
         Width           =   630
      End
      Begin VB.Label lblNotes 
         BackStyle       =   0  'Transparent
         Caption         =   "取消绑定操作,如果需要取消指定卡的绑定,请在取消卡号项中刷卡或输入指定的卡号进行取消绑定。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   975
         TabIndex        =   7
         Top             =   270
         Width           =   5340
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
         Left            =   345
         TabIndex        =   6
         Top             =   2520
         Width           =   660
      End
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   3540
      Left            =   75
      TabIndex        =   12
      Top             =   345
      Width           =   6945
      _Version        =   589884
      _ExtentX        =   12250
      _ExtentY        =   6244
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmPaticurCardCancelBound"
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
Private mblnOK As Boolean
Private mrsInfo As ADODB.Recordset
Private mobjCardObject As clsCardObject
Private mblnFirst As Boolean
Private mblnCheckOldPass As Boolean
Public mstrPrepayPrivs As String  '预交款相关权限
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mobjSquare As Object

Public Function zlCancelBand(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional lng病人ID As Long, Optional strCardNo As String, _
    Optional blnCheckOldPass As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:入口(取消绑定操作)
    '入参:frmMain-调用的主窗体
    '       lngModule -模块号
    '       lngCardTypeId-卡类别ID
    '       lng病人ID-病人ID
    '       strCardNo-卡号
    '返回:取消成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2011-07-29 11:08:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngCardTypeID = lngCardTypeID: mlngModule = lngModule: mlng病人ID = lng病人ID
    mstrCardNo = strCardNo: mblnOK = False
    mblnCheckOldPass = blnCheckOldPass
    On Error Resume Next
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlCancelBand = mblnOK
End Function
Private Sub InitTaskPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化InitTaskPancel
    '编制:刘兴洪
    '日期:2011-06-30 18:20:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    wndTaskPanel.Top = wndTaskPanel.Top + 50
    wndTaskPanel.Height = wndTaskPanel.Height - 150
    wndTaskPanel.HotTrackStyle = xtpTaskPanelHighlightItem
    
    Call wndTaskPanel.SetGroupInnerMargins(2, 0, 2, 0)
    Call wndTaskPanel.SetGroupOuterMargins(2, -10, 2, -10)
    Call wndTaskPanel.SetMargins(2, 16, 2, 10, 30)
    
    Set tkpGroup = wndTaskPanel.Groups.Add(1, "请刷卡")
    Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
   Set Item.Control = picPass
    tkpGroup.CaptionVisible = False
   Call Item.SetMargins(0, -19, 0, -4)
    picPass.BackColor = Item.BackColor
    Me.BackColor = Item.BackColor
    cmdAllType.BackColor = Item.BackColor
    cmdCancel.BackColor = Item.BackColor
    cmdALL.BackColor = Item.BackColor
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
    Dim strExpend As String
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
    cmdAllType.Caption = Replace(cmdAllType.Caption, "[就诊卡]", "[" & mobjCardObject.CardPreporty.名称 & "]")
    
    InitCardInfor = True
    '85565:李南春,2015/7/21,调用刷卡接口
    If mobjSquare Is Nothing Then Set mobjSquare = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then Exit Function
    mobjSquare.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
    Err = 0: On Error Resume Next
    If mobjCardObject.CardPreporty.接口序号 = 0 Or mobjCardObject.CardPreporty.接口程序名 = "" Then Exit Function
    If Not (mobjCardObject.CardPreporty.是否刷卡 Or mobjCardObject.CardPreporty.是否扫描) Then Exit Function
    If mobjSquare.zlSetBrushCardObject(mobjCardObject.CardPreporty.接口序号, txt卡号, strExpend, _
                mobjCardObject.CardPreporty.消费卡) Then
        Call mobjSquare.zlInitEvents(Me.hWnd, mobjCommEvents)
    End If
End Function
   
Private Sub cmdALL_Click()
    Call SaveData(2)
End Sub

Private Sub cmdAllType_Click()
    Call SaveData(1)
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub
Private Function CheckBindCard(ByVal lng病人ID As Long, ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否绑定卡
    '编制:刘兴洪
    '日期:2011-07-31 05:37:48
    '检查标准:
    '  1.在住院费用记录中无记录,就表示绑定卡
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim str缺省卡类别ID As String
    str缺省卡类别ID = IIf(mobjCardObject.CardPreporty.名称 = "就诊卡" And mobjCardObject.CardPreporty.系统, mobjCardObject.接口序号, "")
    strSQL = "" & _
    "   Select  1 " & _
    "   From 住院费用记录  " & _
    "   Where 病人id = [1] And 记录性质 = 5 and nvl(结论,[3])=[2] and 实际票号=[4] And RowNum=1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, str缺省卡类别ID, Trim(CStr(mlngCardTypeID)), strCardNo)
    CheckBindCard = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function isValied(ByVal intType As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的数据是否有效
    '入参:intType:0-当前卡号;1-当前类别;2-当前病人所有
    '返回:数据有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-29 11:15:42
    '---------------------------------------------------------------------------------------------------------------------------------------------\
    Dim str称呼 As String
    str称呼 = IIf(glngSys Like "8??", "客户", "病人")
    
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then
        MsgBox "不能读取" & str称呼 & "信息，请确定是否正确刷卡！", vbInformation, gstrSysName
         txt卡号.SetFocus: Exit Function
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "不能读取" & str称呼 & "信息，请确定是否正确刷卡！", vbInformation, gstrSysName
         txt卡号.SetFocus: Exit Function
    End If
    If intType = 0 Then
       If Trim(txt卡号.Text) = "" Then
            MsgBox "未输入需要取消的卡号,不能取消绑定操作!", vbOKOnly + vbInformation, gstrSysName
            txt卡号.SetFocus: Exit Function
       End If
       If mlng病人ID <> mrsInfo!病人ID Then
            MsgBox "当前卡号的持有人,不是你选择的病人,不能取消绑定操作!", vbOKOnly + vbInformation, gstrSysName
            txt卡号.SetFocus: Exit Function
       End If
       '检查当前卡号是否绑定操作
       If CheckBindCard(mlng病人ID, Trim(txt卡号)) = False Then
            MsgBox "当前卡号不是绑定的卡号,请使用退卡功能!", vbOKOnly + vbInformation, gstrSysName
            txt卡号.SetFocus: Exit Function
       End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function BlandCancel(ByVal intType As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消绑定卡
    '入参:intType:0-当前卡号;1-当前类别;2-当前病人所有
    '返回:取消成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-29 11:18:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim lng病人ID As Long, Curdate As Date, cllPro As Collection
   Dim strSQL As String, strPassWord As String
   
    On Error GoTo errHandle
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    If intType = 0 Then
        '如果当前卡号取消绑定,预交提醒
        If Not IsCheckCancel退预交(lng病人ID) Then
            Exit Function
        End If
    End If
    Set cllPro = New Collection
    Curdate = zlDatabase.Currentdate
    '105590:李南春,2017/3/10，取消绑定时填写操作员姓名
      'Zl_医疗卡变动_Insert
       strSQL = "Zl_医疗卡变动_Insert("
      '      变动类型_In   Number,
      '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
      strSQL = strSQL & "" & 14 & ","
      '      病人id_In     住院费用记录.病人id%Type,
      strSQL = strSQL & "" & lng病人ID & ","
      '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
      strSQL = strSQL & "" & IIf(intType = 2, "NULL", mlngCardTypeID) & ","
      '      原卡号_In     病人医疗卡信息.卡号%Type,
      strSQL = strSQL & "NULL,"
      '      医疗卡号_In   病人医疗卡信息.卡号%Type,
      strSQL = strSQL & IIf(intType <> 0, "NULL", "'" & txt卡号.Text & "'") & ","
      '      变动原因_In   病人医疗卡变动.变动原因%Type,
      strSQL = strSQL & "'取消卡号绑定',"
      '      密码_In       病人信息.卡验证码%Type,
      strSQL = strSQL & "NULL,"
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
    BlandCancel = True
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

Private Function IsCheckCancel退预交(ByVal lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消卡绑定时时检查病人是否有预交款未退
     '返回:有效,返回true,否则返回False
    '编制:王吉
    '日期:2012-07-16 18:50:36
    '问题号:51537
    '问题号:50891
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim msgBoxResult As String
    Dim strSQL As String
    Dim rsBill As Recordset, rsCard As Recordset
    '69483,刘尔旋,2014-01-15,病人医疗卡退卡退款处理
    strSQL = "Select Count(1) As 医疗卡数 From 病人医疗卡信息 Where 状态=0 And 病人ID=[1]"
    Set rsCard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    strSQL = _
            "Select 预交余额,费用余额 From 病人余额 Where 性质=1 And 类型=1 And 病人ID=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    '问题:48249
    If InStr(1, mstrPrepayPrivs, ";预交退款;") > 0 Then
        If rsBill.RecordCount > 0 Then
            If Format(Nvl(rsBill!预交余额, 0) - Nvl(rsBill!费用余额, 0), "0.00") > 0 Then
                
                '问题号:51537
                '问题号:50891
                msgBoxResult = zlCommFun.ShowMsgbox(gstrSysName, "该病人尚有预交余额未退!" & " " & "是否先进行余额退款操作?", "退预交后继续,继续,取消", Me, vbQuestion)
                If msgBoxResult = "退预交后继续" Then '退预交余额操作
                    '病人余额退款
                     IsCheckCancel退预交 = zlPrepayFunc(2, lng病人ID)
                     Exit Function
                ElseIf msgBoxResult = "继续" Then
                    If rsCard!医疗卡数 = 1 Then
                        MsgBox "该病人尚有预交余额，不能对病人唯一的医疗卡进行取消绑定操作!", vbInformation, gstrSysName
                        IsCheckCancel退预交 = False
                        Exit Function
                    End If
                    IsCheckCancel退预交 = True
                ElseIf msgBoxResult = "取消" Or msgBoxResult = "" Then
                     IsCheckCancel退预交 = False
                     Exit Function
                End If
            End If
'        Else
'        '问题号:51537
'        '问题号:50891
'           If ZL9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "是否继续进行取消卡绑定操作?", "退卡,取消", Me, vbQuestion) = "取消" Then
'                IsCheckCancel退预交 = False
'                Exit Function
'           End If
        End If
    Else
        If rsBill.RecordCount > 0 Then
            If Format(Nvl(rsBill!预交余额, 0) - Nvl(rsBill!费用余额, 0), "0.00") > 0 Then
                If rsCard!医疗卡数 = 1 Then
                    MsgBox "您没有预交退款权限，不能对病人唯一的医疗卡进行取消绑定操作!", vbInformation, gstrSysName
                    IsCheckCancel退预交 = False
                    Exit Function
                End If
            End If
        End If
        If MsgBox("您没有预交退款权限,是否继续进行取消卡绑定操作?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then IsCheckCancel退预交 = False: Exit Function
    End If
        IsCheckCancel退预交 = True
End Function
Private Sub SaveData(ByVal intType As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消绑定卡
    '入参:intType:0-当前卡号;1-当前类别;2-当前病人所有
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If isValied(intType) = False Then Exit Sub
    If BlandCancel(intType) = False Then Exit Sub
    MsgBox "取消成功!", vbOKOnly + vbInformation, gstrSysName
    mblnOK = True
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Call SaveData(0)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call ClearFace
    If GetPatient("-" & mlng病人ID) = False Then Unload Me: Exit Sub
    If mstrCardNo <> "" Then
        If GetPatient(mstrCardNo) = False Then
              If txt卡号.Enabled Then txt卡号.SetFocus
            Exit Sub
        End If
    End If
    If txt卡号.Enabled Then txt卡号.SetFocus
End Sub

Private Sub Form_Load()
    mblnFirst = True
    If glngSys Like "8??" Then lbl病人.Caption = "客户"
    Set mobjCommEvents = New zl9CommEvents.clsCommEvents
    If InitCardInfor = False Then
        '74539,冉俊明,2014-6-27,在收费处发院外卡后，到医疗卡发放管理，取消绑定时，取消绑定窗口一闪而过，无法取消
        MsgBox "该卡设备未启用，您不能进行取消绑定操作，请到【参数设置>设备配置】中启用！", vbInformation, gstrSysName
        mblnFirst = False: Unload Me: Exit Sub
    End If
    Call InitTaskPancel
End Sub
 Private Function GetPatient(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '编制:刘兴洪
    '日期:2011-07-29 11:34:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, strWhere As String, blnReadPatiInfor As Boolean
    blnReadPatiInfor = Left(strInput, 1) = "-"
    
    Set mrsInfo = Nothing
    On Error GoTo errH
    If Not blnReadPatiInfor Then
        '其他类别的号码
        If GetPatiID(mlngCardTypeID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID = 0 Then GoTo NotFoundPati:
        mstrCardNo = strInput
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        txt卡号.Text = strInput
    Else
        lng病人ID = Val(Mid(strInput, 2))
    End If
    
    strSQL = "" & _
    "   Select 病人ID,门诊号,住院号,就诊卡号,姓名,性别,年龄" & _
    "   From 病人信息 " & _
    "   Where 病人ID=[1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If Not blnReadPatiInfor Then
        GetPatient = True: Exit Function
    End If
    If mrsInfo.EOF Then Exit Function

    txtPati.Text = Nvl(mrsInfo!姓名)
    txtPati.Tag = Val(mrsInfo!病人ID)
    txtSex.Text = Nvl(mrsInfo!性别)
    txtAge.Text = Nvl(mrsInfo!年龄)
    Set mrsInfo = Nothing
'    If mblnCheckOldPass Then
'        If zlCommFun.VerifyPassWord(Me, strPassWord, txtPati.Text, txtSex.Text, txtAge.Text, True) = False Then
'            Call ClearFace
'            Exit Function
'        End If
'    End If
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    txtPati.Text = ""
    txtSex.Text = "": txtAge.Text = ""
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjSquare = Nothing
    Set mobjCommEvents = Nothing
End Sub

Private Sub lbl卡号_Click()
    Dim strExpand As String, strCardNo As String, strOutXml As String
    If Not mobjCardObject.CardPreporty.是否接触式读卡 Then Exit Sub
'    If mobjICCard Is Nothing Then
'        Set mobjICCard = CreateObject("zlICCard.clsICCard")
'        Set mobjICCard.gcnOracle = gcnOracle
'    End If
    
'    If Not mobjICCard Is Nothing Then
'        txt卡号.Text = mobjICCard.Read_Card()
'        If txt卡号.Text <> "" Then
'            mblnICCard = True
'            Call CheckFreeCard(txt卡号.Text)
'        End If
'    End If
  
    If mobjCardObject.CardObject Is Nothing Then Exit Sub
    If mobjCardObject.CardObject.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    txt卡号.Text = Trim(strCardNo)
    If txt卡号.Text <> "" Then
        If Not GetPatient(txt卡号.Text) Then
            Call txt卡号_GotFocus
            If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
            Exit Sub
        End If
        cmdOK.SetFocus
    End If
End Sub

Private Sub txt卡号_GotFocus()
    zlControl.TxtSelAll txt卡号
    txt卡号.PasswordChar = IIf(mobjCardObject.CardPreporty.卡号密文规则 <> "", "*", "")
End Sub
Private Sub txt卡号_KeyPress(KeyAscii As Integer)
     If (Len(txt卡号.Text) = mobjCardObject.CardPreporty.卡号长度 - 1 And KeyAscii <> 8) Or (KeyAscii = 13 And Trim(txt卡号.Text) <> "") Then
            If KeyAscii <> 13 Then
                txt卡号.Text = txt卡号.Text & Chr(KeyAscii)
                txt卡号.SelStart = Len(txt卡号.Text)
            End If
            KeyAscii = 0
            If Not GetPatient(txt卡号.Text) Then
                Call txt卡号_GotFocus
                If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
                Exit Sub
            End If
            cmdOK.SetFocus
        End If
End Sub


Private Function zlPrepayFunc(ByVal intFunc As Integer, ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:缴预存款
    '入参:intFunc-1-缴预存;2-退预款;3-作废,4-门诊转住院;5-住院转门诊;
    '编制:刘兴洪
    '日期:2011-07-24 18:25:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun As Object, int预交类型 As Integer
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Function
    'byt预交类型: 0-收预交款(缺省,可切换到退),1-浏览单据(1),2-作废状态(1); 3-余额退款(37770), 4-门诊转住院;5-住院转门诊
    Select Case intFunc
    Case 1  '1.缴预存
        int预交类型 = 0
    Case 2 '退款
        int预交类型 = 3
    Case 3: int预交类型 = 2
    Case 4: int预交类型 = 4
    Case 5: int预交类型 = 5
    End Select
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能： 调用预交款收款窗口
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '   strDBUser:当前数据库登录用户名
    '  bytCallObject:刘兴洪加入(0-预交款调用(缺省的);1-病人费用查询调用,2-医疗卡调用)
    '  lng病人ID-缺省的病人ID
    '  lng主页ID-缺省的主页ID
    '  dblDefPrePayMoney-缺省的预付金额
    Set gfrmCardMgr = Me
    '问题:48249
    If objFun.PlusDeposit(glngSys, gcnOracle, Me, gstrDBUser, 2, lng病人ID, 0, 0, int预交类型) = False Then
        zlPrepayFunc = False
        Set gfrmCardMgr = Nothing
        Exit Function
    End If
    Set gfrmCardMgr = Nothing
    zlPrepayFunc = True
End Function

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt卡号.Text = Trim(strCardNo)
    If txt卡号.Text <> "" Then
        If Not GetPatient(txt卡号.Text) Then
            Call txt卡号_GotFocus
            If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
            Exit Sub
        End If
        cmdOK.SetFocus
    End If
End Sub

Private Sub txt卡号_LostFocus()
    If Not mobjSquare Is Nothing Then mobjSquare.SetEnabled False
End Sub
