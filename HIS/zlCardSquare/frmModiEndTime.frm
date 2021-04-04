VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmModiEndTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "使用时间调整"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8730
   Icon            =   "frmModiEndTime.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picEndDate 
      BorderStyle     =   0  'None
      Height          =   3705
      Left            =   210
      ScaleHeight     =   3705
      ScaleWidth      =   6525
      TabIndex        =   10
      Top             =   450
      Width           =   6525
      Begin VB.CommandButton cmdDefualtSet 
         Caption         =   "增加XX天"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3690
         TabIndex        =   17
         Top             =   3240
         Visible         =   0   'False
         Width           =   2265
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
         Left            =   1275
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   975
         Width           =   4665
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   -90
         TabIndex        =   11
         Top             =   810
         Width           =   7245
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
         Left            =   1275
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1545
         Width           =   4665
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
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1425
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
         Left            =   4305
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   435
         Left            =   4260
         TabIndex        =   6
         Top             =   2730
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   248446978
         UpDown          =   -1  'True
         CurrentDate     =   .999988425925926
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   435
         Left            =   2235
         TabIndex        =   5
         Top             =   2730
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483646
         CustomFormat    =   "yyyy-MM-dd hh:mm"
         Format          =   248446979
         CurrentDate     =   401769
      End
      Begin VB.CheckBox chkEndDate 
         Caption         =   "终止时间"
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
         Left            =   555
         TabIndex        =   4
         Top             =   2790
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   150
         Picture         =   "frmModiEndTime.frx":6852
         Top             =   30
         Width           =   720
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
         Left            =   510
         TabIndex        =   16
         Top             =   1020
         Width           =   660
      End
      Begin VB.Label lblNotes 
         BackStyle       =   0  'Transparent
         Caption         =   "请将[XX]从刷卡器上轻轻划过，  然后选择需要更改的日期！"
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
         Left            =   510
         TabIndex        =   14
         Top             =   1605
         Width           =   630
      End
      Begin VB.Label label3 
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
         Index           =   0
         Left            =   510
         TabIndex        =   13
         Top             =   2190
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
         Left            =   3555
         TabIndex        =   12
         Top             =   2190
         Width           =   690
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
      Height          =   360
      Left            =   7380
      TabIndex        =   8
      Top             =   900
      Width           =   1230
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
      Height          =   360
      Left            =   7380
      TabIndex        =   7
      Top             =   360
      Width           =   1230
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   4755
      Left            =   120
      TabIndex        =   9
      Top             =   90
      Width           =   7095
      _Version        =   589884
      _ExtentX        =   12515
      _ExtentY        =   8387
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmModiEndTime"
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
Private mobjCard As Card

Private mblnFirst As Boolean
Private mblnCheckOldPass As Boolean
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private WithEvents mobjIDCard As clsIDCard '问题号:54278
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents '问题号:56597
Attribute mobjCommEvents.VB_VarHelpID = -1
 
Private mstrPrivs As String
Public Function zlModifyEndTime(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional lng病人ID As Long, Optional strCardNo As String, _
    Optional strPrivs As String) As Boolean
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
    mstrPrivs = strPrivs
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlModifyEndTime = mblnOk
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
    Set tkpGroup = wndTaskPanel.Groups.Add(1, "请刷卡选择修改日期")
    Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
    Set Item.Control = picEndDate
    tkpGroup.CaptionVisible = False
   ' Call Item.SetMargins(0, -19, 0, -4)
   
    picEndDate.BackColor = Item.BackColor
    Me.BackColor = Item.BackColor
    cmdOK.BackColor = Item.BackColor
    chkEndDate.BackColor = Item.BackColor
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
    Dim rsTemp As ADODB.Recordset
    On Error Resume Next
    
    If gobjOneCardComLib.zlGetCard(mlngCardTypeID, False, mobjCard) = False Then Exit Function
    If mobjCard Is Nothing Then Exit Function
    
    If mobjCard.名称 = "就诊卡" And mobjCard.系统 Then
             lbl卡号.BorderStyle = 1: lbl卡号.Tag = "1"
    Else
         If mobjCard.是否接触式读卡 Then
             lbl卡号.BorderStyle = 1: lbl卡号.Tag = "1"
         Else
             lbl卡号.BorderStyle = 0: lbl卡号.Tag = "0"
         End If
     End If
     
    lblNotes.Caption = Replace(lblNotes.Caption, "[XX]", "[" & mobjCard.名称 & "]")
    If InStr(mobjCard.缺省有效时间, "天") Then
        cmdDefualtSet.Caption = "增加" & Val(mobjCard.缺省有效时间) & "天(&A)"
        cmdDefualtSet.Tag = Val(nvl(rsTemp!缺省有效时间)) & "天"
    ElseIf InStr(mobjCard.缺省有效时间, "月") Then
        cmdDefualtSet.Caption = "增加" & Val(mobjCard.缺省有效时间) & "月(&A)"
        cmdDefualtSet.Tag = Val(mobjCard.缺省有效时间) & "月"
    End If
    If mobjCard.缺省有效时间 <> "" And Val(mobjCard.缺省有效时间) > 0 Then cmdDefualtSet.Visible = True
       
     InitCardInfor = True
End Function

Private Sub chkEndDate_Click()
    dtpDate.Enabled = chkEndDate.value
    dtpTime.Enabled = chkEndDate.value
End Sub

Private Sub chkEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
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
    Dim str称呼 As String, Curdate As Date, CardEndDate As Date
    str称呼 = IIf(glngSys Like "8??", "客户", "病人")

    On Error GoTo errHandle
    
    If Not zlstr.IsHavePrivs(mstrPrivs, "使用时间调整") Then
        MsgBox "您没有权限更改使用时间，如需更改，请与系统管理员联系！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mrsInfo Is Nothing Then
        MsgBox "不能读取" & str称呼 & "信息，请确定是否正确刷卡！", vbInformation, gstrSysName
        Call ClearFace: txt卡号.SetFocus: Exit Function
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "不能读取" & str称呼 & "信息，请确定是否正确刷卡！", vbInformation, gstrSysName
        Call ClearFace: txt卡号.SetFocus: Exit Function
    End If
    Curdate = zlDatabase.Currentdate
    CardEndDate = Format(CStr(dtpDate.value) & " " & CStr(dtpTime.value), "YYYY-MM-DD HH:MM:SS")
    If CardEndDate < Curdate And chkEndDate.value = vbChecked Then
        MsgBox "请选择大于当前时间的终止日期进行更改！", vbInformation, gstrSysName
        Exit Function
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ModifCardEndTime() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改卡有效终止时间
    '返回:修改成功,返回true,否则返回False
    '编制:
    '日期:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, Curdate As Date, cllPro As Collection
    Dim strSQL As String, strPassWord As String, strEndDate As String

    On Error GoTo errHandle
    lng病人ID = Val(nvl(mrsInfo!病人ID))
    Set cllPro = New Collection
    Curdate = zlDatabase.Currentdate
    strEndDate = CStr(dtpDate.value) & " " & CStr(dtpTime.value)
    
    'Zl_医疗卡变动_Insert_S
     strSQL = "Zl_医疗卡变动_Insert_S("
    '  变动类型_In     Number,
    strSQL = strSQL & "7,"
    '  病人id_In       病人医疗卡信息.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  卡类别id_In     病人医疗卡信息.卡类别id%Type,
    strSQL = strSQL & "" & mlngCardTypeID & ","
    '  原卡号_In       病人医疗卡信息.卡号%Type,
    strSQL = strSQL & "'" & mstrCardNo & "',"
    '  医疗卡号_In     病人医疗卡信息.卡号%Type,
    strSQL = strSQL & "'" & mstrCardNo & "',"
    '  变动原因_In     病人医疗卡变动.变动原因%Type,
    strSQL = strSQL & "'" & "终止时间调整" & "',"
    '  密码_In         病人医疗卡信息.密码%Type,
    strSQL = strSQL & "'" & strPassWord & "',"
    '  操作员姓名_In   病人医疗卡变动.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  变动时间_In     病人医疗卡变动.登记时间%Type,
    strSQL = strSQL & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  挂失方式_In     病人医疗卡变动.挂失方式%Type := Null,
    strSQL = strSQL & "NULL,"
    '  终止使用时间_In 病人医疗卡信息.终止使用时间%Type := Null,
    strSQL = strSQL & IIf(chkEndDate.value = vbUnchecked, "Null", "to_date('" & Format(strEndDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')") & ")"
    '  卡费_In         病人医疗卡变动.卡费%Type := Null,
    '  病历费_In       病人医疗卡变动.病历费%Type := Null,
    '  费用单号_In     病人医疗卡变动.费用单号%Type := Null,
    '  预交单号_In     病人结算异常记录.预交单号%Type := Null,
    '  变动id_In       病人医疗卡变动.Id%Type := Null,
    '  异常标志_In     Number := 0,
    '  异常id_In       病人结算异常记录.Id%Type := Null,
    '  预交金额_In     病人结算异常记录.预交金额%Type := Null
    Call zlAddArray(cllPro, strSQL)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    ModifCardEndTime = True
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

Private Sub cmdDefualtSet_Click()
    If Format(dtpDate, "yyyy-MM-dd") < Format("3000-01-01", "yyyy-MM-dd") Then
        If InStr(cmdDefualtSet.Tag, "天") Then
            dtpDate = DateAdd("D", Val(cmdDefualtSet.Tag), dtpDate)
        ElseIf InStr(cmdDefualtSet.Tag, "月") Then
            dtpDate = DateAdd("M", Val(cmdDefualtSet.Tag), dtpDate)
        End If
        cmdOK.SetFocus
    End If
End Sub

Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    If ModifCardEndTime = False Then Exit Sub
    MsgBox "终止使用时间修改成功!", vbOKOnly + vbInformation, gstrSysName
    mblnOk = True
    mstrCardNo = ""
    Unload Me
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim Curdate As Date
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If InitCardInfor = False Then Unload Me: Exit Sub
    Call ClearFace
    If mstrCardNo <> "" Then
        If GetPatient(mstrCardNo) = False Then
            Call ClearFace: If txt卡号.Enabled Then txt卡号.SetFocus
            Exit Sub
        End If
    Else
        If txt卡号.Enabled Then txt卡号.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    If glngSys Like "8??" Then lbl病人.Caption = "客户"
    
    Call InitTaskPancel
    '问题号:56597
    Set mobjCommEvents = New zl9CommEvents.clsCommEvents
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Set mobjCommEvents = Nothing
    Set mrsInfo = Nothing
End Sub

Private Sub lbl卡号_Click()
    Dim strCardNo As String, strOutXml As String, strExpand As String
  
    If mlngCardTypeID = 0 Then Exit Sub
    If mobjCard.CardObject Is Nothing Then Exit Sub
    If Not mobjCard.是否接触式读卡 Then Exit Sub
    
    If mobjCard.名称 Like "IC卡*" And mobjCard.系统 = True Then
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
    If mobjCard.名称 Like "*身份证*" And mobjCard.接口程序名 = "" Then
        If mobjIDCard Is Nothing Then
            Set mobjIDCard = CreateObject("zlIDCard.clsIDCard")
        End If
        mobjIDCard.SetEnabled True
        Exit Sub
    End If
    If gobjOneCardComLib.zlReadCard(Me, mlngModule, mobjCard.接口序号, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    
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
    If GetPatiID(mlngCardTypeID, strInput, False, lng病人ID, strPassWord, strErrMsg, , , , , , , , , , True) = False Then GoTo NotFoundPati:
    If lng病人ID = 0 Then GoTo NotFoundPati:
    mstrCardNo = strInput
    
    If lng病人ID <= 0 Then GoTo NotFoundPati:
    strSQL = "" & _
        "Select a.病人id, a.门诊号, a.住院号, a.就诊卡号, a.姓名, a.性别, a.年龄, b.终止使用时间" & vbNewLine & _
        "From 病人信息 a, 病人医疗卡信息 b" & vbNewLine & _
        "Where a.病人id = b.病人id And b.卡号 = [1] And b.卡类别id = [2] And a.病人id = [3]"

    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, mlngCardTypeID, lng病人ID)
    If mrsInfo.EOF Then Exit Function
    txtPati.Text = nvl(mrsInfo!姓名)
    txtPati.Tag = Val(mrsInfo!病人ID)
    txtSex.Text = nvl(mrsInfo!性别)
    txtAge.Text = nvl(mrsInfo!年龄)
    If nvl(mrsInfo!终止使用时间) <> "" Then
        dtpDate = Format(nvl(mrsInfo!终止使用时间), "yyyy-MM-dd")
        dtpTime = Format(nvl(mrsInfo!终止使用时间), "HH:mm:ss")
        chkEndDate.value = 1
    Else
        dtpDate = Format("3000-01-01", "yyyy-MM-dd")
        dtpTime = Format("23:59:59", "HH:mm:ss")
        chkEndDate.value = 0
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除界面
    '编制:刘兴洪
    '日期:2018-11-23 14:21:13
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    txt卡号.PasswordChar = IIf(mobjCard.卡号密文规则 <> "", "*", "")
    txt卡号.Text = ""
    txtSex.Text = "": txtAge.Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt卡号_Change()
    Dim strExpend As String
    On Error GoTo Errhand
   
    txt卡号.PasswordChar = IIf(mobjCard.卡号密文规则 <> "", "*", "")
    '问题号:56597
    '初始化IC卡
    If mobjCard.名称 Like "IC卡*" And mobjCard.系统 = True Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        mobjICCard.SetEnabled Trim(txt卡号.Text) = ""
        Exit Sub
    End If
    '初始化二代身份证
    If mobjCard.名称 Like "*身份证*" And mobjCard.接口程序名 = "" Then
        If mobjIDCard Is Nothing Then
            Set mobjIDCard = CreateObject("zlIDCard.clsIDCard")
        End If
        mobjIDCard.SetEnabled Trim(txt卡号.Text) = ""
        Exit Sub
    End If
    
    gobjOneCardComLib.SetEnabled Trim(txt卡号.Text) = ""
    
    If mobjCard.接口序号 = 0 Or mobjCard.接口程序名 = "" Then Exit Sub
    If Not (mobjCard.是否刷卡 Or mobjCard.是否扫描) Then Exit Sub
    
    Call gobjOneCardComLib.zlSetBrushCardObject(mobjCard.接口序号, txt卡号, strExpend, mobjCard.消费卡)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txt卡号_GotFocus()
    Dim strExpend As String
    
    On Error GoTo Errhand
    zlControl.TxtSelAll txt卡号
    txt卡号.PasswordChar = IIf(mobjCard.卡号密文规则 <> "", "*", "")
    '问题号:56597
    '初始化IC卡
    If mobjCard.名称 Like "IC卡*" And mobjCard.系统 = True Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        mobjIDCard.SetEnabled Trim(txt卡号.Text) = ""
        Exit Sub
    End If
    '初始化二代身份证
    If mobjCard.名称 Like "*身份证*" And mobjCard.接口程序名 = "" Then
        If mobjIDCard Is Nothing Then
            Set mobjIDCard = CreateObject("zlIDCard.clsIDCard")
        End If
        mobjIDCard.SetEnabled Trim(txt卡号.Text) = ""
        Exit Sub
    End If
    
    gobjOneCardComLib.SetEnabled Trim(txt卡号.Text) = ""
    If mobjCard.接口序号 = 0 Or mobjCard.接口程序名 = "" Then Exit Sub
    If Not (mobjCard.是否刷卡 Or mobjCard.是否扫描) Then Exit Sub
    
    Call gobjOneCardComLib.zlSetBrushCardObject(mobjCard.接口序号, txt卡号, strExpend, mobjCard.消费卡)
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
     
    If (Len(txt卡号.Text) = mobjCard.卡号长度 - 1 And KeyAscii <> 8) Or (KeyAscii = 13 And Trim(txt卡号.Text) <> "") Then
        If KeyAscii <> 13 Then
            txt卡号.Text = txt卡号.Text & Chr(KeyAscii)
            txt卡号.SelStart = Len(txt卡号.Text)
        End If
        KeyAscii = 0
        If Not GetPatient(txt卡号.Text) Then
            Call ClearFace
            txt卡号.SetFocus: Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt卡号_LostFocus()
    '问题号:56597
   If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled False
   If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
   If Not gobjOneCardComLib Is Nothing Then gobjOneCardComLib.SetEnabled False
   
End Sub
