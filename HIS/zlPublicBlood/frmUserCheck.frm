VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUserCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "发血确认"
   ClientHeight    =   4785
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmUserCheck.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6915
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4875
      ScaleHeight     =   270
      ScaleWidth      =   1890
      TabIndex        =   26
      Top             =   2010
      Width           =   1920
      Begin MSComCtl2.DTPicker DTPTime 
         Height          =   330
         Left            =   -30
         TabIndex        =   5
         Top             =   -30
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   265420803
         CurrentDate     =   43019
      End
   End
   Begin VB.PictureBox picBlood 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   885
      ScaleHeight     =   270
      ScaleWidth      =   2955
      TabIndex        =   23
      Top             =   450
      Width           =   2985
      Begin VB.ComboBox cboBlood 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   -30
         Width           =   3030
      End
   End
   Begin VB.PictureBox picOper 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   6540
      Picture         =   "frmUserCheck.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   225
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox TXT密码 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   4875
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1590
      Width           =   1920
   End
   Begin VB.PictureBox picOper 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   6540
      Picture         =   "frmUserCheck.frx":06C3
      ScaleHeight     =   240
      ScaleWidth      =   225
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3285
      Width           =   255
   End
   Begin VB.TextBox txt姓名 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   4875
      TabIndex        =   7
      Top             =   3270
      Width           =   1920
   End
   Begin VB.TextBox txt姓名 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   4875
      TabIndex        =   2
      Top             =   1185
      Width           =   1920
   End
   Begin VB.TextBox txt用户 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   4875
      TabIndex        =   6
      Top             =   2850
      Width           =   1920
   End
   Begin VB.TextBox TXT密码 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4875
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   3705
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   15
      TabIndex        =   13
      Top             =   4080
      Width           =   7350
   End
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5715
      TabIndex        =   11
      Top             =   4335
      Width           =   1100
   End
   Begin VB.CommandButton CMD确认 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4530
      TabIndex        =   10
      Top             =   4335
      Width           =   1100
   End
   Begin VB.TextBox txt用户 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   4875
      TabIndex        =   1
      Top             =   765
      Width           =   1920
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
      Height          =   3210
      Left            =   120
      TabIndex        =   21
      Top             =   795
      Width           =   3750
      _cx             =   6615
      _cy             =   5662
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   12
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmUserCheck.frx":0A7C
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblBlood 
      AutoSize        =   -1  'True
      Caption         =   "血液信息"
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   495
      Width           =   720
   End
   Begin VB.Label Lbl口令 
      AutoSize        =   -1  'True
      Caption         =   "密      码"
      Height          =   180
      Index           =   0
      Left            =   3930
      TabIndex        =   22
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label lblInfo 
      Caption         =   "提示："
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   150
      TabIndex        =   20
      Top             =   4305
      Width           =   4065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "身份验证"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   4395
      TabIndex        =   19
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "信息核对(3查8对)"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   180
      Width           =   1440
   End
   Begin VB.Label Lbl姓名 
      AutoSize        =   -1  'True
      Caption         =   "取血人姓名"
      Height          =   180
      Index           =   1
      Left            =   3930
      TabIndex        =   17
      Top             =   3330
      Width           =   900
   End
   Begin VB.Label Lbl姓名 
      AutoSize        =   -1  'True
      Caption         =   "发血人姓名"
      Height          =   180
      Index           =   0
      Left            =   3930
      TabIndex        =   16
      Top             =   1245
      Width           =   900
   End
   Begin VB.Label Lbl用户名 
      AutoSize        =   -1  'True
      Caption         =   "取血人帐号"
      Height          =   180
      Index           =   1
      Left            =   3930
      TabIndex        =   15
      Top             =   2910
      Width           =   900
   End
   Begin VB.Label Lbl口令 
      AutoSize        =   -1  'True
      Caption         =   "密      码"
      Height          =   180
      Index           =   1
      Left            =   3930
      TabIndex        =   14
      Top             =   3765
      Width           =   900
   End
   Begin VB.Image imgFlag 
      Height          =   345
      Left            =   3945
      Picture         =   "frmUserCheck.frx":0AE6
      Stretch         =   -1  'True
      Top             =   105
      Width           =   405
   End
   Begin VB.Label lbl日期 
      AutoSize        =   -1  'True
      Caption         =   "发血日期"
      Height          =   180
      Index           =   0
      Left            =   4110
      TabIndex        =   12
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label Lbl用户名 
      AutoSize        =   -1  'True
      Caption         =   "发血人帐号"
      Height          =   180
      Index           =   0
      Left            =   3930
      TabIndex        =   0
      Top             =   825
      Width           =   900
   End
End
Attribute VB_Name = "frmUserCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean  '为True表示已经正常显示出
Private mblnOk As Boolean
Private mstrOper As String '核对人
Private mstrSendTime As String '核对时间
Private mstrCheckResult As String
Private mblnTakeVerification As Boolean '取血人身份验证
Private mlngDeptID As Long '病人科室ID
Private mlngSendDeptID As Long '发血科室ID
Private mstr申请时间 As String
Private mstr完成时间 As String
Private mblnSelectSendUser As Boolean
Private mintMode As Integer
Private marrHTitle(0 To 2) As String  '核对人
Private marrFTitle(0 To 2) As String  '复查人
Private mstrIDs As String   '血液收发ID信息
Private mBloodResult As Collection  '血液核对结果信息
Private mlngPreBoodID As Long  '上一次选择的血液ID
Private mlngModul As Long '调用模块

Private Enum Vsf_COL
    COL_序号 = 0
    COL_名称 = 1
    COL_结果 = 2
End Enum

Public Property Get SendAndTakeOper() As String '发血人'取血人/接受人'核对人/核对人'复查人
    SendAndTakeOper = mstrOper
End Property

Public Property Get BloodResult() As Collection
    Set BloodResult = mBloodResult
End Property

Public Property Get CheckResult() As String '检查12项目结果内容
    CheckResult = mstrCheckResult
End Property

Public Property Get SendTime() As String '发血时间
    SendTime = Format(mstrSendTime, "YYYY-MM-DD HH:mm")
End Property

Public Function ShowMe(ByVal frmParent As Object, ByVal lngModul As Enum_Inside_Program, ByVal lngDeptID As Long, ByVal lngSendDeptId As Long, ByVal str申请时间 As String, ByVal str完成时间 As String, _
    Optional ByVal blnSelectSendUser As Boolean = True, Optional ByVal intMode As Enum_CheckType = 发血核对, Optional ByVal strIDs As String = "") As Boolean
'功能：发血、接收、执行过程的双核对
'参数: strIDs 血液收发ID，格式以逗号分割(传入则可对没袋血液的核对结果进行设置,否则返回统一结果),传入ID获取结果通过属性“BloodResult”，否则通过属性"CheckResult"获取
'1、对于发血功能而言：发血时间不能小于申请时间，不能超出病人就诊时间(完成时间未空则不检查)
'2、对于接收功能而言：接收时间不能小于发血时间，不能超出病人就诊时间(完成时间未空则不检查)
'3、对于执行核对而言：核对时间不能小于接收时间或上次执行时间，不能大于下次执行时间(完成时间为空则不检查)
    mlngModul = lngModul
    mlngDeptID = lngDeptID
    mlngSendDeptID = lngSendDeptId
    mstr申请时间 = str申请时间
    mstr完成时间 = str完成时间
    mblnSelectSendUser = blnSelectSendUser
    mintMode = intMode
    mstrIDs = strIDs
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    CMD放弃.Enabled = BlnState
    CMD确认.Enabled = BlnState
End Sub

Private Sub cboBlood_Click()
    Dim strTmp As String, strCheck As String
    Dim i As Integer, j As Integer
    With cboBlood
        If mlngPreBoodID = cboBlood.ListIndex Then Exit Sub
        If mlngPreBoodID = -1 Then '首次进入(默认全选)
            For i = vsfCheck.FixedRows To vsfCheck.Rows - 1
                strTmp = "11111111111"
            Next
        Else
            '保存上一次的选择
            For i = vsfCheck.FixedRows To vsfCheck.Rows - 1
                strCheck = strCheck & IIf(Abs(Val(vsfCheck.TextMatrix(i, COL_结果))) = 1, "1", "0")
            Next
            mBloodResult.Remove ("B_" & IIf(.ItemData(mlngPreBoodID) = -1, 0, .ItemData(mlngPreBoodID)))
            mBloodResult.Add strCheck, "B_" & IIf(.ItemData(mlngPreBoodID) = -1, 0, .ItemData(mlngPreBoodID))
            '刷新本次的选择
            strTmp = CStr(mBloodResult("B_" & IIf(.ItemData(.ListIndex) = -1, 0, .ItemData(.ListIndex))))
        End If
        j = 1
        For i = vsfCheck.FixedRows To vsfCheck.Rows - 1
            vsfCheck.TextMatrix(i, COL_结果) = Mid(strTmp, j, 1)
            j = j + 1
        Next
        mlngPreBoodID = cboBlood.ListIndex
    End With
End Sub

Private Sub CMD放弃_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub CMD确认_Click()
    Dim strNote As String
    Dim strUserName As String, strPassword As String
    Dim arrUserName(0 To 1) As String, arrPassword(0 To 1) As String
    Dim strServerName As String
    Dim intCheck As Integer, blnSendUserCheck As Boolean
    Dim strCheck As String
    Dim i As Integer
    On Error GoTo InputError
    
    Call Me.ValidateControls
    
    SetConState False
    '取血人和发血人验证检查
    '------检验用户是否oracle合法用户----------------
    blnSendUserCheck = Val(Lbl用户名(intCheck).Tag) <> UserInfo.id
    For intCheck = 0 To 1
        If blnSendUserCheck = True And intCheck = 0 Or intCheck = 1 Then
            strUserName = Trim(txt用户(intCheck).Text)
            strPassword = Trim(TXT密码(intCheck).Text)
            
            '有效字符串效验
            If Len(Trim(txt用户(intCheck))) = 0 Then
                strNote = "请输入" & IIf(intCheck = 0, marrHTitle(mintMode) & "人", marrFTitle(mintMode) & "人") & "帐号"
                Call gobjControl.ControlSetFocus(txt用户(1))
                GoTo InputError
            End If
            
            If Len(strUserName) <> 1 Then
                If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
                    strNote = IIf(intCheck = 0, marrHTitle(mintMode) & "人", marrFTitle(mintMode) & "人") & "帐号错误"
                    Call gobjControl.ControlSetFocus(txt用户(intCheck))
                    SetConState
                    Exit Sub
                End If
            End If
            If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
                If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
                    strNote = IIf(intCheck = 0, marrHTitle(mintMode) & "人", marrFTitle(mintMode) & "人") & "帐号密码错误"
                    Call gobjControl.ControlSetFocus(TXT密码(intCheck))
                    GoTo InputError
                End If
            End If
            
            If Len(Trim(strPassword)) = 0 Then
                strNote = "请输入" & IIf(intCheck = 0, marrHTitle(mintMode) & "人", marrFTitle(mintMode) & "人") & "帐号密码"
                Call gobjControl.ControlSetFocus(TXT密码(intCheck))
                GoTo InputError
            End If
        End If
        arrUserName(intCheck) = strUserName
        arrPassword(intCheck) = strPassword
     Next
    
'    If IsDate(TXT日期(0).Text) = False Then
'        strNote = marrHTitle(mintMode) & "日期不是有效的日期格式，请检查！"
'        Call gobjControl.ControlSetFocus(TXT日期(0))
'        GoTo InputError
'    End If
    '发血人和取血人不能是同一个
    If txt姓名(1).Text = "" Then
        strNote = "请输入" & marrHTitle(mintMode) & "人"
        Call gobjControl.ControlSetFocus(txt姓名(1))
        GoTo InputError
    End If
    If txt姓名(0).Text = txt姓名(1).Text Then
        strNote = marrHTitle(mintMode) & "人不能和" & marrFTitle(mintMode) & "人是同一个人，请重新确定" & marrFTitle(mintMode) & "人！"
        Call gobjControl.ControlSetFocus(txt姓名(1))
        GoTo InputError
    End If
    '用户登录验证
    If blnSendUserCheck = True Then
        '用户登录验证
        If GetObjectRegister = False Then Exit Sub
        strServerName = gobjRegister.GetServerName
        If gobjRegister.LoginValidate(strServerName, arrUserName(0), arrPassword(0), strNote) = False Then
            TXT密码(0).Text = ""
            Call gobjControl.ControlSetFocus(TXT密码(0))
            SetConState
            GoTo InputError
        End If
    End If
    
    '用户登录验证
    If GetObjectRegister = False Then Exit Sub
    strServerName = gobjRegister.GetServerName
    If gobjRegister.LoginValidate(strServerName, arrUserName(1), arrPassword(1), strNote) = False Then
        TXT密码(1).Text = ""
        Call gobjControl.ControlSetFocus(TXT密码(1))
        SetConState
        GoTo InputError
    End If
    
    
    mstrOper = txt姓名(0).Text & "'" & txt姓名(1).Text
    mstrSendTime = DTPTime.Value
    strCheck = ""
    For i = vsfCheck.FixedRows To vsfCheck.Rows - 1
        strCheck = strCheck & IIf(Abs(Val(vsfCheck.TextMatrix(i, COL_结果))) = 1, "1", "0")
    Next
    mstrCheckResult = strCheck
    If Not mBloodResult Is Nothing Then
        If mBloodResult.Count > 0 Then
            If cboBlood.ItemData(cboBlood.ListIndex) = -1 Then '表示设置所有
                Call mBloodResult.Remove("B_0")
                For i = 1 To cboBlood.ListCount - 1
                    Call mBloodResult.Remove("B_" & cboBlood.ItemData(i))
                    mBloodResult.Add strCheck, "B_" & cboBlood.ItemData(i)
                Next
            Else
                Call mBloodResult.Remove("B_" & cboBlood.ItemData(cboBlood.ListIndex))
                mBloodResult.Add strCheck, "B_" & cboBlood.ItemData(cboBlood.ListIndex)
            End If
        End If
    End If
    
    mblnOk = True
    Unload Me
    Exit Sub
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbInformation, gstrSysName
    End If
    SetConState
    Exit Sub
End Sub

Private Sub DTPTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim lngScrH As Long
    If mblnFirst = False Then
        '显示窗体
        lngScrH = GetSystemMetrics(17) * 15 '屏幕可用高度
        If Me.Top + Me.Height > lngScrH Then
            Me.Top = lngScrH - Me.Height
        End If
    
        If Trim(txt用户(1).Text) = "" Then
            CMD确认.Default = False
            txt用户(1).SetFocus
        Else
            If TXT密码(1).Enabled Then
                TXT密码(1).SetFocus
            Else
                CMD确认.SetFocus
            End If
        End If
        mblnFirst = True
        If Trim(txt用户(1).Text) <> "" And Trim(TXT密码(1).Text) <> "" Then Call CMD确认_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strName As String, arrName
    Dim i As Integer
    
    marrHTitle(0) = "发血"
    marrHTitle(1) = "接收"
    marrHTitle(2) = "核查"
    marrFTitle(0) = "取血"
    marrFTitle(1) = "核查"
    marrFTitle(2) = "复查"
    
    If mstrIDs = "" Then
        lblBlood.Visible = False
        picBlood.Visible = False
        vsfCheck.Height = vsfCheck.Height + vsfCheck.Top - picBlood.Top
        vsfCheck.Top = picBlood.Top
    End If
    mlngPreBoodID = -1
    Call LoadBoold
    With vsfCheck
        .Clear
        .Rows = 12
        .Cols = 3
        .ColWidth(COL_序号) = 500
        .ColWidth(COL_名称) = 2000
        .ColWidth(COL_结果) = 500
        .TextMatrix(0, COL_序号) = "序号"
        .TextMatrix(0, COL_名称) = "核查项目"
        .TextMatrix(0, COL_结果) = "结果"
        strName = "血液制品有效期'血液制品质量'输血装置是否完好'患者姓名'患者住院号'患者病室'患者床号'患者血型'血袋号'血液制品种类'剂量"
        arrName = Split(strName, "'")
        For i = 0 To UBound(arrName)
            .TextMatrix(i + 1, COL_序号) = i + 1
            .TextMatrix(i + 1, COL_名称) = arrName(i)
        Next i
        .ColDataType(COL_结果) = flexDTBoolean
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, COL_结果) = 1
        Next
        .Editable = flexEDKbdMouse
    End With
    
    Lbl用户名(0).Tag = UserInfo.id
    txt用户(0).Text = UserInfo.用户名
    txt用户(0).Tag = txt用户(0).Text
    txt用户(0).locked = Not mblnSelectSendUser: txt用户(0).ForeColor = IIf(mblnSelectSendUser = False, COLOR.深灰色, COLOR.黑色)
    txt姓名(0).Text = UserInfo.姓名
    txt姓名(0).Tag = txt姓名(0).Text
    txt姓名(0).locked = Not mblnSelectSendUser: txt姓名(0).ForeColor = IIf(mblnSelectSendUser = False, COLOR.深灰色, COLOR.黑色)
    TXT密码(0).Text = "123"
    TXT密码(0).locked = True: TXT密码(0).ForeColor = COLOR.深灰色: TXT密码(0).Enabled = mblnSelectSendUser
    picOper(0).Visible = mblnSelectSendUser: picOper(0).Enabled = mblnSelectSendUser
    If IsDate(mstr完成时间) Then
        DTPTime.Value = Format(mstr完成时间, "YYYY-MM-DD HH:mm")
    Else
        DTPTime.Value = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    End If
    DTPTime.Tag = DTPTime.Value
    
    Lbl用户名(0).Caption = marrHTitle(mintMode) & "人帐号"
    Lbl姓名(0).Caption = marrHTitle(mintMode) & "人姓名"
    lbl日期(0).Caption = marrHTitle(mintMode) & "日期"
    Lbl用户名(1).Caption = marrFTitle(mintMode) & "人帐号"
    Lbl姓名(1).Caption = marrFTitle(mintMode) & "人姓名"
    
    If mintMode = 发血核对 Then
        Me.Caption = "发血核对"
    ElseIf mintMode = 接收核对 Then
        Me.Caption = "接收核对"
    Else
        Me.Caption = "执行核对"
    End If
    mblnFirst = False
    mblnOk = False
End Sub

Private Sub picOper_Click(Index As Integer)
    If GetUserName(txt姓名(Index), Index) = True Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Function GetUserName(ByVal objControl As TextBox, ByVal intIndex As Integer, Optional ByVal StrInput As String = "") As Boolean
    Dim rsUser As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim vPoint As RECT, blnCancel As Boolean
    Dim str部门性质 As String, str人员性质 As String
    
    On Error GoTo ErrHand
    If objControl.locked = True Then GetUserName = True: Exit Function
    If StrInput <> "" Then
         If IsNumeric(StrInput) Then
            strWhere = " And b.编号 Like [2]"
         ElseIf gobjCommFun.IsCharAlpha(StrInput) Then
            strWhere = " And b.简码 Like [2]"
            StrInput = UCase(StrInput)
         Else
            strWhere = " And b.姓名 Like [2]"
         End If
    End If
    
    '以技站则不区分部门性质和人员
    If Not mlngModul = p医技工作站 Then
        '发血人为血库人员即可，取血人为临床护士
         strWhere = strWhere & _
            "   And Exists  (Select 1 From 部门性质说明 Where 部门id = d.部门id And Instr([3], ',' || 工作性质 || ',', 1) <> 0 And 服务对象 In (0, 1, 2, 3))"
        If Not (mintMode = 0 And intIndex = 0) Then  '发血
            strWhere = strWhere & _
                "   And Exists  (Select 1 From 人员性质说明 Where 人员id = b.Id And Instr([4], ',' || 人员性质 || ',', 1) <> 0) "
        End If
        
        If mintMode = 0 Then
            str部门性质 = IIf(intIndex = 0, ",血库,", ",临床,护理,")
            str人员性质 = IIf(intIndex = 0, "", ",护士,")
        ElseIf mintMode = 1 Then
            str部门性质 = ",临床,护理,"
            str人员性质 = ",医生,护士,"
        Else
            str部门性质 = ",临床,护理,"
            str人员性质 = ",医生,护士,"
        End If
    End If

    vPoint = GetControlRect(objControl.hWnd)
    strSQL = _
        " Select  Rownum || '-' || b.Id as id, c.用户名,b.编号, b.姓名,b.简码,a.名称 as 科室" & vbNewLine & _
        " From 部门表 a, 人员表 b, 上机人员表 c, 部门人员 d" & vbNewLine & _
        " Where a.Id = d.部门id  And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) " & vbNewLine & _
        " " & strWhere & " And b.Id = c.人员id  And c.人员id = d.人员id And d.部门id = [1]"
    Set rsUser = gobjDatabase.ShowSQLSelect(Me, strSQL, 0, "", False, txt用户(intIndex).Text, "请选择一个" & IIf(intIndex = 0, marrHTitle(mintMode), marrFTitle(mintMode)) & "人员", False, False, True, vPoint.Left, vPoint.Top, objControl.Height, blnCancel, False, False, _
                    IIf(intIndex = 0, mlngSendDeptID, mlngDeptID), StrInput & "%", str部门性质, str人员性质)
    If Not rsUser Is Nothing Then
        If blnCancel = False Then
            If rsUser.EOF Then Exit Function
            Lbl用户名(intIndex).Tag = Split(rsUser!id, "-")(1)
            txt用户(intIndex).Text = Nvl(rsUser!用户名)
            txt用户(intIndex).Tag = txt用户(intIndex).Text
            objControl.Text = Nvl(rsUser!姓名)
            objControl.Tag = objControl.Text
            objControl.SetFocus
            If intIndex = 0 Then '发血人
                If Lbl用户名(intIndex).Tag = UserInfo.id Then
                    TXT密码(intIndex).Text = "123"
                    TXT密码(intIndex).ForeColor = COLOR.深灰色
                    TXT密码(intIndex).locked = True
                Else
                    TXT密码(intIndex).Text = ""
                    TXT密码(intIndex).ForeColor = COLOR.黑色
                    TXT密码(intIndex).locked = False
                End If
            End If
            GetUserName = True
        End If
    Else
        If StrInput = "" And blnCancel = False Then
            If mlngModul = p医技工作站 Then
                MsgBox "没有对应的医技人员信息，请在人员管理中设置！", vbInformation, gstrSysName
            Else
                If mintMode = 0 Then
                    If intIndex = 0 Then
                        MsgBox "没有对应的血库人员信息，请在人员管理中设置！", vbInformation, gstrSysName
                    Else
                        MsgBox "没有对应的临床护士信息，请在人员管理中设置！", vbInformation, gstrSysName
                    End If
                Else
                    MsgBox "没有对应的临床护士和医生信息，请在人员管理中设置！", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub TXT密码_GotFocus(Index As Integer)
    GetFocus TXT密码(Index)
End Sub

Private Sub TXT密码_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub DTPTime_Validate(Cancel As Boolean)
    Dim blnOk As Boolean
    Dim strMsg As String, strCurDate As String

    '日期合法性检查
    blnOk = True: strMsg = ""
    If IsDate(mstr申请时间) Then
        If Format(DTPTime.Value, "YYYY-MM-DD HH:mm") < Format(mstr申请时间, "YYYY-MM-DD HH:mm") Then
            blnOk = False
            If mintMode = 发血核对 Then
                strMsg = "发血日期只能在申请日期[" & Format(mstr申请时间, "YYYY-MM-DD HH:mm") & "]之后"
            ElseIf mintMode = 接收核对 Then
                strMsg = "接收日期只能在发血日期[" & Format(mstr申请时间, "YYYY-MM-DD HH:mm") & "]之后"
            Else
                strMsg = "核对日期只能在接收日期或上次执行日期[" & Format(mstr申请时间, "YYYY-MM-DD HH:mm") & "]之后"
            End If
            GoTo ShowMsg
        End If
    End If
    If IsDate(mstr完成时间) = False Then
        strCurDate = gobjDatabase.Currentdate
        If Format(DTPTime.Value, "YYYY-MM-DD HH:mm") > Format(strCurDate, "YYYY-MM-DD HH:mm") Then
            blnOk = False
            strMsg = marrHTitle(mintMode) & "日期不能大于当前日期[" & Format(strCurDate, "YYYY-MM-DD HH:mm") & "]"
            GoTo ShowMsg
        End If
    Else
        If Format(DTPTime.Value, "YYYY-MM-DD HH:mm") > Format(mstr完成时间, "YYYY-MM-DD HH:mm") Then
            blnOk = False
            If mintMode = 发血核对 Then
                strMsg = "发血日期不能大于病人完成就诊日期[" & Format(mstr完成时间, "YYYY-MM-DD HH:mm") & "]"
            ElseIf mintMode = 接收核对 Then
                strMsg = "接收日期不能大于病人完成就诊日期[" & Format(mstr申请时间, "YYYY-MM-DD HH:mm") & "]"
            Else
                strMsg = "核对日期不能大于结束日期或下次执行日期[" & Format(mstr申请时间, "YYYY-MM-DD HH:mm") & "]"
            End If
            GoTo ShowMsg
        End If
    End If
ShowMsg:
    If blnOk = False Then
        MsgBox strMsg, vbInformation, gstrSysName
        Cancel = True
        DTPTime.Value = DTPTime.Tag
        DTPTime.SetFocus
        Exit Sub
    End If
    DTPTime.Tag = DTPTime.Value
End Sub

Private Sub txt姓名_GotFocus(Index As Integer)
    GetFocus txt姓名(Index)
End Sub

Private Sub txt姓名_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If Index = 1 Then
            DTPTime.SetFocus
        End If
    End If
End Sub

Private Sub txt姓名_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim StrInput As String
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        StrInput = txt姓名(Index).Text
        If StrInput <> "" And txt姓名(Index).Text <> txt姓名(Index).Tag Then
            If GetUserName(txt姓名(Index), Index, StrInput) = False Then Exit Sub
        End If
        gobjCommFun.PressKey vbKeyTab
    Else
        If KeyAscii = 39 Then KeyAscii = 0
    End If
End Sub

Private Sub txt姓名_Validate(Index As Integer, Cancel As Boolean)
    If Index = 1 Then
        If txt姓名(Index).Tag <> "" And txt姓名(Index).Tag <> txt姓名(Index).Text Then txt姓名(Index).Text = txt姓名(Index).Tag
    End If
End Sub

Private Sub txt用户_Change(Index As Integer)
    If Not mblnFirst Then Exit Sub
    CMD确认.Default = False
End Sub

Private Sub txt用户_GotFocus(Index As Integer)
    GetFocus txt用户(Index)
End Sub

Private Sub txt用户_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub txt用户_Validate(Index As Integer, Cancel As Boolean)
    Dim strText As String
    Dim lngUserID As Long, strUser As String, strName As String
    Dim rsOper As New ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim str部门性质 As String, str人员性质 As String
    Dim str部门名称 As String
    
    strText = txt用户(Index).Text
    
    On Error GoTo ErrHand
    If Index = 0 Or Index = 1 Then
        '人员提取
        If txt用户(Index).locked = True Then Exit Sub
        If strText = "" Then txt用户(Index).Tag = txt用户(Index).Text: Exit Sub
        strSQL = "Select a.Id, a.姓名,B.用户名 From 人员表 a, 上机人员表 b Where a.Id = b.人员id And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And b.用户名 = [1]"
        Set rsOper = gobjDatabase.OpenSQLRecord(strSQL, "", UCase(strText))
        If rsOper.EOF Then
            txt用户(Index).Text = txt用户(Index).Tag
            lblInfo.Caption = "提示：" & IIf(Index = 0, marrHTitle(mintMode), marrFTitle(mintMode)) & "人帐号不正确"
            Call txt用户_GotFocus(Index)
            Cancel = True
            Exit Sub
        End If
        lngUserID = Val(rsOper!id)
        strUser = "" & rsOper!用户名
        strName = "" & rsOper!姓名
        
        If Not mlngModul = p医技工作站 Then
            strWhere = strWhere & _
              "   And Exists  (Select 1 From 部门性质说明 Where 部门id = a.id And Instr([3], ',' || 工作性质 || ',', 1) <> 0 And 服务对象 In (0, 1, 2, 3))"
            If Not (mintMode = 0 And Index = 0) Then '发血操作的发血人不用指定人员性质
                strWhere = strWhere & _
                    "   And Exists  (Select 1 From 人员性质说明 Where 人员id = b.人员id And Instr([4], ',' || 人员性质 || ',', 1) <> 0) "
            End If
            
            If mintMode = 0 Then
                str部门性质 = IIf(Index = 0, ",血库,", ",临床,护理,")
                str人员性质 = IIf(Index = 0, "", ",护士,")
            ElseIf mintMode = 1 Then
                str部门性质 = ",临床,护理,"
                str人员性质 = ",医生,护士,"
            Else
                str部门性质 = ",临床,护理,"
                str人员性质 = ",医生,护士,"
            End If
        End If
        
        '执行核对，通过输入帐号提取人员，则不限制是否是当前科室的(应为输血病人，可能存在两个不同科室的验证)
        If Not mintMode = 执行核对 Then
            strSQL = "Select a.名称, b.人员id From 部门表 a, 部门人员 b Where a.Id = b.部门id And a.Id = [1]   And b.人员id = [2] " & strWhere
            Set rsOper = gobjDatabase.OpenSQLRecord(strSQL, "", IIf(Index = 0, mlngSendDeptID, mlngDeptID), lngUserID, str部门性质, str人员性质)
            If rsOper.EOF Then
                strSQL = " Select 名称 from 部门表 where ID=[1]"
                Set rsOper = gobjDatabase.OpenSQLRecord(strSQL, "获取部门名称", IIf(Index = 0, mlngSendDeptID, mlngDeptID))
                If rsOper.EOF Then
                    str部门名称 = ""
                Else
                    str部门名称 = "[" & rsOper!名称 & "]"
                End If
                txt用户(Index).Text = txt用户(Index).Tag
                If mlngModul = p医技工作站 Then
                    lblInfo.Caption = "提示：该" & marrFTitle(mintMode) & "人并非属于科室" & str部门名称 & "的医技人员"
                Else
                    If mintMode = 0 Then
                        If Index = 0 Then
                            lblInfo.Caption = "提示：该" & marrHTitle(mintMode) & "人并非属于当前发血科室" & str部门名称
                        Else
                            lblInfo.Caption = "提示：该" & marrFTitle(mintMode) & "人并非属于病人当前科室" & str部门名称 & "的护士"
                        End If
                    Else
                        lblInfo.Caption = "提示：该" & marrFTitle(mintMode) & "人并非属于科室" & str部门名称 & "的医生或护士"
                    End If
                End If
                Call txt用户_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
        End If
        Lbl用户名(Index).Tag = lngUserID
        txt用户(Index).Text = strUser
        txt用户(Index).Tag = txt用户(Index).Text
        txt姓名(Index).Text = strName
        txt姓名(Index).Tag = txt姓名(Index).Text
        
        If Index = 0 Then  '发血人
            If Lbl用户名(Index).Tag = UserInfo.id Then
                TXT密码(Index).Text = "123"
                TXT密码(Index).ForeColor = COLOR.深灰色
                TXT密码(Index).locked = True
            Else
                TXT密码(Index).Text = ""
                TXT密码(Index).ForeColor = COLOR.黑色
                TXT密码(Index).locked = False
            End If
        End If
            
        lblInfo.Caption = ""
    End If
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfCheck_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> COL_结果 Then Cancel = True
End Sub

Private Sub LoadBoold()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    If mstrIDs = "" Then Exit Sub
    Set mBloodResult = New Collection
    picBlood.Enabled = True
    On Error GoTo ErrHand
    strSQL = _
        " Select /*+CARDINALITY(b 10)*/" & vbNewLine & _
        " a.Id, a.血袋编号, c.名称, c.规格" & vbNewLine & _
        " From 收费项目目录 c, 血液收发记录 a, Table(f_Num2list([1])) b" & vbNewLine & _
        " Where  c.Id = a.血液id And a.Id = b.Column_Value"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "提取血液信息", mstrIDs)
    With cboBlood
        .Clear
        If rsTemp.RecordCount > 1 Then
            .AddItem "本次所有血液统一设置"
            .ItemData(.NewIndex) = -1
            mBloodResult.Add "11111111111", "B_0"
        End If
        Do While Not rsTemp.EOF
            '默认核对结果都正常
            mBloodResult.Add "11111111111", "B_" & rsTemp!id
            .AddItem "编号:" & rsTemp!血袋编号 & "   名称:" & rsTemp!名称 & "   规格" & rsTemp!规格
            .ItemData(.NewIndex) = rsTemp!id
        rsTemp.MoveNext
        Loop
        gobjComlib.cbo.SetListHeight cboBlood, 360
        gobjComlib.cbo.SetListWidthAuto cboBlood
        If .ListIndex = -1 And .ListCount > 0 Then .ListIndex = 0
    End With
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub
