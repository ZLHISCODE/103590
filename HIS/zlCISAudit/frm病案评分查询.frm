VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm病案评分查询 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病案评分检索"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   Icon            =   "frm病案评分查询.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm病案评分查询.frx":000C
   ScaleHeight     =   7875
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo 
      Height          =   300
      Left            =   1620
      TabIndex        =   38
      Top             =   4455
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   10
      Left            =   1620
      TabIndex        =   19
      Tag             =   "出院科室"
      Top             =   4050
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   9
      Left            =   1620
      TabIndex        =   17
      Tag             =   "审核人"
      Top             =   3615
      Width           =   3525
   End
   Begin zl9CISAudit.tipPopup tipPopup1 
      Height          =   420
      Left            =   1140
      Top             =   7095
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "缺省(&W)"
      Height          =   350
      Left            =   360
      TabIndex        =   36
      Top             =   7410
      Width           =   1100
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   8
      Left            =   1620
      TabIndex        =   26
      Tag             =   "入院科室"
      Top             =   5640
      Width           =   3525
   End
   Begin MSComCtl2.DTPicker dt出院日期 
      Height          =   300
      Index           =   0
      Left            =   2655
      TabIndex        =   22
      Top             =   4845
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   136314880
      CurrentDate     =   38373
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   7
      Left            =   1620
      TabIndex        =   15
      Tag             =   "评分人"
      Top             =   3180
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   6
      Left            =   1620
      TabIndex        =   13
      Tag             =   "责任护士"
      Top             =   2790
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   5
      Left            =   1620
      TabIndex        =   11
      Tag             =   "门诊医师"
      Top             =   2415
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   4
      Left            =   1620
      TabIndex        =   9
      Tag             =   "住院医师"
      Top             =   2025
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   3
      Left            =   1620
      TabIndex        =   7
      Tag             =   "姓名"
      Top             =   1650
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   2
      Left            =   1620
      TabIndex        =   5
      Tag             =   "主页ID"
      Top             =   1260
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   1
      Left            =   1620
      TabIndex        =   3
      Tag             =   "住院号"
      Top             =   885
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   0
      Left            =   1620
      TabIndex        =   1
      Tag             =   "病人ID"
      Top             =   495
      Width           =   3525
   End
   Begin VB.OptionButton optOr 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "满足任一条件(&V)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3450
      TabIndex        =   33
      Top             =   6855
      Width           =   1665
   End
   Begin VB.OptionButton optAnd 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "满足所有条件(&U)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1620
      TabIndex        =   32
      Top             =   6855
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2922
      TabIndex        =   34
      Top             =   7410
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4272
      TabIndex        =   35
      Top             =   7410
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dt出院日期 
      Height          =   300
      Index           =   1
      Left            =   2655
      TabIndex        =   24
      Top             =   5265
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   136314880
      CurrentDate     =   38373
   End
   Begin MSComCtl2.DTPicker dt入院日期 
      Height          =   300
      Index           =   0
      Left            =   2655
      TabIndex        =   29
      Top             =   6030
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   136314880
      CurrentDate     =   38373
   End
   Begin MSComCtl2.DTPicker dt入院日期 
      Height          =   300
      Index           =   1
      Left            =   2655
      TabIndex        =   31
      Top             =   6420
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   136314880
      CurrentDate     =   38373
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病理类型(&S)"
      Height          =   180
      Left            =   525
      TabIndex        =   39
      Top             =   4500
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "出院科室(&L)"
      Height          =   180
      Index           =   16
      Left            =   525
      TabIndex        =   18
      Top             =   4110
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "审 核 人(&K)"
      Height          =   180
      Index           =   15
      Left            =   525
      TabIndex        =   16
      Top             =   3675
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病 人 ID(&A)"
      Height          =   180
      Index           =   0
      Left            =   525
      TabIndex        =   0
      Top             =   555
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住 院 号(&B)"
      Height          =   180
      Index           =   1
      Left            =   525
      TabIndex        =   2
      Top             =   945
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住院次数(&D)"
      Height          =   180
      Index           =   2
      Left            =   525
      TabIndex        =   4
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人姓名(&E)"
      Height          =   180
      Index           =   3
      Left            =   525
      TabIndex        =   6
      Top             =   1710
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "主治医师(&F)"
      Height          =   180
      Index           =   4
      Left            =   525
      TabIndex        =   8
      Top             =   2085
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "门诊医师(&G)"
      Height          =   180
      Index           =   5
      Left            =   525
      TabIndex        =   10
      Top             =   2475
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "责任护士(&I)"
      Height          =   180
      Index           =   6
      Left            =   525
      TabIndex        =   12
      Top             =   2850
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "评 分 人(&J)"
      Height          =   180
      Index           =   7
      Left            =   525
      TabIndex        =   14
      Top             =   3240
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "出院日期(&M)"
      Height          =   180
      Index           =   8
      Left            =   525
      TabIndex        =   20
      Top             =   4935
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "入院科室(&Q)"
      Height          =   180
      Index           =   9
      Left            =   525
      TabIndex        =   25
      Top             =   5700
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "入院日期(&R)"
      Height          =   180
      Index           =   10
      Left            =   525
      TabIndex        =   27
      Top             =   6090
      Width           =   990
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "检索条件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   180
      TabIndex        =   37
      Top             =   90
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   5940
      Y1              =   7245
      Y2              =   7245
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   -105
      X2              =   5940
      Y1              =   7260
      Y2              =   7260
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&T)"
      Height          =   180
      Index           =   14
      Left            =   1620
      TabIndex        =   30
      Top             =   6480
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&S)"
      Height          =   180
      Index           =   13
      Left            =   1620
      TabIndex        =   28
      Top             =   6090
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&P)"
      Height          =   180
      Index           =   12
      Left            =   1620
      TabIndex        =   23
      Top             =   5325
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&N)"
      Height          =   180
      Index           =   11
      Left            =   1620
      TabIndex        =   21
      Top             =   4935
      Width           =   990
   End
End
Attribute VB_Name = "frm病案评分查询"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnForce                As Boolean
Public mstrReturn               As String
Public mblnOK                   As Boolean
Public mbln编目后评分           As Boolean
Public mblnCancel               As Boolean
Private intPara                 As Integer

Private mlngSickID              As Long             '病人ID
Private mlngHospitalID          As Long             '住院号
Private mlngHospitalTimes       As Long             '住院次数
Private mstrSickName            As String           '病人姓名
Private mstrMainDoctor          As String           '主治医师
Private mstrOutpatientDoctor    As String           '门诊医师
Private mstrNurses              As String           '责任护士
Private mstrRatingMan           As String           '评分人
Private mstrAuditMan            As String           '审核人
Private mstrOutDept             As String           '出院科室
Private mstrInDept              As String           '入院科室
Private mdatStarOutDate         As Date             '出院开始日期
Private mdatEndOutDate          As Date             '出院开始日期
Private mdatStarInDate          As Date             '入院开始日期
Private mdatEndInDate           As Date             '入院开始日期
Private mstrSickType            As String           '病理类型

'病人ID
Public Property Get lngSickID() As Long
    lngSickID = mlngSickID
End Property

'住院号
Public Property Get lngHospitalID() As Long
    lngHospitalID = mlngHospitalID
End Property

'住院次数
Public Property Get lngHospitalTimes() As Long
    lngHospitalTimes = mlngHospitalTimes
End Property

'病人姓名
Public Property Get strSickName() As String
    strSickName = mstrSickName
End Property

'主治医师
Public Property Get strMainDoctor() As String
    strMainDoctor = mstrMainDoctor
End Property

'门诊医师
Public Property Get strOutpatientDoctor() As String
    strOutpatientDoctor = mstrOutpatientDoctor
End Property

'责任护士
Public Property Get strNurses() As String
    strNurses = mstrNurses
End Property

'评分人
Public Property Get strRatingMan() As String
    strRatingMan = mstrRatingMan
End Property

'审核人
Public Property Get strAuditMan() As String
    strAuditMan = mstrAuditMan
End Property

'出院科室
Public Property Get strOutDept() As String
    strOutDept = mstrOutDept
End Property

'入院科室
Public Property Get strInDept() As String
    strInDept = mstrInDept
End Property

'出院开始日期
Public Property Get datStarOutDate() As Date
    datStarOutDate = mdatStarOutDate
End Property

'出院结束日期
Public Property Get datEndOutDate() As Date
    datEndOutDate = mdatEndOutDate
End Property

'入院开始日期
Public Property Get datStarInDate() As Date
    datStarInDate = mdatStarInDate
End Property

'入院结束日期
Public Property Get datEndInDate() As Date
    datEndInDate = mdatEndInDate
End Property

'病理类型
Public Property Get strSickType() As String
    strSickType = mstrSickType
End Property

'==============================================================================
'=功能：取消退出
'==============================================================================
Private Sub CmdCancel_Click()
    On Error GoTo errH
    txtInfo(0).SetFocus
    mblnCancel = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：设置缺省值
'==============================================================================
Private Sub cmdDefault_Click()
    Dim i           As Long
    
    On Error GoTo errH
    
    For i = 0 To 10
        txtInfo(i).Text = ""
    Next
    
    dt出院日期(0) = DateAdd("M", -1, Now)
    dt出院日期(1) = Now
    dt入院日期(0) = ""
    dt入院日期(1) = ""
    optAnd.Value = True
    txtInfo(0).SetFocus
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：确定查询退出
'==============================================================================
Private Sub CmdOK_Click()
    Dim i           As Long
    
    On Error GoTo errH
    
    intPara = 1
        
    mstrReturn = " and (1=1 "
    For i = 0 To 10
        If Trim(txtInfo(i)) <> "" Then
            If txtInfo(i).Tag = "出院科室" Then '出院科室
                If InStrRev(txtInfo(i), ",") > 0 Then
                    mstrReturn = mstrReturn & IIf(optAnd.Value = True, " And ", " Or ") & txtInfo(i).Tag & "  In (" & Get所属部门(UserInfo.ID, 1) & ")"
                Else
                    mstrReturn = mstrReturn & IIf(optAnd.Value = True, " And ", " Or ") & txtInfo(i).Tag & " = [" & intPara & "] "
                End If
                
            Else
                mstrReturn = mstrReturn & IIf(optAnd.Value = True, " And ", " Or ") & txtInfo(i).Tag & " = [" & intPara & "] "
            End If
        End If
        intPara = intPara + 1
    Next
    
    If cbo.Text <> "" Then
        mstrReturn = mstrReturn & IIf(optAnd.Value = True, " And ", " Or ") & "病理类型=[16] "
    End If
    
    If Not IsNull(dt出院日期(0).Value) Then
        If IsNull(dt出院日期(1)) Then
            mstrReturn = mstrReturn & IIf(optAnd.Value = True, " and ", " or ") & _
                "出院日期 >= [" & intPara & "]"
        Else
            mstrReturn = mstrReturn & IIf(optAnd.Value = True, " and ", " or ") & _
                "出院日期>= [" & intPara & "] and " & _
                "出院日期<= [" & intPara + 1 & "]"
        End If
    End If
    
    intPara = intPara + 2
    
    If Not IsNull(dt入院日期(0).Value) Then
        If IsNull(dt入院日期(1)) Then
            mstrReturn = mstrReturn & IIf(optAnd.Value = True, " and ", " or ") & _
                "入院日期 >= [" & intPara & "]"
        Else
            mstrReturn = mstrReturn & IIf(optAnd.Value = True, " and ", " or ") & _
                "入院日期>= [" & intPara & "] and " & _
                "入院日期<= [" & intPara + 1 & "]"
        End If
    End If
    
    '保存参数
    mlngSickID = Val(txtInfo(0).Text)               '病人ID
    mlngHospitalID = Val(txtInfo(1).Text)           '住院号
    mlngHospitalTimes = Val(txtInfo(2).Text)        '住院次数
    mstrSickName = txtInfo(3).Text                  '病人姓名
    mstrMainDoctor = txtInfo(4).Text                '主治医师
    mstrOutpatientDoctor = txtInfo(5).Text          '门诊医师
    mstrNurses = txtInfo(6).Text                    '责任护士
    mstrRatingMan = txtInfo(7).Text                 '评分人
    mstrInDept = txtInfo(8).Text                    '入院科室
    mstrAuditMan = txtInfo(9).Text                  '审核人
    mstrOutDept = txtInfo(10).Text                  '出院科室
    mstrSickType = cbo.Text                         '病理类型
    
    If Not IsNull(dt出院日期(0).Value) Then
        mdatStarOutDate = Format(dt出院日期(0).Value, "yyyy-mm-dd 00:00:00")      '出院开始日期
    End If
    If Not IsNull(dt出院日期(1).Value) Then
        mdatEndOutDate = Format(dt出院日期(1).Value, "yyyy-mm-dd 23:59:59")       '出院开始日期
    End If
    If Not IsNull(dt入院日期(0).Value) Then
        mdatStarInDate = Format(dt入院日期(0).Value, "yyyy-mm-dd 00:00:00")       '入院开始日期
    End If
    If Not IsNull(dt入院日期(1).Value) Then
        mdatEndInDate = Format(dt入院日期(1).Value, "yyyy-mm-dd 23:59:59")        '入院开始日期
    End If
    If mstrReturn = " and (1=1 " Then
        MsgBox "请输入检索条件！", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    mblnOK = True
    
    If mbln编目后评分 Then
        mstrReturn = mstrReturn & ") and 编目日期 is not null "
    Else
        mstrReturn = mstrReturn & ") "
    End If
    mblnCancel = False
    Unload Me
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：出院日期类型判断
'==============================================================================
Private Sub dt出院日期_Change(Index As Integer)
    On Error GoTo errH
    If dt出院日期(0).Value > dt出院日期(1).Value Then MsgBox "开始日期应该比结束日期早！", vbExclamation, gstrSysName: dt出院日期(0).SetFocus: Exit Sub
    If Index = 0 Then
        If IsNull(dt出院日期(0).Value) Then dt出院日期(1).Value = Null
    Else
        If IsNull(dt出院日期(0).Value) And Not IsNull(dt出院日期(1).Value) Then dt出院日期(1).Value = Null
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：出院日期数据检测
'==============================================================================
Private Sub dt出院日期_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：入院日期类型判断
'==============================================================================
Private Sub dt入院日期_Change(Index As Integer)
    On Error GoTo errH
    If dt入院日期(0).Value > dt入院日期(1).Value Then MsgBox "开始日期应该比结束日期早！", vbExclamation, gstrSysName: dt入院日期(0).SetFocus: Exit Sub
    If Index = 0 Then
        If IsNull(dt入院日期(0).Value) Then dt入院日期(1).Value = Null
    Else
        If IsNull(dt入院日期(0).Value) And Not IsNull(dt入院日期(1).Value) Then dt入院日期(1).Value = Null
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：入院日期数据检测
'==============================================================================
Private Sub dt入院日期_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：窗口初始化'获取系统参数：是否编目后才能评分
'==============================================================================
Private Sub Form_Load()
    On Error GoTo errH
    mblnCancel = True
    mbln编目后评分 = Val(zlDatabase.GetPara(91, glngSys, 0)) = 1
    mblnForce = False
    
    Call Fill病理类型
    
    dt出院日期(0) = DateAdd("M", -1, Now)
    dt出院日期(1) = Now
    dt入院日期(0) = DateAdd("M", -1, Now)
    dt入院日期(1) = Now
    dt入院日期(0) = ""
    dt入院日期(1) = ""
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：页面关闭数据处理
'==============================================================================
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errH
    If mblnForce Then
        '在病案管理程序中强制关闭
        If IsCompiled = True Then
            Call SetWindowLong(Me.hWnd, GWL_WNDPROC, OldWindowProc)
        End If
    Else
        '用户占击了关闭按钮，并不关闭但隐藏
        Cancel = 1
        Me.Hide
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：根据条件生成查询语句
'==============================================================================
Public Function GetFilter(ByVal strPrivs As String, ByVal txtDept As String) As String
    On Error GoTo errH
    If Not IsPrivs(strPrivs, "所有科室") Then
        If txtDept <> UserInfo.部门名称 Then
            txtInfo(10).Text = txtDept
        Else
            txtInfo(10).Text = UserInfo.部门名称
        End If
        txtInfo(10).Locked = True
        txtInfo(10).BackColor = &H80000000
    
    End If
    mblnOK = False
    Me.Show vbModal
    If mblnOK Then
        GetFilter = mstrReturn
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能：条件和 选择
'==============================================================================
Private Sub optAnd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：条件或 选择
'==============================================================================
Private Sub optOr_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：条件录入控制
'==============================================================================
Private Sub txtInfo_Change(Index As Integer)
    On Error GoTo errH
    If InStr(txtInfo(Index), "'") <> 0 Then txtInfo(Index) = Replace(txtInfo(Index), "'", "")
    If InStr(txtInfo(Index), "|") <> 0 Then txtInfo(Index) = Replace(txtInfo(Index), "|", "")
    txtInfo(Index).SelStart = Len(txtInfo(Index))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：条件输入法控制
'==============================================================================
Private Sub txtInfo_GotFocus(Index As Integer)
    On Error GoTo errH
    zlControl.TxtSelAll txtInfo(Index)
    Select Case Index
        Case 0, 1, 2
            Call zlCommFun.OpenIme(False)
        Case Else
            Call zlCommFun.OpenIme(True)
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：条件按键控制
'==============================================================================
Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo errH
    If InStr("'|", Chr(KeyAscii)) <> 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    Select Case Index
        Case 0, 1, 2
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If InStr("1234567890." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：Shift回车 确定查询
'==============================================================================
Private Sub txtInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = vbKeyReturn And Shift = 2 Then
        Call CmdOK_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：控件通用Tips显示
'==============================================================================
Private Sub ShowTips(ctl As Control, str内容 As String, Optional str标题 As String = "提示信息", Optional lng时间 As Long = 2500)
    Dim X           As Single
    Dim Y           As Single
    
    On Error GoTo errH
    
    X = (ctl.Left + ctl.Width / 2) / Screen.TwipsPerPixelX
    Y = (ctl.Top + ctl.Height) / Screen.TwipsPerPixelY
    If Len(str内容) > 0 Then
        tipPopup1.Hide
        tipPopup1.StandardIcon = IDI_INFORMATION
        tipPopup1.ShowCloseButton = True
        tipPopup1.TimeOut = lng时间
        tipPopup1.Title = str标题
        tipPopup1.Text = str内容
        tipPopup1.Show Me.hWnd, X, Y
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Fill病理类型()
    Dim rs As New ADODB.Recordset
    On Error GoTo errH
    gstrSQL = "" & _
        "Select 编码,名称,简码,缺省标志 From 病理类型"
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    With cbo
        .Clear
        .AddItem ""
        .ItemData(.NewIndex) = 1
        
        If Not rs.EOF Then
            rs.MoveFirst
            Do Until rs.EOF
                .AddItem zlCommFun.NVL(rs!名称)
                 .ItemData(.NewIndex) = .NewIndex + 1

                rs.MoveNext
            Loop
        End If
        
        If .ListCount > 0 Then .ListIndex = 0
        
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function Get所属部门(ByVal lng人员Id As Long, ByVal lngMode As Long) As String
    Dim strSQL As String
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    ' lngMode =0 显示方式 lngMode =1 用于查询方式
    
    strSQL = "SELECT  distinct C.名称 AS 科室" & vbNewLine & _
                "      FROM 人员表 A,人员性质说明 B,部门表 C,部门人员 D" & vbNewLine & _
                "      WHERE A.ID=B.人员id AND C.ID=D.部门id AND D.人员id=A.ID And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
                "      AND A.id =[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng人员Id)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        Do Until rsTemp.EOF
            If lngMode = 0 Then
                If Len(strTmp) = 0 Then
                    strTmp = NVL(rsTemp!科室)
                Else
                    strTmp = strTmp & "," & NVL(rsTemp!科室)
                End If
            Else
                If Len(strTmp) = 0 Then
                    strTmp = "'" & NVL(rsTemp!科室) & "'"
                Else
                    strTmp = strTmp & ",'" & NVL(rsTemp!科室) & "'"
                End If
            End If
            
            rsTemp.MoveNext
        Loop
        
        Get所属部门 = strTmp
    Else
        Get所属部门 = UserInfo.部门名称
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Function
End Function

