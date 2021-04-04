VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlidkind.ocx"
Begin VB.Form frmTriageRoomRegNum 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdNewPati 
      Height          =   345
      Left            =   3945
      Picture         =   "frmTriageRoomRegNum.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "新增病人(F4)"
      Top             =   90
      Width           =   375
   End
   Begin VB.CommandButton cmdGetNum 
      Caption         =   "取号(&O)"
      Height          =   405
      Left            =   9990
      TabIndex        =   13
      Top             =   465
      Width           =   1065
   End
   Begin zlIDKind.CommandEx cmdExRoom 
      Height          =   285
      Left            =   9570
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   525
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin zlIDKind.CommandEx cmdExDoctor 
      Height          =   285
      Left            =   6690
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   510
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin zlIDKind.TextEx txtExDoctor 
      Height          =   360
      Left            =   4530
      TabIndex        =   8
      Top             =   495
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483645
      Appearance      =   0
      Text            =   ""
   End
   Begin zlIDKind.CommandEx cmdExDept 
      Height          =   285
      Left            =   3660
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   525
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin zlIDKind.PatiIdentify PatiIdentify 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmTriageRoomRegNum.frx":058A
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindAppearance=   0
      InputAppearance =   0
      ShowSortName    =   -1  'True
      DefaultCardType =   "0"
      IDKindWidth     =   555
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowAutoCommCard=   -1  'True
      AllowAutoICCard =   -1  'True
      AllowAutoIDCard =   -1  'True
      NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
   End
   Begin zlIDKind.TextEx txtExDept 
      Height          =   360
      Left            =   705
      TabIndex        =   5
      Top             =   495
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483645
      Appearance      =   0
      Text            =   ""
   End
   Begin zlIDKind.TextEx txtExRoom 
      Height          =   360
      Left            =   7530
      TabIndex        =   11
      Top             =   495
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483645
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.Label lblBookingNO 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "预约单:A0001"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   9750
      TabIndex        =   3
      Top             =   135
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H80000003&
      X1              =   11220
      X2              =   -30
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别：  年龄：   门诊号： 费别："
      Height          =   180
      Left            =   4380
      TabIndex        =   2
      Top             =   180
      Width           =   2880
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "科室"
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   585
      Width           =   360
   End
   Begin VB.Label lblDoctor 
      AutoSize        =   -1  'True
      Caption         =   "医生"
      Height          =   180
      Left            =   4125
      TabIndex        =   7
      Top             =   585
      Width           =   360
   End
   Begin VB.Label lblRoom 
      AutoSize        =   -1  'True
      Caption         =   "诊室"
      Height          =   180
      Left            =   7080
      TabIndex        =   10
      Top             =   585
      Width           =   360
   End
   Begin XtremeSuiteControls.ShortcutCaption srtcBack 
      Height          =   990
      Left            =   15
      TabIndex        =   14
      Top             =   0
      Width           =   11100
      _Version        =   589884
      _ExtentX        =   19579
      _ExtentY        =   1746
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmTriageRoomRegNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************************************************************************************************
'功能:取号界面(主要应用场景不存在挂号窗口，直接在分诊取号看病)
'对外接口:zlInitVar-初始化相关变量信息,程序必须首先调用
'         LockedScreen-锁屏事件(需要主界面处理相关事性)
'         GetNumSucces-取号成功事件
'内部接口:
'     CreateDeptStructure-创建部门集结构:ID,编码，名称,简码
'编制:刘兴洪
'日期:2018-01-03 11:29:56
'数据处理规则说明:
'   1.先生成挂号记录费用为0,再在医生接诊时，生成划价费用，然后收费
'   2.必须设置参数为：免挂号模式
'**********************************************************************************************************************************************
Private mlngModule As Long, mstrPrivs As String
Private mstrNo As String '单据号
Private mbytMode As Byte '0-取号;1-预约接收取号;2-回诊取号

Private mfrmMain As Object
Private mobjPati As PatiInfor
Private mstr分诊科室 As String
Private mobjCardSqure As Object
Private mrsRegData As ADODB.Recordset
Private mrsBookData As ADODB.Recordset

Private mobjRegister As clsRegist
Private mbytRegMode As Byte '1-出诊表模式;0-传统模式
Private mlngPreDeptID As Long  '上次选择的科室ID
Private mlngPreItemID As Long  '上次选择的项目ID
Private mstrPreDoctorName As String '上次选择的医生
Private mstrPreRoomName As String '上次选择的诊室
Private mrsDept As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mrsRooms As ADODB.Recordset '获取诊室集
Private mobjSetFocus As Object '当前光标移动的控件
Private mblnFirst  As Boolean
Public Event LockedScreen(ByVal blnLocked As Boolean, blnCancel As Boolean)   '锁屏操作，主要是在保存数据时，需要禁止其他操作
Public Event GetNumSucces(ByVal strNO As String)    '保存成功后刷新数据
Private Type ty_Para
    blnBusy  As Boolean ' 诊室忙时允许分诊
    int预约失效次数 As Integer  '预约失约次数
    int预约有效时间 As Integer  '预约有效时间
End Type
Private mblnNotChange As Boolean
Private mPara As ty_Para
Private Sub LoadPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载参数信息
    '入参:
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-15 15:46:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mPara
        .blnBusy = Val(zlDatabase.GetPara("诊室忙时允许分诊", glngSys, mlngModule, 0)) = 1
        .int预约失效次数 = Val(zlDatabase.GetPara("预约失约次数", glngSys, 1111, 0))
        .int预约有效时间 = Val(zlDatabase.GetPara("预约有效时间", glngSys, 1111, 0))
    End With
End Sub

Public Function zlInitVar(ByVal frmMain As Object, ByVal str分诊科室 As String, ByVal lngModule As Long, ByVal strPrivs As String, Optional ByVal objCardSqure As Object, Optional objRegister As clsRegist) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量。
    '入参:str分诊科室-分诊科室(IDs)
    '     objCardSqure-如果主界面中存在该对象，需要传入，否则会自动再创建
    '     objRegister-挂号对象
    '返回 :如果加载成功，返回true,否则返回False(返回Flase时，主窗体需要立即关闭)
    '编制:刘兴洪
    '日期:2018-01-03 11:27:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstr分诊科室 = str分诊科室: mlngModule = lngModule: mbytRegMode = 0: mbytMode = 0
    mstrPrivs = strPrivs: Set mfrmMain = frmMain
    On Error GoTo errHandle
    
    Call LoadPara   '加载参数
   
    Set mobjCardSqure = objCardSqure
    Call PatiIdentify.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, mobjCardSqure, , gstrProductName)
    Set mobjRegister = objRegister
    If mobjRegister Is Nothing Then
         If CreateRegisterObject = False Then Exit Function
         Set mobjRegister = gobjRegist
    End If
    
    cmdNewPati.ToolTipText = "新增病人(F4)"
    cmdNewPati.Visible = InStr(mstrPrivs, ";病案修改;") > 0
    lblPati.Left = cmdNewPati.Left + IIf(InStr(mstrPrivs, ";病案修改;") > 0, cmdNewPati.Width + 50, 0)
    zlInitVar = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdExDept_Click()
    If SelectDept("") = False Then
        DoEvents
        If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        zlControl.TxtSelAll txtExDept
        Exit Sub
    End If
    DoEvents
    If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
End Sub
 
Private Sub cmdExDoctor_Click()
    Dim lngItemID As Long, lngDeptID As Long, varTemp As Variant
    varTemp = Split(txtExDept.Tag & ":", ":")
    
    lngDeptID = Val(varTemp(0))
    lngItemID = Val(varTemp(1))
    
    If SelectDoctor(lngDeptID, lngItemID, "") = False Then
        DoEvents
        If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
        zlControl.TxtSelAll txtExDoctor
        Exit Sub
    End If
    
    If mbytMode = 0 Then Call LoadRoomsData   '加载诊室数据
    DoEvents
    If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
End Sub
 

Private Sub cmdExRoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If SelectRooms("") = False Then
        DoEvents
        If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
        zlControl.TxtSelAll txtExRoom
        Exit Sub
    End If
    DoEvents
    If cmdGetNum.Enabled And cmdGetNum.Visible Then cmdGetNum.SetFocus
End Sub

Private Sub cmdGetNum_Click()
    Dim lng安排ID As Long, strNO As String
    '锁屏
    If LockedScreen(True) = False Then Exit Sub
    Select Case mbytMode
    Case 0  '正常取号
        If CheckDataValied(lng安排ID) = False Then Exit Sub
        If SaveData(lng安排ID, strNO) = False Then Call LockedScreen(False): Exit Sub
        RaiseEvent GetNumSucces(strNO)
        Call LockedScreen(False)
    Case 1 '预约取号
        If SaveBooking(strNO) = False Then Call LockedScreen(False): Exit Sub
    Case 2  '回诊取号
        If SaveHzGetNum(strNO) = False Then Call LockedScreen(False): Exit Sub
    End Select
    
    
    '打印凭条
    Call PrintBill(strNO)
    '清除界面信息
    txtExDept.Text = ""
    txtExDoctor.Text = ""
    txtExRoom.Text = ""
    PatiIdentify.Text = ""
    cmdExDept.Tag = ""
    cmdExDoctor.Tag = ""
    cmdExRoom.Tag = ""
    cmdNewPati.ToolTipText = "新增病人(F4)"
    lblPati.Caption = "性别：  年龄：   门诊号： 费别："
    mstrNo = "": mbytMode = 0
    Set mobjPati = Nothing
    lblBookingNO.Visible = False
    Call LockedScreen(False)
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
End Sub


Private Function LockedScreen(ByVal blnLocked As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:锁屏或解锁
    '入参:blnLocked-true-表示锁屏;False-就是解锁
    '返回:锁屏成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-01-09 16:05:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean
    On Error GoTo errHandle
    
    txtExDept.Enabled = Not blnLocked And cmdExDept.Tag <> "F"
    txtExDoctor.Enabled = Not blnLocked And cmdExDoctor.Tag <> "F"
    txtExRoom.Enabled = Not blnLocked And cmdExRoom.Tag <> "F"
    cmdExDept.Enabled = Not blnLocked And cmdExDept.Tag <> "F"
    cmdExDoctor.Enabled = Not blnLocked And cmdExDoctor.Tag <> "F"
    cmdExRoom.Enabled = Not blnLocked And cmdExRoom.Tag <> "F"
    cmdGetNum.Enabled = Not blnLocked
    PatiIdentify.Enabled = Not blnLocked
    
    
    blnCancel = False
    RaiseEvent LockedScreen(blnLocked, blnCancel)
    LockedScreen = Not blnCancel
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdNewPati_Click()
    Dim lng病人ID As Long, lng病人ID_Out As Long
    If mobjPati Is Nothing Then
        lng病人ID = 0
    Else
        lng病人ID = mobjPati.病人ID
    End If
    If mobjRegister.zlPatiEdit(mfrmMain, lng病人ID, lng病人ID_Out) = False Then
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Exit Sub
    End If
    
    PatiIdentify.Text = "-" & lng病人ID_Out
    If GetPatient(PatiIdentify.GetCurCard, PatiIdentify.Text, False, mobjPati) = False Then
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Exit Sub
    End If
    
    cmdNewPati.ToolTipText = "修改病人(F4)"
    If SelectBooking(mobjPati.病人ID, "") = False Then
        Call ReadRegData    '读取挂号安排数据
    End If
    
    If txtExDept.Enabled And txtExDept.Visible Then
        txtExDept.SetFocus
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

 
Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If cmdGetNum.Enabled And cmdGetNum.Visible Then Call cmdGetNum_Click
        Case vbKeyF3
            If PatiIdentify.Visible = True And PatiIdentify.Enabled Then
                Call PatiIdentify.SetFocus
            End If
        Case vbKeyF4
            If cmdNewPati.Enabled And cmdNewPati.Visible Then Call cmdNewPati_Click
        Case Else
            PatiIdentify.ActiveFastKey
    End Select
End Sub
Private Sub Form_Load()
    Set mobjPati = Nothing
    mblnFirst = True
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With srtcBack
        lnTop.X1 = ScaleWidth
        lnTop.X2 = 0
        lnTop.Y1 = 0
        lnTop.Y2 = 0
        
        .Left = ScaleLeft
        .Top = ScaleTop + 15
        .Height = ScaleHeight - .Top
        .Width = ScaleWidth
        
        lblBookingNO.Left = ScaleHeight - lblBookingNO.Width - 50
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    Set mobjPati = Nothing
End Sub

Private Sub PatiIdentify_Change()
    If mblnNotChange Then Exit Sub
    Set mobjPati = Nothing
    cmdNewPati.ToolTipText = "新增病人(F4)"
    PatiIdentify.Tag = ""
    cmdExDept.Tag = ""
    cmdExDoctor.Tag = ""
    cmdExRoom.Tag = ""
    Call LockedScreen(False)
End Sub
 
 
Private Function CheckPatiCheck(ByVal lng病人ID As Long) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:黑名单检查
    '入参:lng病人ID-病人ID
    '返回:合法返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 16:26:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    If CreatePlugInOK(mlngModule) = False Then CheckPatiCheck = True: Exit Function
    On Error Resume Next
    'PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, _
    '    ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, _
    '    ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
    ''功能：检查当前病人是否是指定的特殊病人
    ''返回：true时允许继续操作，False时不允许操作
    ''参数：
    ''      lngSys,lngModual=当前调用接口的主程序系统号及模块号
    ''      lngType 操作类型：1－门诊挂号，2－住院入院，3－门诊收费，4－住院结帐。
    ''      lngPatiID-病人ID: 新建档的，为0,否则传入建档病人ID
    ''      lngPageID-主页ID: 新建档的，为0,否则传入建档主页ID(住院传入主页ID) 特殊说明：仅 lngType=4 时才传入 lngPageID，其它均传0
    ''      strPatiInforXML-病人信息:针对未建档病人传入，"姓名，性别，年龄，出生日期，医保号，身份证号"，出生日期 格式:2016-11-11 12:12:12
    ''                      固定格式：<XM></XM><XB></XB><NL></NL><CSRQ></CSRQ><YBH></YBH><SFZH></SFZH>
    ''      strReserve=保留参数,用于扩展使用
    Dim blnChecked As Boolean
    blnChecked = gobjPlugIn.PatiValiedCheck(glngSys, mlngModule, 1, lng病人ID, 0, "<YSXM>" & txtExDoctor.Text & "</YSXM>")
    
    If Err <> 0 Then
        Call zlPlugInErrH(Err, "PatiValiedCheck"): Err.Clear
        On Error GoTo 0
        CheckPatiCheck = True: Exit Function
    End If
    CheckPatiCheck = blnChecked
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim strNO As String
    
    If PatiIdentify.Tag <> "" Then blnFindPatied = True: Exit Sub
    
    blnFindPatied = False
    If GetPatient(objCard, strShowText, blnCancel, objCardData, strNO) = False Then
        DoEvents
        If Me.Enabled Then Me.SetFocus
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        blnCancel = True
        Exit Sub
    End If
    cmdNewPati.ToolTipText = "修改病人(F4)"
    strShowText = objCardData.姓名
    PatiIdentify.Tag = objCardData.病人ID
    Set mobjPati = objCardData
    
    blnFindPatied = True
    If strNO = "" Then
        If SelectBooking(objCardData.病人ID, strNO) = False Then
            Call ReadRegData    '读取挂号安排数据
        End If
    End If
    
    DoEvents
    If Me.Enabled Then Me.SetFocus
    If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
End Sub


Private Function GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, objPati As zlIDKind.PatiInfor, Optional ByRef strBookNo_out As String) As Boolean
                        
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:objCard-当前卡对象
    '     strInput-当前输入串
    '     blnCard-当前是否刷卡
    '出参:strBookNo_out-以预约单查找时，返回预约单据号
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-10 13:50:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strOtherName As String, strOtherValue As String, blnCancel As Boolean
    Dim lng病人ID As Long, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim vRect As RECT, rsTmp As ADODB.Recordset
    
    Set objPati = Nothing
    
    On Error GoTo errHandle
    
    strBookNo_out = ""
    If objCard Is Nothing Then Exit Function
    strOtherName = "": strOtherValue = "": lng病人ID = 0
    If blnCard And (objCard.名称 Like "姓名*" Or objCard.是否模糊查找) And InStr("-+*.", Left(strInput, 1)) = 0 Then     '刷卡
        
        If PatiIdentify.Cards.按缺省卡查找 And Not PatiIdentify.GetfaultCard Is Nothing Then
            lng卡类别ID = PatiIdentify.GetfaultCard.接口序号
        ElseIf PatiIdentify.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = PatiIdentify.GetCurCard.接口序号
        Else
            If lng卡类别ID = 0 Then lng卡类别ID = -1
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        
        If PatiIdentify.IsMobileNO(strInput) And lng病人ID = 0 Then
            If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        End If
        If lng病人ID <= 0 Then GoTo NotFoundPati:
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then        '门诊号
        strOtherName = "门诊号": strOtherValue = Val(Mid(strInput, 2))
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then         '病人ID
        lng病人ID = Val(Mid(strInput, 2))
    ElseIf Left(strInput, 1) = "." Then
        strBookNo_out = Mid(strInput, 2)
        strBookNo_out = GetFullNO(strBookNo_out, 12)
        PatiIdentify.Text = strBookNo_out
        If ReadBooking(strBookNo_out, True, objPati) = False Then Exit Function
        GetPatient = True: Exit Function
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                 If zlCommFun.ActualLen(strInput) <= 2 Then '小于一个汉字时，不进行过滤
                    MsgBox "输入的条件太简单，请输入2个字以上进行查找病人。", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                 End If
                 
                 strSQL = _
                     " Select /*+Rule */distinct A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位 " & _
                     " From 病人信息 A " & _
                     " Where Rownum <101 And A.停用时间 is NULL And A.姓名 Like [1]" & _
                     " Order by  姓名"
                 vRect = zlControl.GetControlRect(PatiIdentify.Hwnd)
                 Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, PatiIdentify.Height, blnCancel, False, True, strInput & "%")
                 
                 If blnCancel = True Then Exit Function
                 
                 If rsTmp Is Nothing Then GoTo NotFoundPati:
                 If rsTmp.EOF Then GoTo NotFoundPati:
                 lng病人ID = Val(Nvl(rsTmp!病人ID))
            Case "医保号"
                strInput = UCase(strInput)
                strOtherName = "医保号": strOtherValue = strInput
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                If objCard.接口序号 <> 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.接口序号, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                
                If lng病人ID = 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                If lng病人ID = 0 Then GoTo NotFoundPati:
                 
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If objCard.接口序号 <> 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.接口序号, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                If lng病人ID = 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                If lng病人ID = 0 Then GoTo NotFoundPati:
            Case "预约单"
                strBookNo_out = UCase(strInput)
                strBookNo_out = GetFullNO(strBookNo_out, 12)
                PatiIdentify.Text = strBookNo_out
                
                If ReadBooking(strBookNo_out, True, objPati) = False Then Exit Function
                GetPatient = True: Exit Function
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strOtherName = "门诊号": strOtherValue = Val(strInput)
             Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
        End Select
    End If
    
    If PatiIdentify.zlGetPatiObjectFromPatiID(lng病人ID, objPati, strErrMsg, strOtherName, strOtherValue) = False Then GoTo NotFoundPati:
    
    If objPati Is Nothing Then GoTo NotFoundPati:
    If objPati.病人ID = 0 Then GoTo NotFoundPati:
    If CheckPatiCheck(objPati.病人ID) = False Then Exit Function
    Call SetPatiColor(PatiIdentify, objPati.病人类型, IIf(objPati.险类 = 0, lblPati.ForeColor, vbRed))
    mblnNotChange = True
    PatiIdentify.Text = objPati.姓名
    mblnNotChange = False
    
    lblPati.Caption = "性别:" & objPati.性别 & Space(4)
    lblPati.Caption = lblPati.Caption & "年龄:" & objPati.年龄 & Space(4)
    lblPati.Caption = lblPati.Caption & "门诊号:" & objPati.门诊号 & Space(4)
    lblPati.Caption = lblPati.Caption & "费别:" & objPati.费别 & Space(4)
    lblPati.Caption = lblPati.Caption & "病人类型:" & objPati.病人类型 & Space(4)
    lblPati.Caption = lblPati.Caption & "付款方式:" & objPati.医疗付款方式 & Space(4)
    lblPati.Caption = lblPati.Caption & "身份证号:" & objPati.身份证号 & Space(4)
    lblPati.Caption = lblPati.Caption & IIf(objPati.险类名称 = "", "", "险类名称:" & objPati.险类名称 & Space(4))
  
    GetPatient = True
    Exit Function
NotFoundPati:
    mblnNotChange = False
    MsgBox "未找到符合条件的病人", vbInformation + vbOKOnly, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotChange = False
    Call SaveErrLog
End Function



Private Sub PatiIdentify_GotFocus()
    Call zlControl.TxtSelAll(PatiIdentify.objTxtInput)
End Sub
Private Function ReadRegData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取挂号数据
    '入参:
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-04 09:55:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str医生姓名 As String, lng医生ID As Long, lngDeptID As Long, lngItemID As Long
    Dim str科室IDs As String, strTemp As String
    Dim lngID As Long
    
    On Error GoTo errHandle
    txtExDept.Text = ""
    txtExDoctor.Text = ""
    txtExRoom.Text = ""
    
    Call CreateDeptStructure
    If mobjRegister Is Nothing Then Exit Function
    If mobjRegister.zlGetRegisterData(mrsRegData, mstr分诊科室, , False, mbytRegMode) = False Then Exit Function
     
    If mrsRegData Is Nothing Then Exit Function
    If mrsRegData.RecordCount = 0 Then Exit Function
    lngID = 1
    txtExDept.Tag = "": txtExDoctor.Tag = ""
    With mrsRegData
        .MoveFirst
        Do While Not .EOF
            lngDeptID = Val(Nvl(mrsRegData!科室ID))
            lngItemID = Val(Nvl(mrsRegData!项目ID))
            
            strTemp = lngDeptID & ":" & lngItemID
            
            If strTemp = mlngPreDeptID & ":" & mlngPreItemID Then
                If txtExDept.Text = "" Then
                    txtExDept.Text = Nvl(mrsRegData!科室编码) & "-" & Nvl(mrsRegData!科室名称) & "【" & mrsRegData!项目名称 & "】"
                    txtExDept.Tag = strTemp
                End If
                
                str医生姓名 = Nvl(mrsRegData!医生姓名): lng医生ID = Val(Nvl(mrsRegData!医生ID))
                If Nvl(!医生姓名) = mstrPreDoctorName Then
                     txtExDoctor.Text = Nvl(mrsRegData!医生姓名)
                     txtExDoctor.Tag = Val(Nvl(mrsRegData!医生ID)) & ":" & Nvl(mrsRegData!医生姓名)
                End If
            End If
            
            If InStr(str科室IDs & ",", "," & strTemp & ",") = 0 Then
                mrsDept.AddNew
                mrsDept!ID = lngID
                mrsDept!科室ID = lngDeptID
                mrsDept!编码 = CStr(Nvl(mrsRegData!科室编码))
                mrsDept!名称 = CStr(Nvl(mrsRegData!科室名称))
                mrsDept!简码 = CStr(Nvl(mrsRegData!科室简码))
                mrsDept!项目ID = lngItemID
                mrsDept!项目编码 = CStr(Nvl(mrsRegData!项目编码))
                mrsDept!项目名称 = CStr(Nvl(mrsRegData!项目名称))
                mrsDept!是否原科室 = 0
                mrsDept.Update
                str科室IDs = str科室IDs & "," & strTemp
                lngID = lngID + 1
            End If
            .MoveNext
        Loop
    End With
    
    mrsRegData.MoveFirst
    If txtExDept.Tag = "" Then
        mlngPreDeptID = Val(Nvl(mrsRegData!科室ID))
        mlngPreItemID = Val(Nvl(mrsRegData!项目ID))
        strTemp = mlngPreDeptID & ":" & mlngPreItemID
        
        txtExDept.Text = Nvl(mrsRegData!科室编码) & "-" & Nvl(mrsRegData!科室名称) & "【" & mrsRegData!项目名称 & "】"
        txtExDept.Tag = strTemp
    End If
    If txtExDoctor.Tag = "" Then
        str医生姓名 = Nvl(mrsRegData!医生姓名): lng医生ID = Val(Nvl(mrsRegData!医生ID))
        txtExDoctor.Text = str医生姓名
        txtExDoctor.Tag = lng医生ID & ":" & str医生姓名
        mstrPreDoctorName = str医生姓名
    End If
    ReadRegData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CreateDeptStructure() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结构
    '返回:初始化成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-04 18:19:09
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    
    Set mrsDept = New ADODB.Recordset
    With mrsDept
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "ID", adBigInt, , adFldIsNullable
            .Append "科室ID", adBigInt, , adFldIsNullable
            .Append "编码", adVarChar, 30, adFldIsNullable
            .Append "名称", adVarChar, 100, adFldIsNullable
            .Append "简码", adVarChar, 50, adFldIsNullable
            .Append "项目ID", adVarChar, 50, adFldIsNullable
            .Append "项目编码", adVarChar, 50, adFldIsNullable
            .Append "项目名称", adVarChar, 200, adFldIsNullable
            .Append "是否原科室", adBigInt, , adFldIsNullable
        End With
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    CreateDeptStructure = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SelectDept(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择科室
    '入参:strInput-为空时，表示查询所有的
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-05 16:35:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset, intCount As Integer
    Dim lngDeptID As Long, lngItemID As Long, str科室编码 As String, str科室名称 As String, str项目名称 As String, str项目编码 As String
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    On Error GoTo errHandle
    
    Set rsTemp = Nothing
    
    If Trim(strInput) = "" Then GoTo GoSel:
    strInput = UCase(strInput)
    
    strCompents = Replace(gstrLike, "%", "*") & strInput & "*"
    
    If mrsDept Is Nothing Then Exit Function
    If mrsDept.RecordCount = 0 Then Exit Function
    If mrsDept.RecordCount = 1 Then
        txtExDept.Text = mrsDept!编码 & "-" & mrsDept!名称 & "【" & mrsDept!项目名称 & "】"
        lngDeptID = Val(Nvl(mrsDept!科室ID))
        lngItemID = Val(Nvl(mrsDept!项目ID))
        mlngPreDeptID = lngDeptID: mlngPreItemID = lngItemID
        
        txtExDept.Tag = lngDeptID & ":" & lngItemID
        Set mrsDoctor = LoadDoctorData(lngDeptID, lngItemID, mbytMode, Nvl(mrsDept!项目编码), Nvl(mrsDept!项目名称))
        Call LoadDefaultDoctor(lngDeptID)
        SelectDept = True
        Exit Function
    End If
     
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsDept)
    '需要检查是否有多条满足条件的记录
    If IsNumeric(strInput) Then     '输入的是全数字
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strInput) Then     '输入的是全字母
        intInputType = 1
    Else
        intInputType = 2   ' 2-其他
    End If
    
    lngDeptID = 0
    With mrsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编码) = strInput Then
                    txtExDept.Text = Nvl(!编码) & "-" & Nvl(!名称) & "【" & !项目名称 & "】"
                    lngDeptID = Val(Nvl(!科室ID)): lngItemID = Val(Nvl(mrsDept!项目ID))
                    
                    txtExDept.Tag = lngDeptID & ":" & lngItemID
                    
                    Call LoadDoctorData(lngDeptID, lngItemID, , Nvl(!项目名称), Nvl(!项目编码))
                    Call LoadDefaultDoctor(lngDeptID)
                    SelectDept = True
                    Exit Function
                End If
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编码)) = Val(strInput) Then
                    If intCount = 0 Then
                        str科室编码 = Nvl(!编码): lngDeptID = Val(Nvl(!科室ID)): lngItemID = Val(Nvl(!项目ID))
                        str科室名称 = Nvl(!名称): str项目名称 = Nvl(!项目名称): str项目编码 = Nvl(!项目编码)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Val(Nvl(!编码)) Like strInput & "*" Then
                        Call zlDatabase.zlInsertCurrRowData(mrsDept, rsTemp)
                 End If
                 
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = UCase(strInput) Then
                    If intCount = 0 Then
                         str科室编码 = Nvl(!编码): lngDeptID = Val(Nvl(!科室ID)): lngItemID = Val(Nvl(!项目ID))
                        str科室名称 = Nvl(!名称):: str项目名称 = Nvl(!项目名称): str项目编码 = Nvl(!项目编码)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsDept, rsTemp)
                    intCount = intCount + 1
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strInput Or Trim(!简码) = strInput Or UCase(Trim(Nvl(!名称))) = strInput Then
                    If intCount = 0 Then
                         
                        str科室编码 = Nvl(!编码): lngDeptID = Val(Nvl(!科室ID)): lngItemID = Val(Nvl(!项目ID))
                        str科室名称 = Nvl(!名称):  str项目名称 = Nvl(!项目名称): str项目编码 = Nvl(!项目编码)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If Trim(Nvl(!编码)) Like strInput & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!名称))) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsDept, rsTemp)
                End If
            End Select
            .MoveNext
        Loop
    End With
    
    If intCount > 1 Then lngDeptID = 0
GoSel:
    If Trim(strInput) = "" Then Set rsTemp = mrsDept
    If rsTemp Is Nothing Then
        If PatiIdentify.Text = "" Then
            MsgBox "请先选择需要取号的病人", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        Else
            MsgBox "未找到符合条件的科室", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
    End If
    If lngDeptID = 0 And rsTemp.RecordCount = 1 Then
        rsTemp.MoveFirst
        str科室编码 = Nvl(rsTemp!编码): lngDeptID = Val(Nvl(rsTemp!科室ID)): lngItemID = Val(Nvl(rsTemp!项目ID))
        str科室名称 = Nvl(rsTemp!名称):: str项目名称 = Nvl(rsTemp!项目名称): str项目编码 = Nvl(rsTemp!项目编码)
    End If
    
    '直接定位
    If lngDeptID <> 0 Then
        If Trim(strInput) <> "" Then rsTemp.Close: Set rsTemp = Nothing
        txtExDept.Text = str科室编码 & "-" & str科室名称 & "【" & str项目名称 & "】"
        txtExDept.Tag = lngDeptID & ":" & lngItemID
        mlngPreDeptID = lngDeptID: mlngPreItemID = lngItemID
        
        Call LoadDoctorData(lngDeptID, lngItemID, , str项目名称, str项目编码)
        Call LoadDefaultDoctor(lngDeptID)
        SelectDept = True
        Exit Function
    End If


    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = "简码"
    Case Else
        rsTemp.Sort = "编号"
    End Select
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "未找到符合条件的科室", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

    
    '弹出选择器
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, txtExDept, rsTemp, True, "", "ID,科室ID,项目ID,是否原科室", rsReturn) = False Then Exit Function
    If Trim(strInput) <> "" Then rsTemp.Close: Set rsTemp = Nothing
    
    If rsReturn Is Nothing Then Exit Function
    If rsReturn.RecordCount = 0 Then Exit Function
    
    lngDeptID = Val(Nvl(rsReturn!科室ID)): lngItemID = Val(Nvl(rsReturn!项目ID))
    txtExDept.Text = Nvl(rsReturn!编码) & "-" & Nvl(rsReturn!名称) & "【" & rsReturn!项目名称 & "】"
    txtExDept.Tag = lngDeptID & ":" & lngItemID
    mlngPreDeptID = lngDeptID: mlngPreItemID = lngItemID
    
    Set mrsDoctor = LoadDoctorData(lngDeptID, lngItemID, , Nvl(rsReturn!项目名称), Nvl(rsReturn!项目编码))
    Call LoadDefaultDoctor(lngDeptID)
    
    rsReturn.Close: Set rsReturn = Nothing
    SelectDept = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SelectDoctor(ByVal lngDeptID As Long, ByVal lngItemID As Long, strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择医生
    '入参:lngDeptID-科室ID
    '     lngItemID-项目ID
    '     strInput-输入查找的值
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-04 18:48:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset, intCount As Integer
    Dim lng医生ID As Long, str医生姓名 As String, lng科室ID As Long, str科室编码 As String, str科室名称 As String
    Dim lng项目id As Long, str项目名称 As String, str项目编码 As String
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    
    
    On Error GoTo errHandle
    
    strInput = UCase(strInput)  '全部大写区配
    strCompents = Replace(gstrLike, "%", "*") & strInput & "*"
    
    If mrsDoctor Is Nothing Then
        Set mrsDoctor = LoadDoctorData(lngDeptID, lngItemID, mbytMode)
        If mrsDoctor Is Nothing Then Exit Function
    End If
    If mrsDoctor.State <> 1 Then
         Set mrsDoctor = LoadDoctorData(lngDeptID, lngItemID, mbytMode)
         If mrsDoctor Is Nothing Then Exit Function
    End If
    
    If mrsDoctor.RecordCount = 0 Then Exit Function
    
    If Trim(strInput) = "" Then GoTo GoSel:
    
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsDoctor)
           
            
    '需要检查是否有多条满足条件的记录
    If IsNumeric(strInput) Then     '输入的是全数字
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strInput) Then     '输入的是全字母
        intInputType = 1
    Else
        intInputType = 2   ' 2-其他
    End If
    
    With mrsDoctor
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not mrsDoctor.EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编号) = strInput Then
                    txtExDoctor.Text = Nvl(!姓名)
                    txtExDoctor.Tag = Val(Nvl(!医生ID)) & ":" & Nvl(!姓名)
                    mstrPreDoctorName = Nvl(!姓名)
                    If txtExDept.Tag = "" Then
                        txtExDept.Text = Nvl(!科室编码) & "-" & Nvl(!科室名称) & "【" & !项目名称 & "】"
                        mlngPreDeptID = Val(Nvl(!科室ID)): mlngPreItemID = Val(Nvl(!项目ID))
                        txtExDept.Tag = mlngPreDeptID & ":" & mlngPreItemID
                    End If
                    SelectDoctor = True
                    Exit Function
                End If
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编号)) = Val(strInput) Then
                    If intCount = 0 Then
                        str医生姓名 = Nvl(!姓名): lng医生ID = Val(Nvl(!医生ID))
                        lng科室ID = Val(Nvl(!科室ID)): lng项目id = Val(Nvl(!项目ID))
                         str科室编码 = Nvl(!科室编码): str科室名称 = Nvl(!科室名称)
                         str项目编码 = Nvl(!项目编码): str项目名称 = Nvl(!项目名称)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Val(Nvl(!编号)) Like strInput & "*" Then
                        Call zlDatabase.zlInsertCurrRowData(mrsDoctor, rsTemp)
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If UCase(Trim(Nvl(!简码))) = strInput Then
                    If intCount = 0 Then
                        str医生姓名 = Nvl(!姓名): lng医生ID = Val(Nvl(!医生ID))
                        lng科室ID = Val(Nvl(!科室ID)): lng项目id = Val(Nvl(!项目ID))
                        str科室编码 = Nvl(!科室编码): str科室名称 = Nvl(!科室名称)
                        str项目编码 = Nvl(!项目编码): str项目名称 = Nvl(!项目名称)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.根据参数来匹配相同数据
                If UCase(Trim(Nvl(!简码))) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsDoctor, rsTemp)
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编号) = strInput Or UCase(Trim(!简码)) = strInput Or UCase(Trim(!姓名)) = strInput Then
                    If intCount = 0 Then
                        str医生姓名 = Nvl(!姓名): lng医生ID = Val(Nvl(!医生ID))
                        lng科室ID = Val(Nvl(!科室ID)): lng项目id = Val(Nvl(!项目ID))
                        str科室编码 = Nvl(!科室编码): str科室名称 = Nvl(!科室名称)
                        str项目编码 = Nvl(!项目编码): str项目名称 = Nvl(!项目名称)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If Trim(!编号) Like strInput & "*" Or UCase(Trim(Nvl(!简码))) Like strCompents Or UCase(Trim(Nvl(!姓名))) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsDoctor, rsTemp)
                End If
            End Select
            .MoveNext
        Loop
    End With
    
    If intCount > 1 Then str医生姓名 = ""

GoSel:
    If Trim(strInput) = "" Then Set rsTemp = mrsDoctor
    If str医生姓名 = "" And rsTemp.RecordCount = 1 Then
        rsTemp.MoveFirst
        str医生姓名 = Nvl(rsTemp!姓名): lng医生ID = Val(Nvl(rsTemp!医生ID))
        lng科室ID = Val(Nvl(rsTemp!科室ID)): lng项目id = Val(Nvl(rsTemp!项目ID))
        str科室编码 = Nvl(rsTemp!科室编码): str科室名称 = Nvl(rsTemp!科室名称)
        str项目编码 = Nvl(rsTemp!项目编码): str项目名称 = Nvl(rsTemp!项目名称)
    End If
    
    '直接定位
    If str医生姓名 <> "" Then
        rsTemp.Close: Set rsTemp = Nothing
        txtExDoctor.Text = str医生姓名
        txtExDoctor.Tag = lng医生ID & ":" & str医生姓名
        mstrPreDoctorName = str医生姓名
        
        If txtExDept.Tag = "" And str科室名称 <> "" Then
            txtExDept.Text = str科室编码 & "-" & str科室名称 & "【" & str项目名称 & "】"
            txtExDept.Tag = lng科室ID & ":" & lng项目id
            mlngPreDeptID = lng科室ID: mlngPreItemID = lng项目id
        End If
        SelectDoctor = True
        Exit Function
    End If
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = "编号"
    Case 1 '输入全拼音
        rsTemp.Sort = "简码"
    Case Else
        rsTemp.Sort = "编号"
    End Select
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "未找到符合条件的医生", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, txtExDoctor, rsTemp, True, "", "ID,科室ID,医生ID", rsReturn) = False Then Exit Function
     If Trim(strInput) <> "" Then rsTemp.Close: Set rsTemp = Nothing
    
    If rsReturn Is Nothing Then Exit Function
    If rsReturn.State <> 1 Then Exit Function
    
    If rsReturn.RecordCount = 0 Then Exit Function
 
    
    txtExDoctor.Text = Nvl(rsReturn!姓名)
    txtExDoctor.Tag = Val(Nvl(rsReturn!医生ID)) & ":" & Nvl(rsReturn!姓名)
    mstrPreDoctorName = Nvl(rsReturn!姓名)
    If txtExDept.Tag = "" And Nvl(rsReturn!科室名称) <> "" Then
        txtExDept.Text = Nvl(rsReturn!科室编码) & "-" & Nvl(rsReturn!科室名称) & "【" & rsReturn!项目名称 & "】"
        txtExDept.Tag = Val(Nvl(rsReturn!科室ID)) & ":" & Val(Nvl(rsReturn!项目ID))
        mlngPreDeptID = Val(Nvl(rsReturn!科室ID)):  mlngPreItemID = Val(Nvl(rsReturn!项目ID))
    End If
    
    rsReturn.Close: Set rsReturn = Nothing
    SelectDoctor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDefaultDoctor(ByVal lngDeptID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载缺省医生
    '入参:lngDeptID-部门ID
    '返回:缺省成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 09:39:49
    '说明:
    '   缺省方式:1.只有一个,缺省这个医生;2.缺省上一个选择医生;3.缺省第一个医生(按编号)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng医生ID As Long, str医生姓名 As String
    On Error GoTo errHandle
    txtExDoctor.Text = ""
    txtExDoctor.Tag = ""
    If mrsDoctor Is Nothing Then Exit Function
    If mrsDoctor Is Nothing Then Exit Function
    If mrsDoctor.RecordCount = 0 Then Exit Function
    If mrsDoctor.RecordCount = 1 Then
        txtExDoctor.Text = Nvl(mrsDoctor!姓名)
        txtExDoctor.Tag = Val(Nvl(mrsDoctor!医生ID)) & ":" & Nvl(mrsDoctor!姓名)
        mstrPreDoctorName = Nvl(mrsDoctor!姓名)
        LoadDefaultDoctor = True: Exit Function
    End If
    If mstrPreDoctorName <> "" Then
        mrsDoctor.Filter = "姓名='" & mstrPreDoctorName & "'"
        If mrsDoctor.RecordCount <> 0 Then
            txtExDoctor.Text = Nvl(mrsDoctor!姓名)
            txtExDoctor.Tag = Val(Nvl(mrsDoctor!医生ID)) & ":" & Nvl(mrsDoctor!姓名)
            mstrPreDoctorName = Nvl(mrsDoctor!姓名)
            mrsDoctor.Filter = 0
            LoadDefaultDoctor = True: Exit Function
        End If
    End If
    
    mrsDoctor.Filter = 0: mrsDoctor.Sort = "编号": mrsDoctor.MoveFirst '缺省第一个
    txtExDoctor.Text = Nvl(mrsDoctor!姓名)
    txtExDoctor.Tag = Val(Nvl(mrsDoctor!医生ID)) & ":" & Nvl(mrsDoctor!姓名)
    mstrPreDoctorName = Nvl(mrsDoctor!姓名)
    
    LoadDefaultDoctor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDoctorData(ByVal lngDeptID As Long, ByVal lngItemID As Long, _
    Optional bytMode As Byte, Optional str项目名称 As String, Optional str项目编码 As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载医生数据集
    '入参:lngDeptID-部门ID,0时，表示查所有安排医生
    '     lngItemID-项目
    '     bytMode-0-普通;1-预约取号;2-回诊取号
    '返回:返回医生集
    '编制:刘兴洪
    '日期:2018-01-04 18:19:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsDeptDoctor As ADODB.Recordset, strDoctors As String, strTemp As String, i As Long
    Dim blnLoadDoctorFromDept As Boolean '是否加载缺省的科室医生,以安排中是否有只安排到部门的，有，就加载，无就以医生为准
    
    
    On Error GoTo errHandle
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "ID", adBigInt, , adFldIsNullable
            .Append "医生ID", adBigInt, , adFldIsNullable
            .Append "编号", adVarChar, 20, adFldIsNullable
            .Append "姓名", adVarChar, 100, adFldIsNullable
            .Append "简码", adVarChar, 50, adFldIsNullable
            .Append "科室ID", adBigInt, , adFldIsNullable
            .Append "科室编码", adVarChar, 50, adFldIsNullable
            .Append "科室名称", adVarChar, 100, adFldIsNullable
            .Append "项目ID", adBigInt, , adFldIsNullable
            .Append "项目编码", adVarChar, 50, adFldIsNullable
            .Append "项目名称", adVarChar, 200, adFldIsNullable
        End With
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    If bytMode = 1 Then blnLoadDoctorFromDept = True: GoTo gotoDoctor: '预约模式，只能取部门所对应的医生
    
    
    If mrsRegData Is Nothing Then
        If ReadRegData = False Then Exit Function
    End If
    
    blnLoadDoctorFromDept = False
    mrsRegData.Filter = IIf(lngDeptID = 0, "", "科室ID=" & lngDeptID & " And 项目ID=" & lngItemID)
    strDoctors = "": i = 1
    
    With mrsRegData
        Do While Not .EOF
           
            strTemp = Val(Nvl(mrsRegData!医生ID)) & ":" & Nvl(!医生姓名)
            If lngDeptID <> 0 And Val(Nvl(!医生ID)) = 0 And Nvl(!医生姓名) = "" Then blnLoadDoctorFromDept = True ' 存在只安排到科室的号,这种情况需要将该科室的的医生全部显示出来，供选择
            
            If InStr(strDoctors & ",", "," & strTemp & ",") = 0 And Nvl(mrsRegData!医生姓名) <> "" Then
                rsTemp.AddNew
                rsTemp!ID = i
                rsTemp!医生ID = Val(Nvl(mrsRegData!医生ID))
                rsTemp!编号 = CStr(Nvl(mrsRegData!医生编号))
                rsTemp!姓名 = CStr(Nvl(mrsRegData!医生姓名))
                If Val(Nvl(mrsRegData!医生ID)) = 0 Then
                    rsTemp!简码 = zlCommFun.SpellCode(CStr(Nvl(mrsRegData!医生姓名)))
                Else
                    rsTemp!简码 = CStr(Nvl(mrsRegData!医生简码))
                End If
                rsTemp!科室ID = Val(Nvl(mrsRegData!科室ID))
                rsTemp!科室编码 = CStr(Nvl(mrsRegData!科室编码))
                rsTemp!科室名称 = CStr(Nvl(mrsRegData!科室名称))
                rsTemp!项目ID = Val(Nvl(mrsRegData!项目ID))
                rsTemp!项目编码 = CStr(Nvl(mrsRegData!项目编码))
                rsTemp!项目名称 = CStr(Nvl(mrsRegData!项目名称))
                rsTemp.Update
                i = i + 1
                strDoctors = strDoctors & "," & strTemp
            End If
            
            .MoveNext
        Loop
    End With
     mrsRegData.Filter = 0
gotoDoctor:
    
    If blnLoadDoctorFromDept Or bytMode = 2 Then
       '无医生
       rsTemp.AddNew
       rsTemp!ID = i
       rsTemp!医生ID = 0
       rsTemp!姓名 = ""
       rsTemp.Update
       i = i + 1
       If mobjRegister.zlGetDoctorFromDeptID(lngDeptID, rsDeptDoctor) Then  '根据部门ID来获取所涉及的医生集
            With rsDeptDoctor
                Do While Not .EOF
                    strTemp = Val(Nvl(!ID)) & ":" & Nvl(!姓名)
                    If InStr(strDoctors & ",", "," & strTemp & ",") = 0 And Nvl(!姓名) <> "" Then
                        rsTemp.AddNew
                        rsTemp!ID = i
                        rsTemp!医生ID = Val(Nvl(!ID))
                        rsTemp!编号 = CStr(Nvl(!编号))
                        rsTemp!姓名 = CStr(Nvl(!姓名))
                        If Val(Nvl(!ID)) = 0 Then
                            rsTemp!简码 = zlCommFun.SpellCode(CStr(Nvl(!姓名)))
                        Else
                            rsTemp!简码 = CStr(Nvl(!简码))
                        End If
                        rsTemp!科室ID = Val(Nvl(!科室ID))
                        rsTemp!科室编码 = CStr(Nvl(!科室编码))
                        rsTemp!科室名称 = CStr(Nvl(!科室名称))
                        rsTemp!项目ID = lngItemID
                        rsTemp!项目编码 = str项目编码
                        rsTemp!项目名称 = str项目名称
                        i = i + 1
                        rsTemp.Update
                        strDoctors = strDoctors & "," & strTemp
                    End If
                    .MoveNext
                Loop
            End With
       End If
    End If
    Set LoadDoctorData = rsTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Set LoadDoctorData = rsTemp
End Function

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    If txtExDept.Enabled And txtExDept.Visible Then
        txtExDept.SetFocus
    Else
        
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtExDept_Change()
    txtExDept.Tag = ""
    txtExDoctor.Text = ""
    Set mrsRooms = Nothing
    
End Sub
Private Sub txtExDept_GotFocus()
    zlControl.TxtSelAll txtExDept
End Sub

Private Sub txtExDept_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtExDept, KeyAscii, m文本式)
    If KeyAscii <> 13 Then Exit Sub
    If txtExDept.Tag = "" Then
        If SelectDept(Trim(txtExDept.Text)) = False Then
            DoEvents
            If Me.Enabled Then Me.SetFocus
            If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
            zlControl.TxtSelAll txtExDept
            Exit Sub
        End If
    End If
    DoEvents
    If Me.Enabled Then Me.SetFocus
    If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
End Sub
Private Sub txtExDept_LostFocus()
    zlCommFun.OpenIme False
End Sub
Private Sub txtExDoctor_Change()
    txtExDoctor.Tag = ""
    Set mrsRooms = Nothing
    txtExRoom.Text = ""
End Sub

Private Sub txtExDoctor_GotFocus()
    zlControl.TxtSelAll txtExDoctor
End Sub

Private Sub txtExDoctor_KeyPress(KeyAscii As Integer)
    Dim lngItemID As Long, lngDeptID As Long, varTemp As Variant
    
    Call zlControl.TxtCheckKeyPress(txtExDoctor, KeyAscii, m文本式)
    If KeyAscii <> 13 Then Exit Sub
    
    If txtExDoctor.Tag <> "" Then
        If txtExRoom.Enabled And txtExRoom.Visible Then
            txtExRoom.SetFocus: Exit Sub
        Else
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
    End If
    varTemp = Split(txtExDept.Tag & ":", ":")
    lngDeptID = Val(varTemp(0))
    lngItemID = Val(varTemp(1))
        
    If SelectDoctor(lngDeptID, lngItemID, Trim(txtExDoctor.Text)) = False Then
        DoEvents
        If Me.Enabled Then Me.SetFocus
        If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
        zlControl.TxtSelAll txtExDoctor
        Exit Sub
    End If
    
    If mbytMode = 0 Then Call LoadRoomsData  '加载诊室数据
    DoEvents
    If Me.Enabled Then Me.SetFocus
    If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
End Sub

Private Sub txtExDoctor_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txtExRoom_Change()
    txtExRoom.Tag = ""
End Sub

Private Sub txtExRoom_GotFocus()
    zlControl.TxtSelAll txtExRoom
End Sub

Private Sub txtExRoom_KeyPress(KeyAscii As Integer)

    Call zlControl.TxtCheckKeyPress(txtExRoom, KeyAscii, m文本式)
    If KeyAscii <> 13 Then Exit Sub
    If txtExRoom.Tag = "" And Trim(txtExRoom.Text) <> "" Then
        If SelectRooms(Trim(txtExRoom.Text)) = False Then
            DoEvents
            If Me.Enabled Then Me.SetFocus
            Set mobjSetFocus = txtExRoom
            If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
            zlControl.TxtSelAll txtExRoom
            Exit Sub
        End If
    End If
    DoEvents
    If Me.Enabled Then Me.SetFocus
    Set mobjSetFocus = cmdGetNum
    If cmdGetNum.Enabled And cmdGetNum.Visible Then cmdGetNum.SetFocus
End Sub

Private Sub txtExRoom_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Function GetRegisterPlanID(ByRef lng安排Id_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据选择的部门，医生来获取具体安排的ID
    '入参:
    '出参:lng安排ID_Out-安排ID(新版本为记录ID)
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 14:18:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngDoctorID As Long, str医生姓名 As String, varData As Variant
    Dim lngItemID As Long, lngDeptID As Long, varTemp As Variant
    Dim lng安排ID As Long, lng计划ID As Long
    
    On Error GoTo errHandle
    If mbytMode = 1 Then
      If mrsBookData Is Nothing Then Exit Function
      If mrsBookData.State <> 1 Then Exit Function
      If mrsBookData.RecordCount = 0 Then Exit Function
      lng安排Id_Out = Val(Nvl(mrsBookData!出诊记录ID))
      If lng安排Id_Out <> 0 Then GetRegisterPlanID = True: Exit Function
      
       GetRegisterPlanID = mobjRegister.zlGetRegisterPlanID_Tradition(Nvl(mrsBookData!号别), lng安排Id_Out, lng计划ID)
       Exit Function
    End If
        
     
     
    If mrsRegData Is Nothing Then Exit Function
    
    varTemp = Split(txtExDept.Tag & ":", ":")
    
    lngDeptID = Val(varTemp(0))
    lngItemID = Val(varTemp(1))
     
    If txtExDept.Tag = "" Then Exit Function
    
    varData = Split(txtExDoctor.Tag & ":", ":")
    lngDoctorID = Val(varData(0))
    str医生姓名 = varData(1)
    mrsRegData.Filter = "科室ID=" & lngDeptID & " And 项目ID=" & lngItemID
    
    With mrsRegData
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Val(Nvl(!医生ID)) = lngDoctorID And Nvl(!医生姓名) = str医生姓名 Then
                lng安排Id_Out = Val(Nvl(mrsRegData!ID)): Exit Do
            End If
            If Val(Nvl(!医生ID)) = 0 And Nvl(!医生姓名) = "" Then lng安排ID = Val(Nvl(mrsRegData!ID))
            .MoveNext
        Loop
    End With
    
    mrsRegData.Filter = 0
    If lng安排Id_Out = 0 Then lng安排Id_Out = lng安排ID
    If lng安排Id_Out = 0 Then Exit Function
    GetRegisterPlanID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadRoomsData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载诊室信息集
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 14:11:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng安排ID As Long
    Dim lngItemID As Long, lngDeptID As Long, varTemp As Variant
    varTemp = Split(txtExDept.Tag & ":", ":")
    
    lngDeptID = Val(varTemp(0))
    lngItemID = Val(varTemp(1))
    
    On Error GoTo errHandle
    If GetRegisterPlanID(lng安排ID) = False Then
         If mobjRegister.zlGetRegRoomsFromDeptid(lngDeptID, lngItemID, Trim(txtExDoctor.Text), mrsRooms) = False Then Set mrsRooms = Nothing: Exit Function
    Else
        If mobjRegister.zlGetRegRoomsFromPlanID(lng安排ID, mrsRooms, mPara.blnBusy) = False Then Set mrsRooms = Nothing: Exit Function
    End If
    LoadRoomsData = True
    If mrsRooms Is Nothing Then Exit Function
    If mrsRooms.RecordCount = 0 Then Exit Function
    
    '设置缺省值
    txtExRoom.Text = mrsRooms!名称
    txtExRoom.Tag = mrsRooms!名称
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SelectRooms(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择诊室
    '入参:strInput-为空时，表示查询所有的
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-05 16:35:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngRoomID As Long, str编码 As String, str名称 As String, intCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    On Error GoTo errHandle
    
    If mrsRooms Is Nothing Then Call LoadRoomsData  '加载诊室集
    
    strInput = UCase(Trim(strInput))
    If Trim(strInput) = "" Then GoTo GoSel:
    
    strCompents = Replace(gstrLike, "%", "*") & strInput & "*"
    
    If mrsRooms Is Nothing Then
        Call LoadRoomsData
        If mrsRooms Is Nothing Then
            MsgBox "未找到符合条件的诊室", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    If mrsRooms.State <> 1 Then
         Call LoadRoomsData
        If mrsRooms Is Nothing Then
            MsgBox "未找到符合条件的诊室", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If mrsRooms.RecordCount = 0 Then
        MsgBox "未找到符合条件的诊室", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'b.ID, b.编码, b.名称,b.简码, b.位置
    If mrsRooms.RecordCount = 1 Then
        lngRoomID = Val(Nvl(mrsRooms!ID))
        txtExRoom.Text = mrsRooms!名称
        txtExRoom.Tag = mrsRooms!名称
        SelectRooms = True
        Exit Function
    End If
     
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsRooms)
    '需要检查是否有多条满足条件的记录
    If IsNumeric(strInput) Then     '输入的是全数字
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strInput) Then     '输入的是全字母
        intInputType = 1
    Else
        intInputType = 2   ' 2-其他
    End If
    
    lngRoomID = 0
    With mrsRooms
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编码) = strInput Then
                    lngRoomID = Val(Nvl(!ID))
                    txtExRoom.Text = Nvl(!名称)
                    txtExRoom.Tag = Nvl(!名称)
                    SelectRooms = True
                    Exit Function
                End If
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编码)) = Val(strInput) Then
                    If intCount = 0 Then
                        str编码 = Nvl(!编码): lngRoomID = Val(Nvl(!ID))
                        str名称 = Nvl(!名称)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Val(Nvl(!编码)) Like strInput & "*" Then
                        Call zlDatabase.zlInsertCurrRowData(mrsRooms, rsTemp)
                 End If
                 
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strInput Then
                    If intCount = 0 Then
                         str编码 = Nvl(!编码): lngRoomID = Val(Nvl(!ID))
                        str名称 = Nvl(!名称)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.根据参数来匹配相同数据
                If UCase(Trim(Nvl(!简码))) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsRooms, rsTemp)
                    intCount = intCount + 1
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strInput Or UCase(Trim(!简码)) = strInput Or UCase(Trim(!名称)) = strInput Then
                    If intCount = 0 Then
                        str编码 = Nvl(!编码): lngRoomID = Val(Nvl(!ID))
                        str名称 = Nvl(!名称)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If Trim(Nvl(!编码)) Like strInput & "*" Or Trim(Nvl(!简码)) Like strCompents Or Trim(Nvl(!名称)) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsRooms, rsTemp)
                    intCount = intCount + 1
                End If
            End Select
            .MoveNext
        Loop
    End With
    
    If intCount > 1 Then lngRoomID = 0
GoSel:
    If Trim(strInput) = "" Then Set rsTemp = mrsRooms
    If rsTemp Is Nothing Then
        MsgBox "未找到符合条件的诊室", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsTemp.State <> 1 Then
         MsgBox "未找到符合条件的诊室", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If lngRoomID = 0 And rsTemp.RecordCount = 1 Then
        rsTemp.MoveFirst
        str编码 = Nvl(rsTemp!编码): lngRoomID = Val(Nvl(rsTemp!ID))
        str名称 = Nvl(rsTemp!名称)
    End If
    
    '直接定位
    If lngRoomID <> 0 Then
        If Trim(strInput) <> "" Then rsTemp.Close: Set rsTemp = Nothing
        txtExRoom.Text = str名称
        txtExRoom.Tag = str名称
        SelectRooms = True
        Exit Function
    End If
    

    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = "简码"
    Case Else
        rsTemp.Sort = "编码"
    End Select
    If rsTemp.RecordCount = 0 Then
        MsgBox "未找到符合条件的诊室", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, txtExRoom, rsTemp, True, "", "ID", rsReturn) = False Then Exit Function
    If Trim(strInput) <> "" Then rsTemp.Close: Set rsTemp = Nothing
    
    If rsReturn Is Nothing Then Exit Function
    If rsReturn.RecordCount = 0 Then Exit Function
    
    lngRoomID = Val(Nvl(rsReturn!ID))
    txtExRoom.Text = Nvl(rsReturn!名称)
    txtExRoom.Tag = Nvl(rsReturn!名称)
    rsReturn.Close: Set rsReturn = Nothing
    SelectRooms = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckDataValied(ByRef lng安排Id_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入数据的合法性
    '入参:
    '出参:lng安排Id_Out-返回当前安排ID
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 14:54:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnExist As Boolean
    Dim strErrMsg As String
    
    lng安排Id_Out = 0
    
    On Error GoTo errHandle
    If mobjPati Is Nothing Then
       MsgBox "未选择病人，不能取号!", vbInformation + vbOKOnly, gstrSysName
        Call LockedScreen(False)
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Exit Function
    End If
    If mobjPati.病人ID = 0 Or PatiIdentify.Text = "" Then
        MsgBox "未选择病人，不能取号!", vbInformation + vbOKOnly, gstrSysName
        Call LockedScreen(False)
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Exit Function
    End If
    
    If txtExDept.Tag = "" Then
        MsgBox "未选的需要取号的科室，不能取号!", vbInformation + vbOKOnly, gstrSysName
        Call LockedScreen(False)
        If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        Exit Function
    End If
    If txtExDoctor.Tag = "" And txtExDoctor.Text <> "" Then
        MsgBox "医生选择错误,请选择正确的医生!", vbInformation + vbOKOnly, gstrSysName
        Call LockedScreen(False)
        If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
        Exit Function
    End If
    If txtExRoom.Tag = "" And txtExRoom.Text <> "" Then
        MsgBox "诊室选择错误,请选择正确的诊室!", vbInformation + vbOKOnly, gstrSysName
        Call LockedScreen(False)
        If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
        Exit Function
    End If
    
    If Not mrsRooms Is Nothing Then
        If mrsRooms.RecordCount <> 0 And txtExDept.Tag = "" Then
            MsgBox "你还未选择诊室,不允许取号！!", vbInformation + vbOKOnly, gstrSysName
            Call LockedScreen(False)
            If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
            Exit Function
        End If
    End If

    '黑名单检查
    If CheckPatiCheck(mobjPati.病人ID) = False Then
        Call LockedScreen(False)
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Exit Function
    End If
    
    If mbytMode = 0 Then
        '检查是否存在安排
        blnExist = GetRegisterPlanID(lng安排Id_Out)
    End If
    
    
    If blnExist Then blnExist = lng安排Id_Out <> 0
    If Not blnExist Then
        If txtExDoctor.Tag <> "" Then
            MsgBox "未找到科室为" & txtExDept.Text & "且医生为" & txtExDoctor.Text & " 的安排，不能取号!", vbInformation + vbOKOnly, gstrSysName
            Call LockedScreen(False)
            If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        Else
            MsgBox "未找到科室为" & txtExDept.Text & "的安排，不能取号!", vbInformation + vbOKOnly, gstrSysName
             Call LockedScreen(False)
            If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        End If
        Exit Function
    End If
    
    '检查是否超号
    If mobjRegister.zlRegisterCheckValied(mobjPati.病人ID, lng安排Id_Out, strErrMsg) = False Then
        If strErrMsg <> "" Then
            ShowMsgbox txtExDept.Text & " " & strErrMsg & ",请选择其他科室就诊!"
            Call LockedScreen(False)
            DoEvents
            If Me.Enabled Then Me.SetFocus
            If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
            Exit Function
        End If
    End If
    
    CheckDataValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call LockedScreen(False)
End Function

Public Function SaveHzGetNum(ByRef strNo_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存回诊取号
    '返回:取号成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-16 16:17:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng挂号ID As Long, lng科室ID As Long, lng项目id As Long, str诊室 As String, str医生 As String, lng医生ID As Long
    Dim blnYes As Boolean, strSQL As String, varTemp As Variant
        
    On Error GoTo errHandle
    
    If mrsBookData Is Nothing Then Exit Function
    If mrsBookData.State <> 1 Then Exit Function
    If mrsBookData.RecordCount = 0 Then Exit Function
    mrsBookData.MoveFirst
    lng挂号ID = Val(Nvl(mrsBookData!挂号ID))
    strNo_Out = Nvl(mrsBookData!NO)
    
    If Val(Nvl(mrsBookData!记录标志)) <> 2 Then
        MsgBox "挂号单为" & mrsBookData!NO & "不是回诊单据，不能取号!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    varTemp = Split(txtExDept.Tag & ":", ":")
    lng科室ID = Val(varTemp(0))
    lng项目id = Val(varTemp(1))
     
    str诊室 = txtExRoom.Text
    str医生 = txtExDoctor.Text
    lng医生ID = Val(Split(txtExDoctor.Tag & ":", ":")(0))
    
    If lng科室ID = -1 Then
        MsgBox "请确定要回诊的科室。", vbInformation, gstrSysName
        If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        Exit Function
    End If
    If lng科室ID <> mrsBookData!科室ID Then
        If MsgBox("注意:" & vbCrLf & "  你选择的科室与回诊的科室不一致,你是否要调整病人回诊科室?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        Exit Function
        End If
        blnYes = True
    End If
    If str诊室 <> Nvl(mrsBookData!诊室) And blnYes = False Then
        If MsgBox("注意:" & vbCrLf & "  你选择的诊室与回诊的诊室不一致,你是否要调整病人的回诊诊室?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
            Exit Function
        End If
        blnYes = True
    End If
    
    If str医生 <> Nvl(mrsBookData!医生姓名) And blnYes = False Then
        If MsgBox("注意:" & vbCrLf & "  你选择的医生与回诊的医生不一致,你是否要调整病人的回诊医生?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
            Exit Function
        End If
        blnYes = True
    End If
    
    'Zl_病人挂号记录_回诊
    strSQL = "Zl_病人挂号记录_回诊("
    '  Id_In         病人挂号记录.ID%Type,
    strSQL = strSQL & "" & lng挂号ID & ","
    '  新执行科室_In 病人挂号记录.执行部门id%Type,
    strSQL = strSQL & "" & lng科室ID & ","
    '  新诊室_In     病人挂号记录.诊室%Type,
    strSQL = strSQL & "'" & str诊室 & "',"
    '  新医生_In     病人挂号记录.执行人%Type,
    strSQL = strSQL & "'" & str医生 & "',"
    '  需回诊_In Integer:=0
    strSQL = strSQL & "0)"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveHzGetNum = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
 End Function

Private Function SaveBooking(ByRef strNo_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:预约接收取号
    '入参:
    '出参:strNo_Out-返回取号的单据号
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-15 17:31:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng挂号ID As Long
    On Error GoTo errHandle
    If mrsBookData Is Nothing Then Exit Function
    If mrsBookData.State <> 1 Then Exit Function
    If mrsBookData.RecordCount = 0 Then Exit Function
    
    mrsBookData.MoveFirst
    strNo_Out = Nvl(mrsBookData!NO)
    If strNo_Out = "" Then Exit Function
    If mrsBookData!记录状态 = 1 Then
        '付完款的取号，直接采用签道方式
        lng挂号ID = Val(Nvl(mrsBookData!挂号ID))
 
        ' Zl_病人挂号记录_签到
        strSQL = "Zl_病人挂号记录_签到("
        '  Id_In       病人挂号记录.Id%Type,
        strSQL = strSQL & "" & lng挂号ID & ","
        '  操作类型_In Integer := 0,
        strSQL = strSQL & "" & 0 & ","
        '  预约方式_In 预约方式.名称%Type := Null,
        strSQL = strSQL & "NULL,"
        '  诊室_In     病人挂号记录.诊室%Type := Null,
        strSQL = strSQL & "'" & txtExRoom.Text & "',"
        '  医生_In     病人挂号记录.执行人%Type := Null
        strSQL = strSQL & "'" & txtExDoctor.Text & "')"
 
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        SaveBooking = True: Exit Function
    End If
    '    Zl_分诊预约接收_取号
    strSQL = "Zl_分诊预约接收_取号("
    '  No_In         门诊费用记录.No%Type
    strSQL = strSQL & "'" & strNo_Out & "',"
    '  诊室_In       门诊费用记录.发药窗口%Type,
    strSQL = strSQL & "'" & txtExRoom.Text & "',"
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & mobjPati.病人ID & ","
    '  医生姓名_In   门诊费用记录.执行人 %Type,
    strSQL = strSQL & "'" & txtExDoctor.Text & "',"
    '  操作员编号_In 门诊费用记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  登记时间_In   门诊费用记录.登记时间%Type := Null,
    strSQL = strSQL & "sysdate,"
    '  摘要_In       病人挂号记录.摘要%Type := Null,
    strSQL = strSQL & "NULL,"
    '  险类_In       病人挂号记录.险类%Type := Null
    strSQL = strSQL & "NULL)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveBooking = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function SaveData(ByVal lng安排ID As Long, ByRef strNo_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存(只是存划挂号费0的记录)
    '入参:lng安排ID-当前的安排ID
    '出参:strNo_out-保存成功后，返回的单据号
    '返回:保存成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-08 16:38:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strNO As String
    
    
    strNO = zlDatabase.GetNextNo(12)
    
    On Error GoTo errHandle
    
    'Zl_门诊分诊取号_Insert
    strSQL = "Zl_门诊分诊取号_Insert("
    '  病人id_In     病人信息.病人id%Type,
    strSQL = strSQL & "" & mobjPati.病人ID & ","
    '  记录id_In     临床出诊记录.Id%Type,
    strSQL = strSQL & "" & IIf(mbytRegMode = 1, lng安排ID, "NULL") & ","
    '  安排id_In     挂号安排.Id%Type,
    strSQL = strSQL & "" & IIf(mbytRegMode <> 1, lng安排ID, "NULL") & ","
    '  单据号_In     病人挂号记录.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  诊室_In       门诊诊室.名称%Type,
    strSQL = strSQL & "" & IIf(Trim(txtExRoom.Text) = "", "NULL", "'" & Trim(txtExRoom.Text) & "'") & ","
    '  医生姓名_In   挂号安排.医生姓名%Type,
    strSQL = strSQL & "" & IIf(Trim(txtExDoctor.Text) = "", "NULL", "'" & Trim(txtExDoctor.Text) & "'") & ","
    '  医生id_In     挂号安排.医生id%Type,
    strSQL = strSQL & "" & IIf(Trim(txtExDoctor.Tag) = "", "NULL", "'" & Val(Split(txtExDoctor.Tag & ":", ":")(0)) & "'") & ","
    '  开单部门id_In 门诊费用记录.开单部门id%Type,
    strSQL = strSQL & "" & UserInfo.部门ID & ","
    '  操作员编号_In 门诊费用记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  退号重用_In   Integer := 0,
    strSQL = strSQL & "0,"
    '  站点_In Varchar2:=Null
    strSQL = strSQL & "" & IIf(Trim(gstrNodeNo) = "", "NULL", "'" & Trim(gstrNodeNo) & "'") & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    strNo_Out = strNO
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function PrintBill(ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印单据
    '入参:strNo-挂号单号
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-09 15:39:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    Select Case Val(zlDatabase.GetPara("挂号凭条打印方式", glngSys, 9000, "0"))
        Case 0    '不打印
           Exit Function
        Case 1    '自动打印
        Case 2    '选择打印
            If MsgBox("要打印取号凭条吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    End Select
    strSQL = "select ID From 病人挂号记录 where NO =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then
        MsgBox "未找到单据号为" & strNO & "的取号记录,请检查"
    End If
    
    '暂定为可以重复打印
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1113", Me, "挂号ID=" & Val(Nvl(rsTemp!ID)), "发票号=无", 2)
    PrintBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SelectBooking(ByVal lng病人ID As Long, ByRef strNo_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择预约单据
    '入参:lng病人ID-指定病人ID
    '出参:strNO_out-返回预约具体单据号
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-15 16:29:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmNew As frmSelRegist, rsTemp As ADODB.Recordset, strNO As String
    Dim lng病人ID1 As Long
    
    On Error GoTo errHandle
    Set mrsBookData = Nothing
    
    lblBookingNO.Visible = False    '隐藏预约单
    
    'bytType-0-仅包含预约单据;1-仅包含已经支付但没在签道的预约单;2-仅包含回诊病人;3-包含(0,1,2)
    If mobjRegister.zlGetRegisterBookData(lng病人ID, rsTemp, , mstr分诊科室, 3) = False Then Exit Function
    
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.EOF Then Exit Function
    
    strNo_Out = Nvl(rsTemp!NO)
    lng病人ID1 = Val(Nvl(rsTemp!病人ID))
    
    If rsTemp.RecordCount > 1 Then
        '先找回诊病人
        rsTemp.Filter = "记录标志=2"
        If rsTemp.EOF = False Then strNo_Out = Nvl(rsTemp!NO): GoTo LoadPati:
        rsTemp.Filter = "记录状态<>0 "  '先取已经付费的
        If rsTemp.EOF = False Then strNo_Out = Nvl(rsTemp!NO): GoTo LoadPati:
        
         '检查该病人是否有预约单据
        rsTemp.Filter = 0
        Set frmNew = New frmSelRegist
        If frmNew.ShowRegist(Me, mstrPrivs, False, mPara.int预约失效次数, strNo_Out, rsTemp, lng病人ID, 1, mstr分诊科室) = False Then
            If Not frmNew Is Nothing Then Unload frmNew
            Set frmNew = Nothing
            Exit Function
        End If
        If Not frmNew Is Nothing Then Unload frmNew
        Set frmNew = Nothing
        lng病人ID1 = Val(Nvl(rsTemp!病人ID))
    End If
    
LoadPati:
    If lng病人ID <> lng病人ID1 Then
        '加载病人信息
        If GetPatient(PatiIdentify.GetCurCard, "-" & Val(Nvl(rsTemp!病人ID)), False, mobjPati) = False Then Exit Function
        cmdNewPati.ToolTipText = "修改病人(F4)"
    End If
    
    '加载预约单
    If ReadBooking(strNo_Out) = False Then Exit Function

    SelectBooking = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function ReadBooking(ByVal strNO As String, Optional blnReadPati As Boolean, Optional objPati As PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取预约单据
    '入参: strNo-预约单据,为空时，表示根据病人ID来检查预约单据,否则根据预约单来查找预约单据
    '     blnReadPati-是否需要根据挂号单中的病人信息重新读取病信息
    '出参:objPati-返回病人信息,blnReadPati为true时，返回,否则为Nothing
    '返回:获取成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-15 15:39:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean, dtDate As Date
    On Error GoTo errHandle
    
    
    If mobjRegister.zlGetRegisterBookData(0, mrsBookData, strNO, , 3) = False Then Exit Function
    
    If mrsBookData Is Nothing Then
        If strNO <> "" Then MsgBox "预约单:" & strNO & "不存在,可能被其他人接收或无效的预约单号!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mrsBookData.EOF Then
        If strNO <> "" Then MsgBox "预约单:" & strNO & "不存在,可能被其他人接收或无效的预约单号!", vbInformation + vbOKOnly, gstrSysName
        Set mrsBookData = Nothing
        Exit Function
    End If
    
    
    mbytMode = IIf(Val(Nvl(mrsBookData!记录标志)) = 2, 2, 1)
       
    If InStr("," & mstr分诊科室 & ",", "," & Val(Nvl(mrsBookData!科室ID)) & ",") = 0 And mstr分诊科室 <> "" Then
        MsgBox strNO & "的" & IIf(mbytMode <> 2, "预约", "回诊") & "单据不能在本分诊台取号，请检查！", vbInformation + vbOKOnly, gstrSysName
        Set mrsBookData = Nothing
        mbytMode = 0
        Exit Function
    End If
    
    If mPara.int预约有效时间 <> 0 And mbytMode <> 2 Then
        dtDate = DateAdd("n", 1 * mPara.int预约有效时间, zlDatabase.Currentdate)
        If Format(dtDate, "yyyy-MM-dd hh:mm:ss") > Format(mrsBookData!预约时间, "yyyy-MM-dd hh:mm:ss") Then
           dtDate = DateAdd("n", -1 * mPara.int预约有效时间, CDate(Format(mrsBookData!预约时间, "yyyy-MM-dd hh:mm:ss")))
           MsgBox "该预约号已过预约最后接收时间 " & Format(dtDate, "yyyy-MM-dd hh:mm:00") & ",不能接收", vbInformation, gstrSysName
           Set mrsBookData = Nothing: mbytMode = 0
           Exit Function
        End If
    End If
    
    If blnReadPati Then
        If GetPatient(PatiIdentify.GetCurCard, "-" & Val(Nvl(mrsBookData!病人ID)), False, objPati) = False Then Exit Function
    End If
    
    
    '暂不存在接收时换号的处理
    'If ReadRegData = False Then
        Call CreateDeptStructure    '如果读取失败，只是预约挂号单的科室
    'End If
    If mrsDept Is Nothing Then Exit Function
    If mrsDept.State <> 1 Then Exit Function
     
    mbytMode = IIf(Val(Nvl(mrsBookData!记录标志)) = 2, 2, 1)
    If mbytMode <> 2 Then
        lblBookingNO.Caption = "预约单:" & strNO
    Else
        lblBookingNO.Caption = "回诊(单号:" & strNO & ")"
    End If
    lblBookingNO.Left = lblPati.Left + lblPati.Width + 200
    If mbytMode = 2 Then
        '回诊可以选择科室
        Call ReadRegData
    End If
    
    mrsDept.Filter = "科室ID=" & Val(Nvl(mrsBookData!科室ID)) & " And 项目ID=" & Val(Nvl(mrsBookData!项目ID))
    If mrsDept.EOF Then '加上本科室的相关选择
        mrsDept.Filter = 0
        mrsDept.AddNew
        mrsDept!ID = mrsDept.RecordCount + 1
        mrsDept!科室ID = Val(Nvl(mrsBookData!科室ID))
        mrsDept!编码 = CStr(Nvl(mrsBookData!科室编码))
        mrsDept!名称 = CStr(Nvl(mrsBookData!科室名称))
        mrsDept!简码 = CStr(Nvl(mrsBookData!科室简码))
        
        mrsDept!项目ID = CStr(Nvl(mrsBookData!项目ID))
        mrsDept!项目编码 = CStr(Nvl(mrsBookData!项目编码))
        mrsDept!项目名称 = CStr(Nvl(mrsBookData!挂号项目))
        
        mrsDept!是否原科室 = 1
        mrsDept.Update
    End If

    '加载缺省科室
    txtExDept.Text = mrsDept!编码 & "-" & mrsDept!名称 & "【" & mrsDept!项目名称 & "】"
    txtExDept.Tag = Val(Nvl(mrsDept!科室ID)) & ":" & Val(Nvl(mrsDept!项目ID))
    
    '加载缺省医生
    Set mrsDoctor = LoadDoctorData(Val(Nvl(mrsDept!科室ID)), Val(Nvl(mrsDept!项目ID)), mbytMode)
    mrsDept.Filter = 0
    
    txtExDoctor.Text = Trim(Nvl(mrsBookData!医生姓名))
    txtExDoctor.Tag = Val(Nvl(mrsBookData!医生ID)) & ":" & Trim(Nvl(mrsBookData!医生姓名))
    
    Call LoadRoomsData  '加载诊室
    '加载诊室
    txtExRoom.Text = Nvl(mrsBookData!诊室)
    txtExRoom.Tag = Nvl(mrsBookData!诊室)
    lblBookingNO.Visible = True
    
    
    txtExDept.Enabled = mbytMode = 2: cmdExDept.Enabled = mbytMode = 2: cmdExDept.Tag = IIf(mbytMode = 2, "", "F")
    
    
    blnEnabled = txtExDoctor.Text = "" Or mbytMode = 2
    If mbytMode <> 2 Then
        If blnEnabled Then
            If mrsDoctor Is Nothing Then
                blnEnabled = False
            ElseIf mrsDoctor.State <> 1 Then
                blnEnabled = False
            ElseIf mrsDoctor.RecordCount = 0 Then
                 blnEnabled = False
            End If
        End If
    End If
    txtExDoctor.Enabled = blnEnabled: cmdExDoctor.Enabled = blnEnabled: cmdExDoctor.Tag = IIf(blnEnabled, "", "F")
    
    
    blnEnabled = txtExRoom.Text <> "" Or mbytMode = 2
    If mbytMode <> 2 Then
        If blnEnabled Then
            If mrsRooms Is Nothing Then
                blnEnabled = False
            ElseIf mrsRooms.State <> 1 Then
                blnEnabled = False
            ElseIf mrsRooms.RecordCount = 0 Then
                 blnEnabled = False
            End If
        End If
    End If
    txtExRoom.Enabled = blnEnabled: cmdExRoom.Enabled = blnEnabled: cmdExRoom.Tag = IIf(blnEnabled, "", "F")
    

    ReadBooking = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function





