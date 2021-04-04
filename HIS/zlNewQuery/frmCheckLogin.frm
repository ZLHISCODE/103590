VERSION 5.00
Begin VB.Form frmCheckLogin 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "密码验证"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6420
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCheckLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrError 
      Interval        =   6000
      Left            =   0
      Top             =   0
   End
   Begin zl9NewQuery.ctlButton ctlCancel 
      Height          =   720
      Left            =   4500
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1270
      Caption         =   "取消"
      AutoSize        =   0   'False
      ButtonHeight    =   600
   End
   Begin VB.Timer Time 
      Interval        =   4000
      Left            =   1170
      Top             =   2700
   End
   Begin VB.TextBox TxtCardID 
      Height          =   435
      Left            =   2595
      TabIndex        =   0
      Top             =   1830
      Width           =   3135
   End
   Begin VB.TextBox Txtpwd 
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   2595
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2430
      Width           =   3135
   End
   Begin zl9NewQuery.ctlButton ctlOK 
      Height          =   720
      Left            =   2745
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1270
      Caption         =   "确定"
      AutoSize        =   0   'False
      ButtonHeight    =   600
   End
   Begin zl9NewQuery.ctlButton ctlReset 
      Height          =   720
      Left            =   960
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1270
      Caption         =   "重置"
      AutoSize        =   0   'False
      ButtonHeight    =   600
   End
   Begin VB.Label Lblreg 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCheckLogin.frx":1E26
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   2025
      TabIndex        =   7
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Lblinfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   450
      Left            =   105
      TabIndex        =   6
      Top             =   405
      Width           =   5685
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "你选择的挂号项目为："
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1290
      TabIndex        =   5
      Top             =   30
      Width           =   3105
   End
   Begin VB.Label LBLErr 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "密码错误,重输"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2010
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label lblCardID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卡号  "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1845
      TabIndex        =   3
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label Lblpwd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1845
      TabIndex        =   2
      Top             =   2475
      Width           =   660
   End
   Begin VB.Image Imgbak 
      Height          =   2130
      Left            =   180
      Picture         =   "frmCheckLogin.frx":1E64
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1605
   End
End
Attribute VB_Name = "frmCheckLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private Type PARAM_IN
    RegisterMode As String                  '当前采用的挂号方式
    Depart As String
    RegisterItem As String
    DoctorName As String
    DoctorID As Long
    领用ID As Long
    BillNo As String
    RegisterPrice As Double
    DetailID As Long
    DepartID As Long
    号别 As String
End Type
Private mParamIn As PARAM_IN

Private mCurPayNeed As Currency          '病人需要的费用
Private mCurLeft As Currency
Private mlngTime As Long

Private Type PATIENTINFO
    PatientID As String          '病人的ID
    Name As String
    Sex As String                '记录病人的姓名、性别
    DoorPost As String
    Age As String                '记录病人的门诊号和年龄
    FareClass As String
    strIDCard As String           '身份证号
    str出生日期 As String
    str出生地址 As String
    str民族 As String
End Type
Private mPatient As PATIENTINFO
Private mBrushIDCardPatiInfor As PATIENTINFO    '刷卡时的病人信息
Private mlng卡类别ID As Long '医疗卡类别ID 刷医疗卡时传入
Private mblnCanCommit As Boolean         '是否能够提交数据
Private mblnCharge As Boolean
Private mblnNoChange As Boolean
Private mblnBrushCard As Boolean  '刷卡

Private mobjICCard As Object 'IC卡对象
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1

'######################################################################################################################

Public Function ShowLogin(ByVal frmMain As Object, ByVal strRegisterMode As String, _
                            ByVal StrDepart As String, ByVal strRegisterItem As String, ByVal StrDoctorName As String, ByVal lng领用ID As Long, ByVal strBillNo As String, _
                            ByVal lngDoctorID As Long, ByVal dbRegisterPrice As Double, ByVal lngDetailID As Long, ByVal lngDepartID As Long, ByVal str号别 As String, _
                            Optional ByVal lng卡类别ID As Long) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    mParamIn.RegisterMode = strRegisterMode
    mParamIn.Depart = StrDepart
    mParamIn.RegisterItem = strRegisterItem
    mParamIn.DoctorName = StrDoctorName
    mParamIn.领用ID = lng领用ID
    mParamIn.BillNo = strBillNo
    mParamIn.DoctorID = lngDoctorID
    mParamIn.RegisterPrice = dbRegisterPrice
    mParamIn.DetailID = lngDetailID
    mParamIn.DepartID = lngDepartID
    mParamIn.号别 = str号别
    mlng卡类别ID = lng卡类别ID
    
    Me.Show 1, frmMain
    ShowLogin = True
    
End Function

Private Function ShowErrorInfo(ByVal strError As String) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim strErrorInfo As String
    
    mblnNoChange = True

    Select Case strError
    Case "余额不足"
        strErrorInfo = "注意：你的余额不足，请先缴费"
    Case "病人在院"
        strErrorInfo = "注意：你已经在院,不能挂号"
    Case "身份验证"
        
        Select Case mParamIn.RegisterMode
        Case "身份证挂号"
            strErrorInfo = "你的身份证号码错误，请重试！"
        Case "就诊卡挂号"
            strErrorInfo = "你的卡号或密码错误，请重试！"
        Case "ＩＣ卡挂号"
            strErrorInfo = "你的身份证号码错误，请重试！"
        End Select
    Case "就诊卡"
            strErrorInfo = "病人信息不存在"
    End Select
    
    TxtCardID.Text = ""
    Txtpwd.Text = ""
    LBLErr.Caption = strErrorInfo
    LBLErr.Visible = True
    Lblreg.Visible = False
    If TxtCardID.Enabled And TxtCardID.Visible Then TxtCardID.SetFocus
            
    mblnNoChange = False
    
End Function

Private Function ShowPatientInfo(ByVal lng病人ID As Long) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
        
    '将各个变量初始化
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    mlngTime = Val(GetPara("密码验证窗体停留时间"))
    mPatient.Name = "null"
    mPatient.Sex = "nul"
    mPatient.DoorPost = "null"
    mPatient.Age = "null"
    mPatient.FareClass = "null"
    mPatient.PatientID = lng病人ID
    
    '将病人的简单信息求出
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select 姓名,性别,门诊号,年龄,费别,Trunc(出生日期) as 出生日期 from 病人信息 where 病人ID=[1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mPatient.PatientID))
    If rs.BOF = False Then
        mPatient.Name = zlCommFun.Nvl(rs("姓名").Value)
        mPatient.DoorPost = zlCommFun.Nvl(rs("门诊号").Value, 0)
        mPatient.Sex = zlCommFun.Nvl(rs("性别").Value)
        mPatient.Age = zlCommFun.Nvl(rs("年龄").Value)
        mPatient.FareClass = zlCommFun.Nvl(rs("费别").Value)
    End If

    '将一些控件的可见属性改变
    '------------------------------------------------------------------------------------------------------------------
    ctlOK.Visible = True
    '68550,刘尔旋,2014-01-08,读卡后按钮显示内容错误的问题
    ctlOK.Caption = "确定"
    ctlReset.Visible = False
    TxtCardID.Visible = False
    Lblpwd.Visible = False
    Txtpwd.Visible = False
    Lblreg.Caption = Chr(10) + Chr(13) + "如想重新选择，请按“取消”；" + Chr(10) + Chr(13) + "若进行挂号，请按“确认”"
    If mPatient.Age <> "" And mPatient.Age <> "0" Then strTmp = "/" + mPatient.Age Else mPatient.Age = ""
    lblCardID.Caption = "你的信息:" + mPatient.Name + "/" + mPatient.Sex + strTmp
    
    ShowPatientInfo = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitData() As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    
    '将挂号的信息进行显示

    lblInfo.Caption = mParamIn.Depart + "/" + mParamIn.RegisterItem + "/" + mParamIn.DoctorName
    
    ctlOK.Visible = False
    ctlReset.Visible = True
    mblnCanCommit = False
    Txtpwd.Text = ""
    TxtCardID.Text = ""
    
    Select Case mParamIn.RegisterMode
    Case "身份证挂号"
        Txtpwd.Visible = False
        Lblpwd.Visible = False
        ctlReset.Visible = False
        ctlOK.Visible = False
        Lblreg.Caption = "如想重新选择，请按“取消”；否则请正确放置身份证卡"
    Case "ＩＣ卡挂号"
        Txtpwd.Visible = False
        Lblpwd.Visible = False
        ctlReset.Visible = False
        ctlOK.Visible = True
        ctlOK.Caption = "读卡"
        Lblreg.Caption = "如想重新选择，请按“取消”；否则请正确放置ＩＣ卡并按“读卡”"
    Case Else
        Txtpwd.Visible = True
        Lblpwd.Visible = True
        ctlReset.Visible = True
        ctlOK.Visible = True
        ctlOK.Caption = "确定"
        Lblreg.Caption = "如想重新选择，请按“取消”；如确定，请刷卡，并输入密码"
        If mParamIn.RegisterMode = "就诊卡挂号" Then Me.Txtpwd.MaxLength = 0
    End Select
    
    If GetPara("密文显示卡号") = "1" Then
        TxtCardID.PasswordChar = "*"
    End If

End Function

Private Function CheckIdentify(ByVal strMode As String, ByRef lng病人ID As Long, Optional ByVal strUser As String, Optional ByVal strPsw As String) As Boolean
    '******************************************************************************************************************
    '功能:身份验证
    '参数:
    '返回:
    '******************************************************************************************************************
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim strDentify As String
    Dim varAry As Variant
    
    On Error GoTo errHand
    
    Select Case strMode
    '------------------------------------------------------------------------------------------------------------------
    Case "医保卡挂号"
        
        strDentify = gclsInsure.Identify2(UCase(strUser), strPsw, 3, , gintInsure)
        If strDentify = "" Then Exit Function

        varAry = Split(strDentify, ";")
        If UBound(varAry) >= 8 Then
            lng病人ID = Val(varAry(8))
        Else
            Exit Function
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case "就诊卡挂号"
        
        
        If strPsw = "" And mlng卡类别ID = 0 Then
            strSQL = "Select 病人ID From 病人信息 Where 就诊卡号 = [1] And 卡验证码 Is Null"
        ElseIf strPsw <> "" And mlng卡类别ID = 0 Then
            strSQL = "Select 病人ID From 病人信息 Where 就诊卡号 = [1] And 卡验证码 = [2]"
        ElseIf strPsw = "" And mlng卡类别ID <> 0 Then
            strSQL = "Select 病人ID From 病人医疗卡信息 Where 卡号 = [1] And 密码 Is Null And 卡类别id= " & mlng卡类别ID
        Else
            strSQL = "Select 病人ID From 病人医疗卡信息 Where 卡号 = [1] And 密码 = [2] And 卡类别id= " & mlng卡类别ID
        End If
                
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strUser), strPsw)
        If rs.BOF Then Exit Function
        lng病人ID = rs("病人id").Value
        
    '------------------------------------------------------------------------------------------------------------------
    Case "身份证挂号"
        
        strSQL = "Select  病人ID From 病人信息 Where 身份证号 = [1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strUser))
        If rs.BOF Then Exit Function
        lng病人ID = rs("病人id").Value
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ＩＣ卡挂号"
    
        strSQL = "Select  病人ID From 病人信息 Where IC卡号 = [1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strUser))
        If rs.BOF Then Exit Function
        lng病人ID = rs("病人id").Value
        
    End Select
    
    CheckIdentify = True
    
    Exit Function
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function CheckIsHosptial(ByVal lng病人ID As Long) As Boolean
    '******************************************************************************************************************
    '功能:判断当前病人是否在院
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    CheckIsHosptial = True
    
    strSQL = "select 当前科室ID,当前病区ID from 病人信息 where 病人ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If rs.BOF = False Then
        If ((Not IsNull(rs("当前科室ID").Value)) Or (Not IsNull(rs("当前病区ID").Value))) Then Exit Function
    End If
    
    CheckIsHosptial = False
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckMoney(ByVal lng病人ID As Long) As Boolean
    '******************************************************************************************************************
    '功能:通过费别计算需要的费用
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim aryItem As Variant
    Dim strSQL As String
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    strSQL = "Select Nvl(费别,'全费') As 费别,nvl(C.预交余额,0)-nvl(C.费用余额,0) as 余额 From 病人信息 A,病人余额 C Where A.病人ID=[1] and A.病人ID=C.病人ID(+)  And C.性质(+)=1 And C.类型(+)=1 "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If rs.BOF Then Exit Function
    
    If rs("费别").Value = "全费" Then
        mCurPayNeed = CCur(mParamIn.RegisterPrice)
    Else
        aryItem = GetRegistPrice(CLng(mParamIn.DetailID))
        mCurPayNeed = 0
        For intLoop = 0 To UBound(aryItem)
            mCurPayNeed = mCurPayNeed + ActualMoney(CStr(rs("费别").Value), aryItem(intLoop, 1), aryItem(intLoop, 0))
        Next
    End If

    If Val(Nvl(rs!余额)) < mCurPayNeed Then
        CheckMoney = False
    Else
        CheckMoney = True
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ctlCancel_CommandClick()
    Unload Me
End Sub

Private Sub ctlOK_CommandClick()
    
    Dim lng病人ID As Long
    
    If mblnCanCommit = False Then
        '身份验证
        Select Case mParamIn.RegisterMode
        '--------------------------------------------------------------------------------------------------------------
        Case "身份证挂号", "就诊卡挂号", "ＩＣ卡挂号", "医保卡挂号"
            
            If mParamIn.RegisterMode = "ＩＣ卡挂号" Then
                TxtCardID.Text = ""
                If Not (mobjICCard Is Nothing) Then
                    TxtCardID.Text = mobjICCard.Read_Card(Me)
                End If
                If TxtCardID.Text = "" Then Exit Sub
            End If
            
            '将特殊字符去掉
            If mParamIn.RegisterMode = "就诊卡挂号" Then
                TxtCardID.Text = Replace(TxtCardID.Text, ":", "")
                TxtCardID.Text = Replace(TxtCardID.Text, "：", "")
                TxtCardID.Text = Replace(TxtCardID.Text, ";", "")
                TxtCardID.Text = Replace(TxtCardID.Text, "；", "")
                TxtCardID.Text = Replace(TxtCardID.Text, "?", "")
                TxtCardID.Text = Replace(TxtCardID.Text, "？", "")
            End If
        
            If CheckIdentify(mParamIn.RegisterMode, lng病人ID, TxtCardID.Text, Txtpwd.Text) = False Then
                Call ShowErrorInfo("身份验证")
                Exit Sub
            End If
            
            If CheckIsHosptial(lng病人ID) = True Then
                Call ShowErrorInfo("病人在院")
                Exit Sub
            End If
            
            If mblnCharge = False Then
                If CheckMoney(lng病人ID) = False Then
                    Call ShowErrorInfo("余额不足")
                    Exit Sub
                End If
            End If
            
            Call ShowPatientInfo(lng病人ID)
            
            mblnCanCommit = True
            
        End Select
                
    Else
        '提交
        Call CommitData
    End If

End Sub

Private Sub ctlReset_CommandClick()
    Call InitData
    If TxtCardID.Enabled Then TxtCardID.SetFocus
End Sub

Private Sub Form_Activate()
'    If mBlnUse = True Then Unload Me
End Sub

Private Sub Form_Load()

    '将挂号的信息进行显示
    mblnCharge = (Val(GetPara("挂号时生成划价单", "1")) = 1)
    mlngTime = Val(GetPara("密码验证窗体停留时间")) / 2
    If Dir(App.Path & "\图形\挂号确认窗体左面背景.pic") <> "" Then
        Imgbak.Picture = LoadPicture(App.Path & "\图形\挂号确认窗体左面背景.pic")
    End If
    
    ctlReset.Picture = frmselectinfo.ilsImage.ListImages("reset")
    ctlOK.Picture = frmselectinfo.ilsImage.ListImages("ok")
    ctlCancel.Picture = frmselectinfo.ilsImage.ListImages("close")
    
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    On Error Resume Next
    Set mobjICCard = CreateObject("zlICCard.clsICCard")
    On Error GoTo 0
    ctlOK.Width = ctlReset.Width
    
    Call InitData

End Sub

Private Sub Form_Paint()
    Call DrawColorToColor(Me, Me.BackColor, &HFFC0C0, , True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mobjIDCard Is Nothing) Then
        On Error Resume Next
        Call mobjIDCard.SetEnabled(False)
        On Error GoTo 0
    End If
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
   Dim lngPreIDKind As Long
   
    If Not TxtCardID.Locked And TxtCardID.Text = "" And Me.ActiveControl Is TxtCardID Then
        With mBrushIDCardPatiInfor
            .strIDCard = strID
            .Name = strName
            .Sex = strSex
            .str出生日期 = Format(datBirthDay, "yyyy-mm-dd")
            .str出生地址 = strAddress
            .str民族 = strNation
        End With
        TxtCardID.Text = strID
        Call zlSave病人信息
        mblnCanCommit = False
        Call ctlOK_CommandClick
    Else
        mBrushIDCardPatiInfor.strIDCard = ""
    End If
        
End Sub

Private Sub Time_Timer()
    On Error Resume Next
   
    Time.Tag = Val(Time.Tag) - 1
    If Val(Time.Tag) = 0 Then Unload Me
   
End Sub

Private Sub ResetTime()
    Time.Tag = mlngTime
    If LBLErr.Visible = True Then
        LBLErr.Visible = False
        Lblreg.Visible = True
    End If
    
End Sub

Private Sub tmrError_Timer()
    If LBLErr.Visible = True Then
        LBLErr.Visible = False
        Lblreg.Visible = True
    End If
End Sub

Private Sub TxtCardID_Change()
    Dim strTmp As String
    Dim intLen As Integer
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
    
    If mblnNoChange Then Exit Sub
    
    Call ResetTime
    
    Select Case mParamIn.RegisterMode
    Case "就诊卡挂号"

'        strTmp = zlDatabase.GetPara(20, glngSys, , "")
'        If UBound(Split(strTmp, "|")) >= 4 Then intLen = Val(Split(strTmp, "|")(4))
'
'        If Len(TxtCardID.Text) = intLen Then
'
'            '求出密码是否为空
'            gstrSQL = "Select 1 From 病人信息 Where 就诊卡号 = [1] And 卡验证码 Is Null"
'            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(TxtCardID.Text))
'            If rs.BOF = False Then
'                mblnCanCommit = False
'                Call ctlOK_CommandClick
'            Else
'               Txtpwd.SetFocus
'            End If
'        End If
    Case "身份证挂号"
    
        If Me.ActiveControl Is TxtCardID Then
            If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(TxtCardID.Text = "")
        End If
    End Select

    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub TxtCardID_GotFocus()
    
    If mParamIn.RegisterMode = "身份证挂号" Then
        If Not (mobjIDCard Is Nothing) Then
            On Error Resume Next
            Call mobjIDCard.SetEnabled(True)
            On Error GoTo 0
        End If
    Else
        If Not (mobjIDCard Is Nothing) Then
            On Error Resume Next
            Call mobjIDCard.SetEnabled(False)
            On Error GoTo 0
        End If
    End If
    
End Sub

Private Sub TxtCardID_KeyPress(KeyAscii As Integer)
    Dim lng病人ID As Long, blnCard As Boolean
    
    If KeyAscii = 13 Then
    
        Select Case mParamIn.RegisterMode
        '--------------------------------------------------------------------------------------------------------------
        Case "医保卡挂号", "就诊卡挂号"
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        Case "身份证挂号", "ＩＣ卡挂号"
            mblnCanCommit = False
            Call ctlOK_CommandClick
        End Select
    Else
        If mParamIn.RegisterMode = "就诊卡挂号" Then
            Select Case Chr(KeyAscii)
            Case ":", "：", ";", "；", "?", "？"
                KeyAscii = 0
            Case Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
        End If
    End If
    
End Sub

Private Sub TxtCardID_LostFocus()
    Dim rs As ADODB.Recordset
    Dim strPwd As String
    If Not (mobjIDCard Is Nothing) Then
        On Error Resume Next
        Call mobjIDCard.SetEnabled(True)
        On Error GoTo 0
    End If
     If Me.ActiveControl Is Me.ctlCancel Or Me.ActiveControl Is Me.ctlReset Then Exit Sub
     If Trim(TxtCardID.Text) = "" Then Exit Sub
     Select Case mParamIn.RegisterMode
     Case "就诊卡挂号"
          '求出密码是否为空
            Me.Txtpwd.SetFocus
            If mlng卡类别ID = 0 Then
                 gstrSQL = "Select 1 From 病人信息 Where 就诊卡号 = [1] And 卡验证码 Is Null"
            Else
                 gstrSQL = "Select 1 From 病人医疗卡信息 Where 卡号 = [1] And 密码 Is Null" & IIf(mlng卡类别ID = 0, "", " And 卡类别ID=[2] ")
            End If
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(TxtCardID.Text), mlng卡类别ID)
            If rs.BOF = False Then
                mblnCanCommit = False
                Call ctlOK_CommandClick
            Else
                If mlng卡类别ID = 0 Then
                     gstrSQL = "Select 卡验证码 From 病人信息 Where 就诊卡号 = [1]"
                Else
                     gstrSQL = "Select 密码 as  卡验证码  From 病人医疗卡信息 Where 卡号 = [1] " & IIf(mlng卡类别ID = 0, "", " And 卡类别ID=[2] ")
                End If
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(TxtCardID.Text), mlng卡类别ID)
                If rs.EOF Then ShowErrorInfo "就诊卡": Exit Sub
                If frmCardPass.ShowCardPass(Nvl(rs!卡验证码)) Then
                    Txtpwd.Text = Nvl(rs!卡验证码)
                    Call ctlOK_CommandClick
                Else
                    Call ctlReset_CommandClick
                End If
            End If
      Case "医保卡挂号"
        If frmCardPass.GetCardPass(strPwd) Then
            Txtpwd.Text = Nvl(strPwd):    Call ctlOK_CommandClick: Exit Sub
        Else
            Call ctlReset_CommandClick
        End If
     End Select
End Sub

 

Private Sub Txtpwd_Change()
    If mblnNoChange Then Exit Sub
     Call ResetTime
End Sub

Private Sub Txtpwd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Txtpwd.Visible = False Then Exit Sub
        mblnCanCommit = False
        Call ctlOK_CommandClick
    End If
End Sub

Private Sub CommitData()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset, rsPati As New ADODB.Recordset, rs As New ADODB.Recordset
    Dim aryItem As Variant, Str结帐ID As String, strNo As String, i As Integer, str收据费目 As String, str结算方式 As String
    Dim Cur预交支付 As Currency, Cur医保支付 As Currency, cur实收 As Currency, cur应收 As Currency
    Dim Arr医保 As Variant, str保险大类ID As String, str保险项目是否 As String, str统筹金额 As String
    Dim StrRoom As String, strBed As String, str费别 As String, strTmp As String, strNow As String, str划价NO As String
    Dim cllProBefor As Collection, cllPro As Collection, cllproAfter As Collection, strSQL As String
    If mblnCanCommit = False Then Exit Sub
    
    strNow = "To_Date('" & CStr(Format(zlDatabase.Currentdate, "yyyy-mm-dd")) & "','yyyy-mm-dd')"
    '求出当前诊断科室
    '------------------------------------------------------------------------------------------------------------------
    StrRoom = GetRoom(mParamIn.号别)
    If StrRoom = "" Then
        StrRoom = "null"
    Else
        StrRoom = "'" + StrRoom + "'"
    End If
        
    Set cllProBefor = New Collection: Set cllPro = New Collection: Set cllproAfter = New Collection
    
    '求出结算方式
    '------------------------------------------------------------------------------------------------------------------
    str保险大类ID = "Null": str保险项目是否 = "Null": str统筹金额 = "Null"
    
    If mParamIn.RegisterMode = "医保卡挂号" Then
        str结算方式 = "医保基金": Cur医保支付 = CCur(mCurPayNeed)
    Else
        str结算方式 = "个人帐户": Cur预交支付 = CCur(mCurPayNeed)
    End If
    
    '求出序列号
    '------------------------------------------------------------------------------------------------------------------
    strNo = zlDatabase.GetNextNo(12)

    '求出结帐ID
    '------------------------------------------------------------------------------------------------------------------
    Str结帐ID = CStr(zlDatabase.GetNextId("病人结帐记录"))
    aryItem = GetRegistPrice(mParamIn.DetailID)
    
    gstrSQL = "Select C.编码 as 付款码" & _
                " From 病人信息 A,医疗付款方式 C" & _
                " Where A.病人ID=[1] " & _
                " And A.医疗付款方式=C.名称(+)"
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(mPatient.PatientID))
    If rsPati.BOF = False Then strBed = zlCommFun.Nvl(rsPati("付款码").Value)
    
                
    On Error GoTo ErrHandle
            
    '------------------------------------------------------------------------------------------------------------------
    For i = 0 To UBound(aryItem)
    
        '获取通过医保得到的数据
        If mParamIn.RegisterMode = "医保卡挂号" Then
            str结算方式 = "医保基金"
            If gintInsure > 0 Then Arr医保 = Split(gclsInsure.GetItemInsure(CLng(mPatient.PatientID), Val(aryItem(i, 5)), ActualMoney(mPatient.FareClass, aryItem(i, 1), aryItem(i, 4) * aryItem(i, 0)), True, gintInsure), ";")
            str保险项目是否 = CStr(Arr医保(0))
            If CStr(Arr医保(1)) <> "0" Then str保险大类ID = CStr(Arr医保(1))
            str统筹金额 = CStr(CCur(Arr医保(2)))
        Else
            str结算方式 = "个人帐户"
        End If
        
        gstrSQL = "Select 收据费目 From 收入项目 where ID =[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(aryItem(i, 1)))
        If Not rsTmp.BOF And Not IsNull(rsTmp("收据费目")) Then str收据费目 = CStr(rsTmp("收据费目"))
        
        
        cur应收 = CCur(Val(Format(aryItem(i, 0) * aryItem(i, 4), "0.00")))
'        cur应收 = IIf(mblnCharge, "0.00", CCur(Val(Format(aryItem(i, 0), "0.00"))))
        cur实收 = cur应收
        str费别 = mPatient.FareClass
        
        '获取费别及实收金额
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) As 结果 From Dual"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, str费别, CLng(aryItem(i, 5)), CLng(aryItem(i, 1)), cur应收, 0, 0)
        If rs.BOF = False Then
            strTmp = Trim(zlCommFun.Nvl(rs("结果").Value))
            If strTmp <> "" Then
                If InStr(strTmp, ":") > 0 Then
                    cur实收 = Format(Val(Mid(strTmp, InStr(strTmp, ":") + 1)), "0.00")
                    str费别 = Trim(Mid(strTmp, 1, InStr(strTmp, ":") - 1))
                End If
            End If
        End If
            

        If i > 0 Then
            Cur医保支付 = 0
            Cur预交支付 = 0
        End If
        

        gstrSQL = VB_病人挂号记录_Insert(mPatient.PatientID, mPatient.DoorPost, mPatient.Name, mPatient.Sex, mPatient.Age, Val(strBed), str费别, strNo, mParamIn.BillNo, _
                                       i + 1, Val(aryItem(i, 4)), CLng(aryItem(i, 5)), Format(ActualMoney(mPatient.FareClass, aryItem(i, 1), aryItem(i, 0)), "0.00"), CLng(aryItem(i, 1)), str收据费目, _
                                    str结算方式, IIf(mblnCharge, 0, cur应收), IIf(mblnCharge, 0, cur实收), _
                                    mParamIn.DepartID, mParamIn.DepartID, strNow, strNow, mParamIn.DoctorName, mParamIn.DoctorID, mParamIn.号别, StrRoom, Val(Str结帐ID), mParamIn.领用ID, Cur预交支付, _
                                    Cur医保支付, str保险大类ID, str保险项目是否, str统筹金额, Val(aryItem(i, 6)), Val(aryItem(i, 7)))
        zlAddArray cllPro, gstrSQL
        
        '问题:31187:主要是将挂号汇总单独出来
        If mParamIn.号别 <> "" And i + 1 = 1 Then
           strSQL = "zl_病人挂号汇总_Update("
           '  医生姓名_In   挂号安排.医生姓名%Type,
           strSQL = strSQL & "'" & mParamIn.DoctorName & "',"
           '  医生id_In     挂号安排.医生id%Type,
           strSQL = strSQL & "" & IIf(mParamIn.DoctorID = 0, "NULL", mParamIn.DoctorID) & ","
           '  收费细目id_In 门诊费用记录.收费细目id%Type,
           strSQL = strSQL & "" & mParamIn.DetailID & ","
           '  执行部门id_In 门诊费用记录.执行部门id%Type,
           strSQL = strSQL & "" & mParamIn.DepartID & ","
           '  发生时间_In   门诊费用记录.发生时间%Type,
           strSQL = strSQL & "" & strNow & ","
           '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收
           strSQL = strSQL & "" & 0 & ","
            '  号码_In       挂号安排.号码%Type := Null
            strSQL = strSQL & "'" & mParamIn.号别 & "')"
           
           Call zlAddArray(cllproAfter, strSQL)
        End If
        '----------------------------------------------------------------------------------------------------------
        If mblnCharge Then
            If str划价NO = "" Then str划价NO = zlDatabase.GetNextNo(13)
                gstrSQL = "zl_门诊划价记录_Insert("
                '    No_In         门诊费用记录.NO%Type,
                gstrSQL = gstrSQL & "'" & str划价NO & "',"
                '    序号_In       门诊费用记录.序号%Type,
                gstrSQL = gstrSQL & "" & i + 1 & ","
                '    病人id_In     门诊费用记录.病人id%Type,
                gstrSQL = gstrSQL & "" & mPatient.PatientID & ","
                '    主页id_In     住院费用记录.主页id%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    标识号_In     门诊费用记录.标识号%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    付款方式_In   门诊费用记录.付款方式%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    姓名_In       门诊费用记录.姓名%Type,
                gstrSQL = gstrSQL & "'" & mPatient.Name & "',"
                '    性别_In       门诊费用记录.性别%Type,
                gstrSQL = gstrSQL & "'" & mPatient.Sex & "',"
                '    年龄_In       门诊费用记录.年龄%Type,
                gstrSQL = gstrSQL & "'" & mPatient.Age & "',"
                '    费别_In       门诊费用记录.费别%Type,
                gstrSQL = gstrSQL & "'" & str费别 & "',"
                '    加班标志_In   门诊费用记录.加班标志%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    病人科室id_In 门诊费用记录.病人科室id%Type,
                gstrSQL = gstrSQL & "" & mParamIn.DepartID & ","
                '    开单部门id_In 门诊费用记录.开单部门id%Type,
                gstrSQL = gstrSQL & "" & UserInfo.部门ID & ","
                '    开单人_In     门诊费用记录.开单人%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
                '    从属父号_In   门诊费用记录.从属父号%Type,
                gstrSQL = gstrSQL & IIf(Val(aryItem(i, 7)) = 0, "NULL", Val(aryItem(i, 7))) & ","   '57045
                '    收费细目id_In 门诊费用记录.收费细目id%Type,
                gstrSQL = gstrSQL & "" & CLng(aryItem(i, 5)) & ","
                '    收费类别_In   门诊费用记录.收费类别%Type,
                gstrSQL = gstrSQL & "'1',"
                '    计算单位_In   门诊费用记录.计算单位%Type,
                gstrSQL = gstrSQL & "'" & CStr(aryItem(i, 3)) & "',"
                '    发药窗口_In   门诊费用记录.发药窗口%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    付数_In       门诊费用记录.付数%Type,
                gstrSQL = gstrSQL & "1,"
                '    数次_In       门诊费用记录.数次%Type,
                gstrSQL = gstrSQL & "" & CLng(aryItem(i, 4)) & ","
                '    附加标志_In   门诊费用记录.附加标志%Type,
                gstrSQL = gstrSQL & "0,"
                '    执行部门id_In 门诊费用记录.执行部门id%Type,
                gstrSQL = gstrSQL & "" & mParamIn.DepartID & ","
                '    价格父号_In   门诊费用记录.价格父号%Type,
                gstrSQL = gstrSQL & IIf(Val(aryItem(i, 6)) = 0, "NULL", Val(aryItem(i, 6))) & ","
                '    收入项目id_In 门诊费用记录.收入项目id%Type,
                gstrSQL = gstrSQL & "" & CLng(aryItem(i, 1)) & ","
                '    收据费目_In   门诊费用记录.收据费目%Type,
                gstrSQL = gstrSQL & "'" & str收据费目 & "',"
                '    标准单价_In   门诊费用记录.标准单价%Type,
                gstrSQL = gstrSQL & "" & CLng(aryItem(i, 0)) & ","
                '    应收金额_In   门诊费用记录.应收金额%Type,
                gstrSQL = gstrSQL & "" & cur应收 & ","
                '    实收金额_In   门诊费用记录.实收金额%Type,
                gstrSQL = gstrSQL & "" & cur实收 & ","
                '    发生时间_In   门诊费用记录.发生时间%Type,
                gstrSQL = gstrSQL & "" & strNow & ","
                '    登记时间_In   门诊费用记录.登记时间%Type,
                gstrSQL = gstrSQL & "" & strNow & ","
                '    药品摘要_In   药品收发记录.摘要%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    操作员姓名_In 门诊费用记录.操作员姓名%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
                '    类别id_In     药品单据性质.类别id%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '    费用摘要_In   门诊费用记录.摘要%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '    医嘱序号_In   门诊费用记录.医嘱序号%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '    频次_In       药品收发记录.频次%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '    单量_In       药品收发记录.单量%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '    用法_In       药品收发记录.用法%Type := Null, --用法[|煎法]
                gstrSQL = gstrSQL & "NULL,"
                '    期效_In       药品收发记录.扣率%Type := Null,
                 gstrSQL = gstrSQL & "1,"
                '    计价特性_In   药品收发记录.扣率%Type := Null,
                gstrSQL = gstrSQL & "0,"
                '    病人来源_In   Number := 1,
                gstrSQL = gstrSQL & "4)"
                '    保险编码_In   门诊费用记录.保险编码%Type := Null,
                '    费用类型_In   门诊费用记录.费用类型%Type := Null,
                '    保险项目否_In 门诊费用记录.保险项目否%Type := Null,
                '    保险大类id_In 门诊费用记录.保险大类id%Type := Null,
                '    中药形态_In       门诊费用记录.结论%Type := Null
            zlAddArray cllPro, gstrSQL
        End If
    Next
    
    '修改医保的接口数据
    '------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrFirst:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    Err = 0: On Error GoTo ErrHandle:
    If mParamIn.RegisterMode = "医保卡挂号" And gintInsure > 0 Then
         If gclsInsure.RegistSwap(CLng(Str结帐ID), Cur医保支付, gintInsure) = False Then
            gcnOracle.RollbackTrans
            Unload Me
            Exit Sub
         End If
    End If
    zlExecuteProcedureArrAy cllproAfter, Me.Caption, False, True
    
    
    Err = 0: On Error GoTo ErrEnd:
    '打印单据
    '------------------------------------------------------------------------------------------------------------------
    'If mblnCharge = False Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1111", Me, "NO=" & strNo, 2)
    'End If
 
    Call frmClose.ShowForm(Me, mPatient.Name, strNo, str划价NO)
        
    Unload Me
    
    Exit Sub
    '-----------------------------------------------------------------------------------------------------------------
ErrFirst:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Sub
ErrEnd:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Sub
ErrHandle:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
    
End Sub

Private Function VB_病人挂号记录_Insert(ByVal lng病人ID As String, ByVal lng门诊号 As String, ByVal str姓名 As String, ByVal str性别 As String, ByVal str年龄 As String, _
    ByVal str床号 As String, ByVal str费别 As String, ByVal str单据号 As String, ByVal str票据号 As String, ByVal int序号 As String, ByVal lng数次 As Long, ByVal lng收费细目id As String, _
    ByVal db标准单价 As String, ByVal lng收入项目id As String, ByVal str收据费目 As String, ByVal str结算方式 As String, ByVal db应收金额 As String, ByVal db实收金额 As String, _
    ByVal lng病人科室id As String, ByVal lng执行部门id As String, ByVal str发生时间 As String, ByVal str登记时间 As String, ByVal str医生姓名 As String, ByVal lng医生id As String, _
    ByVal str号别 As String, ByVal str发药窗口 As String, ByVal lng结帐id As String, ByVal lng领用ID As String, _
    ByVal Cur预交支付 As String, ByVal Cur医保支付 As String, ByVal str保险大类ID As String, ByVal str保险项目是否 As String, _
    ByVal str统筹金额 As String, ByVal int价格父号 As Integer, ByVal int从属父号 As Integer) As String
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    
    Dim strSQL As String, bln生成队列 As Boolean
    bln生成队列 = Val(zlDatabase.GetPara("排队叫号模式", glngSys, 1113)) <> 0
 
    'Zl_病人挂号记录_Insert
    strSQL = "zl_病人挂号记录_Insert("
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  门诊号_In     门诊费用记录.标识号%Type,
    strSQL = strSQL & "" & lng门诊号 & ","
    '  姓名_In       门诊费用记录.姓名%Type,
    strSQL = strSQL & "'" & str姓名 & "',"
    '  性别_In       门诊费用记录.性别%Type,
    strSQL = strSQL & "'" & str性别 & "',"
    '  年龄_In       门诊费用记录.年龄%Type,
    strSQL = strSQL & "'" & str年龄 & "',"
    '  付款方式_In   门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
    strSQL = strSQL & "'" & Val(str床号) & "',"
    '  费别_In       门诊费用记录.费别%Type,
    strSQL = strSQL & "'" & str费别 & "',"
    '  单据号_In     门诊费用记录.NO%Type,
    strSQL = strSQL & "'" & str单据号 & "',"
    '  票据号_In     门诊费用记录.实际票号%Type,
    strSQL = strSQL & "'" & str票据号 & "',"
    '  序号_In       门诊费用记录.序号%Type,
    strSQL = strSQL & "" & int序号 & ","
    '  价格父号_In   门诊费用记录.价格父号%Type,
    strSQL = strSQL & "" & IIf(int价格父号 = 0, "NULL", int价格父号) & ","
    '  从属父号_In   门诊费用记录.从属父号%Type,
    strSQL = strSQL & "" & IIf(int从属父号 = 0, "NULL", int从属父号) & ","
    '  收费类别_In   门诊费用记录.收费类别%Type,
    strSQL = strSQL & "'1',"
    '  收费细目id_In 门诊费用记录.收费细目id%Type,
    strSQL = strSQL & "" & lng收费细目id & ","
    '  数次_In       门诊费用记录.数次%Type,
    strSQL = strSQL & "" & lng数次 & ","
    '  标准单价_In   门诊费用记录.标准单价%Type,
    strSQL = strSQL & "" & db标准单价 & ","
    '  收入项目id_In 门诊费用记录.收入项目id%Type,
    strSQL = strSQL & "" & lng收入项目id & ","
    '  收据费目_In   门诊费用记录.收据费目%Type,
    strSQL = strSQL & "'" & str收据费目 & "',"
    '  结算方式_In   病人预交记录.结算方式%Type, --现金的结算名称
    strSQL = strSQL & "'" & str结算方式 & "',"
    '  应收金额_In   门诊费用记录.应收金额%Type,
    strSQL = strSQL & "" & db应收金额 & ","
    '  实收金额_In   门诊费用记录.实收金额%Type,
    strSQL = strSQL & "" & db实收金额 & ","
    '  病人科室id_In 门诊费用记录.病人科室id%Type,
    strSQL = strSQL & "" & lng病人科室id & ","
    '  开单部门id_In 门诊费用记录.开单部门id%Type,
    strSQL = strSQL & "" & UserInfo.部门ID & ","
    '  执行部门id_In 门诊费用记录.执行部门id%Type,
    strSQL = strSQL & "" & lng执行部门id & ","
    '  操作员编号_In 门诊费用记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  发生时间_In   门诊费用记录.发生时间%Type,
    strSQL = strSQL & "" & str发生时间 & ","
    '  登记时间_In   门诊费用记录.登记时间%Type,
    strSQL = strSQL & "" & str登记时间 & ","
    '  医生姓名_In   挂号安排.医生姓名%Type,
    strSQL = strSQL & "'" & str医生姓名 & "',"
    '  医生id_In     挂号安排.医生id%Type,
    strSQL = strSQL & "" & IIf(lng医生id = 0, "NULL", lng医生id) & ","
    '  病历费_In Number, --该条记录是否病历工本费
    strSQL = strSQL & "0,"
    '  急诊_In       Number,
    strSQL = strSQL & "0,"
    '  号别_In       挂号安排.号码%Type,
    strSQL = strSQL & "'" & str号别 & "',"
    '问题:48508
    '  诊室_In       病人费用记录.发药窗口%Type,
    strSQL = strSQL & "'" & Replace(str发药窗口, "'", "") & "',"
    '  结帐id_In     门诊费用记录.结帐id%Type,
    strSQL = strSQL & "" & lng结帐id & ","
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "" & IIf(lng领用ID = 0, "NULL", lng领用ID) & ","
    '  预交支付_In   病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
    strSQL = strSQL & "" & Round(Val(Cur预交支付), 2) & ","
    '  现金支付_In   病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
    strSQL = strSQL & "" & 0 & ","
    '  个帐支付_In   病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
    strSQL = strSQL & "" & Round(Val(Cur医保支付), 2) & ","
    '  保险大类id_In 门诊费用记录.保险大类id%Type,
    strSQL = strSQL & "" & IIf(Val(str保险大类ID) = 0, "NULL", str保险大类ID) & ","
    '  保险项目否_In 门诊费用记录.保险项目否%Type,
    strSQL = strSQL & "" & Val(str保险项目是否) & ","
    '  统筹金额_In   门诊费用记录.统筹金额%Type,
    strSQL = strSQL & "" & Val(str统筹金额) & ","
    '  摘要_In       门诊费用记录.摘要%Type, --预约挂号摘要信息
    strSQL = strSQL & "NULL,"
    '  预约挂号_In   Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
    strSQL = strSQL & "0,"
    '  收费票据_In   Number := 0, --挂号是否使用收费票据
    strSQL = strSQL & "0,"
    '  保险编码_In   门诊费用记录.保险编码%Type,
    strSQL = strSQL & "NULL,"
    '  复诊_In       病人挂号记录.复诊%Type := 0,
    strSQL = strSQL & "0,"
    '  号序_In       挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
    strSQL = strSQL & "NULL,"
    '  社区_In       病人挂号记录.社区%Type := Null,
    strSQL = strSQL & "NULL,"
    '  预约接收_In   Number := 0,
    strSQL = strSQL & "0,"
    '  预约方式_In   预约方式.名称%Type := Null,
    strSQL = strSQL & "NULL,"
    '  生成队列_In Number:=0
    strSQL = strSQL & IIf(bln生成队列, 1, 0) & ")"
    VB_病人挂号记录_Insert = strSQL
    
End Function

Private Function zlSave病人信息() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:用身份证挂号时,当用身份证刷卡时,肯定要先建档
    '入参:
    '出参:
    '返回:成功,返回true,否则False
    '编制:刘兴洪
    '日期:2009-12-04 14:05:53
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strSQL As String, lng病人ID As Long, intType As Integer
    Dim rsTemp As New ADODB.Recordset, str出生日期 As String, str年龄 As String
    
    '未刷卡,不处理
    If mBrushIDCardPatiInfor.strIDCard = "" Then Exit Function
    
    
    strSQL = "Select 病人ID,姓名,性别,门诊号,年龄,费别,Trunc(出生日期) as 出生日期,家庭地址,民族 From 病人信息  where 身份证号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mBrushIDCardPatiInfor.strIDCard)
    If rsTemp.EOF = False Then
       '存在病人信息
        mPatient.Name = zlCommFun.Nvl(rsTemp!姓名)
        mPatient.DoorPost = zlCommFun.Nvl(rsTemp!门诊号, 0)
        mPatient.Sex = zlCommFun.Nvl(rsTemp!性别)
        mPatient.Age = zlCommFun.Nvl(rsTemp!年龄)
        mPatient.FareClass = zlCommFun.Nvl(rsTemp!费别)
        mPatient.strIDCard = mBrushIDCardPatiInfor.strIDCard
        mPatient.PatientID = zlCommFun.Nvl(rsTemp!病人id)
        mPatient.str出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
        mPatient.str出生地址 = Nvl(rsTemp!家庭地址)
        mPatient.str民族 = Nvl(rsTemp!民族)
        '存在的话，就不保存了
        zlSave病人信息 = True
        Exit Function
    End If
 
    
    '新病人,先建档
    lng病人ID = zlDatabase.GetNextNo(1): mPatient.PatientID = lng病人ID
    If IsDate(mBrushIDCardPatiInfor.str出生日期) Then
        strSQL = "Select (Sysdate-to_date([1],'yyyy-mm-dd'))/365 As 岁 From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mBrushIDCardPatiInfor.str出生日期)
        str年龄 = Format(Val(Nvl(rsTemp!岁)), "###0.00") & "岁"
    Else
        str年龄 = ""
    End If
    
    '  --处理类型：
    '  --             1=新建病人信息及门诊病案(用于新挂号病人)
    '  --             2=修改病人信息，新建门诊病案(用于无病案的病人)
    '  --             3=修改病人信息，不处理门诊病案(用于有病案的病人,但可能修改了病案的门诊号)
    '  --过敏药物：分隔格式串"ID~名称~~ID~名称...",新增或修改病人信息时用。
    
    'Zl_挂号病人病案_Insert
    strSQL = "Zl_挂号病人病案_Insert("
    '  处理类型_In     Number,
    strSQL = strSQL & "" & 1 & ","
    '  病人id_In       病人信息.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  门诊号_In       病人信息.门诊号%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  就诊卡号_In     病人信息.就诊卡号%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  卡验证码_In     病人信息.卡验证码%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  姓名_In         病人信息.姓名%Type,
    strSQL = strSQL & "'" & mBrushIDCardPatiInfor.Name & "',"
    '  性别_In         病人信息.性别%Type,
    strSQL = strSQL & "'" & mBrushIDCardPatiInfor.Sex & "',"
    '  年龄_In         病人信息.年龄%Type,
    strSQL = strSQL & "" & IIf(str年龄 = "", "NULL", "'" & str年龄 & "'") & ","
    '  费别_In         病人信息.费别%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  医疗付款方式_In 病人信息.医疗付款方式%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  国籍_In         病人信息.国籍%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  民族_In         病人信息.民族%Type,
    strSQL = strSQL & "'" & mBrushIDCardPatiInfor.str民族 & "',"
    '  婚姻_In         病人信息.婚姻状况%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  职业_In         病人信息.职业%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  身份证号_In     病人信息.身份证号%Type,
    strSQL = strSQL & "'" & mBrushIDCardPatiInfor.strIDCard & "',"
    '  工作单位_In     病人信息.工作单位%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  合同单位id_In   病人信息.合同单位id%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  单位电话_In     病人信息.单位电话%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  单位邮编_In     病人信息.单位邮编%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  家庭地址_In     病人信息.家庭地址%Type,
    strSQL = strSQL & "'" & mBrushIDCardPatiInfor.str出生地址 & "',"
    '  家庭电话_In     病人信息.家庭电话%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  户口邮编_In     病人信息.户口邮编%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  登记时间_In     病人信息.登记时间%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  过敏药物_In     Varchar2,
    strSQL = strSQL & "" & "NULL" & ","
    '  挂号单_In       病人挂号记录.NO%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  出生日期_In     病人信息.出生日期%Type := Null,
    If IsDate(mBrushIDCardPatiInfor.str出生日期) Then
        strSQL = strSQL & "to_date('" & mBrushIDCardPatiInfor.str出生日期 & "','yyyy-mm-dd'),"
    Else
        strSQL = strSQL & "" & "NULL" & ","
    End If
    '  医保号_In       病人信息.医保号%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  Ic卡号_In       病人信息.Ic卡号%Type := Null
    strSQL = strSQL & "" & "NULL" & ")"
    
    Err = 0: On Error GoTo errHand:
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    With mPatient
        .Name = mBrushIDCardPatiInfor.Name
        .PatientID = lng病人ID
        .Sex = mBrushIDCardPatiInfor.Sex
        .Age = str年龄
        .str出生日期 = mBrushIDCardPatiInfor.str出生日期
        .strIDCard = mBrushIDCardPatiInfor.strIDCard
        .str出生地址 = mBrushIDCardPatiInfor.str出生地址
        .str民族 = mBrushIDCardPatiInfor.str民族
    End With
    zlSave病人信息 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
