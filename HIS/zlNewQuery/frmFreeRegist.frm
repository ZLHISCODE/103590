VERSION 5.00
Begin VB.Form frmFreeRegist 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   " lblHospital.Caption = GetUnitName + Chr(10) + Chr(13) + ""病人自助挂号系统"""
   ClientHeight    =   8070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrReload 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   6240
   End
   Begin VB.Timer tmrReadCardState 
      Left            =   2400
      Top             =   7560
   End
   Begin VB.PictureBox picReg 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   -720
      ScaleHeight     =   4935
      ScaleWidth      =   9975
      TabIndex        =   6
      Top             =   1080
      Width           =   9975
      Begin zl9NewQuery.ctlButton ctlBack 
         Height          =   915
         Left            =   1920
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4080
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1614
         Caption         =   "返回"
         AutoSize        =   0   'False
         ButtonHeight    =   800
      End
      Begin zl9NewQuery.ctlButton ctlOK 
         Height          =   915
         Left            =   1920
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1614
         Caption         =   "取号"
         AutoSize        =   0   'False
         ButtonHeight    =   800
      End
      Begin VB.Timer Time 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   2040
         Top             =   1680
      End
      Begin VB.Label lblCard 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   720
         Left            =   3960
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Label lblCardID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请刷第二代身份证取号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   42
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   840
         Left            =   -120
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   8550
      End
      Begin VB.Image imgBackgroundImg 
         Height          =   4095
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.PictureBox picLine 
      BackColor       =   &H0000C000&
      Height          =   45
      Index           =   1
      Left            =   960
      ScaleHeight     =   45
      ScaleWidth      =   9135
      TabIndex        =   5
      Top             =   6240
      Width           =   9135
   End
   Begin VB.PictureBox picLine 
      BackColor       =   &H0000C000&
      Height          =   45
      Index           =   0
      Left            =   360
      ScaleHeight     =   45
      ScaleWidth      =   9135
      TabIndex        =   4
      Top             =   1080
      Width           =   9135
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   4680
   End
   Begin VB.Image imgExit 
      Height          =   1380
      Left            =   8040
      MouseIcon       =   "frmFreeRegist.frx":0000
      MousePointer    =   4  'Icon
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   2700
   End
   Begin VB.Label lblHospital 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " 简易挂号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2445
   End
   Begin VB.Label lblNoBIll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对不起，票据已经使用完，请到窗口挂号。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   5040
      Width           =   15960
   End
   Begin VB.Label lblNoBIll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对不起，票据已经使用完，请到窗口挂号。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Index           =   1
      Left            =   -3360
      TabIndex        =   0
      Top             =   2040
      Width           =   15960
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   525
      Left            =   1320
      TabIndex        =   3
      Top             =   6840
      Width           =   285
   End
End
Attribute VB_Name = "frmFreeRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlng领用ID As Long
Private mlng刷新时间 As Long
Private mStrBillNo As String
'Private mrsReg As ADODB.Recordset
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mdblUpBgColor As Double, mdblDownBgColor As Double
Private mlngTime As Long
Private Type FreeRegistInfo
    lng科室ID As Long
    str医生姓名 As String
    lng医生id As Long
    lng急诊   As Long
    lng号别 As Long
    str科室 As String
    str项目 As String
    lng项目ID  As Long
End Type
Private mRegistInfo As FreeRegistInfo
Private Type PatientIdCard
        Name As String
        str年龄 As String
        Sex As String
        Address As String
        CardId As String
        Nation As String
        Birthday As String
End Type
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
Private mIdCardInfo As PatientIdCard
Private mPatient As PATIENTINFO

Private Sub ctlBack_CommandClick()
    Me.Time.Enabled = False
    Call setControlEnabled(False)
End Sub
Private Sub Form_Paint()
 If mdblUpBgColor = 0 And mdblDownBgColor = 0 Then
    Call DrawColorToColor(Me, Me.BackColor, &HFF8080, , True)
 Else
    Call DrawColorToColor(Me, mdblUpBgColor, mdblDownBgColor, , True)
 End If
End Sub

Private Sub Form_Resize()
        On Error Resume Next
        Me.lblHospital.Left = 0
        Me.lblHospital.Top = 20 * Screen.TwipsPerPixelY
        Me.lblHospital.Width = Me.ScaleWidth
        Me.picLine(0).Left = Me.lblHospital.Left
        Me.picLine(0).Top = lblHospital.Top + Me.lblHospital.Height + 20 * Screen.TwipsPerPixelY
        Me.picLine(0).Width = Me.ScaleWidth
        Me.Lblinfo.Left = 0
        Me.Lblinfo.Top = Me.ScaleHeight - 20 * Screen.TwipsPerPixelY - Me.Lblinfo.Height
        Me.Lblinfo.Width = Me.ScaleWidth
        Me.picLine(1).Left = Me.Lblinfo.Left
        Me.picLine(1).Top = Lblinfo.Top - 20 * Screen.TwipsPerPixelY
        Me.picLine(1).Width = Me.ScaleWidth
        With picReg
            .Left = 0
            .Top = picLine(0).Top + picLine(0).Height
            .Height = picLine(1).Top - picLine(0).Top - picLine(0).Height
            .Width = Me.ScaleWidth
        End With
        With lblCardID
            .Left = (Me.picReg.ScaleWidth - .Width) / 2
            .Top = (Me.picReg.ScaleHeight - .Height) / 2
        End With
        Me.imgExit.Top = Me.Lblinfo.Top
     '   Me.imgExit.Height = Me.ScaleHeight - Me.picLine(1).Top - Me.picLine(1).Height - 1 * Screen.TwipsPerPixelY
        Me.imgExit.Left = Me.ScaleWidth - Me.imgExit.Width
        Call SetMsgState
End Sub

Private Sub SetMsgState()
    On Error Resume Next
    Dim wd As Long
    With Me.lblCard
        .Left = IIf(.Width < Me.picReg.ScaleWidth, (Me.picReg.ScaleWidth - .Width) / 2, 0)
        .Top = lblCardID.Top - .Height
    End With
    wd = ctlOK.Width + 20 * Screen.TwipsPerPixelX + ctlBack.Width
    With Me.ctlOK
        .Left = (Me.picReg.ScaleWidth - wd) / 2
        .Top = lblCardID.Top + lblCardID.Height
    End With
    With Me.ctlBack
        .Left = ctlOK.Left + 20 * Screen.TwipsPerPixelX + ctlOK.Width
        .Top = lblCardID.Top + lblCardID.Height
    End With
End Sub
 
 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
   Call InitPicture
   Call Reload
End Sub

Private Sub Reload()
    On Error GoTo hErr
    Call InitFreeRegistMsg
    Call Form_Resize
    mlng刷新时间 = Val(GetPara("主窗体刷新周期", 0)) * 60
    If mlng刷新时间 = 0 Then mlng刷新时间 = 3000
    mlngTime = Val(GetPara("密码验证窗体停留时间")) / 2
    Me.tmrReload.Interval = 1000: Me.tmrReload.Enabled = True
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hwnd)
    End If
    ctlBack.ShowPicture = False
    ctlOK.ShowPicture = False
     If InitBill() = False Then Exit Sub
     If InitFreeRegist() = False Then Exit Sub
     If Me.Lblinfo.Caption = "" Then Lblinfo.Caption = " 挂号项目为-" & _
                    Nvl(mRegistInfo.str科室) & "/" & Nvl(mRegistInfo.str项目) & IIf(IsNull(mRegistInfo.str医生姓名), "", "/" & Nvl(mRegistInfo.str医生姓名))
 
     Call setControlEnabled(False)
     Me.tmrReadCardState.Enabled = False
     Me.tmrReadCardState.Interval = 300
  
 Exit Sub
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub
Public Sub ShowMe(objParent As Object)
    Me.Hide
    Me.Show , objParent
End Sub

Private Function setControlEnabled(blnAllowReg As Boolean)
    If mRegistInfo.lng号别 < 0 And (Not mobjIDCard Is Nothing) Then mobjIDCard.SetEnabled False
    Me.lblCardID.Visible = Not blnAllowReg
    lblCard.Visible = blnAllowReg
    Time.Tag = Val(mlngTime)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Not blnAllowReg)
    Time.Enabled = blnAllowReg
    lblCard.Visible = blnAllowReg
    Me.ctlOK.Visible = blnAllowReg
    Me.ctlBack.Visible = blnAllowReg
    Me.tmrReadCardState.Enabled = blnAllowReg
     
End Function

 

 

Private Sub imgExit_DblClick()
    If frmExitPsw.ShowPsw(Me, True) Then
         Unload frmMainQuery
    End If
End Sub

Private Sub picReg_Paint()
     If Me.imgBackgroundImg.Picture = 0 Then
            Call DrawColorToColor(Me.picReg, mdblUpBgColor, mdblDownBgColor, , True)
    End If
'   Call InitPicture
  'Call ShowPicture
End Sub

Private Sub picReg_Resize()
 On Error Resume Next
 Me.imgBackgroundImg.Left = 0
 Me.imgBackgroundImg.Top = 0
 Me.imgBackgroundImg.Width = Me.picReg.ScaleWidth
 Me.imgBackgroundImg.Height = Me.picReg.ScaleHeight
' Me.picBack.Left = Me.imgBackgroundImg.Left
' Me.picBack.Top = Me.imgBackgroundImg.Top
' Me.picBack.Width = Me.imgBackgroundImg.Width
' Me.picBack.Height = Me.imgBackgroundImg.Height
' LoadBackGroundPicture
' Me.picShow.Width = picReg.Width
' Me.picShow.Height = Me.picReg.Height
End Sub

 Private Sub Time_Timer()
 '-----------------------
 '刷身份证后 信息默认显示时间
 '超过时间 默认为取消挂号
 '-----------------------
   
    Time.Tag = Val(Time.Tag) - 1
    If Val(Time.Tag) <= 0 Then
       setControlEnabled False: Time.Enabled = False
    End If
End Sub
Private Function InitBill() As Boolean
'票据领用检查及初始
     Dim i As Integer
      mlng领用ID = CheckUsedBill(4, IIf(mlng领用ID > 0, mlng领用ID, glng挂号ID))
      If mlng领用ID <= 0 Then
          picReg.Visible = False
          ShowErrMsg "对不起，票据已经使用完，请到窗口挂号。"
          InitBill = False
          Exit Function
      End If
      LblNoBill(0).Visible = False
      LblNoBill(1).Visible = False
      picReg.Visible = True
      InitBill = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not (mobjIDCard Is Nothing) Then
        Call mobjIDCard.SetEnabled(False)
    End If
    Set mobjIDCard = Nothing
    tmrReadCardState.Enabled = False
    Time.Enabled = False
    Timer1.Enabled = False
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
     On Error GoTo hErr
    If Me.Time.Enabled Then Exit Sub
    Me.lblCard.Caption = "欢迎" & strName & "(" & IIf(strSex = "女", "女士", "先生") & ") 到本院就诊！" & _
              IIf(Screen.Width / Screen.TwipsPerPixelX <= 800, vbCrLf, "") & "请取号"
    If GetControlTextWidth(lblCard, Me.lblCard.Caption) > Me.ScaleWidth Then
            Me.lblCard.Caption = GetNewLineString(Me.lblCard.Caption, TextWidth(Me.lblCard.Caption))
    End If
    With mIdCardInfo
           .CardId = strID
           .Name = strName
           .Sex = strSex
           .Address = strAddress
           .Birthday = Format(datBirthday, "yyyy-mm-dd")
           .Nation = strNation
    End With
    Call SetMsgState
    setControlEnabled True
    Exit Sub
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub ctlOK_CommandClick()
    On Error GoTo hErr
    Me.ctlOK.Enabled = False
    Me.ctlBack.Enabled = False
    Time.Enabled = False
    If Save病人信息() = False Then frmShowMessage.ShowMe Me, "病人信息保存失败！":   GoTo ctlState
    If SaveData() = False Then frmShowMessage.ShowMe Me, "挂号失败了！": GoTo ctlState
    frmClose.ShowForm Me, mPatient.Name, mStrBillNo
    Call InitBill
   GoTo ctlState
Exit Sub
hErr:
If ErrCenter() = 1 Then Resume
    SaveErrLog
 Exit Sub
ctlState:
   setControlEnabled False: Me.ctlOK.Enabled = True: Me.ctlBack.Enabled = True:
End Sub
Private Sub Timer1_Timer()
'对下面的提示性
 On Error Resume Next
    If LblNoBill(0).Left + LblNoBill(0).Width > 0 Then
        LblNoBill(0).Left = LblNoBill(0).Left - 100
    Else
        LblNoBill(0).Left = LblNoBill(1).Left + LblNoBill(1).Width
    End If
    If LblNoBill(1).Left + LblNoBill(1).Width > 0 Then
        LblNoBill(1).Left = LblNoBill(1).Left - 100
    Else
        LblNoBill(1).Left = LblNoBill(0).Left + LblNoBill(0).Width
    End If
End Sub

Private Sub ShowErrMsg(strMsg As String)
    Me.picReg.Visible = True
    Me.lblCardID.Caption = strMsg
    With lblCardID
        .Left = (Me.picReg.ScaleWidth - .Width) / 2
        .Top = (Me.picReg.ScaleHeight - .Height) / 2
    End With
    Me.lblCardID.Visible = True
End Sub

 Private Function GetNewLineString(ByVal strMsg As String, ByVal lngWidth As Long) As String
  Dim strTmp As String, lngFontWidth As Long
  Dim lngTmp As Long
  If TextWidth(strMsg) < lngWidth Then GetNewLineString = strMsg: Exit Function
  lngFontWidth = TextWidth("啊")
  For lngTmp = 1 To TextWidth(strMsg) / lngWidth
      strTmp = strTmp & IIf(lngTmp = 1, "", vbCrLf) & Mid$(strMsg, 1, lngWidth / lngFontWidth - IIf(lngWidth / lngFontWidth > 2, 1, 0))
      strMsg = Mid$(strMsg, lngWidth / lngFontWidth + IIf(lngWidth / lngFontWidth > 2, 0, 1))
  Next
   If strMsg <> "" Then strTmp = strTmp & vbCrLf & strMsg
   GetNewLineString = strTmp
 End Function
 
Private Function InitFreeRegist() As Boolean
   Dim strSQL As String, strMsg As String
   Dim rsReg As ADODB.Recordset
      mRegistInfo.lng号别 = Val(GetPara("简单挂号号别", -1))
   If mRegistInfo.lng号别 < 0 Then
       ShowErrMsg "当前没有可以使用的挂号项目！"
        InitFreeRegist = False
        Exit Function
   End If
   
    strSQL = "" & _
    "      Select a.Id, a.号类, a.号码 As 号别, a.科室id, a.项目id,d.名称 as 科室, b.名称 as 项目, a.医生姓名, a.医生id, b.名称, c.现价,Nvl(b.项目特性,0) as 急诊" & _
    "      From 挂号安排 A, 收费项目目录 B, 收费价目 C,部门表 D " & _
    "      Where a.号码 = [1] And a.项目id = b.ID And b.ID = c.收费细目id And a.科室Id=d.Id And Nvl(a.停用日期, Sysdate + 1) > Sysdate " & vbNewLine & _
    "        And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=A.ID and Sysdate between 开始停止时间 and 结束停止时间 )" & vbNewLine & _
    "        And sysDate Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
    "      Union All" & vbNewLine & _
    "      Select a.Id, a.号类, a.号码 As 号别, a.科室id, a.项目id, e.名称 as 科室, b.名称 as 项目, a.医生姓名, a.医生id, b.名称, c.现价,nvl(b.项目特性,0) as 急诊" & _
    "      From 挂号安排 A, 收费项目目录 B, 收费价目 C, 收费从属项目 D,部门表 E " & _
    "      Where a.号码 = [1] And a.科室ID=E.Id And a.项目id = d.主项id And d.从项id = b.ID And b.ID = c.收费细目id And Nvl(a.停用日期, Sysdate + 1) > Sysdate" & vbNewLine & _
    "        And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=A.ID and Sysdate between 开始停止时间 and 结束停止时间 )" & vbNewLine & _
    "        And sysDate Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))"
   On Error GoTo hErr
    Set rsReg = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mRegistInfo.lng号别)
     If rsReg.EOF Then
          ShowErrMsg "对不起，当前号别已经停用，请到窗口挂号。": mRegistInfo.lng号别 = -1: Exit Function
     End If
     With mRegistInfo
        .lng号别 = Val(Nvl(rsReg!号别, 0))
        .lng急诊 = Val(rsReg!急诊)
        .lng科室ID = Val(Nvl(rsReg!科室Id, 0))
        .lng医生id = Val(Nvl(rsReg!医生ID, 0))
         .str医生姓名 = Nvl(rsReg!医生姓名)
        .str科室 = Nvl(rsReg!科室)
        .lng项目ID = Val(Nvl(rsReg!项目Id))
        .str项目 = Nvl(rsReg!项目)
     End With
     InitFreeRegist = True
  Exit Function
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

 


Private Function Save病人信息() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人是否存在 不存在便给病人建立档案
    '入参:
    '出参:
    '返回:成功,返回true,否则False
    '---------------------------------------------------------------------------------------------------------------------------------------------
     
    Dim strSQL As String, lng病人ID As Long, intType As Integer, str门诊号 As String
    Dim rsTemp As New ADODB.Recordset, str出生日期 As String, str年龄 As String
    Dim strNow As String, str医疗付款方式 As String, blnUpdatePatient As Boolean
    Dim lng险类 As Long, str医保 As String
    '未刷卡,不处理
    If mIdCardInfo.CardId = "" Then Exit Function
    strSQL = " " & _
   "   Select " & _
   "          病人ID,姓名,性别,门诊号,年龄,费别,Trunc(出生日期) as 出生日期,家庭地址,民族,险类,费别,医疗付款方式,医保号  " & _
   "   From 病人信息  where 身份证号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mIdCardInfo.CardId)
    If rsTemp.EOF = False Then
        mPatient.Name = zlCommFun.Nvl(rsTemp!姓名)
        mPatient.DoorPost = zlCommFun.Nvl(rsTemp!门诊号, 0)
        mPatient.Sex = zlCommFun.Nvl(rsTemp!性别)
        mPatient.Age = zlCommFun.Nvl(rsTemp!年龄)
        mPatient.FareClass = zlCommFun.Nvl(rsTemp!费别)
        mPatient.strIDCard = mIdCardInfo.CardId
        mPatient.PatientID = zlCommFun.Nvl(rsTemp!病人id)
        mPatient.str出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
        mPatient.str出生地址 = Nvl(rsTemp!家庭地址)
        mPatient.str民族 = Nvl(rsTemp!民族)
        '存在的话，就不保存了
        If mPatient.DoorPost <> 0 Then Save病人信息 = True: Exit Function
        lng病人ID = zlCommFun.Nvl(rsTemp!病人id, 0)
        str医疗付款方式 = Nvl(rsTemp!医疗付款方式)
        mPatient.FareClass = Nvl(rsTemp!费别)
        mPatient.PatientID = lng病人ID
        mIdCardInfo.str年龄 = zlCommFun.Nvl(rsTemp!年龄)
        str年龄 = mIdCardInfo.str年龄
        lng险类 = zlCommFun.Nvl(rsTemp!险类, 0)
        str医保 = zlCommFun.Nvl(rsTemp!医保号, "")
        blnUpdatePatient = True
          
    End If
    strNow = "To_Date('" & CStr(Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:MM:ss")) & _
                             "','YYYY-MM-DD HH24:MI:SS')"
    If Not blnUpdatePatient Then   '新建病人信息
         '获取默认的费别
         strSQL = "Select 名称 From 费别 Where 缺省标志 = 1 And 服务对象 In (1, 3) Order By 服务对象,编码"
         Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
         If Not rsTemp.EOF Then
             mPatient.FareClass = Nvl(rsTemp!名称)
         End If
         strSQL = "Select 名称 From 医疗付款方式 Where 缺省标志 = 1"
         Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
         If Not rsTemp.BOF Then
             str医疗付款方式 = Nvl(rsTemp!名称)
         End If
         '新病人,先建档
         lng病人ID = zlDatabase.GetNextNo(1): mPatient.PatientID = lng病人ID
          If IsDate(mIdCardInfo.Birthday) Then
             strSQL = "Select (Sysdate-to_date([1],'yyyy-mm-dd'))/365 As 岁 From dual"
             Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mIdCardInfo.Birthday)
             str年龄 = Format(Int(Val(Nvl(rsTemp!岁))), "###0") & "岁"
             mIdCardInfo.str年龄 = str年龄
         Else
             str年龄 = ""
         End If
         If Exist病人ID(lng病人ID) = False Then
            '之所以检查 是否存在相同的ID 是为了减少并发的原因
            '造成的相同的病人ID
            lng病人ID = zlDatabase.GetNextNo(1): mPatient.PatientID = lng病人ID
         End If
    End If
    str门诊号 = Nvl(zlDatabase.GetNextNo(3), 0)
    '为了避免并发的原因
    If Exist门诊号(str门诊号, lng病人ID) Then str门诊号 = Nvl(zlDatabase.GetNextNo(3), 0)
    
    '  --处理类型：
    '  --             1=新建病人信息及门诊病案(用于新挂号病人)
    '  --             2=修改病人信息，新建门诊病案(用于无病案的病人)
    '  --             3=修改病人信息，不处理门诊病案(用于有病案的病人,但可能修改了病案的门诊号)
    '  --过敏药物：分隔格式串"ID~名称~~ID~名称...",新增或修改病人信息时用。
    
    'Zl_挂号病人病案_Insert
    strSQL = "Zl_挂号病人病案_Insert("
    '  处理类型_In     Number,
    strSQL = strSQL & "" & IIf(blnUpdatePatient, 2, 1) & ","
    '  病人id_In       病人信息.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  门诊号_In       病人信息.门诊号%Type,
    strSQL = strSQL & "" & str门诊号 & ","
    '  就诊卡号_In     病人信息.就诊卡号%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  卡验证码_In     病人信息.卡验证码%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  姓名_In         病人信息.姓名%Type,
    strSQL = strSQL & "'" & mIdCardInfo.Name & "',"
    '  性别_In         病人信息.性别%Type,
    strSQL = strSQL & "'" & mIdCardInfo.Sex & "',"
    '  年龄_In         病人信息.年龄%Type,
    strSQL = strSQL & "" & IIf(str年龄 = "", "NULL", "'" & str年龄 & "'") & ","
    '  费别_In         病人信息.费别%Type,
    strSQL = strSQL & "'" & mPatient.FareClass & "',"
    '  医疗付款方式_In 病人信息.医疗付款方式%Type,
    strSQL = strSQL & "" & IIf(str医疗付款方式 = "", "NULL", "'" & str医疗付款方式 & "'") & ","
    '  国籍_In         病人信息.国籍%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  民族_In         病人信息.民族%Type,
    strSQL = strSQL & "'" & mIdCardInfo.Nation & "',"
    '  婚姻_In         病人信息.婚姻状况%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  职业_In         病人信息.职业%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  身份证号_In     病人信息.身份证号%Type,
    strSQL = strSQL & "'" & mIdCardInfo.CardId & "',"
    '  工作单位_In     病人信息.工作单位%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  合同单位id_In   病人信息.合同单位id%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  单位电话_In     病人信息.单位电话%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  单位邮编_In     病人信息.单位邮编%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  家庭地址_In     病人信息.家庭地址%Type,
    strSQL = strSQL & "'" & mIdCardInfo.Address & "',"
    '  家庭电话_In     病人信息.家庭电话%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  户口邮编_In     病人信息.户口邮编%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  登记时间_In     病人信息.登记时间%Type,
    strSQL = strSQL & "" & strNow & ","
    '  过敏药物_In     Varchar2,
    strSQL = strSQL & "" & "NULL" & ","
    '  挂号单_In       病人挂号记录.NO%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  出生日期_In     病人信息.出生日期%Type := Null,
    If IsDate(mIdCardInfo.Birthday) Then
        strSQL = strSQL & "to_date('" & mIdCardInfo.Birthday & "','yyyy-mm-dd'),"
    Else
        strSQL = strSQL & "" & "null" & ","
    End If
    '  医保号_In       病人信息.医保号%Type := Null,
    strSQL = strSQL & "" & IIf(str医保 = "", "NULL", "'" & str医保 & "'") & ","
    '  Ic卡号_In       病人信息.Ic卡号%Type := Null
    strSQL = strSQL & "" & "NULL" & ","
   '  险类_In         病人信息.险类%Type := Null
    strSQL = strSQL & IIf(blnUpdatePatient, IIf(lng险类 = 0, "null", lng险类), "null") & ")"
    
    Err = 0: On Error GoTo errHand:
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    With mPatient
        .Name = mIdCardInfo.Name
        .PatientID = lng病人ID
        .Sex = mIdCardInfo.Sex
        .Age = str年龄
        .str出生日期 = Format(mIdCardInfo.Birthday, "yyyy-mm-dd")
        .strIDCard = mIdCardInfo.CardId
        .str出生地址 = mIdCardInfo.Address
        .str民族 = mIdCardInfo.Nation
        .DoorPost = str门诊号
        
    End With
    Save病人信息 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function


Private Function SaveData() As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset, rsPati As New ADODB.Recordset, rs As New ADODB.Recordset
    Dim aryItem As Variant, strNo As String, i As Integer, str收据费目 As String
    Dim StrRoom As String, strBed As String, str费别 As String, strTmp As String, strNow As String
    Dim cllReg As Collection, strSQL As String
    Err = 0: On Error GoTo ErrHandle:
   ' If mblnCanCommit = False Then Exit Sub
  
    
    strNow = "To_Date('" & CStr(Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:MM:ss")) & "','YYYY-MM-DD HH24:MI:SS')"
    '求出当前诊断科室
    '------------------------------------------------------------------------------------------------------------------
    StrRoom = GetRoom(Nvl(mRegistInfo.lng号别))
    If StrRoom = "" Then
        StrRoom = "null"
    Else
        StrRoom = "'" + StrRoom + "'"
    End If
        
    Set cllReg = New Collection
    gstrSQL = "Select C.编码 as 付款码" & _
                " From 病人信息 A,医疗付款方式 C" & _
                " Where A.病人ID=[1] " & _
                " And A.医疗付款方式=C.名称(+)"
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(mPatient.PatientID))
    If rsPati.BOF = False Then strBed = zlCommFun.Nvl(rsPati("付款码").Value)
    
    '严格：取下一个号码
     mStrBillNo = GetNextBill(mlng领用ID)
    
     If mStrBillNo = "" Then
        '没有可用的票据
     End If
     strNo = zlDatabase.GetNextNo(12)
     aryItem = GetFreeRegistPrice(Val(mRegistInfo.lng项目ID))
     'On Error GoTo ErrHandle
            
    '------------------------------------------------------------------------------------------------------------------
    For i = 0 To UBound(aryItem)
        
        gstrSQL = "Select 收据费目 From 收入项目 where ID =[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(aryItem(i, 1)))
        If Not rsTmp.BOF And Not IsNull(rsTmp("收据费目")) Then str收据费目 = CStr(rsTmp("收据费目"))
        '获取费别及实收金额
        str费别 = mPatient.FareClass
        gstrSQL = VB_病人挂号记录_Insert(mPatient.PatientID, mPatient.DoorPost, mPatient.Name, mPatient.Sex, mPatient.Age, Val(strBed), mPatient.FareClass, strNo, mStrBillNo, _
                                   i + 1, Val(aryItem(i, 4)), CLng(aryItem(i, 5)), Format(Nvl(aryItem(i, 0)), "0.00"), CLng(aryItem(i, 1)), _
                                   str收据费目, "", Val(aryItem(i, 0)), 0, _
                                 Val(mRegistInfo.lng科室ID), Val(mRegistInfo.lng科室ID), strNow, strNow, Nvl(mRegistInfo.str医生姓名), Val(Nvl(mRegistInfo.lng医生id)), Val(Nvl(mRegistInfo.lng急诊)), Nvl(mRegistInfo.lng号别), StrRoom, 0, mlng领用ID, 0, _
                                0, 0, 0, 0, Val(aryItem(i, 6)), Val(aryItem(i, 7)))
        zlAddArray cllReg, gstrSQL
    Next
    '问题:31187:主要是将挂号汇总单独出来
    If mRegistInfo.lng号别 >= 0 Then
        strSQL = "zl_病人挂号汇总_Update("
        '  医生姓名_In   挂号安排.医生姓名%Type,
        strSQL = strSQL & "'" & Nvl(mRegistInfo.str医生姓名) & "',"
        '  医生id_In     挂号安排.医生id%Type,
        strSQL = strSQL & "" & IIf(Val(Nvl(mRegistInfo.lng医生id)) = 0, "NULL", mRegistInfo.lng医生id) & ","
        '  收费细目id_In 门诊费用记录.收费细目id%Type,
        strSQL = strSQL & "" & Nvl(mRegistInfo.lng项目ID, 0) & ","
        '  执行部门id_In 门诊费用记录.执行部门id%Type,
        strSQL = strSQL & "" & Nvl(mRegistInfo.lng科室ID, 0) & ","
        '  发生时间_In   门诊费用记录.发生时间%Type,
        strSQL = strSQL & "" & strNow & ","
        '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收
        strSQL = strSQL & "" & 0 & ","
        ' 号码_In       挂号安排.号码%Type := Null
        strSQL = strSQL & mRegistInfo.lng号别 & ")"
        
        Call zlAddArray(cllReg, strSQL)
    End If
    gblnBeginTrans = True
    zlExecuteProcedureArrAy cllReg, Me.Caption, False, False
    gblnBeginTrans = False
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1111", Me, "NO=" & strNo, 2)
    'Call frmClose.ShowForm(Me, mPatient.Name, strNo)
    SaveData = True
    Exit Function
    '-----------------------------------------------------------------------------------------------------------------
 
ErrHandle:
    If gblnBeginTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    gblnBeginTrans = False
    Call SaveErrLog
End Function

Private Function VB_病人挂号记录_Insert(ByVal lng病人ID As String, ByVal lng门诊号 As String, ByVal str姓名 As String, ByVal str性别 As String, ByVal str年龄 As String, _
    ByVal str床号 As String, ByVal str费别 As String, ByVal str单据号 As String, ByVal str票据号 As String, ByVal int序号 As String, ByVal lng数次 As Long, ByVal lng收费细目id As String, _
    ByVal db标准单价 As String, ByVal lng收入项目id As String, ByVal str收据费目 As String, ByVal str结算方式 As String, ByVal db应收金额 As String, ByVal db实收金额 As String, _
    ByVal lng病人科室id As String, ByVal lng执行部门id As String, ByVal str发生时间 As String, ByVal str登记时间 As String, ByVal str医生姓名 As String, ByVal lng医生id As String, _
    ByVal lng急诊 As Long, ByVal str号别 As String, ByVal str发药窗口 As String, ByVal lng结帐id As String, ByVal lng领用ID As String, _
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
    strSQL = strSQL & lng急诊 & ","
    '  号别_In       挂号安排.号码%Type,
    strSQL = strSQL & "'" & str号别 & "',"
    '  诊室_In       门诊费用记录.发药窗口%Type,
    strSQL = strSQL & "" & str发药窗口 & ","
    '  结帐id_In     门诊费用记录.结帐id%Type,
    strSQL = strSQL & "" & IIf(lng结帐id = 0, "NULL", lng结帐id) & ","
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

 
Private Sub tmrReadCardState_Timer()
    '----------------------------------
    '功能设置身份证读卡器是否自动读卡
    '因Idcard会在ShowIDCardInfo事件后
    '会启动自动刷卡
    '----------------------------------
    If Me.ctlOK.Visible Then
        mobjIDCard.SetEnabled False
    End If
End Sub

Private Sub tmrReload_Timer()
'---------------------------------
'刷新挂号信息
'---------------------------------
    Static lngTime As Double
    If mlng刷新时间 = 0 Then mlng刷新时间 = 6000
    If lngTime >= mlng刷新时间 Then
        If Me.Time.Enabled Then Exit Sub
         'frmFlash.Show , Me
          Call Reload
         lngTime = 0
         'Unload frmFlash
    Else
        If Me.Time.Enabled = False Then lngTime = lngTime + 1
    End If
End Sub

Private Function GetControlTextWidth(objControl As Control, strTxt As String) As Double
    '--------------------------------------------------------------------------------
    '获取控件中文本本应该有的宽度
    '--------------------------------------------------------------------------------
    
    Dim lngFont As Long
    lngFont = Me.Font.Size
    Me.Font.Size = objControl.Font.Size
    GetControlTextWidth = TextWidth(strTxt)
    If lngFont <> 0 Then Me.Font.Size = lngFont
End Function

Private Function GetFreeRegistPrice(ByVal lng项目ID) As Variant
    '******************************************************************************************************************
    '功能：返回指定挂号类型，在指定时间的价格二维（六列）数组。
    '   第一列为价格，第二列表示收入项目ID，第三列填写收入项目,第四列为计算单位,第五列为数次,第六列为收费细目ID,第七列(价格序号),第八列(从属父号)
    '参数：lng项目ID=挂号项目ID(收费细目ID)
    '返回：数组
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim aryTmp(), i As Integer
    Dim int性质 As Integer, int父号 As Integer, lng收入项目id As Long
    On Error GoTo errH

    gstrSQL = "Select 1 as 性质,A.类别,A.ID as 项目ID,A.计算单位,B.收入项目ID,1 as 数次,C.收据费目,B.现价" & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=[1] " & _
        " And ((To_Char(Sysdate,'YYYY-MM-DD HH24:MI:SS') Between To_Char(B.执行日期,'YYYY-MM-DD HH24:MI:SS') And To_Char(B.终止日期,'YYYY-MM-DD HH24:MI:SS')) or (To_Char(Sysdate,'YYYY-MM-DD HH24:MI:SS')>=To_Char(B.执行日期,'YYYY-MM-DD HH24:MI:SS') And (B.终止日期 is NULL Or B.终止日期=To_Date('3000-01-01','YYYY-MM-DD'))))"
    gstrSQL = gstrSQL & " Union ALL " & _
        "Select 2 as 性质,A.类别,A.ID as 项目ID,A.计算单位,C.ID as 收入项目ID,D.从项数次 as 数次,C.收据费目,B.现价" & _
        " From 收费项目目录 A,收费价目 B,收入项目 C,收费从属项目 D" & _
        " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.从项ID And D.主项ID=[1]" & _
        "        And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
        ""
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", lng项目ID)
    If rs.EOF Then
        GetFreeRegistPrice = Null
    Else
        ReDim aryTmp(rs.RecordCount - 1, 8)
        int性质 = 0: lng收入项目id = 0
        For i = 1 To rs.RecordCount
            If lng项目ID = Val(Nvl(rs!项目Id)) Then
                If lng收入项目id <> Val(Nvl(rs!收入项目ID)) Then
                    int性质 = 1: int父号 = i:
                     lng收入项目id = Val(Nvl(rs!收入项目ID))
                End If
            Else
                int性质 = 2
            End If
            
            aryTmp(i - 1, 0) = zlCommFun.Nvl(rs("现价").Value, 0)
            aryTmp(i - 1, 1) = zlCommFun.Nvl(rs("收入项目ID").Value, 0)
            aryTmp(i - 1, 2) = zlCommFun.Nvl(rs("收据费目").Value)
            aryTmp(i - 1, 3) = zlCommFun.Nvl(rs("计算单位").Value)
            aryTmp(i - 1, 4) = zlCommFun.Nvl(rs("数次").Value)
            aryTmp(i - 1, 5) = zlCommFun.Nvl(rs("项目ID").Value)
            aryTmp(i - 1, 6) = IIf(int性质 = 1 And i <> int父号, int父号, 0)
            aryTmp(i - 1, 7) = IIf(int性质 = 2 And i <> int父号, int父号, 0)
            rs.MoveNext
        Next
        GetFreeRegistPrice = aryTmp
    End If
  Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    GetFreeRegistPrice = Null
End Function

Private Sub InitPicture()
    Dim rsPic As New ADODB.Recordset
    On Error GoTo hErr

    gstrSQL = "select 序号,名称,宽度,高度,类型 from 咨询图片元素 where 性质=7 order by 修改日期 desc,序号 desc "
    Set rsPic = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsPic.BOF = False Then
      On Error Resume Next
       Me.imgBackgroundImg.Picture = ReadPicByFieldNew(rsPic!序号)
    End If

    If rsPic.State <> adStateClosed Then rsPic.Close
    Set rsPic = Nothing
    
Exit Sub
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub
 
 

Private Sub InitFreeRegistMsg()
    Dim strFontName As String, strMsg As String, dblColor As Double, dblSize As Double
    Dim blnBold As Boolean, blnItalic As Boolean
    '提示信息
    If GetRegistParaFont("简单挂号提示信息", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
        With Me.lblCardID
            .Caption = strMsg
            .Font.Name = strFontName
            .Font.Size = IIf(dblSize > 0, dblSize, .Font.Size)
            .ForeColor = dblColor
            .FontBold = blnBold
            .FontItalic = blnItalic
        End With
        With Me.lblCard
            .Caption = strMsg
            .Font.Name = strFontName
            .Font.Size = IIf(dblSize > 0, dblSize, .Font.Size)
            .ForeColor = dblColor
            .FontBold = blnBold
            .FontItalic = blnItalic
        End With
        
    Else
        lblCardID.Caption = "请刷第二代身份证取号"
    End If
    ctlOK.Font.Name = "宋体"
    ctlOK.Font.Size = 40
    ctlOK.Font.Bold = True
    ctlBack.Font.Name = "宋体"
    ctlBack.Font.Bold = True
    ctlBack.Font.Size = 40
    If GetRegistParaFont("简单挂号上标题", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
        With Me.lblHospital
            .Caption = strMsg
            .Font.Name = strFontName
            .Font.Size = IIf(dblSize > 0, dblSize, .Font.Size)
            .ForeColor = dblColor
            .FontBold = blnBold
            .FontItalic = blnItalic
        End With
    Else
        lblHospital.Caption = GetUnitName & "-简易挂号"
    End If
    If GetRegistParaFont("简单挂号下标题", strMsg, strFontName, dblSize, dblColor, blnBold, blnItalic) Then
      With Me.Lblinfo
            .Caption = strMsg
            .Font.Name = strFontName
            .Font.Size = IIf(dblSize > 0, dblSize, .Font.Size)
            .ForeColor = dblColor
            .FontBold = blnBold
            .FontItalic = blnItalic
      End With
    End If
      mdblUpBgColor = CDbl(Me.BackColor): mdblDownBgColor = CDbl(&HFFC0C0)
      Call GetFreeRegistBGColor(mdblUpBgColor, mdblDownBgColor)
    
End Sub
 
 
