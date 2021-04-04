VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmCaseTendBody 
   Caption         =   "体温作图"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaseTendBody.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11085
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   4545
      ScaleHeight     =   6825
      ScaleWidth      =   5145
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   705
      Visible         =   0   'False
      Width           =   5175
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   5370
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   915
         Width           =   5160
         _Version        =   589884
         _ExtentX        =   9102
         _ExtentY        =   9472
         _StockProps     =   0
         BorderStyle     =   1
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton CmdRef 
         Caption         =   "刷新"
         Height          =   315
         Left            =   4485
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "取消"
         Top             =   510
         Width           =   555
      End
      Begin VB.CommandButton cmdFilterUserCancle 
         Height          =   315
         Left            =   4530
         Picture         =   "frmCaseTendBody.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "取消"
         Top             =   6435
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterUserOk 
         Height          =   315
         Left            =   3990
         Picture         =   "frmCaseTendBody.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "确认"
         Top             =   6435
         Width           =   450
      End
      Begin MSComCtl2.DTPicker dtpE 
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   14
         Top             =   510
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127664131
         CurrentDate     =   37068
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   300
         Index           =   0
         Left            =   885
         TabIndex        =   12
         Top             =   495
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127664131
         CurrentDate     =   37068
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000005&
         Caption         =   "系统默认提取存在体温单文件的待入科、在院、转出和出院3天内的病人，对于出院病人操作员可以指定时间范围进行过滤。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   60
         TabIndex        =   0
         Top             =   0
         Width           =   5100
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Index           =   0
         Left            =   2265
         TabIndex        =   1
         Top             =   555
         Width           =   180
      End
      Begin VB.Label lbl出院时间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院时间"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   548
         Width           =   705
      End
   End
   Begin VB.PictureBox picCondition 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   840
      ScaleHeight     =   405
      ScaleWidth      =   7755
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   165
      Width           =   7755
      Begin VB.PictureBox pic标识 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3510
         ScaleHeight     =   345
         ScaleWidth      =   2775
         TabIndex        =   6
         Top             =   0
         Width           =   2775
         Begin VB.Label lbl床号 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "床123456"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label lbl姓名 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "王二麻子王二麻子"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1260
            TabIndex        =   8
            Top             =   60
            Width           =   2040
         End
      End
      Begin VB.PictureBox pic病人 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   825
         ScaleHeight     =   315
         ScaleWidth      =   1725
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1755
         Begin VB.TextBox txt病人 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   15
            TabIndex        =   5
            Top             =   90
            Width           =   1335
         End
         Begin VB.Image img病人列表 
            Height          =   360
            Left            =   1350
            Picture         =   "frmCaseTendBody.frx":13DE
            Tag             =   "弹出本病区有体温单文件的病人列表"
            Top             =   -30
            Width           =   360
         End
      End
      Begin VB.PictureBox pic住院次数 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6330
         ScaleHeight     =   225
         ScaleWidth      =   1335
         TabIndex        =   3
         Top             =   0
         Width           =   1365
         Begin VB.ComboBox cboPages 
            Height          =   315
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   -30
            Width           =   1425
         End
      End
      Begin VB.Label lbl定位 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "定位病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   10
         Top             =   90
         Width           =   720
      End
      Begin VB.Image img上一个 
         Height          =   360
         Left            =   2580
         Picture         =   "frmCaseTendBody.frx":1AE0
         Tag             =   "上一个病人"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image img下一个 
         Height          =   360
         Left            =   2940
         Picture         =   "frmCaseTendBody.frx":21E2
         Tag             =   "下一个病人"
         Top             =   0
         Width           =   360
      End
   End
   Begin zl9TemperatureChartGD.usrBodyEditor BodyEdit 
      Height          =   4425
      Left            =   255
      TabIndex        =   15
      Top             =   840
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   7805
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   7080
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBody.frx":28E4
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16642
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgRPT 
      Left            =   240
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBody.frx":3176
            Key             =   "woman"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBody.frx":99D8
            Key             =   "man"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   195
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************
'病人基本信息
'***************************************************************

Private Enum PATI_TYPE
    pt入院待入住 = 0
    pt转科待入住 = 1
    pt转病区待入住 = 2
    pt在院 = 3
'    pt家庭病床 = 3.1
'    pt预转科 = 3.2
'    pt转病区 = 3.3
    pt预出 = 4
    pt出院 = 5
    pt死亡 = 6
    pt最近转出 = 7
End Enum

Private Type type_Patient
    lng病人ID As Long
    lng主页ID As Long
    lng病区ID As Long '当前执行病区ID
    lng科室ID As Long
    lng出院 As Long
    lng婴儿 As Long
    lng编辑 As Long
    lng护理等级 As Long
    lng文件ID As Long
    lng原始大小 As Long
    lngPage As Long
    排序 As String  'PATI_TYPE
    病区ID As Long '病人当前所在病区
    病案状态 As Long
End Type

Private T_Info As type_Patient    '记录当前病人信息

Private Enum PATI_COLUMN
    c_选择 = 0
    c_图标 = 1
    c_排序 = 2
    c_状态 = 3
    c_床号 = 4
    C_病人ID = 5
    c_主页ID = 6
    c_姓名 = 7
    c_年龄 = 8
    c_住院号 = 9
    c_入院日期 = 10
    c_出院日期 = 11
End Enum

Private mblnChildForm As Boolean
Private mcbrToolBar As CommandBar
Private mcbr查看 As CommandBarControl
Private mstrPrivs As String
Private mstrSQL As String
Private mblnShowing As Boolean
Private mblnChanged As Boolean
Private mfrmMain As Form
Private mIntDataEditor As Integer
Private mblnMove As Boolean
Private mfrmTendBody As Object '体温单对象
Private mintChange As Integer '参数最近转出天数
Private mdtOutEnd As String '参数出院显示终止时间
Private mdtOutBegin As String '参数出院显示开始时间
Private mrsPati As New ADODB.Recordset
Private mintPrePage As Integer
Private mbytFontSize As Byte
Private mbytSize As Byte
Private mblnDoctorStation As Boolean

Public Event AfterPrint()
Public Event CmdClick(ByVal strParam As String)

Public Property Let ReSize(bytSize As Byte)
    mbytSize = bytSize
End Property

Public Property Let DoctorStation(blnDoctorStation As Boolean)
    mblnDoctorStation = blnDoctorStation
End Property

'######################################################################################################################
'自定义函数、过程区域
Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-19 15:16
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
    BodyEdit.FontSize = bytSize
    Call BodyEdit.ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = mbytFontSize
    
    Me.FontSize = bytFontSize
    Me.FontName = "宋体"
    
    Set CtlFont = cbsThis.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsThis.Options.Font = CtlFont
    
    lbl定位.FontSize = Me.FontSize
    lbl定位.Top = pic病人.Top + (pic病人.Height - lbl定位.Height) \ 2
    lbl定位.Left = 60
    txt病人.FontSize = Me.FontSize
    txt病人.Top = (pic病人.Height - txt病人.Height)
    pic病人.Left = lbl定位.Left + lbl定位.Width + 20
    img上一个.Left = pic病人.Left + pic病人.Width
    img下一个.Left = img上一个.Left + img上一个.Width
    pic标识.Left = img下一个.Left + img下一个.Width + TextWidth("刘")
    pic住院次数.Left = pic标识.Left + pic标识.Width + 20
    cboPages.FontSize = Me.FontSize
    pic住院次数.Height = cboPages.Height - 20
    pic住院次数.Top = (picCondition.Height - pic住院次数.Height) \ 2
    If pic住院次数.Top < 0 Then pic住院次数.Top = 0
    picCondition.Width = pic住院次数.Left + pic住院次数.Width + 20
    '病人选择
    lblInfo.FontSize = Me.FontSize
    lbl出院时间.FontSize = Me.FontSize
    Label2(0).FontSize = bytFontSize
    dtpB(0).Font.Size = bytFontSize
    dtpB(0).Width = TextWidth("2012-01-01") + 400
    dtpB(0).Height = TextHeight("刘") * 1.5
    dtpE(0).Font.Size = bytFontSize
    dtpE(0).Width = TextWidth("2012-01-01") + 400
    dtpE(0).Height = TextHeight("刘") * 1.5
    
    Set CtlFont = rptPati.PaintManager.CaptionFont
    CtlFont.Size = bytFontSize
    Set rptPati.PaintManager.CaptionFont = CtlFont
    
    Set CtlFont = rptPati.PaintManager.TextFont
    CtlFont.Size = bytFontSize
    Set rptPati.PaintManager.TextFont = CtlFont
    rptPati.Redraw
           
   
    dtpB(0).Left = lbl出院时间.Left + lbl出院时间.Width + TextWidth("刘") / 2
    Label2(0).Left = dtpB(0).Left + dtpB(0).Width + TextWidth("刘") / 2
    dtpE(0).Left = Label2(0).Left + Label2(0).Width + TextWidth("刘") / 2
    CmdRef.Height = TextHeight("刘") + 100
    CmdRef.Left = picPati.ScaleWidth - CmdRef.Width - 120
    If dtpE(0).Left + dtpE(0).Width + TextWidth("刘") > CmdRef.Left Then
        picPati.Width = dtpE(0).Left + dtpE(0).Width + TextWidth("刘") + CmdRef.Width + 120
        rptPati.Width = picPati.ScaleWidth
        CmdRef.Left = picPati.ScaleWidth - CmdRef.Width - 120
    End If
    lblInfo.Width = picPati.ScaleWidth - TextWidth("刘") / 2
    lblInfo.Height = TextHeight("刘") * (TextWidth(lblInfo.Caption) \ picPati.ScaleWidth + 1) + 20
    dtpB(0).Top = lblInfo.Height + lblInfo.Top + 60
    lbl出院时间.Top = dtpB(0).Top + (dtpB(0).Height - lbl出院时间.Height) \ 2
    Label2(0).Top = lbl出院时间.Top
    dtpE(0).Top = dtpB(0).Top
    CmdRef.Top = dtpE(0).Top
     
    rptPati.Top = CmdRef.Top + CmdRef.Height + 60
    cmdFilterUserOk.Top = rptPati.Top + rptPati.Height + 120
    cmdFilterUserCancle.Top = cmdFilterUserOk.Top
    cmdFilterUserCancle.Left = picPati.ScaleWidth - cmdFilterUserCancle.Width - 120
    cmdFilterUserOk.Left = cmdFilterUserCancle.Left - cmdFilterUserOk.Width - 60
    picPati.Height = cmdFilterUserCancle.Top + cmdFilterUserCancle.Height + 60
End Sub

Public Function ShowEdit(ByVal frmMain As Object, strParam As String, Optional ByVal bytMode As Byte = 1, Optional strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    Dim blnShowing As Boolean
    
    mbytFontSize = IIf(mbytSize = 0, 9, IIf(mbytSize = 1, 12, mbytSize))

    mblnMove = False
    mblnChildForm = False
    mblnChanged = False
    mstrPrivs = strPrivs
    
    blnShowing = mblnShowing
    
    mblnShowing = True
    
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    On Error GoTo Errhand
    
    If blnShowing Then
        If Val(varParam(0)) = T_Info.lng病人ID Or Val(varParam(1)) = T_Info.lng主页ID And T_Info.lng病区ID = Val(varParam(2)) Then
            Call ShowWindow(Me.hWnd, SW_RESTORE)
            Call BringWindowToTop(Me.hWnd)
            Exit Function
        End If
    End If
    
    Set BodyEdit.ParentForm = Me
    Set mfrmMain = frmMain

    '参数格式：病人ID;主页ID;病区ID;文件ID;出院;编辑;婴儿;是否更具窗体大小自动校正体温单格式(1 否 0 校正)页号(默认显示第几页,如果页号超出范围就按缺省显示,0按缺省显示)
    
    '初始化参数
    
    T_Info.lng婴儿 = 0
    T_Info.lng出院 = 0
    T_Info.lng编辑 = 0
    T_Info.lng原始大小 = 0
    T_Info.lngPage = 0
    
    T_Info.lng病人ID = Val(varParam(0))
    T_Info.lng主页ID = Val(varParam(1))
    T_Info.lng病区ID = Val(varParam(2))
    T_Info.lng科室ID = Val(varParam(2))
    T_Info.lng文件ID = Val(varParam(3))
    
    If UBound(varParam) > 3 Then T_Info.lng出院 = Val(varParam(4))
    If UBound(varParam) > 4 Then
        T_Info.lng编辑 = Val(varParam(5))
    Else
        If InStr(1, ";" & mstrPrivs & ";", ";体温单作图;") = 0 Then
            T_Info.lng编辑 = 0
        Else
            T_Info.lng编辑 = 1
        End If
    End If
    If UBound(varParam) > 5 Then T_Info.lng婴儿 = Val(varParam(6))
    If UBound(varParam) > 6 Then T_Info.lng原始大小 = Val(varParam(7))
    If UBound(varParam) > 7 Then
        T_Info.lngPage = Val(varParam(8))
    Else
        T_Info.lngPage = glngCurPage
    End If
    
    mintPrePage = T_Info.lng主页ID
    
    Set RS = New ADODB.Recordset
    If blnShowing = False Then
        Call InitMenuBar
        '体温单原始大小可以切换病人
        Call RefreshPatiList(RS)
        Call AddPages
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 出院科室ID,nvl(数据转出,0) 转出  from 病案主页 Where 病人id=[1] And 主页id=[2] "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng病人ID, T_Info.lng主页ID)
    If RS.BOF = False Then
        T_Info.lng科室ID = Val(zlCommFun.Nvl(RS("出院科室ID").Value))
        If T_Info.lng出院 = 1 Then mblnMove = (Val(RS("转出")) <> 0)
    End If
    
    
    '------------------------------------------------------------------------------------------------------------------
    BodyEdit.FontSize = mbytSize
    If OpenPatientMap = False Then
        Unload Me
        Exit Function
    End If
    
    Call GetTendEidor
    Call ReSetFontSize
    
    If blnShowing = False Then
        Hook Me
        
        If bytMode = 1 Then
            Me.Show , mfrmMain
        Else
            Me.Show 1, mfrmMain
        End If
        ShowEdit = mblnChanged
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlInit() As Boolean
    mblnChildForm = True
End Function

Public Function GetCurvePage() As Long
   GetCurvePage = BodyEdit.intPage
End Function

Public Sub zlDataEditor(ByVal intDataEditor As Integer)
    BodyEdit.DateEditor = intDataEditor
End Sub

Public Function zlRefresh(ByVal frmParent As Form, strParam As String, Optional strPrivs As String) As Boolean

   '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    Dim intBaby As Integer
    
    mblnMove = False
    mstrPrivs = strPrivs
    mblnChildForm = True
    stbThis.Visible = Not mblnChildForm
    cbsThis.ActiveMenuBar.Visible = False
    picCondition.Visible = False
    picCondition.Enabled = False
    cbsThis.RecalcLayout
    
    mblnChanged = False
    
    Set BodyEdit.ParentForm = frmParent
    BodyEdit.FontSize = 0
    
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    On Error GoTo Errhand
    
    '参数格式：病人ID;主页ID;病区ID;文件ID;出院;编辑;婴儿;是否更具窗体大小自动校正体温单格式(1 否 0校正);页号(默认显示第几页,如果页号超出范围就按缺省显示,0按缺省显示)
    If Val(varParam(3)) <> T_Info.lng文件ID Then
        glngCurPage = 0
    Else
        If UBound(varParam) > 5 Then
            intBaby = Val(varParam(6))
        Else
            intBaby = 0
        End If
        
        If T_Info.lng婴儿 <> intBaby Then
            glngCurPage = 0
        End If
    End If
    
    '初始化参数
    T_Info.lng婴儿 = 0
    T_Info.lng出院 = 0
    T_Info.lng编辑 = 0
    T_Info.lng原始大小 = 0
    T_Info.lngPage = 0
    
    T_Info.lng病人ID = Val(varParam(0))
    T_Info.lng主页ID = Val(varParam(1))
    T_Info.lng病区ID = Val(varParam(2))
    T_Info.lng科室ID = Val(varParam(2))
    T_Info.lng文件ID = Val(varParam(3))
    
    If UBound(varParam) > 3 Then T_Info.lng出院 = Val(varParam(4))
    If UBound(varParam) > 4 Then
        T_Info.lng编辑 = Val(varParam(5))
    Else
        If InStr(1, ";" & mstrPrivs & ";", ";体温单作图;") = 0 Then
            T_Info.lng编辑 = 0
        Else
            T_Info.lng编辑 = 1
        End If
    End If
    If UBound(varParam) > 5 Then T_Info.lng婴儿 = Val(varParam(6))
    If UBound(varParam) > 6 Then T_Info.lng原始大小 = Val(varParam(7))
    If UBound(varParam) > 7 Then
        T_Info.lngPage = Val(varParam(8))
    Else
        T_Info.lngPage = glngCurPage
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 出院科室ID,nvl(数据转出,0) 转出 from 病案主页 Where 病人id=[1] And 主页id=[2] "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng病人ID, T_Info.lng主页ID)
    If RS.BOF = False Then
        T_Info.lng科室ID = Val(zlCommFun.Nvl(RS("出院科室ID").Value))
        If T_Info.lng出院 = 1 Then mblnMove = (Val(RS("转出")) <> 0)
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        Unload Me
        Exit Function
    End If
    
    Call GetTendEidor
    
    zlRefresh = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function OpenPatientMap() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim strParam As String
    
    On Error GoTo Errhand
    
    T_Info.lng护理等级 = 3
    gstrSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng病人ID, T_Info.lng主页ID)
    If RS.BOF = False Then T_Info.lng护理等级 = zlCommFun.Nvl(RS("护理等级"), 3)
    
    '重新提取文件ID
    '53880:刘鹏飞,2012-09-19,提取文件ID不应该加科室ID，应为病人可能存在两次不同科室住院，转科的情况。
    gstrSQL = "select A.ID from 病人护理文件 A,病历文件列表 B" & _
       "    where A.病人ID=[1] and A.主页Id=[2] and A.婴儿=[3] and A.格式ID=B.ID and B.种类=3 and B.保留=-1"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng病人ID, T_Info.lng主页ID, T_Info.lng婴儿)
    If mblnMove = True Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
    End If
    
    If RS.BOF = False Then T_Info.lng文件ID = Val(zlCommFun.Nvl(RS("ID")))
    '初始化曲线菜单
    If InitBodyLine = False Then Exit Function
    
    '参数：病人ID;主页ID;病区ID;文件ID;出院标志;编辑标志;婴儿;护理等级;原始大小;页号(默认显示第几页,如果页号超出范围就按缺省显示,0按缺省显示)
    strParam = T_Info.lng病人ID & ";" & T_Info.lng主页ID & ";" & T_Info.lng病区ID & ";" & T_Info.lng文件ID & ";" & _
        T_Info.lng出院 & ";" & T_Info.lng编辑 & ";" & T_Info.lng婴儿 & ";" & T_Info.lng护理等级 & ";" & T_Info.lng原始大小 & ";" & T_Info.lngPage
    Call BodyEdit.zlMenuClick("初始化", strParam)
        
    OpenPatientMap = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitBodyLine() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim cbrItem As CommandBarControl
    Dim strSQL As String
    
    On Error GoTo Errhand
    
    '--曲线设置检查
    mstrSQL = _
            " Select a.记录名, a.项目序号" & vbNewLine & _
            " From 体温记录项目 A, 护理记录项目 B" & vbNewLine & _
            " Where a.记录法 = 1 And a.项目序号 = b.项目序号 And b.护理等级 >= [1] And Nvl(b.应用方式, 0) = 1 And Nvl(b.适用病人, 0) In (0, [2]) And" & vbNewLine & _
            "      (b.适用科室 = 1 Or (b.适用科室 = 2 And Exists (Select 1 From 护理适用科室 C Where b.项目序号 = c.项目序号 And c.科室id = [3])))" & vbNewLine & _
            " Order By a.排列序号"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, T_Info.lng护理等级, IIf(T_Info.lng婴儿 = 0, 1, 2), T_Info.lng科室ID)
    If rsTmp.BOF Then
        MsgBox "无适用此病人的体温单操作曲线项目，请在护理项目管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '--记录频次时间段设置检查
    mstrSQL = " Select Distinct nvl(记录频次,2) 频次,Decode(项目表示,4,2,1) 类别  From 体温记录项目 A,护理记录项目 B" & _
            "   WHERE A.记录法 =2 And A.项目序号<>3 And A.项目序号=B.项目序号 AND B.护理等级>=[1] And Nvl(b.应用方式,0)=1 And Nvl(b.适用病人, 0) In (0, [2]) And" & _
            "      (b.适用科室 = 1 Or (b.适用科室 = 2 And Exists (Select 1 From 护理适用科室 C Where b.项目序号 = c.项目序号 And c.科室id = [3])))" & _
            "   Order by 类别,频次"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, T_Info.lng护理等级, IIf(T_Info.lng婴儿 = 0, 1, 2), T_Info.lng科室ID)
    
    rsTmp.Filter = "类别=1"
    Do While Not rsTmp.EOF
        strSQL = "select Count(*) 记录数 From 护理项目频次 where 频次=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!频次))
        If Val(rsData!记录数) < Val(rsTmp!频次) Then
            MsgBox "护理项目记录频次时段设置不完整，请在护理项目管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    rsTmp.MoveNext
    Loop
    rsTmp.Filter = ""
    
    '--汇总项目时间段设置检查
    mstrSQL = "select 名称,开始,结束,类别 from 护理汇总时段 where 单据=1"
    Set rsData = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    
    rsTmp.Filter = "类别=2"
    Do While Not rsTmp.EOF
        rsData.Filter = ""
        If Val(rsTmp!频次) = 1 Then
            rsData.Filter = "类别=3"
            If rsData.RecordCount = 0 Then
                MsgBox "护理汇总时段设置不完整，请在护理项目管理中设置！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            rsData.Filter = "类别=1 OR 类别=2"
            If rsData.RecordCount < 2 Then
                MsgBox "护理汇总时段设置不完整，请在护理项目管理中设置！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    rsTmp.MoveNext
    Loop
    
    InitBodyLine = True
    
    Exit Function
    
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function PrintData(ByVal bytMode As Byte, Optional ByVal strPrintDevice As String = "") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnCur As Boolean
    Dim lngBeginY As Long
    Dim intBeginPage As Integer
    Dim intPrintRange As Integer
    Dim strPage  As String, strParam As String
    
    '传入了打印机名称,说明是批量打印,自动从第1页开始打印,不进行任何询问
    '返回:0-取消,2-预览,1-打印
    
    If strPrintDevice = "" Then
        'strParam = T_Info.lng文件ID & ";" & T_Info.lng病人ID & ";" & T_Info.lng主页ID & ";" & T_Info.lng科室Id & ";" & T_Info.lng科室Id
        strParam = T_Info.lng文件ID & ";" & Me.BodyEdit.AllPage
        bytMode = frmCaseTendBodyPrintSet.PrintSet(Me, True, strParam, intPrintRange, lngBeginY, intBeginPage, strPage, mstrPrivs, bytMode)
    Else
        bytMode = 2
        intPrintRange = 2
    End If
    If bytMode = 0 Then Exit Function
    If intBeginPage <= 0 Then intBeginPage = -1
    
    '打印当前页传入当前页号
    If intPrintRange = 0 Then
        strPage = Me.BodyEdit.intPage - 1
    End If
    
    Select Case bytMode
    Case 2  '打印
        Call BodyEdit.PrintState(intPrintRange, True, lngBeginY, intBeginPage, strPrintDevice, strPage)
    Case 1  '预览
        Call BodyEdit.PrintState(intPrintRange, False, lngBeginY, intBeginPage, strPrintDevice, strPage)
    End Select

End Function

Public Function zlPrintBody(Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDevice As String) As Long
    '入参:1-预览,2-打印
    '返回值:0-失败;1-成功;2-打印
    gblnPrinted = False
    Call PrintData(IIf(bytMode = 1, 2, 1), strPrintDevice)
    zlPrintBody = IIf(gblnPrinted, 2, 1)
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim objCustom As CommandBarControlCustom
    
    On Error GoTo Errhand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&E)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
       
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存数据(&S)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "恢复数据(&R)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControl.BeginGroup = True
    End With


    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve, "曲线编辑(&Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CurveTable, "表格编辑(&T)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_Show, "曲线显示(&D)")
    End With

    Set mcbr查看 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    With mcbr查看.CommandBar.Controls
                
'       Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
'
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
                
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."):
        cbrControl.BeginGroup = True
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    picCondition.Visible = True
    picCondition.Enabled = True
    
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("条件工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0

    With mcbrToolBar.Controls
          Set objCustom = .Add(xtpControlCustom, 1, "")
          objCustom.Handle = picCondition.hWnd
    End With
    
    '定位工具栏
    '------------------------------------------------------------------------------------------------------------------

    For Each cbrControl In mcbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("Q"), conMenu_Edit_Curve
        .Add FCONTROL, Asc("T"), conMenu_Edit_CurveTable
        .Add FCONTROL, Asc("D"), conMenu_Edit_Curve_Show
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    
    Call InitReprotControl '初始化病人信息列表
    
    InitMenuBar = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub BodyEditCur(ByVal intDataEditor As Integer, Optional ByVal strParam As String = "")
    Call GetTendEidor
    If intDataEditor = 0 Or intDataEditor = -1 Then
        gintEditorCurveState = intDataEditor
        Call BodyEdit.zlMenuClick("体温数据编辑", strParam)
    ElseIf intDataEditor = 1 Then
         Call BodyEdit.zlMenuClick("体温数据显示设置", strParam)
    End If
End Sub

Private Sub BodyEdit_DbClickCur(ByVal intDataEditor As Integer)
    Call BodyEditCur(intDataEditor)
End Sub

Private Sub cboPages_Click()
    Dim blnEdit As Boolean
    Dim lngType As PATI_TYPE
    If cboPages.ListIndex = -1 Then Exit Sub
    If Val(cboPages.ItemData(cboPages.ListIndex)) = mintPrePage Then Exit Sub
    mintPrePage = Val(cboPages.ItemData(cboPages.ListIndex))
    T_Info.lng主页ID = cboPages.ItemData(cboPages.ListIndex)
    
    Call GetPatiInfo
    '刷新数据
    T_Info.lng婴儿 = 0: T_Info.lngPage = 0
    
    '确定当前是否可以编辑
    blnEdit = True
    With T_Info
        lngType = Val(.排序)
        If lngType = pt出院 Or lngType = pt死亡 Then
            If Not (Val(.病案状态) = 0 Or Val(.病案状态) = 2 Or Val(.病案状态) = 999) Then
                '可能是在院抽查反馈状态，出院后并未提交审查
                If Val(.病案状态) = 1 Or Val(.病案状态) = 2 Then blnEdit = False
            End If
        ElseIf lngType = pt转科待入住 Or lngType = pt转病区待入住 Then
            blnEdit = False
        End If
        blnEdit = blnEdit And (T_Info.lng病区ID = .病区ID Or lngType = pt最近转出)
    End With
    
    If InStr(1, ";" & mstrPrivs & ";", ";体温单作图;") > 0 And blnEdit = True And mblnDoctorStation = False Then
        T_Info.lng编辑 = 1
    Else
        T_Info.lng编辑 = 0
    End If
    Call OpenPatientMap
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim lngIndex As Long
    Dim cbrControl As CommandBarControl
    Dim lngKey As Long
    
    On Error GoTo Errhand
    
    ReDim Preserve strSQL(1 To 1)
        
    Select Case Control.Id
        Case conMenu_File_PrintSet   '打印设置
            
            On Error Resume Next
            Call frmPrintSet.ShowMe(Me, 1)
            
        Case conMenu_File_Preview  '打印预览
            
            Call PrintData(2)
            
        Case conMenu_File_Print  '打印
        
            Call PrintData(1)
        
        Case conMenu_View_ToolBar_Button

'            cbsThis(2).Visible = Not cbsThis(2).Visible
'            cbsThis.RecalcLayout

        Case conMenu_View_ToolBar_Text

'            For Each cbrControl In cbsThis(1).Controls
'                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
'            Next
'
'            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
            
        Case conMenu_Edit_Curve '曲线编辑
            Call BodyEditCur(0)
        Case conMenu_Edit_CurveTable '表格编辑
            Call BodyEditCur(-1)
        Case conMenu_Edit_Curve_Show '曲线显示
            Call BodyEditCur(1)
            
        Case conMenu_Edit_Save '保存数据
            
        Case conMenu_Edit_Reuse '数据恢复
            
        Case conMenu_Help_Help
        
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.hWnd)
            
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hWnd)
            
        Case conMenu_Help_Web_Mail
            
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
    End Select
    
    Exit Sub
    
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
        End Select
    End If
    
    Err = 0
    On Error Resume Next
    
    Select Case Control.Id

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Curve, conMenu_Edit_CurveTable, conMenu_Edit_Curve_Show
        
        Control.Enabled = (T_Info.lng编辑 = 1)
        
    Case conMenu_View_ToolBar_Button
    
        Control.Checked = Me.cbsThis(2).Visible
        
    Case conMenu_View_ToolBar_Text
    
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        
    Case conMenu_View_ToolBar_Size
    
        Control.Checked = Me.cbsThis.Options.LargeIcons
        
    Case conMenu_View_StatusBar
    
        Control.Checked = Me.stbThis.Visible
        
    End Select
End Sub

Private Sub BodyEdit_zlAfterPrint()
    gblnPrinted = True
    RaiseEvent AfterPrint
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With BodyEdit
        .mblnResize = True
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Top = lngTop
        .mblnResize = False
        .Height = lngBottom - lngTop
    End With
    picCondition.Width = Me.pic住院次数.Left + Me.pic住院次数.Width + 100
End Sub

Private Sub cmdFilterUserCancle_Click()
    picPati.Visible = False
End Sub

Private Sub cmdFilterUserOk_Click()
     Call rptPati_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub CmdRef_Click()
    Dim RS As New ADODB.Recordset
    Call RefreshPatiList(RS)
    Call img病人列表_MouseDown(1, 0, 0, 0)
End Sub

Private Sub dtpB_Change(Index As Integer)
'时间范围改变时刷新
    If dtpB(Index).Value >= dtpE(Index).Value Then
        MsgBox "出院时间范围的开始时间应小于结束时间", vbInformation, gstrSysName
        dtpB(Index).Value = dtpB(Index).Tag
        dtpB(Index).SetFocus: Exit Sub
    Else
        dtpB(Index).Tag = dtpB(Index).Value
        If Index = 0 Then mdtOutBegin = dtpB(Index).Value
    End If
End Sub

Private Sub dtpE_Change(Index As Integer)
    If dtpB(Index).Value >= dtpE(Index).Value Then
        MsgBox "出院时间范围的开始时间应小于结束时间", vbInformation, gstrSysName
        dtpE(Index).Value = dtpE(Index).Tag
        dtpE(Index).SetFocus: Exit Sub
    Else
        dtpE(Index).Tag = dtpE(Index).Value
        If Index = 0 Then mdtOutEnd = dtpE(Index).Value
    End If
End Sub

Private Sub Form_Load()
    Call GetLocalSetting '读取相关参数
    If Not mblnChildForm Then
         Call RestoreWinState(Me, App.ProductName)
    End If
End Sub

Private Sub GetTendEidor()
    If Not gobjTendEditor Is Nothing Then Set gobjTendEditor = Nothing
    Set gobjTendEditor = Me
End Sub

Private Sub BodyEdit_CmdClick(ByVal strParam As String)
    Dim arrParam() As String
    If mfrmTendBody Is Nothing Then Set mfrmTendBody = New frmCaseTendBody
    
    If mfrmTendBody.ShowEdit(BodyEdit.ParentForm, strParam, 0, mstrPrivs) Then
        arrParam = Split(strParam, ";")
        If UBound(arrParam) > 6 Then arrParam(7) = 0
        If UBound(arrParam) > 7 Then
            strParam = arrParam(0) & ";" & arrParam(1) & ";" & arrParam(2) & ";" & arrParam(3) & ";" & arrParam(4) & ";" & arrParam(5) & ";" & arrParam(6) & ";" & arrParam(7)
        Else
            strParam = Join(arrParam, ";")
        End If
        
        Call zlRefresh(BodyEdit.ParentForm, strParam, mstrPrivs)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHook Me
    
    mblnShowing = False
    Set mfrmTendBody = Nothing
    
    If Not mblnChildForm Then
        Call SaveWinState(Me, App.ProductName)
        mblnChanged = True
    End If
    If Not gobjTendEditor Is Nothing Then Set gobjTendEditor = Nothing
    
    Set mcbrToolBar = Nothing
    Set mcbr查看 = Nothing
    Set mrsPati = Nothing
    '卸载用户控件对象 （窗体关闭时用户控件的 UserControl_Terminate 事件无法进入 所以放在父窗体关闭执行 ）
    Call BodyEdit.ReleaseObj
End Sub

Private Sub img病人列表_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngColor As Long
    Dim lngLoop As Long
    Dim objRow As ReportRow
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strPatient As String '病人列表信息
    Dim lngRow As Long, lngID As Long 'VSF选择的病人ID
    Dim lngLeft As Long, lngTop  As Long, lngRight As Long, lngBottom As Long
    Dim ArrCode() As String
    Dim blnVisible As Boolean
    
    If Button <> 1 Then Exit Sub
    
    On Error GoTo Errhand
    
    If rptPati.Records.Count = 0 And mrsPati.RecordCount > 0 Then
        '显示病人列表供选择
        With mrsPati
            .MoveFirst
            
            Do While Not .EOF
                Set objRecord = rptPati.Records.Add()
                objRecord.Tag = CStr(!病人ID & "," & !主页ID)
                Set objItem = objRecord.AddItem("")
                objItem.HasCheckbox = True
                objItem.Checked = False
                
                Set objItem = objRecord.AddItem(""): objItem.Icon = IIf(!性别 = "男", 1, 0)
                Set objItem = objRecord.AddItem(CStr(!排序))
                objItem.Caption = CStr(!排序 & !类型)
                Set objItem = objRecord.AddItem(CStr(!排序 & !类型))
                objItem.Caption = CStr(!排序 & !类型)
                
                Set objItem = objRecord.AddItem(LPAD(Nvl(!床号), 10, " "))
                objItem.Caption = Trim(Nvl(!床号, " "))
                objRecord.AddItem Val(!病人ID)
                objRecord.AddItem Val(!主页ID)
                objRecord.AddItem CStr(Nvl(!姓名))
                objRecord.AddItem CStr(Nvl(!年龄))
                Set objItem = objRecord.AddItem(CStr(Nvl(!住院号)))
                objItem.Caption = Nvl(!住院号, " ")
                
                Set objItem = objRecord.AddItem(Format(!入院日期, "yyyy-MM-dd HH:mm:ss"))
                objItem.Caption = Format(!入院日期, "yyyy-MM-dd HH:mm:ss")
                Set objItem = objRecord.AddItem(Format(!出院日期, "yyyy-MM-dd HH:mm:ss"))
                objItem.Caption = Format(!出院日期, "yyyy-MM-dd HH:mm:ss")
                
                '提取病人类型的颜色
                lngColor = Nvl(!颜色, 0)
                If lngColor <> 0 Then objRecord.Item(c_姓名).ForeColor = lngColor
                
                .MoveNext
            Loop
        End With
    End If
    
    If mrsPati.RecordCount > 0 Then mrsPati.MoveFirst
    mrsPati.Find "病人ID=" & T_Info.lng病人ID
    
    Call mcbrToolBar.GetWindowRect(lngLeft, lngTop, lngRight, lngBottom)
    rptPati.Populate '缺省不选中任何行
    picPati.Left = picCondition.Left + Me.pic病人.Left
    picPati.Top = lngTop - Me.Top - 60
    picPati.Visible = True
    
    '选中当前病人(先折叠组的话,Rows.Count只有组的个数了,所以先定位,再折叠)
    For lngLoop = 0 To rptPati.Rows.Count - 1
        If Not (rptPati.Rows(lngLoop).Record Is Nothing) Then
            If Val(rptPati.Rows(lngLoop).Record.Item(C_病人ID).Value) = T_Info.lng病人ID Then
                Set rptPati.FocusedRow = rptPati.Rows(lngLoop)
                Exit For
            End If
        End If
    Next
    
    '折叠所有组(选中病人那一组不折叠)
    For Each objRow In rptPati.Rows
        If objRow.GroupRow And objRow.Index <> rptPati.FocusedRow.ParentRow.Index Then
            objRow.Expanded = False
        End If
    Next
    If Not rptPati.FocusedRow Is Nothing Then
        rptPati.FocusedRow.EnsureVisible
    End If
    rptPati.SetFocus
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub img病人列表_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo Me.pic病人.hWnd, img病人列表.Tag
End Sub

Private Sub img上一个_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LocatePati(1)
End Sub

Private Sub img上一个_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hWnd, img上一个.Tag
End Sub

Private Sub img下一个_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LocatePati(2)
End Sub

Private Sub img下一个_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hWnd, img下一个.Tag
End Sub

Private Sub lbl姓名_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo pic标识.hWnd, lbl姓名.Caption
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call cmdFilterUserCancle_Click
    End If
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If rptPati.Records.Count = 0 Then Exit Sub
    If rptPati.FocusedRow.Record Is Nothing Then Exit Sub
    
    T_Info.lng病人ID = Split(rptPati.FocusedRow.Record.Tag, ",")(0)
    T_Info.lng主页ID = Split(rptPati.FocusedRow.Record.Tag, ",")(1)
    '如果需要病人定位后按上一个,下一个时按定位前的顺序,可把该语句屏蔽掉
    mrsPati.Filter = ""
    If mrsPati.RecordCount > 0 Then mrsPati.MoveFirst
    mrsPati.Find "病人ID=" & T_Info.lng病人ID
    picPati.Visible = False
    txt病人.Text = ""
    mintPrePage = -1
    Call AddPages
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptPati_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub txt病人_GotFocus()
    Call zlControl.TxtSelAll(txt病人)
End Sub

Private Sub txt病人_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    strInput = Trim(txt病人.Text)
    If strInput = "" Then Exit Sub
    
    strInput = " 床号='" & LPAD(strInput, 10, " ") & "'"
    mrsPati.Filter = strInput
    If mrsPati.RecordCount = 0 Then
        If Not IsNumeric(Trim(txt病人.Text)) Then
            strInput = " 姓名='" & Trim(txt病人.Text) & "'"
        Else
            strInput = " 住院号=" & Trim(txt病人.Text)
        End If
        mrsPati.Filter = strInput
        
        If mrsPati.RecordCount = 0 Then
            mrsPati.Filter = 0
            MsgBox "未找到该病人的有效数据，请重新输入！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    T_Info.lng病人ID = mrsPati!病人ID
    T_Info.lng主页ID = mrsPati!主页ID
    mrsPati.Filter = 0
    If mrsPati.RecordCount > 0 Then mrsPati.MoveFirst
    mrsPati.Find "病人ID=" & T_Info.lng病人ID
    mintPrePage = -1
    Call AddPages
    
    picPati.Visible = False
End Sub


Private Sub LocatePati(ByVal intType As Integer)
    '参数说明:intType:1-上一个病人;2-下一个病人
    '病人范围:在床病人循环,与老版保持一致
    Dim blnExit As Boolean  '强制退出
    On Error Resume Next
    
redo:
    If mrsPati.RecordCount = 0 Then Exit Sub
    If intType = 1 Then
        mrsPati.MovePrevious
        If mrsPati.BOF Then mrsPati.MoveLast
    Else
        mrsPati.MoveNext
        If mrsPati.EOF Then mrsPati.MoveFirst
    End If
    If mrsPati!病人ID <> 0 Then
        If mrsPati!病人ID <> T_Info.lng病人ID Then
            T_Info.lng病人ID = mrsPati!病人ID
            T_Info.lng主页ID = mrsPati!主页ID
            mintPrePage = -1
            Call AddPages
        Else
            If blnExit Then Exit Sub
            blnExit = True
            GoTo redo
        End If
    Else
        GoTo redo
    End If
    
    picPati.Visible = False
End Sub

Private Sub AddPages()
    Dim i As Integer, j As Integer
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    '根据病人ID读取该病人的住院次数

    '52004,刘鹏飞,2012-08-10,住院次数应该默认定位到当前病人当前住院次数
    strSQL = " Select A.主页ID From 病案主页 A,病人护理文件 B,病历文件列表 C" & _
        " Where A.病人ID=B.病人ID and A.主页ID=B.主页ID And nvl(B.婴儿,0)=0 And B.格式ID=C.ID And C.种类=3 And C.保留=-1 " & _
        " And A.主页ID<>0 And A.病人ID=[1] Order by A.主页ID Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取住院次数", T_Info.lng病人ID)
    cboPages.Clear
    Do While Not rsTemp.EOF
        cboPages.AddItem "第 " & rsTemp!主页ID & " 次"
        cboPages.ItemData(cboPages.NewIndex) = rsTemp!主页ID
        If rsTemp!主页ID = T_Info.lng主页ID Then
            Call zlControl.CboSetIndex(cboPages.hWnd, cboPages.NewIndex)
        End If
        rsTemp.MoveNext
    Loop
    If cboPages.ListIndex = -1 Then
        Call zlControl.CboSetIndex(cboPages.hWnd, 0)
    End If
    Call cboPages_Click
End Sub

Private Sub GetPatiInfo()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo Errhand
    
    strSQL = "Select A.姓名,B.出院病床 床号, C.颜色,B.当前病区ID,B.病案状态,B.出院科室id" & vbNewLine & _
        " From 病人信息 A, 病案主页 B,病人类型 C" & vbNewLine & _
        " Where A.病人ID=B.病人ID And B.病人ID=[1] And B.主页ID=[2] And B.病人类型=C.名称(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取病人信息", T_Info.lng病人ID, T_Info.lng主页ID)
    
    lbl床号.Caption = "床" & Trim(Nvl(rsTemp!床号))
    lbl姓名.Caption = Nvl(rsTemp!姓名)
    lbl姓名.ForeColor = Nvl(rsTemp!颜色, 0)
    T_Info.排序 = mrsPati!排序
    T_Info.病区ID = Nvl(rsTemp!当前病区ID, 0)
    T_Info.病案状态 = Val(Nvl(rsTemp!病案状态, 0))
    T_Info.lng科室ID = Val(Nvl(rsTemp!出院科室ID, 0))
    Me.pic标识.Width = lbl姓名.Width + lbl姓名.Left
    Me.pic住院次数.Width = Me.cboPages.Width - 50
    Me.pic住院次数.Left = pic标识.Left + pic标识.Width + 50
    picCondition.Width = Me.pic住院次数.Left + Me.pic住院次数.Width + 100
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub GetLocalSetting()
'功能：从注册表读取出院病人的时间范围
    Dim i As Integer
    Dim curDate As Date, intDay As Integer

    '病人显示范围
    mintChange = Val(zlDatabase.GetPara("最近转出天数", glngSys, p住院护士站, 7))
    '如果大于30天就取缺省值
    If mintChange > 30 Then mintChange = 7
    
    '出院病人时间范围
    curDate = zlDatabase.Currentdate
    mdtOutEnd = Format(curDate, "yyyy-MM-dd")
    mdtOutBegin = Format(CDate(mdtOutEnd) - 3, "yyyy-MM-dd")
    dtpE(0).Value = mdtOutEnd
    dtpE(0).Tag = mdtOutEnd
    dtpB(0).Value = mdtOutBegin
    dtpB(0).Tag = mdtOutBegin
End Sub

Public Sub RefreshPatiList(Optional ByVal rsThis As ADODB.Recordset)
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo Errhand
    '52004,刘鹏飞,2012-08-10,住院次数应该默认定位到当前病人当前住院次数
    '刷新病人清单,仍定位到当前操作的病人上
    Call LoadPatient(rsThis)
    mrsPati.Filter = "病人ID=" & T_Info.lng病人ID
    If mrsPati.RecordCount > 0 Then
        T_Info.lng主页ID = mrsPati!主页ID
    Else
         '如果找不到再从数据库中提取
        strSQL = "" & _
            "Select /*+ RULE */ Decode(B.出院方式,'死亡',6,5) as 排序," & _
            " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
            " Decode(B.出院方式,'死亡','死亡病人','出院病人') as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,A.姓名,A.性别,A.年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
            " lpad(nvl(B.出院病床,' '),10,' ') 床号,E.名称 as 护理等级,B.费别,B.当前病况,B.入院日期,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(B.路径状态,-1) 路径状态,trunc(b.出院日期)-trunc(b.入院日期)+1 as 住院天数,z.颜色" & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z" & _
            " Where A.病人ID=B.病人ID And B.病人类型=Z.名称(+) And Nvl(B.主页ID,0)<>0 And B.状态=0" & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And B.当前病区ID=[1] And B.病人ID=[2] And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And B.出院日期 Is Not NULL And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
        Set rsTemp = New ADODB.Recordset
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, T_Info.lng病区ID, T_Info.lng病人ID)
        rsTemp.Sort = "出院日期 DESC"
        If rsTemp.RecordCount > 0 Then T_Info.lng主页ID = rsTemp!主页ID
        Call CopyReocrd(rsTemp)
    End If
    
    mrsPati.Filter = ""
    If mrsPati.RecordCount > 0 Then mrsPati.MoveFirst
    mrsPati.Find ("病人ID=" & T_Info.lng病人ID)
    rptPati.Records.DeleteAll
    Call GetPatiInfo
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CopyReocrd(ByVal rsPati As ADODB.Recordset)
    Dim strField As String, strValue As String
    
    rsPati.Filter = 0
    If rsPati.RecordCount <> 0 Then rsPati.MoveFirst
    strField = "排序|排序2|类型|病人ID|主页ID|住院号|姓名|性别|年龄|科室|科室ID|住院医师|责任护士|病案状态|床号|护理等级|费别|当前病况|入院日期|出院日期|住院天数|出院方式|病人类型|状态|险类|就诊卡号|路径状态|颜色"
    Do While Not rsPati.EOF
        strValue = rsPati!排序 & "|" & rsPati!排序2 & "|" & rsPati!类型 & "|" & rsPati!病人ID & "|" & rsPati!主页ID & "|" & Nvl(rsPati!住院号, 0) & "|" & Nvl(rsPati!姓名) & "|" & Nvl(rsPati!性别) & "|" & _
                  Nvl(rsPati!年龄) & "|" & Nvl(rsPati!科室) & "|" & Nvl(rsPati!科室ID, 0) & "|" & Nvl(rsPati!住院医师) & "|" & Nvl(rsPati!责任护士) & "|" & Nvl(rsPati!病案状态, 0) & "|" & Nvl(rsPati!床号) & "|" & _
                  Nvl(rsPati!护理等级, "三级") & "|" & Nvl(rsPati!费别) & "|" & Nvl(rsPati!当前病况, "一般") & "|" & Format(rsPati!入院日期, "yyyy-MM-dd HH:mm:ss") & "|" & Format(rsPati!出院日期, "yyyy-MM-dd HH:mm:ss") & "|" & rsPati!住院天数 & "|" & rsPati!出院方式 & "|" & _
                  Nvl(rsPati!病人类型, "普通病人") & "|" & rsPati!状态 & "|" & Nvl(rsPati!险类, 0) & "|" & Nvl(rsPati!就诊卡号) & "|" & Nvl(rsPati!路径状态, 0) & "|" & Nvl(rsPati!颜色, 0)
        Call Record_Add(mrsPati, strField, strValue)
        rsPati.MoveNext
    Loop
End Sub

Private Sub LoadPatient(ByVal rsThis As ADODB.Recordset)
    Dim strSQL As String
    Dim strField As String, strValue As String
    On Error GoTo Errhand
    Set rsThis = Nothing
    '入院等入科和转科待入科病人(病人科室所属的病区都可接收)
    'c.科室id + 0,说明：通过H表的索引连接过滤后，记录数量很少，再连接B表则更快
    If rsThis Is Nothing Then
ErrGO:
        Set rsThis = New ADODB.Recordset
        strField = "排序," & adDouble & ",2|排序2," & adDouble & ",2|类型," & adLongVarChar & ",50|病人ID," & adDouble & ",18|主页ID," & adDouble & ",18|" & _
                   "住院号," & adDouble & ",18|姓名," & adLongVarChar & ",20|性别," & adLongVarChar & ",10|年龄," & adLongVarChar & ",20|科室," & adLongVarChar & ",50|" & _
                   "科室ID," & adDouble & ",18|住院医师," & adLongVarChar & ",20|责任护士," & adLongVarChar & ",20|病案状态," & adLongVarChar & ",20|" & _
                   "床号," & adLongVarChar & ",20|护理等级," & adLongVarChar & ",50|费别," & adLongVarChar & ",50|当前病况," & adLongVarChar & ",50|" & _
                   "入院日期," & adLongVarChar & ",20|出院日期," & adLongVarChar & ",20|住院天数," & adLongVarChar & ",20|出院方式," & adLongVarChar & ",20|" & _
                   "病人类型," & adLongVarChar & ",50|状态," & adLongVarChar & ",10|险类," & adDouble & ",18|就诊卡号," & adLongVarChar & ",20|路径状态," & adLongVarChar & ",20|颜色," & adDouble & ",18"
        Call Record_Init(mrsPati, strField)
        strSQL = _
            "Select /*+ RULE */Distinct" & vbNewLine & _
            " Decode(B.状态,1,0,Decode(c.开始原因,3,1,2)) As 排序, Decode(Nvl(b.病案状态, 0), 0, 999, b.病案状态) As 排序2," & _
            " Decode(B.状态,1,'入院待入住病人',Decode(c.开始原因,3,'转科待入住病人','转病区待入住病人')) As 类型," & _
            " a.病人id, b.主页id, A.门诊号,B.住院号, a.姓名, a.性别, b.年龄," & vbNewLine & _
            " d.名称 As 科室, c.科室id, c.经治医师 As 住院医师,b.责任护士, b.病案状态, lpad(nvl(C.床号,' '),10,' ') as 床号," & _
            " e.名称 As 护理等级, b.费别,b.当前病况, b.入院日期, b.出院日期,B.出院方式, b.病人类型, b.状态, b.险类, a.就诊卡号," & vbNewLine & _
            " -1 As 路径状态,trunc(sysdate)-trunc(b.入院日期)+1 as 住院天数,Z.颜色" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病人变动记录 C, 部门表 D, 收费项目目录 E,病人类型 Z,在院病人 R" & vbNewLine & _
            "Where B.病人类型=Z.名称(+) And A.病人ID = R.病人ID And a.病人id = b.病人id And Nvl(b.主页id, 0) <> 0 And b.病人id = c.病人id And b.主页id = c.主页id And c.科室id = d.Id" & vbNewLine & _
            "      And (d.站点='" & gstrNodeNo & "' Or d.站点 is Null)" & vbNewLine & _
            "      And b.护理等级id = e.Id(+) And Nvl(c.附加床位, 0) = 0 And c.终止时间 Is Null" & vbNewLine & _
            "      And ((c.开始原因 in(1,3) And Exists(Select 1 From 病区科室对应 H Where c.科室id = h.科室id And h.病区id = [1])) or (c.开始原因=15 And c.病区id = [1]))" & vbNewLine & _
            "      And ((c.开始原因 = 1 And b.状态 = 1) Or (c.开始原因 in (3,15) And c.开始时间 Is Null And b.状态 = 2)) "
    
        '在院病人
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Decode(B.状态,3,4,DECODE(B.出院病床, NULL, 3.1,DECODE(B.状态,2,3.2,3))) as 排序," & _
            " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
            " Decode(B.状态,3,'预出院病人',DECODE(B.出院病床, NULL, '家庭病床',DECODE(B.状态,2,'预转科病人', '在院病人'))) as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,A.姓名,A.性别,B.年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
            " lpad(nvl(B.出院病床,' '),10,' ') as 床号,E.名称 as 护理等级,B.费别,B.当前病况,B.入院日期,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(B.路径状态,-1) 路径状态,trunc(sysdate)-trunc(b.入院日期)+1 as 住院天数,z.颜色" & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z,在院病人 R" & _
            " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And A.住院次数=B.主页ID And Nvl(B.主页ID,0)<>0 And Nvl(B.状态,0)<>1" & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL And A.病人ID=R.病人ID And R.病区ID=[1]"
            
        '出院病人:出院病人可能已有多次住院
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Decode(B.出院方式,'死亡',6,5) as 排序," & _
            " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
            " Decode(B.出院方式,'死亡','死亡病人','出院病人') as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,A.姓名,A.性别,B.年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
            " lpad(nvl(B.出院病床,' '),10,' ') AS 床号,E.名称 as 护理等级,B.费别,B.当前病况,B.入院日期,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(B.路径状态,-1) 路径状态,trunc(b.出院日期)-trunc(b.入院日期)+1 as 住院天数,z.颜色" & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z" & _
            " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.状态=0" & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And B.当前病区ID+0=[1] And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And B.出院日期 Between [2] And [3] And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
        '转出病人:在院,医生和床号显示本科转出前的
    
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Distinct 7 as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,'转出病人' as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,A.姓名,A.性别,B.年龄,D.名称 as 科室,C.科室ID,C.经治医师 as 住院医师,B.责任护士,B.病案状态," & _
            " lpad(nvl(C.床号,' '),10,' ') as 床号,E.名称 as 护理等级,B.费别,B.当前病况,B.入院日期,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(B.路径状态,-1) 路径状态,trunc(sysdate)-trunc(b.入院日期)+1 as 住院天数,z.颜色" & _
            " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D,收费项目目录 E,病人类型 Z" & _
            " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.护理等级ID=E.ID(+)" & _
            " And B.病人ID=C.病人ID And B.主页ID=C.主页ID" & _
            " And B.当前病区ID<>[1] And C.病区ID+0=[1] And C.科室ID=D.ID" & _
            " And Nvl(C.附加床位,0)=0 And C.终止原因 In(3,15) And C.终止时间 Between Sysdate-[4] And Sysdate" & _
            " And Nvl(B.状态,0)<>2 And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
    
        '提取病人信息
        strSQL = "SELECT A.排序,A.排序2,A.类型,A.病人ID,A.主页ID,A.门诊号,A.住院号,A.姓名,A.性别,A.年龄,A.科室,A.科室ID,A.住院医师,A.责任护士,A.病案状态," & _
                " lpad(nvl(A.床号,' '),10,' ') as 床号,A.护理等级,A.费别,A.当前病况,A.入院日期,A.出院日期,A.出院方式,A.病人类型," & _
                " A.状态,A.险类,A.就诊卡号,A.路径状态,A.住院天数,A.颜色" & _
                " From (" & strSQL & ") A,病人护理文件 B,病历文件列表 C" & _
                " Where A.病人ID=B.病人ID and A.主页ID=B.主页ID And nvl(B.婴儿,0)=0 And B.格式ID=C.ID And C.种类=3 And C.保留=-1"
        strSQL = strSQL & " Order by A.排序,A.床号,A.主页ID DESC"
        
        Screen.MousePointer = 11
        On Error GoTo Errhand
        Set rsThis = zlDatabase.OpenSQLRecord(strSQL, "提取病人列表", T_Info.lng病区ID, _
            CDate(Format(mdtOutBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(mdtOutEnd, "yyyy-MM-dd 23:59:59")), _
            mintChange)
        '复制数据
        Call CopyReocrd(rsThis)
        Screen.MousePointer = 0
    Else
        If rsThis.State = 1 Then
            Set mrsPati = rsThis.Clone
        Else
            GoTo ErrGO
        End If
    End If
    Exit Sub
Errhand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitReprotControl()
 '初始化病人选择器
    Dim objCol As ReportColumn
    With rptPati
        Set objCol = .Columns.Add(c_选择, "", 0, False): objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_图标, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_排序, "排序", 0, True)
        Set objCol = .Columns.Add(c_状态, "状态", 0, True)
        Set objCol = .Columns.Add(c_床号, "床号", 40, True)
        Set objCol = .Columns.Add(C_病人ID, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_主页ID, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(c_年龄, "年龄", 60, True)
        Set objCol = .Columns.Add(c_住院号, "住院号", 60, True)
        Set objCol = .Columns.Add(c_入院日期, "入院日期", 120, True)
        Set objCol = .Columns.Add(c_出院日期, "出院日期", 120, True)
        For Each objCol In .Columns
            If objCol.Index <> c_选择 Then
                objCol.Editable = False
            Else
                objCol.Sortable = True
                objCol.Editable = True
            End If
            objCol.Groupable = (objCol.Index = c_状态)
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有病人..."
        End With
        
        .PreviewMode = False
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgRPT
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(c_排序)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(c_床号)
    End With
End Sub

