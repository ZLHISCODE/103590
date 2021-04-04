VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchiveView 
   AutoRedraw      =   -1  'True
   Caption         =   "电子病案查阅"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12660
   Icon            =   "frmArchiveView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   11670
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   44
      Top             =   1605
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   0
      Width           =   3165
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3765
      ScaleHeight     =   975
      ScaleWidth      =   7695
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   135
      Width           =   7695
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " 基本就诊信息 "
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   60
         TabIndex        =   5
         Top             =   75
         Width           =   7500
         Begin VB.Frame fraIn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   195
            TabIndex        =   24
            Top             =   255
            Visible         =   0   'False
            Width           =   7170
            Begin VB.Label lbl类型zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   4770
               TabIndex        =   42
               Top             =   0
               Width           =   1080
            End
            Begin VB.Label lbl类型zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "类型:"
               Height          =   180
               Index           =   0
               Left            =   4305
               TabIndex        =   41
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl住院号zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号:"
               Height          =   180
               Index           =   0
               Left            =   1560
               TabIndex        =   40
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lbl姓名zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl付款zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "付款:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   38
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl床号zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "床号:"
               Height          =   180
               Index           =   0
               Left            =   3150
               TabIndex        =   37
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl医保号zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医保号:"
               Height          =   180
               Index           =   0
               Left            =   5940
               TabIndex        =   36
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lbl入院zy 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院:"
               Height          =   180
               Index           =   0
               Left            =   4305
               TabIndex        =   35
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl病况zy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病况:"
               Height          =   180
               Index           =   0
               Left            =   3150
               TabIndex        =   34
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl护理zy 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "护  理:"
               Height          =   180
               Index           =   0
               Left            =   1560
               TabIndex        =   33
               Top             =   255
               Width           =   630
            End
            Begin VB.Label lbl护理zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2190
               TabIndex        =   32
               Top             =   255
               Width           =   900
            End
            Begin VB.Label lbl病况zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H000000FF&
               Height          =   180
               Index           =   1
               Left            =   3585
               TabIndex        =   31
               Top             =   255
               Width           =   675
            End
            Begin VB.Label lbl入院zy 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   4770
               TabIndex        =   30
               Top             =   255
               Width           =   90
            End
            Begin VB.Label lbl医保号zy 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00008000&
               Height          =   180
               Index           =   1
               Left            =   6600
               TabIndex        =   29
               Top             =   0
               Width           =   90
            End
            Begin VB.Label lbl床号zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3585
               TabIndex        =   28
               Top             =   0
               Width           =   675
            End
            Begin VB.Label lbl付款zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   435
               TabIndex        =   27
               Top             =   255
               Width           =   1080
            End
            Begin VB.Label lbl姓名zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   435
               TabIndex        =   26
               Top             =   0
               Width           =   1080
            End
            Begin VB.Label lbl住院号zy 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2190
               TabIndex        =   25
               Top             =   0
               Width           =   900
            End
         End
         Begin VB.Frame fraOut 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   195
            TabIndex        =   6
            Top             =   255
            Visible         =   0   'False
            Width           =   7170
            Begin VB.Label lbl急 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "急"
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   21.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   435
               Left            =   6705
               TabIndex        =   23
               Top             =   0
               Visible         =   0   'False
               Width           =   435
            End
            Begin VB.Label lbl挂号单mz 
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3870
               TabIndex        =   22
               Top             =   0
               Width           =   1065
            End
            Begin VB.Label lbl挂号单mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "挂号单:"
               Height          =   180
               Index           =   0
               Left            =   3255
               TabIndex        =   21
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lbl医生mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2385
               TabIndex        =   20
               Top             =   0
               Width           =   780
            End
            Begin VB.Label lbl医生mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医生:"
               Height          =   180
               Index           =   0
               Left            =   1935
               TabIndex        =   19
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl社区号mz 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00008000&
               Height          =   180
               Index           =   1
               Left            =   5655
               TabIndex        =   18
               Top             =   255
               Width           =   90
            End
            Begin VB.Label lbl社区号mz 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "社区号:"
               Height          =   180
               Index           =   0
               Left            =   5025
               TabIndex        =   17
               Top             =   255
               Width           =   630
            End
            Begin VB.Label lbl门诊号mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊号:"
               Height          =   180
               Index           =   0
               Left            =   3240
               TabIndex        =   16
               Top             =   255
               Width           =   630
            End
            Begin VB.Label lbl姓名mz 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   450
            End
            Begin VB.Label lbl费别mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "费别:"
               Height          =   180
               Index           =   0
               Left            =   1935
               TabIndex        =   14
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl医保号mz 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医保号:"
               Height          =   180
               Index           =   0
               Left            =   5025
               TabIndex        =   13
               Top             =   0
               Width           =   630
            End
            Begin VB.Label lbl医保号mz 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00008000&
               Height          =   180
               Index           =   1
               Left            =   5655
               TabIndex        =   12
               Top             =   0
               Width           =   90
            End
            Begin VB.Label lbl费别mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   2385
               TabIndex        =   11
               Top             =   255
               Width           =   765
            End
            Begin VB.Label lbl姓名mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   450
               TabIndex        =   10
               Top             =   0
               Width           =   1425
            End
            Begin VB.Label lbl门诊号mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3870
               TabIndex        =   9
               Top             =   255
               Width           =   1095
            End
            Begin VB.Label lbl付款mz 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "付款:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   8
               Top             =   255
               Width           =   450
            End
            Begin VB.Label lbl付款mz 
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "#"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   450
               TabIndex        =   7
               Top             =   255
               Width           =   1455
            End
         End
      End
   End
   Begin MSComctlLib.TreeView tvwArchive 
      Height          =   5865
      Left            =   315
      TabIndex        =   3
      Top             =   1170
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   10345
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   0
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   3660
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   1515
      Width           =   45
   End
   Begin XtremeSuiteControls.TabControl tbcArchive 
      Height          =   6315
      Left            =   3900
      TabIndex        =   1
      Top             =   1605
      Width           =   7365
      _Version        =   589884
      _ExtentX        =   12991
      _ExtentY        =   11139
      _StockProps     =   64
   End
   Begin XtremeSuiteControls.TabControl tbcHistory 
      Height          =   7245
      Left            =   240
      TabIndex        =   0
      Top             =   735
      Width           =   3210
      _Version        =   589884
      _ExtentX        =   5662
      _ExtentY        =   12779
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   120
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":058A
            Key             =   "住院"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":6DEC
            Key             =   "门诊"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":7386
            Key             =   "object_report"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":7920
            Key             =   "object_case"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":7EBA
            Key             =   "object_tend"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":8454
            Key             =   "object_first"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":89EE
            Key             =   "object_advice"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":8F88
            Key             =   "object_file"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":9522
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArchiveView.frx":FD84
            Key             =   "Path"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmArchiveView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjRichEMR As Object
Private mclsOutAdvices As clsDockOutAdvices
Private mclsInAdvices As clsDockInAdvices
Private mclsDockAduits As clsDockAduits
Private mclsPath As zlCISPath.clsDockPath
Private WithEvents mclsTendsNew As zl9TendFile.clsTendFile    '新版护士工作站
Attribute mclsTendsNew.VB_VarHelpID = -1
Private mclsArchive As zlMedRecPage.clsArchive '电子病案查阅窗体类

Private mlng病人ID  As Long
Private mlng就诊ID As Long '病人当前或者最后的就诊ID，门诊为挂号ID,住院号主页ID
Private mstr挂号单 As String
Private mlng科室ID As Long
Private mlng病区ID As Long
Private mblnMoved As Boolean
Private mblnNewTends As Boolean
Private mrsData As ADODB.Recordset

Private mcolSubForm As Collection
Private mblnTabTmp As Boolean
Private mlngPre就诊ID As Long

Public Sub ShowArchive(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, Optional ByVal blnModal As Boolean)
'功能：公共接口方法，类似 ShowMe方法
    
    mstr挂号单 = "": mlngPre就诊ID = 0
    mblnMoved = False: mblnNewTends = False
    mlng病人ID = lng病人ID
    mlng就诊ID = lng就诊ID
    
    Me.Show IIf(blnModal, 1, 0), frmParent
End Sub

Private Sub InitBasicData()
'功能：初始化一些基本数据，如下拉列表加载等
    Dim StrSQL As String
    Dim objTab As TabControlItem
    Dim strTmp As String
    
    Screen.MousePointer = 11
    LockWindowUpdate Me.hwnd
        
    Call cboDept.Clear
    Call tbcHistory.RemoveAll
    
    StrSQL = " Select ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,0 as 数据转出 From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
        " Union ALL" & _
        " Select ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,1 as 数据转出 From H病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
        " Union ALL" & _
        " Select 主页ID as 就诊ID,Null,入院日期 as 开始时间,出院日期,出院科室ID,数据转出 From 病案主页 Where 病人ID=[1] And Nvl(主页ID,0)<>0"
    StrSQL = "Select A.就诊ID,A.NO,A.开始时间,A.结束时间,B.名称 as 科室,A.数据转出 From (" & StrSQL & ") A,部门表 B Where A.科室ID=B.ID Order by 开始时间 Desc"
    
    On Error GoTo errH
    Set mrsData = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID)
    
    Do While Not mrsData.EOF
        strTmp = IIf(IsNull(mrsData!NO), "第" & mrsData!就诊id & "次住院", "门诊就诊") & ":" & mrsData!科室 & "," & Format(mrsData!开始时间, "yyyy-MM-dd HH:mm") & _
            IIf(Not IsNull(mrsData!结束时间), "～" & Format(mrsData!结束时间, "yyyy-MM-dd HH:mm"), "")
        
        If mrsData.AbsolutePosition = 1 Then
            Set objTab = tbcHistory.InsertItem(tbcHistory.ItemCount, strTmp, tvwArchive.hwnd, IIf(IsNull(mrsData!NO), 0, 1))
                objTab.Tag = mrsData!就诊id & "," & mrsData!NO & "," & Nvl(mrsData!数据转出, 0)
        End If
        
        cboDept.AddItem strTmp
        cboDept.ItemData(cboDept.NewIndex) = Val(mrsData!就诊id)
        
        mrsData.MoveNext
    Loop
        
    If cboDept.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboDept.hwnd, 0)
        Call cboDept_Click
    End If
    LockWindowUpdate 0
    Screen.MousePointer = 0
    Exit Sub
errH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim objTab As TabControlItem
    Dim frmTendBody As Object
    Dim intIdx As Integer
     
    picInfo.BackColor = fraLR.BackColor
    fraInfo.BackColor = picInfo.BackColor
    fraIn.BackColor = picInfo.BackColor
    fraOut.BackColor = picInfo.BackColor
    
    '初始对象
    '------------------------------------------------------------------------------------------------------------------
    If Not gobjEmr Is Nothing Then
        Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "新版病历", False)
        If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
    End If
    If mclsArchive Is Nothing Then
        Set mclsArchive = New zlMedRecPage.clsArchive
        Call mclsArchive.InitArchiveMedRec(gcnOracle, glngSys)
    End If
    Set mclsOutAdvices = New clsDockOutAdvices
    Set mclsInAdvices = New clsDockInAdvices
    Set mclsDockAduits = New clsDockAduits
    Set mclsPath = New clsDockPath
    Set mclsTendsNew = New zl9TendFile.clsTendFile
    Call mclsTendsNew.InitTendFile(gcnOracle, glngSys)
    Set frmTendBody = mclsDockAduits.zlGetFormTendBody
    Call zlControl.FormSetCaption(frmTendBody, False, False)
    
    '子窗体
    '-----------------------------------------------------
    Set mcolSubForm = New Collection
    mcolSubForm.Add mclsArchive.zlGetForm(0), "_门诊首页"
    mcolSubForm.Add mclsArchive.zlGetForm(1), "_住院首页"
    mcolSubForm.Add mclsDockAduits.zlGetFormEPR, "_病历信息"
    mcolSubForm.Add mclsOutAdvices.zlGetForm, "_门诊医嘱"
    mcolSubForm.Add mclsInAdvices.zlGetForm, "_住院医嘱"
    mcolSubForm.Add frmTendBody, "_体温记录单"
    mcolSubForm.Add mclsDockAduits.zlGetFormTendFile, "_护理记录单"
    mcolSubForm.Add mclsPath.zlGetForm, "_临床路径"
    mcolSubForm.Add mclsTendsNew.zlGetfrmInTendFile, "_新版护理"
    If Not mobjRichEMR Is Nothing Then mcolSubForm.Add mobjRichEMR.zlGetForm, "_电子病历"
    
    With tbcArchive
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .Color = xtpTabColorOffice2003
            .Layout = xtpTabLayoutAutoSize
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        '隐式出发Form_Load采取添加一个图片方式，切换的时候再依次重新加载
        Set objTab = .InsertItem(intIdx, "门诊首页", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "住院首页", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "病历信息", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "门诊医嘱", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "住院医嘱", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "体温记录单", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "护理记录单", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "临床路径", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        Set objTab = .InsertItem(intIdx, "新版护理", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
            objTab.Visible = False: intIdx = intIdx + 1
        If Not mobjRichEMR Is Nothing Then
            Set objTab = .InsertItem(intIdx, "电子病历", picTmp.hwnd, 0): objTab.Tag = objTab.Caption
                objTab.Visible = False: intIdx = intIdx + 1
        End If
    End With
    
    '就诊历史
    '-----------------------------------------------------
    With tbcHistory
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
            .DisableLunaColors = False
            .BoldSelected = True
            .HotTracking = True
            .ShowIcons = True
        End With
        .SetImageList ils16
    End With
    Call InitBasicData
    Call RestoreWinState(Me, App.ProductName)
    If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Me.cboDept.Width = tbcHistory.Width
    Me.tbcHistory.Top = cboDept.Height
    Me.tbcHistory.Left = 0
    
    Me.tbcHistory.Height = Me.ScaleHeight - cboDept.Height
    
    Me.fraLR.Top = 0
    Me.fraLR.Left = Me.tbcHistory.Width
    Me.fraLR.Height = Me.ScaleHeight
    
    Me.picInfo.Top = 0
    Me.picInfo.Left = Me.fraLR.Left + Me.fraLR.Width
    Me.picInfo.Width = Me.ScaleWidth - Me.tbcHistory.Width - Me.fraLR.Width
    
    Me.tbcArchive.Left = Me.fraLR.Left + Me.fraLR.Width
    Me.tbcArchive.Top = Me.picInfo.Top + Me.picInfo.Height
    Me.tbcArchive.Width = Me.ScaleWidth - Me.tbcHistory.Width - Me.fraLR.Width
    Me.tbcArchive.Height = Me.ScaleHeight - Me.picInfo.Height
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    '强行Unload,不然不会激活子窗体的事件
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next

    Set mclsArchive = Nothing
    Set mclsDockAduits = Nothing
    Set mclsOutAdvices = Nothing
    Set mclsInAdvices = Nothing
    Set mclsPath = Nothing
    Set mclsTendsNew = Nothing
    Set mobjRichEMR = Nothing
    Set mrsData = Nothing
End Sub

Private Sub cboDept_Click()

    If cboDept.Text = "" Then Exit Sub
    
    If mlngPre就诊ID = cboDept.ItemData(cboDept.ListIndex) Then Exit Sub
    
    mlngPre就诊ID = cboDept.ItemData(cboDept.ListIndex)
    
    mlng就诊ID = mlngPre就诊ID
    
    mrsData.Filter = "就诊ID= " & mlng就诊ID
    
    If Not mrsData.EOF Then
        mstr挂号单 = Nvl(mrsData!NO, "")
        mblnMoved = Val(Nvl(mrsData!数据转出, "")) = 1
    End If
    '显示基本信息
    If mstr挂号单 <> "" Then
        Call ShowOutPatiInfo
    Else
        Call ShowInPatiInfo
    End If
    
    fraOut.Visible = mstr挂号单 <> ""
    fraIn.Visible = mstr挂号单 = ""

    '显示档案目录
    Me.tbcHistory(0).Caption = cboDept.Text
    Call ShowArchiveTree
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
    Call Form_Resize
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If fraLR.Left + X < 1000 Or fraLR.Left + X > Me.ScaleWidth - 3000 Then Exit Sub
        
        Me.tbcHistory.Width = tbcHistory.Width + X
        Call Form_Resize
    End If
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left * 3
    
    fraIn.Width = fraInfo.Width - fraIn.Left * 2
    fraOut.Width = fraIn.Width
    lbl急.Left = fraOut.Width - lbl急.Width - 60
End Sub

Private Sub tbcArchive_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：刷新子窗体界面及数据
'说明：仅在人为切换界面卡片激活
    Dim Index As Long, objItem As TabControlItem
    
    If mblnTabTmp Then Exit Sub
    If Item.Tag = "" Then Exit Sub '初始添卡时,还没赋值
    If Item.Handle = picTmp.hwnd Then
        Screen.MousePointer = 11
        Index = Item.Index
        mblnTabTmp = True
        On Error GoTo errH
        Select Case Item.Tag
            Case "门诊首页"
                Set objItem = tbcArchive.InsertItem(Index, "门诊首页", mcolSubForm("_门诊首页").hwnd, 0)
                objItem.Tag = "门诊首页"
            Case "住院首页"
                Set objItem = tbcArchive.InsertItem(Index, "住院首页", mcolSubForm("_住院首页").hwnd, 0)
                objItem.Tag = "住院首页"
            Case "病历信息"
                Set objItem = tbcArchive.InsertItem(Index, "病历信息", mcolSubForm("_病历信息").hwnd, 0)
                objItem.Tag = "病历信息"
            Case "门诊医嘱"
                Set objItem = tbcArchive.InsertItem(Index, "门诊医嘱", mcolSubForm("_门诊医嘱").hwnd, 0)
                objItem.Tag = "门诊医嘱"
            Case "住院医嘱"
                Set objItem = tbcArchive.InsertItem(Index, "住院医嘱", mcolSubForm("_住院医嘱").hwnd, 0)
                objItem.Tag = "住院医嘱"
            Case "体温记录单"
                Set objItem = tbcArchive.InsertItem(Index, "体温记录单", mcolSubForm("_体温记录单").hwnd, 0)
                objItem.Tag = "体温记录单"
            Case "护理记录单"
                Set objItem = tbcArchive.InsertItem(Index, "护理记录单", mcolSubForm("_护理记录单").hwnd, 0)
                objItem.Tag = "护理记录单"
            Case "临床路径"
                Set objItem = tbcArchive.InsertItem(Index, "临床路径", mcolSubForm("_临床路径").hwnd, 0)
                objItem.Tag = "临床路径"
            Case "新版护理"
                Set objItem = tbcArchive.InsertItem(Index, "新版护理", mcolSubForm("_新版护理").hwnd, 0)
                objItem.Tag = "新版护理"
            Case "电子病历"
                Set objItem = tbcArchive.InsertItem(Index, "电子病历", mcolSubForm("_电子病历").hwnd, 0)
                objItem.Tag = "电子病历"
        End Select
        Call tbcArchive.RemoveItem(Index + 1)
        objItem.Selected = True
        mblnTabTmp = False
        Screen.MousePointer = 0
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbcHistory_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If tbcHistory.Tag = "don't refresh" Then Exit Sub
    If Item.Tag = "" Then Exit Sub
    
    mlngPre就诊ID = 0
    mlng就诊ID = Val(Split(Item.Tag, ",")(0))
    mstr挂号单 = Split(Item.Tag, ",")(1)
    mblnMoved = Val(Split(Item.Tag, ",")(2)) = 1
    
    '显示基本信息
    If mstr挂号单 <> "" Then
        Call ShowOutPatiInfo
    Else
        Call ShowInPatiInfo
    End If
    fraOut.Visible = mstr挂号单 <> ""
    fraIn.Visible = mstr挂号单 = ""
    
    '显示档案目录
    Call ShowArchiveTree
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
    Call Form_Resize
End Sub

Private Sub ShowArchiveTab(ByVal strShow As String, ByVal strCaption As String)
'功能：切换显示不同的档案页面，或者清空界面
    Dim i As Long
    
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive(i).Tag = strShow Then
            tbcArchive(i).Caption = strCaption
            If Not tbcArchive(i).Visible Then
                tbcArchive(i).Visible = True
                tbcArchive(i).Selected = True
                Exit For
            End If
        End If
    Next
    
    For i = 0 To tbcArchive.ItemCount - 1
        If tbcArchive(i).Tag <> strShow Then
            If tbcArchive(i).Visible Then tbcArchive(i).Visible = False
        End If
    Next
End Sub

Private Sub tvwArchive_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim arrPar As Variant
    Dim intSel As Integer
        
    If tvwArchive.Tag = Node.Key Then Exit Sub
    
    LockWindowUpdate Me.hwnd
    
    arrPar = Split(Node.Tag, ";")
    
    If Node.Key Like "R1K*" Or Node.Key Like "R2K*" Or Node.Key Like "R4K*" Or Node.Key Like "R5K*" Or Node.Key Like "R6K*" Or Node.Key Like "R7K*" Then
        Call ShowArchiveTab("病历信息", Node.Text)
    End If
    If Node.Key = "R11" Then
        Call ShowArchiveTab(IIf(mstr挂号单 <> "", "门诊首页", "住院首页"), tbcHistory.Selected.Caption)
        Call mclsArchive.zlRefresh(IIf(mstr挂号单 <> "", 0, 1), mlng病人ID, mlng就诊ID, mblnMoved)
    ElseIf Node.Key = "R12" Then '医嘱记录
        If mstr挂号单 <> "" Then
            Call ShowArchiveTab("门诊医嘱", tbcHistory.Selected.Caption)
            Call mclsOutAdvices.zlRefresh(mlng病人ID, mstr挂号单, False, mblnMoved)
        Else
            Call ShowArchiveTab("住院医嘱", tbcHistory.Selected.Caption)
            Call mclsInAdvices.zlRefresh(mlng病人ID, mlng就诊ID, mlng病区ID, mlng科室ID, 0, mblnMoved)
        End If
    ElseIf Node.Key Like "R1K*" Then '门诊病历
        Call mclsDockAduits.zlRefresh(1, Val(arrPar(0)))
    ElseIf Node.Key Like "R2K*" Then '住院病历
        Call mclsDockAduits.zlRefresh(2, Val(arrPar(0)))
    ElseIf Node.Key Like "R3K*" Then '护理记录
        If UBound(arrPar) >= 1 Then
            If mblnNewTends = False Then
                If Val(arrPar(1)) = -1 Then
                    Call ShowArchiveTab("体温记录单", Node.Text)
                    Call mclsDockAduits.zlRefreshTendBody(mlng病人ID, mlng就诊ID, Val(arrPar(0)), 0)
                Else
                    Call ShowArchiveTab("护理记录单", Node.Text)
                    Call mclsDockAduits.zlRefresh(3, Val(arrPar(3)), mlng病人ID, mlng就诊ID, Val(arrPar(0)), CStr(arrPar(2)))
                End If
            Else
                Select Case Val(arrPar(1))
                    Case -1
                        intSel = 0
                    Case 1
                        intSel = 2
                    Case Else
                        intSel = 1
                End Select
                Call ShowArchiveTab("新版护理", Node.Text)
                Call mclsTendsNew.zlRefreshTendFile(mlng病人ID, mlng就诊ID, Val(arrPar(4)), Val(arrPar(0)), False, IIf(glngModul = p住院医生站, True, False), intSel, Val(arrPar(3)), 1)
            End If
        End If
    ElseIf Node.Key Like "R4K*" Then '护理病历
        Call mclsDockAduits.zlRefresh(4, Val(arrPar(0)))
    ElseIf Node.Key Like "R5K*" Then '疾病证明
        Call mclsDockAduits.zlRefresh(5, Val(arrPar(0)))
    ElseIf Node.Key Like "R6K*" Then '知情文件
        Call mclsDockAduits.zlRefresh(6, Val(arrPar(0)))
    ElseIf Node.Key Like "R7K*" Then '诊疗报告
        Call mclsDockAduits.zlRefresh(7, Val(arrPar(0)))
    ElseIf Node.Key = "R8" Then
        If mstr挂号单 = "" Then
            Call ShowArchiveTab("临床路径", Node.Text)
            Call mclsPath.zlRefreshReadOnly(mlng病人ID, mlng就诊ID)
        End If
    ElseIf InStr(Node.Key, "R") = 0 And Len(Node.Tag) >= 32 Then
        'EMR病历预览
        If Not mobjRichEMR Is Nothing Then
            Call ShowArchiveTab("电子病历", Node.Text)
            If InStr(Node.Tag, "|") > 0 Then
                Call mobjRichEMR.zlShowDoc(Split(Node.Tag, "|")(0), Split(Node.Tag, "|")(1))
            Else
                Call mobjRichEMR.zlShowDoc(Node.Tag, "")
            End If
        End If
    Else
        LockWindowUpdate 0
        Exit Sub
    End If
    
    tvwArchive.Tag = Node.Key
    LockWindowUpdate 0
    If tvwArchive.Visible And tvwArchive.Enabled Then tvwArchive.SetFocus
End Sub

Private Function ShowArchiveTree() As Boolean
'功能：显示病人档案树形目录
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, objNode As Node, strSQL1 As String
    Dim blnPath As Boolean
    Dim strSel As String
    
    Screen.MousePointer = 11
    
    If Not tvwArchive.SelectedItem Is Nothing Then
        If tvwArchive.SelectedItem.Key = "R11" Or tvwArchive.SelectedItem.Key = "R12" Then
            strSel = Split(tvwArchive.SelectedItem.Key, "K")(0)
        End If
    End If
    
    '病人科室存在可用的临床路径时，显示临床路径记录
    If mstr挂号单 = "" Then
        If GetInsidePrivs(p临床路径应用) <> "" Then
            blnPath = HavePath(mlng科室ID)
        End If
    End If
    
    On Error GoTo errH
    '1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-疾病证明;6-知情文件;7-诊疗报告,11-首页信息,12-医嘱记录,13-临床路径
    StrSQL = _
        " Select * From (" & _
            " Select 'R11' As ID, '' As 上级id, '首页信息' As 名称, '' As 参数,1 As 末级,'object_first' As 图标,'01' As 排序 From Dual Union All" & _
            " Select 'R12' As ID, '' As 上级id, '医嘱记录' As 名称, '' As 参数,1 As 末级,'object_advice' As 图标,'02' As 排序 From Dual Union All" & _
            " Select 'R1' As ID, '' As 上级id, '门诊病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'03' As 排序 From Dual Where [3]=0 Union All" & _
            " Select 'R2' As ID, '' As 上级id, '住院病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'04' As 排序 From Dual Where [3]=1 Union All" & _
            " Select 'R3' As ID, '' As 上级id, '护理记录' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'05' As 排序 From Dual Where [3]=1 Union All" & _
            " Select 'R4' As ID, '' As 上级id, '护理病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'06' As 排序 From Dual Where [3]=1 Union All" & _
            " Select 'R7' As ID, '' As 上级id, '诊疗报告' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'07' As 排序 From Dual Union All" & _
            " Select 'R5' As ID, '' As 上级id, '疾病证明' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'08' As 排序 From Dual Union All" & _
            " Select 'R6' As ID, '' As 上级id, '知情文件' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'09' As 排序 From Dual" & _
            IIf(blnPath, " Union All Select 'R8' As ID, '' As 上级id, '临床路径' As 名称, '' As 参数,0 As 末级,'Path' As 图标,'10' As 排序 From Dual", "")
    '病历部分
    'ID=上级ID+K病历ID,医嘱ID,0
    '参数=病历ID;医嘱ID
    StrSQL = StrSQL & " Union All" & _
        " Select A.上级id||'K'||Trim(To_Char(A.ID))||','||Trim(To_Char(Nvl(A.医嘱id,0)))||',0' As ID,A.上级id," & _
        "       Decode(A.医嘱id,Null,A.名称||'('||To_Char(A.创建时间, 'YYYY-MM-DD')||')',A.名称||'：'||B.医嘱内容||'('||To_Char(A.创建时间, 'YYYY-MM-DD')||')') As 名称," & _
        "       Trim(To_Char(A.ID))||';'||Decode(A.医嘱id,Null,'0',Trim(To_Char(A.医嘱id))) As 参数," & _
        "       1 As 末级,Decode(病历种类,1,'object_case',2,'object_case',4,'object_case',7,'object_report','object_file') As 图标,排序 " & _
        " From (Select A.ID, 'R'||A.病历种类 As 上级id, A.病历名称 As 名称,C.医嘱id,A.病历种类,A.创建时间,To_Char(A.创建时间,'YYYY-MM-DD HH24:MI:SS') As 排序" & _
        "       From 电子病历记录 A,病人医嘱报告 C " & _
        "       Where A.病人id = [1] And A.主页id = [2] And (A.病人来源=2 And [3]=1 Or Nvl(A.病人来源,0)<>2 And [3]=0)" & _
        "           And C.病历id(+)=A.ID And A.病历种类 In (1, 2, 3, 4, 5, 6, 7)" & _
        "       ) A,病人医嘱记录 B Where A.医嘱id=B.Id(+)"
    '护理部分
    'ID=上级ID+K文件ID,0,科室ID
    '参数=科室ID;保留;开始～截止;文件ID
    '检查本次病人是使用的是老板还是新版
    strSQL1 = "Select 1 From 病人护理记录 A Where a.病人id = [1] And a.主页id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL1, "检查是否存在老板数据", mlng病人ID, mlng就诊ID)
    If rsTmp.RecordCount > 0 Then
        mblnNewTends = False
        StrSQL = StrSQL & " Union All" & _
            " Select 'R3K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.科室Id)) As ID,'R3' As 上级id," & _
            "       A.名称||'('||B.名称||'：'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI') || '～' ||To_Char(A.截止, 'YYYY-MM-DD HH24:MI') || ')' As 名称," & _
            "       Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI')||'～'||To_Char(A.截止, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID)) As 参数," & _
            "       1 As 末级,'object_tend' As 图标,To_Char(a.开始,'YYYY-MM-DD HH24:MI:SS') As 排序" & _
            " From (" & _
            "   Select F.ID, F.编号, F.名称, R.开始, R.截止, R.科室id, 保留" & _
            "   From (" & _
            "       Select ID, 编号, 名称, 3 As 护理级别, 通用, 0 As 科室id, 保留" & _
            "          From 病历文件列表 Where 种类 = 3 And 保留 < 0" & _
            "       Union All" & _
            "       Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, A.科室id, L.保留" & _
            "          From 病历页面格式 F, 病历文件列表 L, 病历应用科室 A" & _
            "          Where L.种类 = 3 And L.保留 = 0 And L.种类 = F.种类 And L.编号 = F.编号 And L.ID = A.文件id(+)" & _
            "       ) F,(" & _
            "       Select R.科室id, Nvl(Min(R.护理级别), 3) As 护理级别, Min(R.发生时间) As 开始, Max(R.发生时间) As 截止" & _
            "          From 病人护理记录 R" & _
            "          Where R.病人来源 = 2 And R.病人id = [1] And Nvl(R.主页id, 0) = [2] And Nvl(R.婴儿, 0) = 0" & _
            "          Group By R.科室id" & _
            "       ) R" & _
            "       Where (F.通用 = 1 Or F.通用 = 2 And F.科室id = R.科室id) And F.护理级别 >= R.护理级别" & _
            "   ) A, 部门表 B Where A.科室id = B.ID)" & _
            "Order By Decode(上级id,Null,' ',上级id),排序"
    Else
        mblnNewTends = True
        StrSQL = StrSQL & " Union All" & _
                " Select 'R3K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.科室Id)) As ID,'R3' As 上级id," & vbNewLine & _
                "     A.名称||'('||B.名称||'：'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI') || '～' ||To_Char(A.截止, 'YYYY-MM-DD HH24:MI') || ')' As 名称," & vbNewLine & _
                "      Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI')||'～'||To_Char(A.截止, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID))||';'||Trim(To_Char(A.婴儿)) As 参数," & vbNewLine & _
                "       1 As 末级,'object_tend' As 图标,To_Char(a.开始,'YYYY-MM-DD HH24:MI:SS') As 排序" & vbNewLine & _
                " From (" & vbNewLine & _
                "   Select R.ID, F.编号, R.名称,R.婴儿, R.开始, NVL(R.截止,nvl(R.时间,R.开始)) 截止, R.科室id, 保留" & vbNewLine & _
                "   From (" & vbNewLine & _
                "       Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, L.保留" & vbNewLine & _
                "          From 病历页面格式 F, 病历文件列表 L" & vbNewLine & _
                "          Where L.种类 = 3 And L.种类 = F.种类 And L.编号 = F.编号 And (L.通用=1 OR L.通用=2)" & vbNewLine & _
                "" & vbNewLine & _
                "       ) F,(" & vbNewLine & _
                "       Select R.ID,R.科室id,R.文件名称 名称,R.格式ID,nvl(R.婴儿,0) 婴儿,Min(R.开始时间) As 开始, Max(R.结束时间) As 截止,MAX(T.发生时间) 时间" & vbNewLine & _
                "          From 病人护理文件 R,病人护理数据 T" & vbNewLine & _
                "          Where R.ID=T.文件ID(+) And R.病人id = [1] And Nvl(R.主页id, 0) = [2]" & vbNewLine & _
                "          Group By R.ID,R.文件名称,R.科室id,R.格式ID,R.婴儿" & vbNewLine & _
                "       ) R" & vbNewLine & _
                "       Where F.ID=R.格式ID" & vbNewLine & _
                "   ) A, 部门表 B Where A.科室id = B.ID And DECODE(A.保留,-1,0,A.婴儿)=A.婴儿)" & vbNewLine & _
                " Order By Decode(上级id,Null,' ',上级id),排序"
    End If
    If mblnMoved Then
        StrSQL = Replace(StrSQL, "电子病历记录", "H电子病历记录")
        StrSQL = Replace(StrSQL, "病人护理记录", "H病人护理记录")
        StrSQL = Replace(StrSQL, "病人医嘱记录", "H病人医嘱记录")
        StrSQL = Replace(StrSQL, "病人医嘱报告", "H病人医嘱报告")
        StrSQL = Replace(StrSQL, "病人护理文件", "H病人护理文件")
        StrSQL = Replace(StrSQL, "病人护理数据", "H病人护理数据")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng就诊ID, IIf(mstr挂号单 = "", 1, 0))
    
    tvwArchive.Tag = ""
    tvwArchive.Nodes.Clear
            
    Do While Not rsTmp.EOF
        If Nvl(rsTmp!上级ID) = "" Then
            Set objNode = tvwArchive.Nodes.Add(, , CStr(rsTmp!ID), rsTmp!名称, Nvl(rsTmp!图标))
        Else
            Set objNode = tvwArchive.Nodes.Add(CStr(rsTmp!上级ID), tvwChild, CStr(rsTmp!ID), rsTmp!名称, Nvl(rsTmp!图标))
        End If
        
        objNode.Tag = Nvl(rsTmp!参数)
        objNode.Expanded = True
        
        If tvwArchive.Nodes.Count = 1 Then
            objNode.Selected = True
        ElseIf objNode.Key = strSel Then
            objNode.Selected = True
        End If
        
        rsTmp.MoveNext
    Loop
    
    Set rsTmp = Nothing
    Set rsTmp = GetEmrCISStruct(mlng病人ID, mlng就诊ID)
    
    If Not rsTmp Is Nothing Then
        If rsTmp.State = ADODB.adStateOpen Then
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Do Until rsTmp.EOF
                    Set objNode = tvwArchive.Nodes.Add(rsTmp!上级ID.Value, tvwChild, rsTmp!ID.Value, rsTmp!名称.Value, rsTmp!图标.Value, rsTmp!图标.Value)
                    objNode.Tag = Nvl(rsTmp!参数) '文档ID[|子文档ID]
                    rsTmp.MoveNext
                Loop
            End If
        End If
    End If
    
    If Not tvwArchive.SelectedItem Is Nothing Then
        tvwArchive.SelectedItem.EnsureVisible
        Call tvwArchive_NodeClick(tvwArchive.SelectedItem)
    End If
    
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowOutPatiInfo() As Boolean
'功能：选择门诊病人某次历史就诊记录时，读取相关的病人信息
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    
    StrSQL = "Select B.Id,B.NO,B.门诊号,B.姓名,B.性别,B.年龄,A.医疗付款方式," & _
        " A.费别,A.险类,A.医保号,B.急诊,B.发生时间,B.执行人,B.执行状态,B.执行时间," & _
        " B.执行部门ID as 科室ID,B.诊室,B.社区,D.社区号,C.名称 as 科室" & _
        " From 病人信息 A,病人挂号记录 B,部门表 C,病人社区信息 D" & _
        " Where A.病人ID=B.病人ID And B.ID=[1] And B.执行部门ID=C.ID" & _
        " And B.病人ID=D.病人ID(+) And B.社区=D.社区(+) And B.记录性质=1 And B.记录状态=1"
    If mblnMoved Then
        StrSQL = Replace(StrSQL, "病人挂号记录", "H病人挂号记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng就诊ID)
    With rsTmp
        '保险病人姓名红色显示
        lbl姓名mz(1).Caption = Nvl(!姓名)
        If Not IsNull(!险类) Then
            lbl姓名mz(1).ForeColor = vbRed
        Else
            lbl姓名mz(1).ForeColor = lbl门诊号mz(1).ForeColor
        End If
        lbl医生mz(1).Caption = Nvl(!执行人)
        lbl挂号单mz(1).Caption = !NO
        lbl门诊号mz(1).Caption = Nvl(!门诊号)
        lbl付款mz(1).Caption = Nvl(!医疗付款方式)
        lbl费别mz(1).Caption = Nvl(!费别)
        lbl医保号mz(1).Caption = Nvl(!医保号)
        lbl社区号mz(1).Caption = Nvl(!社区号)
        lbl急.Visible = Nvl(!急诊, 0) <> 0
        
        mlng科室ID = Nvl(!科室ID, 0)
        mlng病区ID = 0
    End With
    
    ShowOutPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowInPatiInfo() As Boolean
'功能：选择某次住院记录时，读取相关的病人信息
'返回：blnMoved=本次住院病案是否转出了
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    
    StrSQL = "Select NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, NVL(B.年龄,A.年龄) 年龄,B.住院号,B.出院病床,B.医疗付款方式," & _
        " D.信息值 as 医保号,B.险类,B.当前病况,C.名称 as 护理等级,B.入院日期," & _
        " B.出院日期,B.病人类型,B.状态,B.出院科室ID,B.当前病区ID,A.住院次数" & _
        " From 病人信息 A,病案主页 B,收费项目目录 C,病案主页从表 D" & _
        " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2] And B.护理等级ID=C.ID(+)" & _
        " And B.病人ID=D.病人ID(+) And B.主页ID=D.主页ID(+) And D.信息名(+)='医保号'"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng病人ID, mlng就诊ID)
    
    With rsTmp
        '保险病人颜色特殊显示
        lbl姓名zy(1).Caption = Nvl(!姓名)
        lbl姓名zy(1).ForeColor = zlDatabase.GetPatiColor(Nvl(!病人类型))
        
        lbl住院号zy(1).Caption = Nvl(!住院号)
        lbl床号zy(1).Caption = Nvl(!出院病床)
        lbl医保号zy(1).Caption = Nvl(!医保号)
        lbl护理zy(1).Caption = Nvl(!护理等级)
        lbl付款zy(1).Caption = Nvl(!医疗付款方式)
        
        '危重病人病况红色显示
        lbl病况zy(1).Caption = Nvl(!当前病况)
        If Nvl(!当前病况) = "危" Or Nvl(!当前病况) = "重" Or Nvl(!当前病况) = "急" Then
            lbl病况zy(1).ForeColor = vbRed
        Else
            lbl病况zy(1).ForeColor = lbl住院号zy(1).ForeColor
        End If
        
        lbl入院zy(1).Caption = Format(!入院日期, "yyyy-MM-dd HH:mm")
        If Not IsNull(!出院日期) Then
            lbl入院zy(1).Caption = lbl入院zy(1).Caption & "～" & Format(!出院日期, "yyyy-MM-dd HH:mm")
        End If
        
        lbl类型zy(1).Caption = Nvl(!病人类型)
        
        mlng科室ID = Nvl(!出院科室ID, 0)
        mlng病区ID = Nvl(!当前病区ID, 0)
    End With
    
    ShowInPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetEmrCISStruct(ByVal lngPatiID As Long, ByVal lngPageID As Long) As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strExtendTag As String, strReturn As String, StrSQL As String
    
    On Error GoTo errH
    If gobjEmr Is Nothing Then Set GetEmrCISStruct = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtendTag = "" Then Set GetEmrCISStruct = Nothing: Exit Function
    
    '上级ID，ID，名称，参数，图标
    StrSQL = "Select Decode(e.Kind, '02', 'R2', '03', 'R3', '04', 'R7', '05', 'R8', 'R2') 上级id, Nvl(d.Subdoc_Id, Rawtohex(b.Id)) As ID," & vbNewLine & _
                "       d.Subdoc_Id As 子文档id," & vbNewLine & _
                "       Nvl(d.Subdoc_Title, b.Title) ||" & vbNewLine & _
                "        Decode(d.Completor, Null, ''," & vbNewLine & _
                "               '【 ' || d.Completor || ' 在' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || '签名】') As 名称," & vbNewLine & _
                "       Rawtohex(b.Id) || Decode(d.Subdoc_Id, Null, Null, '|' || d.Subdoc_Id) As 参数, 'object_case' As 图标" & vbNewLine & _
                "From Bz_Doc_Log B," & vbNewLine & _
                "     (Select Distinct a.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor" & vbNewLine & _
                "       From Bz_Act_Log A, Bz_Doc_Tasks C" & vbNewLine & _
                "       Where a.Extend_Tag = :etag And a.Id = c.Actlog_Id And c.Real_Doc_Id Is Not Null) D, Antetype_List E" & vbNewLine & _
                "Where b.Actlog_Id = d.Id And d.Real_Doc_Id = b.Id And d.Antetype_Id = e.Id And" & vbNewLine & _
                "      Decode(d.Subdoc_Id, Null, d.Antetype_Id, b.Antetype_Id) = b.Antetype_Id And" & vbNewLine & _
                "      Decode(d.Subdoc_Title, Null, e.Title, d.Subdoc_Title) = e.Title" & vbNewLine & _
                "Order By e.Code, b.Creat_Time, d.Complete_Time"
                
    strReturn = gobjEmr.OpenSQLRecordset(StrSQL, strExtendTag & "^16^etag", rsTemp)
    
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, gstrSysName
        Set GetEmrCISStruct = Nothing: Exit Function
    End If
    
    Set GetEmrCISStruct = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetEMRIn_Tag(ByVal lngPatiID As Long, ByVal lngPageID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
    
    
    StrSQL = "Select Nvl(a.Id, b.Id) ID" & vbNewLine & _
                "From (Select Max(ID) ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 2 And Nvl(附加床位, 0) = 0) A," & vbNewLine & _
                "     (Select Max(ID) ID From 病人变动记录 Where 病人id = [2] And 主页id = [2] And 开始原因 = 1 And Nvl(附加床位, 0) = 0) B"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, "读取病人入院ID", lngPatiID, lngPageID)
    
    If rsTmp Is Nothing Then Exit Function
    If Nvl(rsTmp!ID) = "" Then Exit Function
    GetEMRIn_Tag = "BD_" & rsTmp!ID
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

