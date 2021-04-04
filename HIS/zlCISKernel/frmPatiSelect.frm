VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiSelect 
   Caption         =   "病人选择"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15000
   Icon            =   "frmPatiSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   15000
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消(&E)"
      Height          =   495
      Left            =   12480
      TabIndex        =   26
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   495
      Left            =   10800
      TabIndex        =   25
      Top             =   7680
      Width           =   1575
   End
   Begin VB.PictureBox pic转出 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   5880
      ScaleHeight     =   5415
      ScaleWidth      =   9855
      TabIndex        =   4
      Top             =   4200
      Width           =   9855
      Begin XtremeReportControl.ReportControl rpt转出 
         Height          =   2415
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   3360
         _Version        =   589884
         _ExtentX        =   5927
         _ExtentY        =   4260
         _StockProps     =   0
         BorderStyle     =   2
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.TextBox txtChange 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   900
         MaxLength       =   3
         TabIndex        =   23
         Text            =   "7"
         Top             =   120
         Width           =   285
      End
      Begin VB.Frame fraChange 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   870
         TabIndex        =   22
         Top             =   330
         Width           =   300
      End
      Begin VB.CommandButton cmdRef 
         Caption         =   "刷新"
         Height          =   255
         Left            =   2625
         TabIndex        =   21
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lbl转出 
         AutoSize        =   -1  'True
         Caption         =   "显示最近    天的转出病人"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   150
         Width           =   2160
      End
   End
   Begin VB.PictureBox pic出院 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   3600
      ScaleHeight     =   5415
      ScaleWidth      =   9855
      TabIndex        =   2
      Top             =   2640
      Width           =   9855
      Begin XtremeReportControl.ReportControl rpt出院 
         Height          =   2415
         Left            =   0
         TabIndex        =   17
         Top             =   480
         Width           =   3360
         _Version        =   589884
         _ExtentX        =   5927
         _ExtentY        =   4260
         _StockProps     =   0
         BorderStyle     =   2
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cboSelectTime 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label lbl出院时间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院时间"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.PictureBox picHLDJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1200
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic在院 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   840
      ScaleHeight     =   5415
      ScaleWidth      =   9855
      TabIndex        =   1
      Top             =   360
      Width           =   9855
      Begin VB.PictureBox picIn在院 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   5895
         TabIndex        =   5
         Top             =   0
         Width           =   5895
         Begin VB.PictureBox picPati 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Index           =   999
            Left            =   0
            Picture         =   "frmPatiSelect.frx":6852
            ScaleHeight     =   1455
            ScaleWidth      =   1395
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   1395
            Begin VB.Label lbl病情 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H000080FF&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   210
               Index           =   999
               Left            =   1740
               TabIndex        =   14
               Top             =   1620
               Width           =   105
            End
            Begin VB.Label lbl房间号 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Height          =   180
               Index           =   999
               Left            =   1800
               TabIndex        =   13
               Top             =   840
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label lbl床号 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "09123"
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
               Index           =   999
               Left            =   120
               TabIndex        =   12
               Top             =   120
               Width           =   675
            End
            Begin VB.Label lbl住院号 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "027647132"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   999
               Left            =   120
               TabIndex        =   11
               Top             =   840
               Width           =   810
            End
            Begin VB.Label lbl性别 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "男"
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   999
               Left            =   630
               TabIndex        =   10
               Top             =   1125
               Width           =   180
            End
            Begin VB.Label lbl年龄 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "33"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   999
               Left            =   930
               TabIndex        =   9
               Top             =   1125
               Width           =   525
            End
            Begin VB.Label lbl姓名 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "李四王麻中华人民共和国"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Index           =   999
               Left            =   75
               TabIndex        =   8
               Top             =   375
               Width           =   1215
            End
            Begin VB.Label lblSplit 
               BackColor       =   &H008080FF&
               Height          =   60
               Index           =   999
               Left            =   45
               TabIndex        =   7
               Top             =   735
               Width           =   1260
            End
            Begin VB.Image img护理等级 
               Appearance      =   0  'Flat
               Height          =   240
               Index           =   999
               Left            =   1065
               Picture         =   "frmPatiSelect.frx":DA34
               Stretch         =   -1  'True
               Top             =   60
               Width           =   240
            End
            Begin VB.Label lblSelect 
               BackColor       =   &H00FFC0C0&
               Height          =   360
               Index           =   999
               Left            =   45
               TabIndex        =   15
               Top             =   375
               Visible         =   0   'False
               Width           =   1260
            End
         End
      End
      Begin VB.VScrollBar HScr 
         Height          =   3945
         LargeChange     =   1000
         Left            =   9600
         Max             =   0
         SmallChange     =   20
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.Timer tmrOpen 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrClose 
      Left            =   480
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5265
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   9287
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList imgHLDJ 
      Index           =   999
      Left            =   0
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":DD76
            Key             =   "Pati"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":E310
            Key             =   "Notify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":E8AA
            Key             =   "等待审查"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":EE44
            Key             =   "拒绝审查"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":F3DE
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":F978
            Key             =   "正在抽查"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1038A
            Key             =   "审查反馈"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":10D9C
            Key             =   "抽查反馈"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":11336
            Key             =   "审查整改"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":11D48
            Key             =   "抽查整改"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1275A
            Key             =   "未导入"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":12CF4
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1328E
            Key             =   "不符合"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":13CA0
            Key             =   "正常结束"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1423A
            Key             =   "变异结束"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":147D4
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":14D6E
            Key             =   "单病种"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1B5D0
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1BB6A
            Key             =   "紧急"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1C104
            Key             =   "Fbaby"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPatiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
'Const LWA_COLORKEY = &H1
Private lngAlpha As Integer
Private j As Integer
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As PointAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" _
    (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const ALTERNATE = 1
Private mobjFileSys As New FileSystemObject
Private mlng病区ID As Long
Private mblnCardOrder As Boolean
Private mColBad As New Collection
Private mblnCancle As Boolean
Private mintIndex As Integer
Private mlngColor As Long
Private mstr病人IDs As String
Private mstr住院号 As String
Private mstr床号 As String
Private mstr姓名 As String
Private mblnOK As Boolean
Private Enum PATIREPORT_COLUMN
    COL_序号 = 0
    COL_病人ID = 1
    COL_主页ID = 2
    COL_姓名 = 3
    COL_住院号 = 4
    COL_床号 = 5
    col_性别 = 6
    col_年龄 = 7
    col_病人类型 = 8
    col_入院日期 = 9
    col_出院日期 = 10
    col_住院天数 = 11
End Enum
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mintOutPreTime As Integer

Public Function ShowMe(objParent As Object, ByVal lng病区ID As Long, str病人IDs As String, Optional str姓名 As String, Optional str住院号 As String, Optional str床号 As String) As Boolean
    mlng病区ID = lng病区ID
    mstr病人IDs = ""
    mstr住院号 = ""
    mstr床号 = ""
    mstr姓名 = ""
    mblnOK = False
    Me.Show 1, objParent
    ShowMe = mblnOK
    If mblnOK Then
        str病人IDs = mstr病人IDs
        str住院号 = mstr住院号
        str床号 = mstr床号
        str姓名 = mstr姓名
    End If
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim objSel As ReportRow
    
    For i = 0 To picPati.Count - 2
        If lblSelect(i).Visible Then
            mstr病人IDs = mstr病人IDs & "," & Val(Split(picPati(i).Tag, ",")(0)) & ":" & Val(Split(picPati(i).Tag, ",")(1))
            mstr住院号 = mstr住院号 & "," & lbl住院号(i).Caption
            mstr床号 = mstr床号 & "," & lbl床号(i).Caption
            mstr姓名 = mstr姓名 & "," & lbl姓名(i).Caption
        End If
    Next
    
    For Each objSel In rpt出院.SelectedRows
        mstr病人IDs = mstr病人IDs & "," & Val(objSel.Record.Item(COL_病人ID).value) & ":" & Val(objSel.Record.Item(COL_主页ID).value)
        mstr住院号 = mstr住院号 & "," & objSel.Record.Item(COL_住院号).value
        mstr床号 = mstr床号 & "," & Trim(objSel.Record.Item(COL_床号).value)
        mstr姓名 = mstr姓名 & "," & objSel.Record.Item(COL_姓名).value
    Next
    
    For Each objSel In rpt转出.SelectedRows
        mstr病人IDs = mstr病人IDs & "," & Val(objSel.Record.Item(COL_病人ID).value) & ":" & Val(objSel.Record.Item(COL_主页ID).value)
        mstr住院号 = mstr住院号 & "," & objSel.Record.Item(COL_住院号).value
        mstr床号 = mstr床号 & "," & Trim(objSel.Record.Item(COL_床号).value)
        mstr姓名 = mstr姓名 & "," & objSel.Record.Item(COL_姓名).value
    Next
    
    mstr病人IDs = Mid(mstr病人IDs, 2)
    mstr住院号 = Mid(mstr住院号, 2)
    mstr床号 = Mid(mstr床号, 2)
    mstr姓名 = Mid(mstr姓名, 2)
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdRef_Click()
    load转出出院病人 rpt转出
End Sub

Private Sub Form_Load()
    '设置淡出
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, 0, LWA_ALPHA  '150为透明度(0-255)
    tmrClose.Interval = 10
    tmrOpen.Interval = 10
    tmrOpen.Enabled = True
    tmrClose.Enabled = False
    lngAlpha = 5
    mblnCancle = False
    InitColor
    
    mblnCardOrder = (Val(zlDatabase.GetPara("床位卡片排序方式", glngSys, P新版护士站, 0)) = 0)
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(0, "在院病人", pic在院.hwnd, 0).Tag = "在院"
        .InsertItem(1, "出院病人", pic出院.hwnd, 0).Tag = "出院"
        .InsertItem(2, "最近转出", pic转出.hwnd, 0).Tag = "转出"
       
        .Item(0).Selected = True '新建时就自动选中了这个,不会再激活事件
        '只加载选择的子窗体
        Form_Resize
        Call tbcSub_SelectedChanged(.Selected)
    End With
    InitSelectTime
    InitReportColumn rpt出院
    InitReportColumn rpt转出
End Sub

Private Sub InitColor()
    Dim strValue As String
    Dim lng特级 As Long, lng一级 As Long, lng二级 As Long, lng三级 As Long
    Const c紫色 As Long = 8388736
    Const c红色 As Long = 255
    Const c兰色 As Long = 16711680
    Const c白色 As Long = 16777215
    
    Call DeleteFile
    mintIndex = 0
    imgHLDJ(999).ListImages.Clear
    '读取护理等级现有设置(无则取缺省数据)
    strValue = zlDatabase.GetPara("特级护理颜色", glngSys, 1265, "")
    lng特级 = IIF(strValue = "", c紫色, Val(strValue))
    strValue = zlDatabase.GetPara("一级护理颜色", glngSys, 1265, "")
    lng一级 = IIF(strValue = "", c红色, Val(strValue))
    strValue = zlDatabase.GetPara("二级护理颜色", glngSys, 1265, "")
    lng二级 = IIF(strValue = "", c兰色, Val(strValue))
    strValue = zlDatabase.GetPara("三级护理颜色", glngSys, 1265, "")
    lng三级 = IIF(strValue = "", c白色, Val(strValue))
    
    '绘图
    mlngColor = lng特级
    Call DrawPoly
    mlngColor = lng一级
    Call DrawPoly
    mlngColor = lng二级
    Call DrawPoly
    mlngColor = lng三级
    Call DrawPoly
End Sub

Private Sub AddColor()
    Dim strFile As String
    mintIndex = mintIndex + 1
    '不保存为文件,当创建多个图片时,加入到imagelist里的始终只有最后一个,应该是由于image中保存的是图片ID造成
    
    strFile = App.Path & "\HLDJTMP" & mintIndex & ".BMP"
    SavePicture picHLDJ.Image, strFile
    picHLDJ.Picture = LoadPicture(strFile)
    imgHLDJ(999).ListImages.Add , "K_" & mintIndex, picHLDJ.Picture
End Sub

Private Sub DrawPoly()
    Dim lngRgn As Long, lngBrush As Long
    Dim lngPen As Long, lngOldPen As Long
    Dim PtInPoly() As PointAPI

    '填充区域并划边线
    ReDim PtInPoly(4) As PointAPI
    PtInPoly(1).X = 0
    PtInPoly(1).Y = 0
    PtInPoly(2).X = picHLDJ.ScaleWidth
    PtInPoly(2).Y = 0
    PtInPoly(3).X = picHLDJ.ScaleWidth
    PtInPoly(3).Y = picHLDJ.ScaleHeight
    PtInPoly(4).X = PtInPoly(1).X
    PtInPoly(4).Y = PtInPoly(1).Y
    
    '创建系统刷子
    picHLDJ.Cls
    lngBrush = CreateSolidBrush(mlngColor)

    '如果创建刷子成功,才选入
    If lngBrush <> 0 Then
        lngRgn = CreatePolygonRgn(PtInPoly(1), UBound(PtInPoly), ALTERNATE)
        FillRgn picHLDJ.hDC, lngRgn, lngBrush
        Call DeleteObject(lngRgn)
        Call DeleteObject(lngBrush)
    End If
    picHLDJ.Refresh
    
    Call AddColor
End Sub

Private Function Get护理等级(ByVal str护理等级 As String) As Integer
    '三级或无等级时,返回3
    If InStr(1, str护理等级, "特") <> 0 Or InStr(1, str护理等级, "重") <> 0 Then
        Get护理等级 = 0
    ElseIf InStr(1, str护理等级, "III") <> 0 Then
        Get护理等级 = 3
    ElseIf InStr(1, str护理等级, "二") <> 0 Or InStr(1, str护理等级, "2") <> 0 Or InStr(1, str护理等级, "Ⅱ") <> 0 Or InStr(1, str护理等级, "II") <> 0 Then
        Get护理等级 = 2
    ElseIf InStr(1, str护理等级, "一") <> 0 Or InStr(1, str护理等级, "1") <> 0 Or InStr(1, str护理等级, "Ⅰ") <> 0 Or InStr(1, str护理等级, "I") <> 0 Then
        Get护理等级 = 1
    Else
        Get护理等级 = 3
    End If
End Function

Private Sub DeleteFile()
    Dim objFile As File
    For Each objFile In mobjFileSys.GetFolder(App.Path).Files
        If Left(objFile.Name, 7) = "HLDJTMP" Then
            mobjFileSys.DeleteFile objFile.Path, True
        End If
    Next
End Sub

Private Sub Load在院病人()
    Dim strSQL As String, rsTmp As Recordset
    Dim i As Long, lngWcount As Long, lngWidth As Long, lngHeigh As Long, lngHcount As Long
    Dim int护理等级 As Integer
    
    For i = mColBad.Count To 1 Step -1
        Unload lbl床号(mColBad(i).Index)
        Unload img护理等级(mColBad(i).Index)
        Unload lbl姓名(mColBad(i).Index)
        Unload lbl住院号(mColBad(i).Index)
        Unload lbl性别(mColBad(i).Index)
        Unload lbl年龄(mColBad(i).Index)
        Unload lblSplit(mColBad(i).Index)
        Unload lblSelect(mColBad(i).Index)
        Unload mColBad(i)
        mColBad.Remove i
    Next
    
    strSQL = "Select distinct a.病人id, a.主页id,A.出院病床 as 床号, a.住院号, a.姓名, a.年龄, a.性别, a.入科时间,A.病人类型," & vbNewLine & _
            "       Trunc(Sysdate) - Trunc(Decode(a.入科时间, Null, a.入院日期, a.入科时间)) || '天' As 住院天数,C.名称 as 护理等级" & IIF(mblnCardOrder, "", ",d.编码 as 床位编制") & vbNewLine & _
            "From 病案主页 A, 在院病人 B,收费项目目录 C" & IIF(mblnCardOrder, "", ",床位编制分类 D,床位状况记录 F") & vbNewLine & _
            "Where a.病人id = b.病人id And a.主页id = b.主页id and A.护理等级ID = C.ID(+) " & IIF(mblnCardOrder, "", " and b.病人id=f.病人id(+) and f.床位编制=D.名称(+)") & " And b.病区id = [1]" & vbNewLine & _
            "order by " & IIF(mblnCardOrder, "", "床位编制,") & "LPAD(A.出院病床,10,' ')"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病区ID)
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    lngWidth = picPati(999).Width
    lngHeigh = picPati(999).Height
    picIn在院.Move 0, 0, pic在院.Width - HScr.Width - 250
    
    lngWcount = (picIn在院.Width) \ (lngWidth + 50)
    lngHcount = (pic在院.Height) \ (lngHeigh + 50)

    
    For i = 0 To rsTmp.RecordCount - 1
        '空卡片
        Load picPati(i)
        picPati(i).Move ((i Mod lngWcount)) * lngWidth + ((i Mod lngWcount) + 1) * 50, (i \ lngWcount) * lngHeigh + (i \ lngWcount + 1) * 50
        mColBad.Add picPati(i)
        picPati(i).Tag = rsTmp!病人ID & "," & rsTmp!主页ID
        picPati(i).Visible = True
        
        Load lbl床号(i)
        Set lbl床号(i).Container = picPati(i)
        lbl床号(i).Caption = rsTmp!床号 & ""
        lbl床号(i).Visible = True
        
        Load img护理等级(i)
        img护理等级(i).Visible = True
        Set img护理等级(i).Container = picPati(i)
        Set img护理等级(i).Picture = Nothing
        img护理等级(i).ZOrder 1
        '设置护理等级(特级紫,一级红,二级蓝,三级无)
        int护理等级 = Get护理等级(rsTmp!护理等级 & "")
        Set img护理等级(i).Picture = imgHLDJ(999).ListImages(int护理等级 + 1).Picture
        
        Load lbl姓名(i)
        Set lbl姓名(i).Container = picPati(i)
        lbl姓名(i).Caption = rsTmp!姓名 & ""
        lbl姓名(i).Visible = True
        
        Load lbl住院号(i)
        Set lbl住院号(i).Container = picPati(i)
        lbl住院号(i).Caption = rsTmp!住院号 & ""
        lbl住院号(i).Visible = True
        
        Load lbl性别(i)
        Set lbl性别(i).Container = picPati(i)
        lbl性别(i).Caption = rsTmp!性别 & ""
        lbl性别(i).Visible = True
        
        Load lbl年龄(i)
        Set lbl年龄(i).Container = picPati(i)
        lbl年龄(i).Caption = rsTmp!年龄 & ""
        lbl年龄(i).Visible = True
        
        Load lblSplit(i)
        Set lblSplit(i).Container = picPati(i)
        lblSplit(i).ZOrder 1
        lblSplit(i).BackColor = IIF(NVL(rsTmp!病人类型) = "普通病人", &HFFFFFF, zlDatabase.GetPatiColor(NVL(rsTmp!病人类型)))
        lblSplit(i).Visible = True
        
        Load lblSelect(i)
        Set lblSelect(i).Container = picPati(i)
        
        rsTmp.MoveNext
    Next
    picIn在院_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tbcSub.Move 0, 0, Me.Width, Me.Height - cmdOK.Height - 700
    cmdOK.Move Me.Width - cmdOK.Width - cmdExit.Width - 600, Me.Height - cmdOK.Height - 630
    cmdExit.Move Me.Width - cmdExit.Width - 400, cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    If mblnCancle = False Then
        Cancel = 1
        lngAlpha = 250
        tmrClose.Enabled = True
    End If
    For i = mColBad.Count To 1 Step -1
        mColBad.Remove i
    Next
End Sub

Private Sub HScr_Change()
    picIn在院.Top = -1 * (HScr.value / HScr.Max) * (picIn在院.Height - pic在院.Height + 800)
End Sub

Private Sub img护理等级_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub img护理等级_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lblSelect_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lblSplit_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lblSplit_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl床号_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl床号_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl年龄_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl年龄_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl入院日期_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl入院日期_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl性别_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl性别_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl姓名_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl姓名_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl住院号_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl住院号_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl住院天数_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl住院天数_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub picIn在院_Resize()
    Dim i As Long, lngWcount As Long, lngWidth As Long, lngHeigh As Long, lngHcount As Long
    
    On Error Resume Next
    If mColBad.Count > 0 Then
        lngWidth = picPati(999).Width
        lngHeigh = picPati(999).Height
        
        lngWcount = (picIn在院.Width) \ (lngWidth + 50)
        lngHcount = (pic在院.Height) \ (lngHeigh + 50)
        picIn在院.Height = ((mColBad.Count \ lngWcount) + 1) * (lngHeigh + 50)
        
        For i = 0 To mColBad.Count - 1
            mColBad(i + 1).Move ((i Mod lngWcount)) * lngWidth + ((i Mod lngWcount) + 1) * 50, (i \ lngWcount) * lngHeigh + (i \ lngWcount + 1) * 50
        Next
        If mColBad.Count > lngHcount * lngWcount Then
            HScr.Visible = True
            HScr.Max = HScr.LargeChange * (mColBad.Count / (lngHcount * lngWcount)) - HScr.LargeChange
        Else
            HScr.Visible = False
        End If
    End If
End Sub

Private Sub picPati_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub SelectPati(Index As Integer)
    lblSelect(Index).Visible = Not lblSelect(Index).Visible
End Sub

Private Sub picPati_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub pic出院_Resize()
    rpt出院.Move 0, rpt出院.Top, pic出院.Width, pic出院.Height
End Sub

Private Sub pic在院_Resize()
    On Error Resume Next
    HScr.Move pic在院.Width - HScr.Width - 250, 0, HScr.Width, pic在院.Height
    picIn在院.Move 0, 0, pic在院.Width - HScr.Width - 250
End Sub

Private Sub pic转出_Resize()
    rpt转出.Move 0, rpt转出.Top, pic转出.Width, pic转出.Height
End Sub

Private Sub rpt出院_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdOK_Click
End Sub

Private Sub rpt转出_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdOK_Click
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Tag = "在院" Then
        If pic在院.Tag = "" Then
            Load在院病人
            pic在院.Tag = "1"
        End If
    ElseIf Item.Tag = "出院" Then
        If pic出院.Tag = "" Then
            load转出出院病人 rpt出院
            pic出院.Tag = "1"
        End If
    ElseIf Item.Tag = "转出" Then
        If pic转出.Tag = "" Then
            load转出出院病人 rpt转出
            pic转出.Tag = "1"
        End If
    End If
End Sub

Private Sub load转出出院病人(objrpt As ReportControl)
    Dim strSQL As String, rsPati As Recordset
    Dim i As Long
    Dim objRecord As ReportRecord
    If Me.Visible = False Then Exit Sub
    With objrpt
        On Error GoTo errH
        If objrpt.Name = "rpt出院" Then
            strSQL = "Select distinct a.病人id, a.主页id,LPAD(A.出院病床,10,' ') as 床号, a.住院号, a.姓名, a.年龄, a.性别, a.入科时间,A.出院日期,A.病人类型," & vbNewLine & _
                "       Trunc(A.出院日期) - Trunc(Decode(a.入科时间, Null, a.入院日期, a.入科时间)) As 住院天数" & vbNewLine & _
                "From 病案主页 A" & vbNewLine & _
                "Where A.当前病区ID=[1]" & vbNewLine & _
                " And A.出院日期 Between to_date([2],'YYYY-MM-DD HH24:MI:SS') And to_date([3],'YYYY-MM-DD HH24:MI:SS')" & _
                "order by A.出院日期 Desc"
        ElseIf objrpt.Name = "rpt转出" Then
            strSQL = "Select Distinct a.病人id, a.主页id, LPad(a.出院病床, 10, ' ') As 床号, a.住院号, a.姓名, a.年龄, a.性别, a.入科时间, C.终止时间 as 出院日期, a.病人类型," & vbNewLine & _
                "                Trunc(Sysdate) - Trunc(Decode(a.入科时间, Null, a.入院日期, a.入科时间)) As 住院天数" & vbNewLine & _
                "From 病案主页 A, 病人变动记录 C" & vbNewLine & _
                "Where a.病人id = c.病人id And a.主页id = c.主页id And a.当前病区id <> [1] And c.病区id + 0 = [1] And Nvl(c.附加床位, 0) = 0 And" & vbNewLine & _
                "      c.终止原因 In (3, 15) And  C.终止时间 Between Sysdate-[4] And Sysdate" & vbNewLine & _
                "Order By C.终止时间 Desc"

        End If
            
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病区ID, Format(mdtOutBegin, "yyyy-MM-dd 00:00:00"), Format(mdtOutEnd, "yyyy-MM-dd 23:59:59"), Val(txtChange.Text))
        
        .Records.DeleteAll
        For i = 1 To rsPati.RecordCount
            Set objRecord = .Records.Add()
            objRecord.AddItem i
            objRecord.Tag = Val(rsPati!病人ID) & "," & Val(rsPati!主页ID)
            objRecord.AddItem Val(rsPati!病人ID)
            objRecord.AddItem Val(rsPati!主页ID)
            objRecord.AddItem CStr(NVL(rsPati!姓名))
            objRecord.AddItem CStr(NVL(rsPati!住院号))
            objRecord.AddItem CStr(NVL(rsPati!床号))
            objRecord.AddItem CStr(NVL(rsPati!性别))
            objRecord.AddItem CStr(NVL(rsPati!年龄))
            objRecord.AddItem CStr(NVL(rsPati!病人类型))
            objRecord.AddItem CStr(NVL(rsPati!入科时间))
            objRecord.AddItem CStr(NVL(rsPati!出院日期))
            objRecord.AddItem CStr(NVL(rsPati!住院天数))
            objRecord.Item(COL_姓名).ForeColor = zlDatabase.GetPatiColor(NVL(rsPati!病人类型))
            
            rsPati.MoveNext
        Next
        .Populate '缺省不选中任何行
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitSelectTime()
    Dim datCurr As Date
    
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mdtOutEnd = datCurr
    mdtOutBegin = mdtOutEnd - 1
    
    cboSelectTime.Clear '出院
    With cboSelectTime
        .AddItem "今天内"
        .ItemData(.NewIndex) = 0
        .AddItem "昨天内"
        .ItemData(.NewIndex) = 1
        .AddItem "前天内"
        .ItemData(.NewIndex) = 2
        .AddItem "一周内"
        .ItemData(.NewIndex) = 7
        .AddItem "30天内"
        .ItemData(.NewIndex) = 30
        .AddItem "60天内"
        .ItemData(.NewIndex) = 60
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 0
End Sub

Private Sub cboSelectTime_Click()
'功能：当时间范围是指定是，弹出时间选择窗体
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If cboSelectTime.ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
    If intDateCount = -1 Then
        If Not frmSelectTime.ShowMe(Me, mdtOutBegin, mdtOutEnd, cboSelectTime) Then
            '取消时恢复原来的选择
            Call cbo.SetIndex(cboSelectTime.hwnd, mintOutPreTime)
            Exit Sub
        End If
    Else
        mdtOutEnd = datCurr
        mdtOutBegin = mdtOutEnd - intDateCount
    End If
    If mdtOutBegin = CDate(0) Or mdtOutEnd = CDate(0) Then
        cboSelectTime.ToolTipText = ""
    Else
        cboSelectTime.ToolTipText = "范围：" & Format(mdtOutBegin, "yyyy-MM-dd") & " 至 " & Format(mdtOutEnd, "yyyy-MM-dd")
    End If
    '保存参数，保证每个地方提取的出院病人都是在同一时间范围内（72783）
    mintOutPreTime = cboSelectTime.ListIndex
    load转出出院病人 rpt出院
End Sub

Private Sub tmrOpen_Timer()
    lngAlpha = lngAlpha + 10
     SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
     SetLayeredWindowAttributes Me.hwnd, 0, lngAlpha, LWA_ALPHA  '150为透明度(0-255)
     If lngAlpha >= 255 Then tmrOpen.Enabled = False
End Sub

Private Sub tmrClose_Timer()
    lngAlpha = lngAlpha - 10
     SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
     SetLayeredWindowAttributes Me.hwnd, 0, lngAlpha, LWA_ALPHA  '150为透明度(0-255)
     If lngAlpha <= 5 Then tmrClose.Enabled = False: mblnCancle = True: Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = HScr.value
    lngMin = HScr.Min
    lngMax = HScr.Max
    
    If KeyCode = vbKeyPageDown Then '下
        If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
            HScr.value = lngCur + (lngMax - lngMin) / 10
        Else
            HScr.value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '上
        If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
            HScr.value = lngCur - (lngMax - lngMin) / 10
        Else
            HScr.value = lngMin
        End If
    End If
    
End Sub

Private Sub Form_Activate()
'鼠标滚轮
    If picIn在院.Visible Then
        glngPreHWnd = GetWindowLong(picIn在院.hwnd, GWL_WNDPROC)
        SetWindowLong picIn在院.hwnd, GWL_WNDPROC, AddressOf FlexScroll
    End If
End Sub

Private Sub InitReportColumn(obj As Object)
    Dim objCol As ReportColumn, lngIdx As Long, i As Long

    With obj
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(COL_序号, "序号", 30, True)
        Set objCol = .Columns.Add(COL_病人ID, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_主页ID, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_姓名, "姓名", 80, True)
        Set objCol = .Columns.Add(COL_住院号, "住院号", 90, True)
        Set objCol = .Columns.Add(COL_床号, "床号", 70, True)
        Set objCol = .Columns.Add(col_性别, "性别", 50, True)
        Set objCol = .Columns.Add(col_年龄, "年龄", 50, True)
        Set objCol = .Columns.Add(col_病人类型, "病人类型", 120, True)
        Set objCol = .Columns.Add(col_入院日期, "入院日期", 120, True)
        Set objCol = .Columns.Add(col_出院日期, IIF(obj.Name = "rpt出院", "出院日期", "转出日期"), 120, True)
        Set objCol = .Columns.Add(col_住院天数, "住院天数", 60, True)
      

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        '.MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(col_出院日期)
        .SortOrder(0).SortAscending = False
    End With
    
    
End Sub

Private Sub txtChange_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then Exit Sub
    load转出出院病人 rpt转出
End Sub

Private Sub txtChange_GotFocus()
    Call zlControl.TxtSelAll(txtChange)
End Sub
