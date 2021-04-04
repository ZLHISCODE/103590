VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmScheduleExcel 
   Caption         =   "生成登记表格"
   ClientHeight    =   5460
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   9720
   Icon            =   "frmScheduleExcel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9720
   Begin VB.TextBox txtInfo 
      Height          =   2295
      Left            =   11700
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "frmScheduleExcel.frx":076A
      Top             =   5625
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock sckMail 
      Left            =   3960
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtHead 
      Height          =   2295
      Left            =   11325
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "frmScheduleExcel.frx":0F62
      Top             =   5760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   5100
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmScheduleExcel.frx":175A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12065
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9720
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   24
         Top             =   30
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   1270
         ButtonWidth     =   1482
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&M.发送"
               Key             =   "发送"
               Object.ToolTipText     =   "发送生成的登记表格(Alt+M)"
               Object.Tag             =   "&M.发送"
               ImageKey        =   "SendMail"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.导出"
               Key             =   "导出"
               Object.ToolTipText     =   "将生成的登记表格导出为Excel文件(Alt+S)"
               Object.Tag             =   "&S.导出"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助(Alt+H)"
               Object.Tag             =   "&H.帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出(Alt+X)"
               Object.Tag             =   "&X.退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   9300
      Top             =   4515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":1FEE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":220E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":242E
            Key             =   "SendMail"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":2BA8
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   9975
      Top             =   4515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":2DC2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":2FE2
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":3202
            Key             =   "SendMail"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScheduleExcel.frx":397C
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4425
      Left            =   45
      TabIndex        =   0
      Top             =   660
      Width           =   2835
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   2145
         TabIndex        =   11
         Text            =   "30"
         Top             =   3540
         Width           =   540
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Text            =   "25"
         Top             =   1050
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   2
         Top             =   435
         Width           =   2580
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "&6.保存发送者密码"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   3240
         Width           =   1845
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2865
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   2235
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1635
         Width           =   2580
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&7.等待服务应答间隔(秒)"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   3585
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.端口号"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   27
         Top             =   795
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.邮件服务器"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&5.密  码"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   7
         Top             =   2625
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&4.用户名"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   5
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.发送人地址"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   1425
         Width           =   1080
      End
   End
   Begin VB.Frame fra2 
      Height          =   4320
      Left            =   3210
      TabIndex        =   12
      Top             =   660
      Width           =   7875
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   5
         Left            =   4710
         MaxLength       =   4
         TabIndex        =   30
         Text            =   "500"
         Top             =   525
         Width           =   870
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   8
         Left            =   1125
         TabIndex        =   17
         Top             =   525
         Width           =   2490
      End
      Begin VB.TextBox txt 
         Height          =   1890
         Index           =   7
         Left            =   1125
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   885
         Width           =   6705
      End
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   4
         Left            =   7425
         Picture         =   "frmScheduleExcel.frx":3B96
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H80000006&
         Height          =   300
         Index           =   13
         Left            =   1125
         TabIndex        =   14
         Top             =   165
         Width           =   6300
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1050
         Left            =   1125
         TabIndex        =   21
         Top             =   2835
         Width           =   6735
         _cx             =   11880
         _cy             =   1852
         Appearance      =   1
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
         Begin VB.Line lnY 
            Index           =   0
            Visible         =   0   'False
            X1              =   270
            X2              =   270
            Y1              =   420
            Y2              =   1635
         End
         Begin VB.Line lnX 
            Index           =   0
            Visible         =   0   'False
            X1              =   -4635
            X2              =   -2850
            Y1              =   -1695
            Y2              =   -1695
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&R.团体人数"
         Height          =   180
         Index           =   5
         Left            =   3735
         TabIndex        =   29
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&E.电子邮件"
         Height          =   180
         Index           =   10
         Left            =   105
         TabIndex        =   16
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&P.团体人员"
         Height          =   180
         Index           =   9
         Left            =   105
         TabIndex        =   20
         Top             =   2820
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&T.邮件内容"
         Height          =   180
         Index           =   8
         Left            =   105
         TabIndex        =   18
         Top             =   930
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&N.团体名称"
         Height          =   180
         Index           =   7
         Left            =   105
         TabIndex        =   13
         Top             =   225
         Width           =   900
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   1020
      Top             =   5370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileMail 
         Caption         =   "发送表格(&M)"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "导出表格(&S)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&T)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmScheduleExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mblnMaining As Boolean

Private Enum mCol
    
    姓名 = 66
    性别
    年龄
    出生日期
    婚姻状况
    身份证号
    门诊号
    健康号
    就诊卡号
    工作单位
    电子邮件
    民族
    学历
    职业
    国籍
    体检组
    
End Enum

Private Enum mVsfCol
    
    姓名
    性别
    年龄
    出生日期
    婚姻状况
    身份证号
    电子邮件
    民族
    学历
    职业
    国籍
    健康号
    就诊卡号
    门诊号
    
End Enum

'（２）自定义过程或函数************************************************************************************************

Private Function CreateTmpFile(Optional ByVal strFileType As String = "tmp") As String
    '------------------------------------------------------------------------------------------------------------------
    '
    '功能:
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    strFileTemp = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
    
    strFileTemp = strFileTemp & "体检登记表_" & Format(Now, "yyyymmdd") & Format(Timer, "0") & "." & strFileType
    
    CreateTmpFile = strFileTemp
End Function

Private Function NewExcelFile(ByRef strExcelFile As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '功能:
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim objExcel As Object
    Dim ExWorkbook As Object
    Dim ExWorkSheet As Object
    Dim strParam As String
    Dim lngLoop As Long
    Dim varParam As Variant
    Dim lngRows As Long
    Dim lngCols As Long
    Dim strColChr As String
    
    On Error GoTo errHand
    
    If Val(txt(5).Text) <= 0 Then
        ShowSimpleMsg "必须指定团体人数或人数不能小于1个！"
        LocationObj txt(5)
        Exit Function
    End If

    If Val(txt(5).Text) < vsf.Rows - 1 Then
        ShowSimpleMsg "指定团体人数必须大于已有人数！"
        LocationObj txt(5)
        Exit Function
    End If
    
    frmWait.OpenWait Me, "生成体检登记表"
    frmWait.WaitInfo = "正在创建Excel对象..."
    
    Set objExcel = CreateObject("Excel.Application")
    Set ExWorkbook = Nothing
    Set ExWorkSheet = Nothing
    Set ExWorkbook = objExcel.Workbooks().Add
    
    Set ExWorkSheet = ExWorkbook.Worksheets("sheet1")
    
    ExWorkSheet.Name = "人员资料"
        
    ExWorkSheet.Unprotect "Transaction"                         '解锁
    objExcel.ActiveWindow.DisplayGridlines = False              '取消网格线
    
    '定义列标题
    ExWorkSheet.Columns("A:A").ColumnWidth = 1
    ExWorkSheet.Range("A3").Value = ""
    
    strParam = "姓名*,10;性别,5;年龄,5;出生日期,10;婚姻状况,10;身份证号*,20;门诊号*,18;健康号*,10;就诊卡号*,15;工作单位,15;电子邮件,15;民族,10;学历,10;职业,10;国籍,10;体检组,10"
    lngRows = Val(txt(5).Text) + 3
    
    varParam = Split(strParam, ";")
    lngCols = UBound(varParam)
    For lngLoop = 0 To lngCols
        strColChr = Chr(lngLoop + 66)
        ExWorkSheet.Range(strColChr & "3").Value = Split(varParam(lngLoop), ",")(0)
        ExWorkSheet.Columns(strColChr & ":" & strColChr).ColumnWidth = Val(Split(varParam(lngLoop), ",")(1))
    Next
    ExWorkSheet.Range("B3:" & Chr(lngCols + 66) & "3").Select
    With objExcel.Selection
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .Font.Bold = True
        .Font.Size = 9
    End With
    
    '定义标题
    ExWorkSheet.Range("B1:" & Chr(lngCols + 66) & "1").Select
    With objExcel.Selection
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = -5002
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 18
    End With
    objExcel.ActiveCell.FormulaR1C1 = "团体体检人员登记表"
        
    ExWorkSheet.Range("B2:" & Chr(lngCols + 66) & "2").Select
    With objExcel.Selection
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = -5002
        .MergeCells = True
        .Font.Bold = False
        .Font.Size = 9
        .RowHeight = 30
    End With
    objExcel.ActiveCell.FormulaR1C1 = "注:标题后加*表示必须输入，若没有姓名，则表示此行无效。" & vbCrLf & "   导入时，按姓名和身份证号查对历史资料；否则建立新的档案。"
            
    ExWorkSheet.Range("B4:" & Chr(lngCols + 66) & lngRows).Select
    With objExcel.Selection
        .Locked = False
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = -5002
        .MergeCells = False
        .Font.Size = 9
    End With
    
    '冻结标题
    ExWorkSheet.Range("H4").Select
    objExcel.ActiveWindow.FreezePanes = True
            
    frmWait.WaitInfo = "正在产生输入可选数据..."
    
    '产生输入可选数据
    ExWorkSheet.Range(Chr(mCol.出生日期) & "4:" & Chr(mCol.出生日期) & lngRows).Select
    objExcel.Selection.NumberFormatLocal = "yyyy-mm-dd;@"
    With objExcel.Selection.Validation
        
        .Delete
        .Add 4, 1, 1, "1900-01-01", "3000-01-01"
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "输入提醒"
        .InputMessage = ""
        .ErrorMessage = "必须输入正确有效的日期，如1980-09-21"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    
    ExWorkSheet.Range(Chr(mCol.性别) & "4:" & Chr(mCol.性别) & lngRows).Select
    With objExcel.Selection.Validation
        .Delete
        .Add 3, 1, 1, GetExcelList("SELECT 名称 FROM 性别")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "输入提醒"
        .InputMessage = ""
        .ErrorMessage = "必须从下拉列表中选择性别"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
        
    ExWorkSheet.Range(Chr(mCol.婚姻状况) & "4:" & Chr(mCol.婚姻状况) & lngRows).Select
    With objExcel.Selection.Validation
        .Delete
        .Add 3, 1, 1, GetExcelList("SELECT 名称 FROM 婚姻状况")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "输入提醒"
        .InputMessage = ""
        .ErrorMessage = "必须从下拉列表中选择婚姻状况"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    
'    ExWorkSheet.Range(Chr(mCol.民族) & "4:" & Chr(mCol.民族) & lngRows).Select
'    With objExcel.Selection.Validation
'        .Delete
'        .Add 3, 1, 1, GetExcelList("SELECT 名称 FROM 民族")
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = "输入提醒"
'        .InputMessage = ""
'        .ErrorMessage = "必须从下拉列表中选择民族"
'        .IMEMode = 0
'        .ShowInput = True
'        .ShowError = True
'    End With
    
    ExWorkSheet.Range(Chr(mCol.学历) & "4:" & Chr(mCol.学历) & lngRows).Select
    With objExcel.Selection.Validation
        .Delete
        .Add 3, 1, 1, GetExcelList("SELECT 名称 FROM 学历")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "输入提醒"
        .InputMessage = ""
        .ErrorMessage = "必须从下拉列表中选择学历"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    
    ExWorkSheet.Range(Chr(mCol.职业) & "4:" & Chr(mCol.职业) & lngRows).Select
    With objExcel.Selection.Validation
        .Delete
        .Add 3, 1, 1, GetExcelList("SELECT 名称 FROM 职业")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "输入提醒"
        .InputMessage = ""
        .ErrorMessage = "必须从下拉列表中选择职业"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    
'    ExWorkSheet.Range(Chr(mCol.国籍) & "4:" & Chr(mCol.国籍) & lngRows).Select
'    With objExcel.Selection.Validation
'        .Delete
'        .Add 3, 1, 1, GetExcelList("SELECT 名称 FROM 国籍")
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = "输入提醒"
'        .InputMessage = ""
'        .ErrorMessage = "必须从下拉列表中选择国籍"
'        .IMEMode = 0
'        .ShowInput = True
'        .ShowError = True
'    End With
'
'    ExWorkSheet.Range(Chr(mCol.国籍) & "4").Select
'    With objExcel.Selection.Validation
'        .Delete
'        .Add 7, 1, 1, "=IF(LENB(" & Chr(mCol.国籍) & "4)>30,FALSE,IF(ISNUMBER(FIND(""'""," & Chr(mCol.国籍) & "4)),FALSE,TRUE))"
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = "输入提醒"
'        .InputMessage = ""
'        .ErrorMessage = "国籍不能含有非法字符(')同时长度不能超过30个字符或15个汉字！"
'        .IMEMode = 0
'        .ShowInput = True
'        .ShowError = True
'    End With
'    objExcel.Selection.NumberFormatLocal = "@"
'    objExcel.Selection.AutoFill ExWorkSheet.Range(Chr(mCol.国籍) & "4:" & Chr(mCol.国籍) & lngRows), 0
    
    ExWorkSheet.Range(Chr(mCol.姓名) & "4").Select
    With objExcel.Selection.Validation
        .Delete
        .Add 7, 1, 1, "=IF(LENB(" & Chr(mCol.姓名) & "4)>20,FALSE,IF(ISNUMBER(FIND(""'""," & Chr(mCol.姓名) & "4)),FALSE,TRUE))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "输入提醒"
        .InputMessage = ""
        .ErrorMessage = "姓名不能含有非法字符(')同时长度不能超过20个字符或10个汉字！"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    objExcel.Selection.NumberFormatLocal = "@"
    objExcel.Selection.AutoFill ExWorkSheet.Range(Chr(mCol.姓名) & "4:" & Chr(mCol.姓名) & lngRows), 0
            
    ExWorkSheet.Range(Chr(mCol.身份证号) & "4").Select
    With objExcel.Selection.Validation
        .Delete
        .Add 7, 1, 1, "=IF(LENB(" & Chr(mCol.身份证号) & "4)>20,FALSE,IF(ISNUMBER(FIND(""'""," & Chr(mCol.身份证号) & "4)),FALSE,TRUE))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "输入提醒"
        .InputMessage = ""
        .ErrorMessage = "身份证号不能含有非法字符(')同时长度不能超过20个字符！"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    objExcel.Selection.NumberFormatLocal = "@"
    objExcel.Selection.AutoFill ExWorkSheet.Range(Chr(mCol.身份证号) & "4:" & Chr(mCol.身份证号) & lngRows), 0

    ExWorkSheet.Range(Chr(mCol.就诊卡号) & "4").Select
    With objExcel.Selection.Validation
        .Delete
        .Add 7, 1, 1, "=IF(LENB(" & Chr(mCol.就诊卡号) & "4)>20,FALSE,IF(ISNUMBER(FIND(""'""," & Chr(mCol.就诊卡号) & "4)),FALSE,TRUE))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "输入提醒"
        .InputMessage = ""
        .ErrorMessage = "就诊卡号不能含有非法字符(')同时长度不能超过20个字符！"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    objExcel.Selection.NumberFormatLocal = "@"
    objExcel.Selection.AutoFill ExWorkSheet.Range(Chr(mCol.就诊卡号) & "4:" & Chr(mCol.就诊卡号) & lngRows), 0
    
    '邮件地址
    ExWorkSheet.Range(Chr(mCol.电子邮件) & "4").Select
    With objExcel.Selection.Validation
        .Delete
        .Add 7, 1, 1, "=IF(LENB(" & Chr(mCol.电子邮件) & "4)>50,FALSE,IF(ISNUMBER(FIND(""'""," & Chr(mCol.电子邮件) & "4)),FALSE,IF(ISNUMBER(FIND(""@""," & Chr(mCol.电子邮件) & "4)),TRUE,FALSE)))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "输入提醒"
        .InputMessage = ""
        .ErrorMessage = "电子邮件地址必须含有字符(@)同时长度不能超过50个字符或25个汉字！"
        .IMEMode = 0
        .ShowInput = True
        .ShowError = True
    End With
    objExcel.Selection.NumberFormatLocal = "@"
    objExcel.Selection.AutoFill ExWorkSheet.Range(Chr(mCol.电子邮件) & "4:" & Chr(mCol.电子邮件) & lngRows), 0
        
     '设置网格线
    ExWorkSheet.Range("B3:" & Chr(lngCols + 66) & lngRows).Select
    With objExcel.Selection
        .Borders(5).LineStyle = -4142
        .Borders(6).LineStyle = -4142
        
        .Borders(7).LineStyle = 1
        .Borders(7).Weight = -4138
        .Borders(7).ColorIndex = 48
        
        .Borders(8).LineStyle = 1
        .Borders(8).Weight = -4138
        .Borders(8).ColorIndex = 48
        
        .Borders(9).LineStyle = 1
        .Borders(9).Weight = -4138
        .Borders(9).ColorIndex = 48
        
        .Borders(10).LineStyle = 1
        .Borders(10).Weight = -4138
        .Borders(10).ColorIndex = 48
        
        .Borders(11).LineStyle = 1
        .Borders(11).Weight = 2
        .Borders(11).ColorIndex = 48
        
        .Borders(12).LineStyle = 1
        .Borders(12).Weight = 2
        .Borders(12).ColorIndex = 48
    End With
    
    frmWait.WaitInfo = "正在填写上次体检人员名单..."
    '填写上次人员
    For lngLoop = 1 To vsf.Rows - 1
        '姓名*,10;性别,5;年龄,5;出生日期,10;婚姻状况,10;身份证号*,20;健康号*,10;就诊卡号*,15;电子邮件,15;民族,10;学历,10;职业,10;国籍,10;体检组,10
        '姓名,1080,1,1,1,;性别,600,1,1,1,;年龄,600,1,1,1,;出生日期,990,1,1,1,;婚姻状况,900,1,1,1,;身份证号,1800,1,1,1,;电子邮件,1800,1,1,1,;民族,900,1,1,1,;学历,900,1,1,1,;职业,900,1,1,1,;国籍,900,1,1,1,;健康号,0,1,1,1,;就诊卡号,0,1,1,1,
        '姓名,1080,1,1,1,;性别,600,1,1,1,;年龄,600,1,1,1,;出生日期,990,1,1,1,;婚姻状况,900,1,1,1,;身份证号,1800,1,1,1,;电子邮件,1800,1,1,1,;民族,900,1,1,1,;学历,900,1,1,1,;职业,900,1,1,1,;国籍,900,1,1,1,;健康号,0,1,1,1,;就诊卡号,0,1,1,1,;门诊号,0,1,1,1,"
        
        ExWorkSheet.Range(Chr(mCol.姓名) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.姓名)
        ExWorkSheet.Range(Chr(mCol.性别) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.性别)
        ExWorkSheet.Range(Chr(mCol.年龄) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.年龄)
        ExWorkSheet.Range(Chr(mCol.出生日期) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.出生日期)
        ExWorkSheet.Range(Chr(mCol.婚姻状况) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.婚姻状况)
        ExWorkSheet.Range(Chr(mCol.身份证号) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.身份证号)
        ExWorkSheet.Range(Chr(mCol.门诊号) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.门诊号)
        ExWorkSheet.Range(Chr(mCol.健康号) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.健康号)
        ExWorkSheet.Range(Chr(mCol.就诊卡号) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.就诊卡号)
        ExWorkSheet.Range(Chr(mCol.电子邮件) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.电子邮件)
        ExWorkSheet.Range(Chr(mCol.民族) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.民族)
        ExWorkSheet.Range(Chr(mCol.学历) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.学历)
        ExWorkSheet.Range(Chr(mCol.职业) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.职业)
        ExWorkSheet.Range(Chr(mCol.国籍) & (lngLoop + 3)).Value = vsf.TextMatrix(lngLoop, mVsfCol.国籍)
                
    Next
    
    ExWorkSheet.Range(Chr(mCol.性别) & "3:" & Chr(mCol.体检组) & "3").Select
    objExcel.Selection.AutoFilter
    
    '锁定
    ExWorkSheet.Range(Chr(mCol.姓名) & "4:" & Chr(mCol.姓名) & "4").Select
    ExWorkSheet.Protect "transaction", , , , , , , , , , , , , , True
    
    objExcel.ActiveWorkbook.Protect "transaction", True, False
    
    If strExcelFile <> "" Then ExWorkbook.SaveAs strExcelFile
    
    'objExcel.Visible = True
    objExcel.Quit
    
    NewExcelFile = True
    
    Set objExcel = Nothing
    
    frmWait.CloseWait
    
    Exit Function
    
errHand:
    objExcel.Quit
    frmWait.CloseWait
    If ErrCenter = 1 Then Resume
End Function

Private Function GetExcelList(ByVal strSQL As String) As String
    
    Dim rs As New ADODB.Recordset
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            GetExcelList = GetExcelList & "," & zlCommFun.NVL(rs.Fields(0).Value)
            rs.MoveNext
        Loop
    End If
    If GetExcelList = "" Then
        
    Else
        GetExcelList = Mid(GetExcelList, 2)
    End If
End Function

Private Function ValidData() As Boolean
    '检查
    
    If Not HaveExcel Then
        MsgBox "请安装好Excel后，再使用本功能。", vbCritical, gstrSysName
        Exit Function
    End If
    
    If Trim(txt(4).Text) = "" Then
        MsgBox "必须确定邮件服务器！"
        LocationObj txt(4)
        Exit Function
    End If
    
    If Val(txt(0).Text) = 0 Then
        MsgBox "必须邮件端口号（一般为25）！"
        LocationObj txt(0)
        Exit Function
    End If
    
    If Trim(txt(1).Text) = "" Then
        MsgBox "必须确定发送人的电子邮件地址！"
        LocationObj txt(1)
        Exit Function
    End If
    
    
    If Trim(txt(2).Text) = "" Then
        MsgBox "必须确定用户名！"
        LocationObj txt(2)
        Exit Function
    End If
    
    If Trim(txt(8).Text) = "" Then
        MsgBox "必须确定团体电子邮件地址！"
        LocationObj txt(8)
        Exit Function
    End If
    
    ValidData = True
    
End Function

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFileMail.Enabled = vData
    mnuFileSaveAs.Enabled = vData
    
    tbrThis.Buttons("发送").Enabled = mnuFileMail.Enabled
    tbrThis.Buttons("导出").Enabled = mnuFileSaveAs.Enabled
    
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
        
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      团体id
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
                    
    gstrSQL = "select B.病人id AS ID,B.姓名,B.性别,B.婚姻状况,TO_CHAR(C.出生日期,'yyyy-mm-dd') AS 出生日期,B.电子邮件,C.身份证号,C.年龄,C.民族,C.学历,C.职业,C.国籍,c.健康号,c.就诊卡号,c.门诊号 " & _
                "from 体检登记记录 A,体检人员档案 B,病人信息 C " & _
                "where C.病人id=B.病人id AND A.ID=B.登记ID AND B.病人id>0 AND A.体检号=(select max(体检号) from 体检登记记录 where 合约单位id=[1])"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
        
    strVsf = "姓名,1080,1,1,1,;性别,600,1,1,1,;年龄,600,1,1,1,;出生日期,990,1,1,1,;婚姻状况,900,1,1,1,;身份证号,1800,1,1,1,;电子邮件,1800,1,1,1,;民族,900,1,1,1,;学历,900,1,1,1,;职业,900,1,1,1,;国籍,900,1,1,1,;健康号,0,1,1,1,;就诊卡号,0,1,1,1,;门诊号,0,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    
    Call AppendRows(vsf, lnX, lnY)
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    ValidEdit = True

End Function


Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim rsData As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case Index
    
    Case 4      '打开团体(合同单位)选择器
        lngKey = Val(cmd(Index).Tag)
        
        gstrSQL = GetPublicSQL(SQL.体检团体选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If ShowTxtSelect(Me, txt(13), "编码,900,0,1;名称,1500,0,1;简码,900,0,1;地址,3000,0,1", Me.Name & "\体检团体选择", "请在下表中选择一个团体/单位。", rsData, rs, 8790, 5100) Then
            
            lngKey = zlCommFun.NVL(rs("ID").Value, 0)
            txt(13).Text = zlCommFun.NVL(rs("名称").Value)
            txt(8).Text = zlCommFun.NVL(rs("电子邮件").Value)
              
            cmd(Index).Tag = lngKey
            
            Call ReadData(lngKey)
            
            txt(Index).Tag = ""
        End If
        
        LocationObj txt(13)
        
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyM
            If tbrThis.Buttons("发送").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("发送"))
        Case vbKeyS
            If tbrThis.Buttons("保存").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("保存"))
        Case vbKeyH
            If tbrThis.Buttons("帮助").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("帮助"))
        Case vbKeyX
            If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
        End If
    End If
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Load()
    
    glngFormW = 9840
    glngFormH = 6150
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    txt(0).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "发送人", txt(0).Text)
    txt(1).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "发送人地址", txt(1).Text)
    txt(2).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "用户名", txt(2).Text)
    txt(3).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "密码", txt(3).Text)
    
    txt(4).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "邮件服务器", txt(4).Text)
    
'    txt(5).Text = Val(GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "续发间隔", txt(5).Text))
    txt(6).Text = Val(GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "等待间隔", txt(6).Text))
    
    chk.Value = Val(GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "是否保存密码", chk.Value))
    
    txt(7).Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "邮件内容", txt(7).Text)
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With fra
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With fra2
        .Left = fra.Left + fra.Width + 15
        .Top = fra.Top
        .Width = Me.ScaleWidth - .Left
        .Height = fra.Height
    End With
    
    txt(13).Width = fra2.Width - txt(13).Left - 60 - cmd(4).Width - 30
    cmd(4).Left = txt(13).Left + txt(13).Width + 30
    
'    txt(8).Width = fra2.Width - txt(8).Left - 60
    txt(7).Width = fra2.Width - txt(7).Left - 60
    
    With vsf
        .Width = fra2.Width - .Left - 60
        .Height = fra2.Height - .Top - 60
    End With
    
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If mblnMaining Then
        Cancel = True
        Exit Sub
    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "发送人", txt(0).Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "发送人地址", txt(1).Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "用户名", txt(2).Text)
    
    If chk.Value = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "密码", txt(3).Text)
    Else
        Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "密码", "")
    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "邮件服务器", txt(4).Text)
'    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "续发间隔", Val(txt(5).Text))
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "等待间隔", Val(txt(6).Text))
    
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "是否保存密码", chk.Value)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "邮件内容", txt(7).Text)
    Call SaveWinState(Me, App.ProductName)
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileMail_Click()
    Dim objMail As clsMail
    Dim blnSuccess As Boolean
    Dim strMessage As String
    Dim lngLoop As Long
    Dim strTmpFile As String
    
    '检查
    If ValidData = False Then Exit Sub
    
    Set objMail = New clsMail
    Set objMail.WinSockObj = sckMail
    
    mblnMaining = True
    
    tbrThis.Buttons("发送").Enabled = False
    tbrThis.Buttons("导出").Enabled = False
    tbrThis.Buttons("帮助").Enabled = False
    tbrThis.Buttons("退出").Enabled = False
    
    vsf.Editable = flexEDNone
    mnuFile.Enabled = False
    mnuView.Enabled = False
    mnuHelp.Enabled = False
    
    strTmpFile = CreateTmpFile("xls")
    Call NewExcelFile(strTmpFile)
    DoEvents
    
'    gstrSQL = objMail.GetOracleMail(txt(8).Text, "体检登记表", txt(1).Text, txt(4).Text, txt(2).Text, txt(3).Text, "<font color=""ff6633"">这是html格式邮件内容</font>", strTmpFile, Val(txt(0).Text))
'
'    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    frmWait.OpenWait Me, "发送电子邮件"
    frmWait.WaitInfo = "正在连接邮件服务器..."

    objMail.ResponseInternal = Val(txt(6).Text)

    If objMail.OpenMailServer(txt(4).Text, txt(2).Text, txt(3).Text, Val(txt(0).Text)) Then
'    If objMail.OpenOutLookExMail() Then

        '发送电子邮件处理

        frmWait.WaitInfo = "正在发送团体体检登记表..."
        blnSuccess = objMail.SendHead(txt(8).Text, txt(2).Text, txt(1).Text, "体检登记表", vbMultipartMixed)
        blnSuccess = objMail.SendMessage(txt(7).Text, vbTextPlain)
        blnSuccess = objMail.SendAttach(strTmpFile)
        blnSuccess = objMail.SendOver
'        blnSuccess = objMail.SendOutLookExMail(txt(8).Text, "体检登记表", txt(7).Text, strTmpFile)

    End If

    frmWait.WaitInfo = "正在关闭邮件服务器..."

    Call objMail.CloseMailServer
'    Call objMail.CloseOutLookExMail
    
    tbrThis.Buttons("发送").Enabled = True
    tbrThis.Buttons("导出").Enabled = True
    tbrThis.Buttons("帮助").Enabled = True
    tbrThis.Buttons("退出").Enabled = True
    
    vsf.Editable = flexEDKbdMouse
    mnuFile.Enabled = True
    mnuView.Enabled = True
    mnuHelp.Enabled = True
    mblnMaining = False
    
    frmWait.CloseWait
    
    '不成功，则提示
    If blnSuccess = False Then ShowSimpleMsg "发送电子邮件失败！"
    
End Sub

Private Function GetReportMessageHtml(ByVal lngKey As Long, ByVal lng病人id As Long) As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim lngLoop1 As Long
    Dim lngLoop2 As Long
    Dim lngLoop3 As Long
    Dim strTmp1 As String
    Dim strTmp2 As String
    
    Dim strSQL As String
    
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=4 class=xlTitle style='width:536pt'>体检报告单</td></tr>"
                        
    strSQL = "SELECT A.体检号,A.体检时间,C.姓名,B.体检病历id,B.复查时间,D.书写人 FROM 体检登记记录 A,体检人员档案 B,病人信息 C,病人病历记录 D WHERE D.ID(+)=B.体检病历id AND C.病人id=B.病人id AND A.ID=B.登记id AND A.ID=[1] AND B.病人id=[2]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey, lng病人id)
    If rs.BOF Then Exit Function
    
    txtInfo.Text = txtInfo.Text & _
        "<tr><td class=xl39 style='font-weight:700'>体检人员：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("姓名")) & "</td></tr>" & _
        "<tr><td class=xl39 style='font-weight:700'>体检日期：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & Format(zlCommFun.NVL(rs("体检时间")), "YYYY-MM-DD") & "</td></tr>" & _
        "<tr><td class=xl39 style='font-weight:700'>体检单号：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("体检号")) & "</td></tr>"
        
    txtInfo.Text = txtInfo.Text & _
        "<tr><td colspan=4 class=xl39 style='font-weight:700'>一、项目报告</td></tr>"
            
    '以下是体检项目报告
    
    '1.科室
    strSQL = "select DISTINCT C.名称,C.ID from 体检项目医嘱 A,体检项目清单 B,部门表 C WHERE A.清单ID=B.ID and C.ID=B.执行科室id AND A.病人id=[1] and B.登记id=[2]"
    Set rs1 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人id, lngKey)
    If rs1.BOF Then Exit Function
    
    For lngLoop1 = 1 To rs1.RecordCount
        
        '2.体检项目(填写了报告的)
        strSQL = "select C.名称,B.报告id,D.书写人 " & _
                        "from ( " & _
                             "SELECT * FROM 病人医嘱记录 WHERE 病人id=" & lng病人id & " AND 挂号单=[1] AND 执行科室id=[2] AND 病人来源=4 AND 医嘱状态<>4 AND 诊疗类别='D' AND 相关id IS NULL " & _
                             "Union All " & _
                             "SELECT * FROM 病人医嘱记录 WHERE 病人id=" & lng病人id & " AND 挂号单=[1] AND 执行科室id=[2] AND 病人来源=4 AND 医嘱状态<>4 AND 诊疗类别='C' AND 相关id>0 " & _
                             ") A, " & _
                             "病人医嘱发送 B, " & _
                             "诊疗项目目录 C, " & _
                             "病人病历记录 D " & _
                        "Where A.ID = B.医嘱id " & _
                              "AND B.报告id>0 " & _
                              "AND C.ID=A.诊疗项目ID " & _
                              "AND D.ID=B.报告id "
        
        Set rs2 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(zlCommFun.NVL(rs("体检号"))), Val(zlCommFun.NVL(rs1("ID"))))
        If rs2.BOF = False Then
                
            txtInfo.Text = txtInfo.Text & "<tr><td colspan=4 class=xl39 style='font-weight:700'>" & lngLoop1 & "、" & zlCommFun.NVL(rs1("名称")) & "</td></tr>"
            
            txtInfo.Text = txtInfo.Text & "<tr>"
            
            For lngLoop2 = 1 To rs2.RecordCount
                
                txtInfo.Text = txtInfo.Text & "<td colspan=2 class=xl39 style='font-weight:600'>(" & lngLoop2 & ")" & zlCommFun.NVL(rs2("名称")) & "</td>"
                txtInfo.Text = txtInfo.Text & "<td colspan=2 class=xl39 style='text-align:right'>检查医生：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs2("书写人")) & "</td>"
                txtInfo.Text = txtInfo.Text & "</tr>"
                
                txtInfo.Text = txtInfo.Text & _
                            "<tr><td class=xl25>项目名称</td>" & vbCrLf & _
                            "<td class=xl25>检查结果</td>" & vbCrLf & _
                            "<td class=xl25>参考范围</td>" & vbCrLf & _
                            "<td class=xl25>提示</td></tr>"
                
                '具体检查项目及结果
                strSQL = _
                    "SELECT * FROM ( " & _
                        "SELECT " & _
                               "项目, " & _
                               "内容||DECODE(标志,NULL,'',DECODE(SUBSTR(标志,3,100),'正常','','异常','(+)','偏低','↓','偏高','↑')) AS 内容, " & _
                               "参考, " & _
                               "DECODE(标志,NULL,'',SUBSTR(标志,3,100)) AS 提示, " & _
                               "排列序号, " & _
                               "元素内序号 " & _
                        "FROM ( " & _
                        "SELECT " & _
                               "项目, " & _
                               "内容, " & _
                               "DECODE(SIGN(INSTR(参考,'''')),1,SUBSTR(参考,1,INSTR(参考,'''')-1),'') AS 标志, " & _
                               "DECODE(SIGN(INSTR(参考,'''')),1,SUBSTR(参考,INSTR(参考,'''')+1,1000),'') AS 参考, " & _
                               "排列序号, " & _
                               "元素内序号 " & _
                        "FROM ( " & _
                        "SELECT " & _
                               "项目, " & _
                               "DECODE(SIGN(INSTR(内容,'''')),1,SUBSTR(内容,1,INSTR(内容,'''')-1),内容) AS 内容, " & _
                               "DECODE(SIGN(INSTR(内容,'''')),1,SUBSTR(内容,INSTR(内容,'''')+1,1000),'') AS 参考, " & _
                               "排列序号, " & _
                               "元素内序号 "
                strSQL = strSQL & _
                        "FROM ( " & _
                        "SELECT C.中文名 AS 项目,DECODE(A.所见内容,NULL,NULL,A.所见内容||' '||DECODE(A.计量单位,NULL,'',A.计量单位)) AS 内容,B.排列序号,A.控件号 AS 元素内序号 FROM 病人病历所见单 A,病人病历内容 B,诊治所见项目 C " & _
                        "Where A.病历ID = B.ID " & _
                              "AND B.病历记录ID=[1] " & _
                              "AND C.ID=A.所见项ID " & _
                        "))) " & _
                        "Union All " & _
                        "SELECT B.标题文本 AS 项目,A.内容,'' AS 参考,'' AS 提示,B.排列序号,0 AS 元素内序号 FROM 病人病历文本段 A,病人病历内容 B " & _
                        "Where A.病历ID = B.ID " & _
                                "And B.病历记录ID = [1] " & _
                              "AND 元素类型 IN (0,-5) " & _
                        ") ORDER BY 排列序号,元素内序号"
                        
                Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs2("报告id"))))
                If rs3.BOF = False Then
                    For lngLoop3 = 1 To rs3.RecordCount
                        txtInfo.Text = txtInfo.Text & vbCrLf & _
                                "<tr><td class=xl28>" & zlCommFun.NVL(rs3("项目")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("内容")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("参考")) & "</td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs3("提示")) & "</td></tr>"
                        rs3.MoveNext
                    Next
                Else
                    txtInfo.Text = txtInfo.Text & vbCrLf & _
                                "<tr><td class=xl28 style='mso-height-source:userset;height:15.0pt'></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td>" & vbCrLf & _
                                "<td class=xl28><font class=" & Chr(34) & "font8" & Chr(34) & "></td></tr>"
                End If
                                        
                strTmp1 = ""
                strTmp2 = ""
                
                strSQL = "SELECT * FROM 体检人员结论 WHERE 病历id in (select id from 病人病历内容 where 病历记录id=[1]) ORDER BY 记录性质,记录序号"
                Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs2("报告id"))))
                If rs3.BOF = False Then
                    For lngLoop3 = 1 To rs3.RecordCount
                        
                        If zlCommFun.NVL(rs3("记录性质"), 0) = 0 Then strTmp1 = strTmp1 & zlCommFun.NVL(rs3("结论描述")) & vbCrLf
                        If zlCommFun.NVL(rs3("记录性质"), 0) = 1 Then strTmp2 = zlCommFun.NVL(rs3("参考建议"))
                        
                        rs3.MoveNext
                    Next
                End If
                
                txtInfo.Text = txtInfo.Text & vbCrLf & _
                    "<tr><td colspan=4 class=xl28 style='font-weight:600'>结论：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp1 & "</td></tr>" & vbCrLf & _
                    "<tr><td colspan=4 class=xl28 style='font-weight:600'>建议：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp2 & "</td></tr>"
                    
                txtInfo.Text = txtInfo.Text & vbCrLf & "<tr><td class=xl39 style='mso-height-source:userset;height:15.0pt'></td></tr>"
                
                rs2.MoveNext
            Next
        End If
        
        rs1.MoveNext
    Next
        
    '总检
    strTmp1 = ""
    strTmp2 = ""
    
    strSQL = "SELECT * FROM 体检人员结论 WHERE 病历id in (select id from 病人病历内容 where 病历记录id=[1]) ORDER BY 记录性质,记录序号"
    Set rs3 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(zlCommFun.NVL(rs("体检病历id"))))
    If rs3.BOF = False Then
        For lngLoop3 = 1 To rs3.RecordCount
            
            If zlCommFun.NVL(rs3("记录性质"), 0) = 0 Then strTmp1 = strTmp1 & zlCommFun.NVL(rs3("结论描述")) & vbCrLf
            If zlCommFun.NVL(rs3("记录性质"), 0) = 1 Then strTmp2 = zlCommFun.NVL(rs3("参考建议"))
            
            rs3.MoveNext
        Next
    End If
            
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=2 class=xl39 style='font-weight:700'>二、总检报告</td>" & vbCrLf & _
        "<td colspan=2 class=xl39 style='text-align:right'>总检医生：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & zlCommFun.NVL(rs("书写人")) & "</td></tr>"
        
    txtInfo.Text = txtInfo.Text & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>结论：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp1 & "</td></tr>" & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>建议：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & strTmp2 & "</td></tr>" & vbCrLf & _
        "<tr><td colspan=4 class=xl25 style='text-align:left'>复查：<font class=" & Chr(34) & "font8" & Chr(34) & ">" & Format(zlCommFun.NVL(rs("复查时间")), "yyyy-MM-dd") & "</td></tr>"
                
                
    '完结
    txtInfo.Text = txtInfo.Text & vbCrLf & "</tr></table></BODY></HTML>"
End Function


Private Sub mnuFileSaveAs_Click()
    Dim strFile As String
       
    If Not HaveExcel Then
        MsgBox "请安装好Excel后，再使用本功能。", vbCritical, gstrSysName
        Exit Sub
    End If
    
    dlg.CancelError = True
    
    On Error GoTo ErrHandler
    
    EditChanged = False
    
    dlg.Flags = &H4 Or &H200000 Or &H800 & &H1000
    dlg.Filter = "体检资料(*.xls)| *.xls"
    dlg.FilterIndex = 0
    
    dlg.DialogTitle = "体检资料收集"
    dlg.FileName = App.Path & "\体检资料收集.xls"
    dlg.ShowSave
    If dlg.FileName <> "" Then Call NewExcelFile(dlg.FileName)
            
    EditChanged = True
    
    Exit Sub
    
ErrHandler:
    EditChanged = True
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize

End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "导出"
        Call mnuFileSaveAs_Click
    Case "发送"
        Call mnuFileMail_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 2 Then txt(2).Tag = "Changed"
    
    If Index = 13 Then
            
        If txt(Index).Tag = "" Then
            Call ResetVsf(vsf)
            txt(Index).Tag = "Changed"
        End If
            
        cmd(4).Tag = ""
        
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If Index <> 7 Then zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        '如果是在团体名称中按了Enter,则要查找历史数据
        If txt(Index).Tag = "Changed" And Index = 13 Then
            
            If InStr(txt(Index).Text, "'") Then
                ShowSimpleMsg "在团体名称中有非法字符 ' ！"
                Exit Sub
            End If
            
            gstrSQL = GetPublicSQL(SQL.团体过滤选择)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(txt(Index).Text) & "%")
            
            If ShowTxtFilter(Me, txt(Index), "名称,1800,0,0;编码,900,0,0;简码,900,0,0;联系人,900,0,0;电话,1200,0,0", Me.Name & "\团体过滤选择", "请从下面选择一个团体单位", rsData, rs, , , , False) Then
                
                txt(Index).Text = zlCommFun.NVL(rs("名称"))
                txt(8).Text = zlCommFun.NVL(rs("电子邮件"))
                cmd(4).Tag = zlCommFun.NVL(rs("ID"))
                                                
                Call ReadData(zlCommFun.NVL(rs("ID")))
            Else
                cmd(4).Tag = ""
            End If
            
            txt(Index).Tag = ""

        End If
        
        zlCommFun.PressKey vbKeyTab
        
        If Index = 13 Then
            zlCommFun.PressKey vbKeyTab
        End If
    End If
    
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

