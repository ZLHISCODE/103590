VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form FrmQualityDataQuery 
   Caption         =   "质控数据查询"
   ClientHeight    =   6525
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   9735
   Icon            =   "FrmQualityDataQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chkShowAvg 
      Caption         =   "显示均线"
      Height          =   300
      Left            =   3555
      TabIndex        =   10
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1065
   End
   Begin VB.CheckBox chkAutoSize 
      Caption         =   "纵坐标自适应"
      Height          =   300
      Left            =   2160
      TabIndex        =   9
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin VB.Frame fraLR_s 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   3210
      MousePointer    =   9  'Size W E
      TabIndex        =   18
      Top             =   30
      Width           =   30
   End
   Begin VB.CheckBox chkShowValue 
      Caption         =   "显示数值"
      Height          =   300
      Left            =   1125
      TabIndex        =   8
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox ChkMultiLine 
      Caption         =   "显示分支"
      Height          =   300
      Left            =   90
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox CmbRes 
      Height          =   300
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   780
      Width           =   1905
   End
   Begin C1Chart2D8.Chart2D ChartMain 
      Height          =   1845
      Left            =   4050
      TabIndex        =   12
      Top             =   2160
      Width           =   2025
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   3572
      _ExtentY        =   3254
      _StockProps     =   0
      ControlProperties=   "FrmQualityDataQuery.frx":020A
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   3360
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":0869
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":0A89
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":0CA9
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":0EC9
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":10E9
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":1309
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":1529
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":1749
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":1965
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":1B85
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":1DA5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   3390
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":20BF
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":22DF
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":24FF
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":271F
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":293F
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":2B5F
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":2D7F
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":2F9F
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":31BB
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":33DB
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQualityDataQuery.frx":35FB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LivData 
      Height          =   1845
      Left            =   6090
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   3254
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   6165
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12091
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
   Begin MSComctlLib.ListView LivMain 
      Height          =   3375
      Left            =   90
      TabIndex        =   14
      Top             =   2670
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "项目"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "靶值"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CV"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   5175
      Left            =   3240
      TabIndex        =   15
      Top             =   810
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   9128
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "图形"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "数据"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1270
      BandCount       =   2
      _CBWidth        =   9735
      _CBHeight       =   720
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinWidth1       =   4500
      MinHeight1      =   660
      Width1          =   9000
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "仪器"
      Child2          =   "CmbDevice"
      MinWidth2       =   2100
      MinHeight2      =   300
      Width2          =   2100
      NewRow2         =   0   'False
      Begin VB.ComboBox CmbDevice 
         Height          =   300
         Left            =   7545
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   2100
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   660
         Left            =   165
         TabIndex        =   17
         Top             =   30
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "规则"
               Key             =   "规则"
               Object.ToolTipText     =   "质控规则"
               Object.Tag             =   "规则"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sdf"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DTBginDate 
      Height          =   300
      Left            =   870
      TabIndex        =   4
      Top             =   1110
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月"
      Format          =   25493504
      CurrentDate     =   38210
   End
   Begin MSComCtl2.DTPicker DTEndData 
      Height          =   300
      Left            =   870
      TabIndex        =   6
      Top             =   1440
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月"
      Format          =   25493504
      CurrentDate     =   38210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "质控品"
      Height          =   180
      Left            =   270
      TabIndex        =   1
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "开始日期"
      Height          =   180
      Left            =   90
      TabIndex        =   3
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "结束日期"
      Height          =   180
      Left            =   90
      TabIndex        =   5
      Top             =   1500
      Width           =   720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
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
         Begin VB.Menu mnuViewToolspilt1 
            Caption         =   "-"
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
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewQualityRule 
         Caption         =   "质控规则"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuShort1 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "增加(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "修改(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "删除(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "增加(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "修改(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "删除(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "详细资料(&D)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "FrmQualityDataQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public QualityRule As String                    '质控规则ID","分开
Dim MouseStartX As Single                       '移动前鼠标的位置
Const MINWIDTH = 2500                           '列表控件最小宽度
Const MaxWidth = 7000                           '列表控件最大宽度
Dim NowFocus As Integer                         '当前得到焦点的控件 1=LivMain;2=LivData
Dim livMainColHead  As Integer                  '排序位置的列(Livmain)
Dim livDataColHead  As Integer                  '排序位置的列(LivData)
Private mLastSeries As Long, mLastPoint As Long, mLastSize As Long

Private Sub chkAutoSize_Click()
    If Me.LivMain.ListItems.Count > 0 Then
        '画线
        LoadResData
    End If
End Sub

Private Sub ChkMultiLine_Click()
    If Me.LivMain.ListItems.Count > 0 Then
        '画线
        LoadResData
    End If
End Sub

Private Sub chkShowAvg_Click()
    On Error Resume Next
    With Me.ChartMain.ChartGroups(1)
        .Styles(11).Line.Pattern = IIf(chkShowAvg.Value = 1, oc2dLineSolid, oc2dLineNone)
    End With
End Sub

Private Sub chkShowValue_Click()
    Dim i As Long
    For i = 1 To Me.ChartMain.ChartLabels.Count
        Me.ChartMain.ChartLabels(i).IsShowing = chkShowValue
    Next
End Sub

Private Sub CmbDevice_Click()
    '读入质控品
    If Me.CmbDevice.ListIndex > -1 Then
        LoadRes (Me.CmbDevice.ItemData(Me.CmbDevice.ListIndex))
    Else
        '没有时清空全部
        Me.CmbRes.Clear
        Me.LivMain.ListItems.Clear
    End If
End Sub

Private Sub CmbRes_Click()
    '读入质控项目
    If Me.CmbRes.ListIndex > -1 Then
        LoadItem (Me.CmbRes.ItemData(Me.CmbRes.ListIndex))
    Else
        Me.LivMain.ListItems.Clear
    End If
End Sub

Private Sub DTBginDate_Change()
    If Me.DTBginDate > Me.DTEndData Then
        Me.DTBginDate = Me.DTEndData
    End If

    If Me.CmbDevice.ListIndex > -1 Then
        LoadRes (Me.CmbDevice.ItemData(Me.CmbDevice.ListIndex))
    End If
End Sub

Private Sub DTEndData_Change()
    If Me.DTEndData < Me.DTBginDate Then
        Me.DTEndData = Me.DTBginDate
    End If

    If Me.CmbDevice.ListIndex > -1 Then
        LoadRes (Me.CmbDevice.ItemData(Me.CmbDevice.ListIndex))
    End If
End Sub

Private Sub Form_Load()
    '初使化
    Initialization
    
    If Me.TabStrip.SelectedItem.Index = 1 Then
        Me.Toolbar1.Buttons("Preview").Enabled = False
        Me.mnuFilePreview.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
    
    '暂时屏蔽方便查问题
    On Error Resume Next
    
    'Cmb
    Me.CmbRes.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) + 30
    Me.CmbRes.Width = Me.fraLR_s.Left - Me.CmbRes.Left - 60
    Me.DTBginDate.Top = Me.CmbRes.Top + Me.CmbRes.Height + 30
    Me.DTBginDate.Width = Me.CmbRes.Width
    Me.DTEndData.Top = Me.DTBginDate.Top + Me.DTBginDate.Height + 30
    Me.DTEndData.Width = Me.CmbRes.Width
    Me.ChkMultiLine.Top = Me.DTEndData.Top + Me.DTEndData.Height + 30
    Me.chkShowValue.Top = Me.ChkMultiLine.Top
    Me.chkAutoSize.Top = Me.ChkMultiLine.Top
    Me.chkShowAvg.Top = Me.ChkMultiLine.Top
    
    'Lable
    Me.Label2.Top = Me.CmbRes.Top + 60
    Me.Label3.Top = Me.DTBginDate.Top + 60
    Me.Label4.Top = Me.DTEndData.Top + 60
    
    'LivMain
    Me.LivMain.Top = Me.ChkMultiLine.Top + Me.ChkMultiLine.Height + 20
    Me.LivMain.Left = 0
    Me.LivMain.Width = Me.fraLR_s.Left
    Me.LivMain.Height = Me.ScaleHeight - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0) - (Me.ChkMultiLine.Top + Me.ChkMultiLine.Height + 40)
    
    'fralr_s
    Me.fraLR_s.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) + 40
    Me.fraLR_s.Left = Me.LivMain.Width
    Me.fraLR_s.Height = Me.ScaleHeight - Me.CoolBar1.Height - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0) - 40
    
    'TabStrip
    Me.TabStrip.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) + 40
    Me.TabStrip.Left = Me.LivMain.Width + Me.fraLR_s.Width
    Me.TabStrip.Width = Me.ScaleWidth - Me.LivMain.Width - Me.fraLR_s.Width
    Me.TabStrip.Height = Me.ScaleHeight - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0) - 40
    
    'chartmain
    Me.ChartMain.Visible = False
    Me.ChartMain.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) + 60 + 320
    Me.ChartMain.Left = Me.LivMain.Width + Me.fraLR_s.Width + 60
    Me.ChartMain.Width = Me.ScaleWidth - Me.LivMain.Width - Me.fraLR_s.Width - 120
    Me.ChartMain.Height = Me.ScaleHeight - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0) - 320 - 120
    If Me.TabStrip.SelectedItem.Index = 1 Then
        Me.ChartMain.Visible = True
    End If
        
    'livdata
    Me.LivData.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) + 60 + 320
    Me.LivData.Left = Me.LivMain.Width + Me.fraLR_s.Width + 60
    Me.LivData.Width = Me.ScaleWidth - Me.LivMain.Width - Me.fraLR_s.Width - 120
    Me.LivData.Height = Me.ScaleHeight - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0) - 320 - 120
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '退出时保存私有设置
    SaveWinState Me, App.ProductName
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "质控间隔", Me.DTEndData - Me.DTBginDate
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "显示质控明细", ChkMultiLine.Value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "显示质控值", chkShowValue.Value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "纵坐标自适应", chkAutoSize.Value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "显示均线", chkShowAvg.Value
End Sub

Private Sub LivData_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    '排序
    If livDataColHead = ColumnHeader.Index - 1 Then '仍是刚才那列
        LivData.SortOrder = IIf(LivData.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        livDataColHead = ColumnHeader.Index - 1
        LivData.SortKey = livDataColHead
        LivData.SortOrder = lvwAscending
    End If
    
End Sub

Private Sub LivMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    '排序
    If livMainColHead = ColumnHeader.Index - 1 Then '仍是刚才那列
        LivMain.SortOrder = IIf(LivMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        livMainColHead = ColumnHeader.Index - 1
        LivMain.SortKey = livMainColHead
        LivMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub LivMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Me.LivMain.ListItems.Count > 0 Then
        '画线
        LoadResData
    End If
End Sub

Private Sub mnuFileExcel_Click()
    '输出Excel
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    '退出
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    '预览
    If Me.TabStrip.SelectedItem.Index = 1 Then Exit Sub
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    If Me.TabStrip.SelectedItem.Index = 1 Then
        With Me.ChartMain
            .PrintChart oc2dFormatBitmap, oc2dScaleToFit, 0, 0, 0, 0
        End With
    Else
        subPrint 1
    End If
End Sub

Private Sub mnuFileSet_Click()
    '打印设置
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    '显示关于
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub mnuHelpTopic_Click()
    '显示帮助
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    '显示主页
    Call zlHomePage(Me.Hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送Email
    Call zlMailTo(Me.Hwnd)
End Sub

Private Sub mnuViewQualityRule_Click()
    '质控规则
    FrmQualityDataQueyRule.Show vbModal, Me
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    '显示或隐藏标准按钮
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    '显示或隐藏文字
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    
    For Each buttTemp In Toolbar1.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    
    Form_Resize
End Sub

Private Sub fralr_s_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        MouseStartX = x
    End If
End Sub

Private Sub fralr_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MoveTmp As Single
    '暂时屏蔽方便查问题
    On Error Resume Next
    If Button = 1 Then
        
        '得到移动后的位置
        MoveTmp = Me.fraLR_s.Left + x - MouseStartX
        
        '超过最大或最小宽度时退出
        If MoveTmp <= MINWIDTH Or MoveTmp >= MaxWidth Then Exit Sub
        
        '移动控件
        
        'fralr_s
        Me.fraLR_s.Left = MoveTmp
        
        'LivMain
        Me.LivMain.Width = Me.fraLR_s.Left
        
        'Frame1
        'Me.Frame1.Width = Me.LivMain.Width
        
        'TabStrip
        Me.TabStrip.Width = Me.ScaleWidth - Me.LivMain.Width - Me.fraLR_s.Width
        Me.TabStrip.Left = Me.fraLR_s.Left + Me.fraLR_s.Width
        
        'chartmain
        Me.ChartMain.Width = Me.TabStrip.Width - 120
        Me.ChartMain.Left = Me.TabStrip.Left + 60
        
        'livdata
        Me.LivData.Width = Me.ChartMain.Width
        Me.LivData.Left = Me.ChartMain.Left
        
        'Cmb
        Me.CmbRes.Width = Me.LivMain.Width - Me.CmbRes.Left - 60
        Me.DTBginDate.Width = Me.CmbRes.Width
        Me.DTEndData.Width = Me.CmbRes.Width
        Me.ChkMultiLine.Width = Me.CmbRes.Width
        
    End If
End Sub

Private Sub TabStrip_Click()
    
    '图形和数据间切换
    Select Case Me.TabStrip.SelectedItem.Index
        Case 1
            '图形
            If Me.ChartMain.Visible = False Then
                Me.LivData.Visible = False
                Me.ChartMain.Visible = True
                NowFocus = 1
            End If
            Me.Toolbar1.Buttons("Preview").Enabled = False
            Me.mnuFilePreview.Enabled = False
        Case 2
            '数据
            If Me.LivData.Visible = False Then
                Me.ChartMain.Visible = False
                Me.LivData.Visible = True
                NowFocus = 2
            End If
            Me.Toolbar1.Buttons("Preview").Enabled = True
            Me.mnuFilePreview.Enabled = True
    End Select
End Sub
Sub DrawLine(DayCount As Integer, SDCost As Double, BX As Double, rsData As ADODB.Recordset, Optional ByVal dblMax As Double = 0)
    '''''''''''''''''''''''''''''''''''''''''''''''
    '功能               画线
    '    参数
    '    DayCount       天数
    '    SDCost         SD值
    '    Bx             靶值
    '    rsData         每日质控数据
    '''''''''''''''''''''''''''''''''''''''''''''''
    Dim Bz As String
    Dim SDz As Integer
    Dim IndexTmp As Integer, iSDTimes As Integer '纵轴最大标准差倍数
    Dim i As Long, N As Long
    
    With Me.ChartMain.ChartGroups(1)
        
        '清除
        .Data.IsBatched = True
        .SeriesLabels.RemoveAll
        .PointLabels.RemoveAll
        
        .Data.NumSeries = 0
       
        '设置标注
        Me.ChartMain.Header.Font.Size = 10
        Me.ChartMain.Header.Interior.ForegroundColor = vbBlue
        If Me.LivMain.SelectedItem Is Nothing Then
            Me.ChartMain.Header.Text = "SD质控图"
        Else
            Me.ChartMain.Header.Text = "SD质控图" & "-" & Me.LivMain.SelectedItem.Text
        End If
        
        'X/Y轴标注
        Me.ChartMain.ChartArea.Axes("X").Title.Text = "时间"
        Me.ChartMain.ChartArea.Axes("Y").Title.Text = "SD"
        
        '坐标原点
        Me.ChartMain.ChartArea.Axes("X").OriginPlacement = oc2dOriginZero
        Me.ChartMain.ChartArea.Axes("Y").OriginPlacement = oc2dOriginZero
        
        .Data.Layout = oc2dDataGeneral '数据设置方式为每个Series拥有各自的X Points
        
        .Data.NumSeries = 11           '设置有几条线
        
        '得到项目IndexTmp
        If Me.LivMain.ListItems.Count > 0 Then
            IndexTmp = Me.LivMain.SelectedItem.Index
            Bz = Me.LivMain.ListItems(IndexTmp).SubItems(1)
        End If
        
        '显示X,Y轴的标注
        If dblMax = 0 Or chkAutoSize.Value <> 1 Then
            iSDTimes = 4
        Else
            iSDTimes = CInt(Abs(dblMax - Bz) / SDCost)
            If iSDTimes * SDCost < Abs(dblMax - Bz) Then iSDTimes = iSDTimes + 1
            If iSDTimes < 4 Then iSDTimes = 4
        End If
        Me.ChartMain.ChartArea.Axes("Y").Max = BX + (SDCost * iSDTimes)
        Me.ChartMain.ChartArea.Axes("Y").Min = BX - (SDCost * iSDTimes)
        Me.ChartMain.ChartArea.Axes("Y").Origin = BX - (SDCost * iSDTimes)

        Me.ChartMain.ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
        Me.ChartMain.ChartArea.Axes("Y").ValueLabels.RemoveAll
        Me.ChartMain.ChartArea.Axes("Y").ValueLabels.Add BX + (SDCost * 3), "+3SD"
        Me.ChartMain.ChartArea.Axes("Y").ValueLabels.Add BX + (SDCost * 2), "+2SD"
        Me.ChartMain.ChartArea.Axes("Y").ValueLabels.Add BX + SDCost, "+1SD"
        Me.ChartMain.ChartArea.Axes("Y").ValueLabels.Add BX, "X(" & Right(Space(8) & Bz, 8) & ")"
        Me.ChartMain.ChartArea.Axes("Y").ValueLabels.Add BX - SDCost, "-1SD"
        Me.ChartMain.ChartArea.Axes("Y").ValueLabels.Add BX - (SDCost * 2), "-2SD"
        Me.ChartMain.ChartArea.Axes("Y").ValueLabels.Add BX - (SDCost * 3), "-3SD"
        
        Me.ChartMain.ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
        Me.ChartMain.ChartArea.Axes("X").ValueLabels.RemoveAll
        
        For i = 1 To DayCount
            Me.ChartMain.ChartArea.Axes("X").ValueLabels.Add i, Day(DTBginDate.Value + i - 1)
        Next
        '设置X,Y开始和最大最小值
        
        Me.ChartMain.ChartArea.Axes("X").Max = DayCount
        Me.ChartMain.ChartArea.Axes("X").Origin = 1

        
      
        '画线基本线
        For i = 1 To .Data.NumSeries
            .Data.NumPoints(i) = DayCount

            Select Case i
                Case 1
                    For N = 1 To DayCount
                        .Data.y(i, N) = BX + (SDCost * 3)
                    Next
                Case 2
                    For N = 1 To DayCount
                        .Data.y(i, N) = BX + (SDCost * 2)
                    Next
                Case 3
                    For N = 1 To DayCount
                        .Data.y(i, N) = BX + SDCost
                    Next
                Case 4
                    For N = 1 To DayCount
                        .Data.y(i, N) = BX
                    Next
                Case 5
                    For N = 1 To DayCount
                        .Data.y(i, N) = BX - SDCost
                    Next
                Case 6
                    For N = 1 To DayCount
                        .Data.y(i, N) = BX - (SDCost * 2)
                    Next
                Case 7
                    For N = 1 To DayCount
                        .Data.y(i, N) = BX - (SDCost * 3)
                    Next
            End Select


        Next

        .Data.IsBatched = False '结束批量更新

        '设置线的颜色和类型
        For i = 1 To .Data.NumSeries
            Select Case i
                Case 1
                    .Styles(i).Line.Pattern = oc2dLineDashDot
                    .Styles(i).Symbol.Shape = oc2dShapeNone
                Case 2
                    .Styles(i).Line.Pattern = oc2dLineLongShortLongDash
                    .Styles(i).Line.COLOR = vbYellow
                    .Styles(i).Symbol.Shape = oc2dShapeNone
                Case 3
                    .Styles(i).Line.Pattern = oc2dLineDotted
                    .Styles(i).Line.COLOR = vbCyan
                    .Styles(i).Symbol.Shape = oc2dShapeNone
                Case 4
                    .Styles(i).Line.Pattern = oc2dLineSolid
                    .Styles(i).Line.COLOR = vbBlack
                    .Styles(i).Symbol.Shape = oc2dShapeNone
                Case 5
                    .Styles(i).Line.Pattern = oc2dLineDotted
                    .Styles(i).Line.COLOR = vbCyan
                    .Styles(i).Symbol.Shape = oc2dShapeNone
                Case 6
                    .Styles(i).Line.Pattern = oc2dLineLongShortLongDash
                    .Styles(i).Line.COLOR = vbYellow
                    .Styles(i).Symbol.Shape = oc2dShapeNone
                Case 7
                    .Styles(i).Line.Pattern = oc2dLineDashDot
                    .Styles(i).Line.COLOR = .Styles(1).Line.COLOR
                    .Styles(i).Symbol.Shape = oc2dShapeNone
                Case 8
                    .Styles(i).Line.Pattern = oc2dLineSolid
                    .Styles(i).Line.COLOR = vbBlue
                    .Styles(i).Symbol.Shape = oc2dShapeNone
                Case 9
                    .Styles(i).Line.Pattern = oc2dLineSolid
                    .Styles(i).Line.COLOR = vbMagenta
                    .Styles(i).Symbol.Shape = oc2dShapeNone
                Case 10
                    .Styles(i).Line.Pattern = oc2dLineSolid
                    .Styles(i).Line.COLOR = vbCyan
                    .Styles(i).Symbol.Shape = oc2dShapeNone
            End Select
        Next
        
        
    End With

End Sub

Sub Initialization()
    '''''''''''''''''''''''''
    '功能           初使化
    '''''''''''''''''''''''''
    Dim lngQryInterval As Long '查询时间间隔
    
    '画线
    Me.ChartMain.Visible = False
    DrawLine 20, 20, 60, Nothing
    Me.ChartMain.Visible = True
    
    '调整时间
    lngQryInterval = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "质控间隔", 0))
    If lngQryInterval < 0 Then lngQryInterval = 0
    Me.DTEndData = date
    Me.DTBginDate = DateAdd("d", -1 * lngQryInterval, date)
    
    ChkMultiLine.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "显示质控明细", 0))
    chkShowValue.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "显示质控值", 0))
    chkAutoSize.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "纵坐标自适应", 0))
    chkShowAvg.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "显示均线", 0))
    
    '读入列表头
    LoadColHead
    
    '恢复私有设置
    RestoreWinState Me, App.ProductName
    
    '读入设备
    LoadDevice
    
    NowFocus = 1

    mLastSeries = -1
End Sub

Sub LoadDevice()
    ''''''''''''''''''''''''''''''''''
    '功能              读入设备
    ''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    
    Me.CmbDevice.Clear
    
    '检验仪器
    gstrSql = "SELECT A.编码||'-'||A.名称,ID FROM 检验仪器 A ORDER BY A.编码||'-'||A.名称"
    Call OpenRecord(rsTmp, gstrSql, Me.Caption)
    If rsTmp.BOF = False Then Call AddComboData(CmbDevice, rsTmp, False)
    CmbDevice.ListIndex = FindComboItem(CmbDevice, Val(Split(GetConnectDevs & ";1", ";")(0)))
    If CmbDevice.ListCount > 0 And CmbDevice.ListIndex = -1 Then CmbDevice.ListIndex = 0
End Sub

Sub LoadRes(DeviceID As Long)
    ''''''''''''''''''''''''''''''''''
    '功能               读入质控品
    '    参数
    '    DeviceID       仪器ID
    ''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    
    Me.CmbRes.Clear
    
    gstrSql = "select * from 检验质控品 where 仪器ID = [1] And Not (开始使用日期 > [2] Or 结束使用日期 < [3])"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, DeviceID, Me.DTEndData.Value, Me.DTBginDate.Value)
    Do Until rsTmp.EOF
        With Me.CmbRes
            .AddItem rsTmp("名称") & IIf(IsNull(rsTmp("批号")), "", "(" & rsTmp("批号") & ")")
            .ItemData(i) = Val(rsTmp("ID"))
            i = i + 1
        End With
        rsTmp.MoveNext
    Loop
            
    If Me.CmbRes.ListCount > 0 Then
        Me.CmbRes.ListIndex = 0
    Else
        Me.LivMain.ListItems.Clear
    End If
    
    rsTmp.Close
End Sub


Sub LoadItem(ResID As Long)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能               读入质控项目和取质控规则默认值
    '    参数
    '    ResID          质控品ID
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim ItmX As ListItem
    
    Me.LivMain.ListItems.Clear
    
    gstrSql = "select a.项目ID,b.中文名,b.英文名 As 缩写,a.靶值,a.SD,a.CV from 检验质控品项目 a ,诊治所见项目 b " & _
                "where a.项目ID+0=b.ID" & _
                " and a.质控品ID = [1] "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, ResID)
        
    Do Until rsTmp.EOF
        With Me.LivMain
            Set ItmX = .ListItems.Add(, "A" & rsTmp("项目ID"), rsTmp("中文名") & "(" & rsTmp("缩写") & ")")
            ItmX.SubItems(1) = Format(zlCommFun.Nvl(rsTmp("靶值")), "###0.00")
            ItmX.SubItems(2) = Format(zlCommFun.Nvl(rsTmp("SD")), "###0.00")
            ItmX.SubItems(3) = zlCommFun.Nvl(rsTmp("CV"))
        End With
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    '得到质控规则的缺省ID
    gstrSql = "select * from 检验质控规则 where 缺省规则 = 1  "
    
    Me.QualityRule = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "选用质控规则", "")
    rsTmp.Open gstrSql, gcnOracle
    
    If rsTmp.EOF <> True And rsTmp.BOF <> True Then
        Me.QualityRule = IIf(Me.QualityRule = "", "", Me.QualityRule & "," & rsTmp("Id"))
    End If
    rsTmp.Close
    
    If Me.LivMain.ListItems.Count > 0 Then
        Set Me.LivMain.SelectedItem = Me.LivMain.ListItems(1)
        LivMain_ItemClick Me.LivMain.SelectedItem
    End If
End Sub

Sub LoadColHead()
    ''''''''''''''''''''''''''''''''''''''
    '功能                   读入列表头
    ''''''''''''''''''''''''''''''''''''''
    
    'LivMain
    With Me.LivMain.ColumnHeaders
        .Clear
        .Add , "A1", "项目"
        .Add , "A2", "靶值"
        .Add , "A3", "SD"
        .Add , "A4", "CV"
        Me.LivMain.Sorted = True
    End With
    
        
    'LivData
    With Me.LivData.ColumnHeaders
        .Clear
        .Add , "A1", "检验日期"
        .Add , "A2", "标本序号"
        .Add , "A3", "检验结果"
        .Add , "A4", "检验部门"
        .Add , "A5", "检验人"
        Me.LivData.ColumnHeaders(5).Alignment = lvwColumnCenter
        Me.LivData.Sorted = True
    End With
End Sub
Sub LoadResData()
    ''''''''''''''''''''''''''''''''''''
    '功能           '读入质控数据并画线
    ''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset, rsMax As New ADODB.Recordset
    Dim rsDayTmp As New ADODB.Recordset
    Dim IndexTmp As Long
    Dim BX As Double                        '得到靶值
    Dim SDx As Double                       '得到SD值
    Dim LineIndex As Integer                '当前第几条线
    Dim i As Integer
    Dim j As Integer, iLastPoint As Integer
    Dim DateTmp As Date
    Dim ItemX As ListItem
    Dim dblSum As Double, dblSS As Double, dblSD As Double
    Dim strQcCond As String '质控号条件
    Dim aQcNO() As String, strEnumQc As String, aRange() As String
    
    On Error GoTo DBError
    '生成质控号的查询条件
    gstrSql = "Select 质控标本号 From 检验质控品 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CmbRes.ItemData(CmbRes.ListIndex))
    If Not rsTmp.EOF Then
        If Len(zlCommFun.Nvl(rsTmp(0))) = 0 Then
            strQcCond = " And (a.是否质控品 = 1 or instr(',' || c.质控标本号 || ',' , ',' || a.标本序号 || ',') > 0 )"
        Else
            aQcNO = Split(rsTmp(0), ",")
            strQcCond = "": strEnumQc = ""
            For i = 0 To UBound(aQcNO)
                If InStr(aQcNO(i), "-") > 0 Then '范围
                    aRange = Split(aQcNO(i), "-")
                    If Val(aRange(0)) > 0 And Val(aRange(1)) > 0 Then
                        strQcCond = strQcCond & " Or a.标本序号 Between " & Val(aRange(0)) & " And " & Val(aRange(1))
                    Else
                        If Val(aRange(0)) > 0 Then
                            strQcCond = strQcCond & " Or a.标本序号 >= " & Val(aRange(0))
                        End If
                        If Val(aRange(1)) > 0 Then
                            strQcCond = strQcCond & " Or a.标本序号 <= " & Val(aRange(1))
                        End If
                    End If
                Else
                    strEnumQc = strEnumQc & "," & aQcNO(i)
                End If
            Next
            If Len(strQcCond) > 0 Then strQcCond = Mid(strQcCond, 5)
            If Len(strEnumQc) > 0 Then
                strQcCond = " And (Instr('" & strEnumQc & ",',',' || a.标本序号 || ',')>0" & _
                    IIf(Len(strQcCond) = 0, "", " Or " & strQcCond) & ")"
            Else
                strQcCond = IIf(Len(strQcCond) = 0, "", " And (" & strQcCond & ")")
            End If
        End If
    Else
        strQcCond = " And (a.是否质控品 = 1 or instr(',' || c.质控标本号 || ',' , ',' || a.标本序号 || ',') > 0 )"
    End If
    
'    gstrSql = "    select  trunc(a.检验时间) as 检验时间 ,avg(nvl(replace(b.检验结果,'#',''),0)) as 检验结果 " & _
'                "    from 检验标本记录 a , 检验普通结果 b , 检验仪器 c,检验质控品 D" & _
'                "    Where a.ID+0 = b.检验标本id " & _
'                "    and a.仪器id+0 = c.id " & _
'                "    and a.报告结果 = b.记录类型 " & _
'                "    and ((D.质控标本号 Is Null" & _
'                "          And (a.是否质控品 = 1 or instr(',' || c.质控标本号 || ',' , ',' || a.标本序号 || ',') > 0 )) " & _
'                "         Or Instr(',' || D.质控标本号 || ',' , ',' || a.标本序号 || ',') > 0) " & _
'                "    and a.核收时间 between [1] and [2] " & _
'                "    and a.仪器ID = [3] " & _
'                "    and b.检验项目ID+0 = [4] " & _
'                "    And D.ID=[5] " & _
'                "    group by trunc(a.检验时间) Having avg(nvl(replace(b.检验结果,'#',''),0))<>0" & _
'                "    order by trunc(a.检验时间)"
    gstrSql = "    select  trunc(a.检验时间) as 检验时间 ,avg(nvl(replace(b.检验结果,'#',''),0)) as 检验结果 " & _
                "    from 检验标本记录 a , 检验普通结果 b , 检验仪器 c" & _
                "    Where a.ID+0 = b.检验标本id " & _
                "    and a.仪器id+0 = c.id " & _
                "    and a.报告结果 = b.记录类型 " & strQcCond & _
                "    and a.核收时间 between [1] and [2] " & _
                "    and a.仪器ID = [3] " & _
                "    and b.检验项目ID+0 = [4] And Nvl(a.标本类别,0)=0" & _
                "    group by trunc(a.检验时间) Having avg(nvl(replace(b.检验结果,'#',''),0))<>0" & _
                "    order by trunc(a.检验时间)"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(Format(DTBginDate, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(DTEndData, "yyyy-MM-dd") & " 23:59:59"), _
                CmbDevice.ItemData(Me.CmbDevice.ListIndex), Mid(Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Key, 2), CmbRes.ItemData(CmbRes.ListIndex))
    '查最大、最小值
    gstrSql = "Select Max(检验结果),Min(检验结果) From (" & _
                "select  trunc(a.检验时间) as 检验时间 ,avg(nvl(replace(b.检验结果,'#',''),0)) as 检验结果 " & _
                "    from 检验标本记录 a , 检验普通结果 b , 检验仪器 c" & _
                "    Where a.ID+0 = b.检验标本id " & _
                "    and a.仪器id+0 = c.id " & _
                "    and a.报告结果 = b.记录类型 " & strQcCond & _
                "    and a.核收时间 between [1] and [2] " & _
                "    and a.仪器ID = [3] " & _
                "    and b.检验项目ID+0 = [4] And Nvl(a.标本类别,0)=0" & _
                "    group by trunc(a.检验时间) Having avg(nvl(replace(b.检验结果,'#',''),0))<>0" & _
                "    order by trunc(a.检验时间))"

    Set rsMax = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(Format(DTBginDate, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(DTEndData, "yyyy-MM-dd") & " 23:59:59"), _
                CmbDevice.ItemData(Me.CmbDevice.ListIndex), Mid(Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Key, 2), CmbRes.ItemData(CmbRes.ListIndex))
    
    '查询明细质控数据
    gstrSql = "select  a.id, a.标本序号, a.检验时间,b.检验结果,a.核收时间,a.检验人,e.名称 " & _
                " from 检验标本记录 a , 检验普通结果 b , 检验仪器 c,部门表 E" & _
                " Where a.ID+0 = b.检验标本id " & _
                " and a.报告结果 = b.记录类型 " & _
                " and a.仪器id+0 = c.id and a.执行科室ID+0=e.ID " & strQcCond & _
                " and a.核收时间 between [1] and [2] " & _
                " and a.仪器ID = [3] " & _
                " and b.检验项目ID+0 = [4] And Nvl(a.标本类别,0)=0" & _
                " order by trunc(a.检验时间),a.标本序号"
    Set rsDayTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(Format(DTBginDate, "yyyy-MM-dd 00:00:00")), CDate(Format(DTEndData, "yyyy-MM-dd 23:59:59")), _
                CmbDevice.ItemData(Me.CmbDevice.ListIndex), Mid(Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Key, 2), CmbRes.ItemData(CmbRes.ListIndex))
        
    Me.ChartMain.Visible = False
    If rsTmp.EOF <> True And rsTmp.BOF <> True Then
        
        '选择一个质控品并画线
        IndexTmp = Me.LivMain.SelectedItem.Index
        BX = Me.LivMain.ListItems(IndexTmp).SubItems(1)
        SDx = Me.LivMain.ListItems(IndexTmp).SubItems(2)
        DrawLine DTEndData.Value - DTBginDate.Value + 1, SDx, BX, rsTmp, IIf(Abs(rsMax(0) - BX) > Abs(rsMax(1) - BX), rsMax(0), rsMax(1))
        rsTmp.MoveFirst
        
        j = 0: iLastPoint = 0
        i = 0
        
        '开始画线
        dblSum = 0: dblSS = 0
        Me.ChartMain.ChartLabels.RemoveAll
        Do Until rsTmp.EOF
            j = rsTmp("检验时间") - DTBginDate + 1
            If Me.ChkMultiLine.Value = 1 Then
                i = 0
                Do Until rsDayTmp.EOF
                    If Format(rsDayTmp("核收时间"), "yyyy-MM-dd") <> Format(rsTmp("检验时间"), "yyyy-MM-dd") Then Exit Do
                    '最多画3条
                    i = i + 1
                    If i < 4 Then
                        With Me.ChartMain.ChartGroups(1)
                            .PointStyles.Add 7 + i, j
                            .PointStyles(7 + i, j).Symbol.COLOR = vbGreen
                            .PointStyles(7 + i, j).Symbol.Shape = oc2dShapeBox
                            .Data.y(7 + i, j) = zlCommFun.Nvl(rsDayTmp("检验结果"), 0)
                            
                            '显示数值
                            Me.ChartMain.ChartLabels.Add
                            Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).AttachMethod = oc2dAttachDataIndex
                            Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).AttachDataIndex.ChartGroup = 1
                            Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).AttachDataIndex.Point = j
                            Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).AttachDataIndex.Series = 7 + i
                            Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).Adjust = oc2dAdjustRight
                            Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).Text = Format(zlCommFun.Nvl(rsDayTmp("检验结果"), 0), "0.00")
                            Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).IsShowing = chkShowValue.Value
                        End With
                        DrawConnectLine 7 + i, iLastPoint, j
                    End If
    
                    rsDayTmp.MoveNext
                Loop
            Else
                '画一条线
                With Me.ChartMain.ChartGroups(1)
                    .PointStyles.Add 8, j
                    .PointStyles(8, j).Symbol.COLOR = vbGreen
                    .PointStyles(8, j).Symbol.Shape = oc2dShapeBox
                    .Data.y(8, j) = zlCommFun.Nvl(rsTmp("检验结果"), 0)
                    '显示数值
                    Me.ChartMain.ChartLabels.Add
                    Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).AttachMethod = oc2dAttachDataIndex
                    Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).AttachDataIndex.ChartGroup = 1
                    Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).AttachDataIndex.Point = j
                    Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).AttachDataIndex.Series = 8
                    Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).Adjust = oc2dAdjustRight
                    Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).Text = Format(zlCommFun.Nvl(rsTmp("检验结果"), 0), "0.00")
                    Me.ChartMain.ChartLabels(Me.ChartMain.ChartLabels.Count).IsShowing = chkShowValue.Value
                End With
                DrawConnectLine 8, iLastPoint, j
            End If
            
            iLastPoint = j
            '计算实际均值、SD、CV
            dblSum = dblSum + zlCommFun.Nvl(rsTmp("检验结果"), 0)
            dblSS = dblSS + zlCommFun.Nvl(rsTmp("检验结果"), 0) ^ 2
            rsTmp.MoveNext
        Loop
        
        '处理只有一条记录时的情况
        If rsTmp.RecordCount = 1 Then
            '计算实际均值、SD、CV
            dblSD = 0
            Me.ChartMain.Header.Text = Me.ChartMain.Header.Text & vbCrLf & _
                "实际均值：" & Right(Space(10) & Format(dblSum / rsTmp.RecordCount, "0.00"), 10) & _
                "  实际标准差：" & Right(Space(10) & Format(dblSD, "0.000"), 10) & _
                "  实际CV：" & Right(Space(10) & Format(dblSD / (dblSum / rsTmp.RecordCount), "0.0000"), 10)
            '画线
            Me.ChartMain.Visible = False
            DrawLine DTEndData.Value - DTBginDate.Value + 1, 20, 60, Nothing
        Else
            '计算实际均值、SD、CV
            dblSD = Sqr((dblSS - dblSum ^ 2 / rsTmp.RecordCount) / (rsTmp.RecordCount - 1))
            Me.ChartMain.Header.Text = Me.ChartMain.Header.Text & vbCrLf & _
                "实际均值：" & Right(Space(10) & Format(dblSum / rsTmp.RecordCount, "0.00"), 10) & _
                "  实际标准差：" & Right(Space(10) & Format(dblSD, "0.000"), 10) & _
                "  实际CV：" & Right(Space(10) & Format(dblSD / (dblSum / rsTmp.RecordCount), "0.0000"), 10)            '质控检验，并显示出来
        End If
        rsDayTmp.MoveFirst
    
        '显示均线
        With Me.ChartMain.ChartGroups(1)
            For i = 1 To DTEndData.Value - DTBginDate.Value + 1
                .Data.y(11, i) = dblSum / rsTmp.RecordCount
            Next
            .Styles(11).Line.COLOR = vbGreen
            .Styles(11).Line.Pattern = IIf(chkShowAvg = 1, oc2dLineSolid, oc2dLineNone)
            .Styles(11).Symbol.Shape = oc2dShapeNone
        End With
    Else
        '画线
        DrawLine DTEndData.Value - DTBginDate.Value + 1, 20, 60, Nothing
        '计算实际均值、SD、CV
        Me.ChartMain.Header.Text = Me.ChartMain.Header.Text & vbCrLf & _
            "实际均值：" & Space(10) & _
            "  实际标准差：" & Space(10) & _
            "  实际CV：" & Space(10)
        '质控检验，并显示出来
    End If
    '判断并显示失控状态
    If rsTmp.RecordCount = 0 Then
        ShowCheckRule 0
    Else
        ShowCheckRule dblSum / rsTmp.RecordCount
    End If
    If Me.TabStrip.SelectedItem.Index = 1 Then
        Me.ChartMain.Visible = True
    End If
    
    rsTmp.Close
    rsMax.Close
    '写入数据到列表
    Me.LivData.ListItems.Clear
    Do Until rsDayTmp.EOF
        Set ItemX = Me.LivData.ListItems.Add(, "A" & rsDayTmp("ID"), Format(rsDayTmp("检验时间"), "yyyy-MM-dd hh:mm:ss"))
        ItemX.SubItems(1) = rsDayTmp("标本序号")
        ItemX.SubItems(2) = Format(rsDayTmp("检验结果"), "###0.00")
        ItemX.SubItems(3) = zlCommFun.Nvl(rsDayTmp("名称"))
        ItemX.SubItems(4) = zlCommFun.Nvl(rsDayTmp("检验人"))
        rsDayTmp.MoveNext
    Loop
    rsDayTmp.Close
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DrawConnectLine(ByVal lngSeries As Long, ByVal lngStartPoint As Long, ByVal lngEndPoint As Long)
    '如果两个连续的值之间有断点，则直接连接这两个值
    Dim i As Long, dblAdd As Double
    If lngStartPoint = 0 Then Exit Sub
    If lngEndPoint - lngStartPoint <= 1 Then Exit Sub
    
    With Me.ChartMain.ChartGroups(1)
        dblAdd = (.Data(lngSeries, lngEndPoint) - .Data(lngSeries, lngStartPoint)) / (lngEndPoint - lngStartPoint)
        For i = 1 To lngEndPoint - lngStartPoint - 1
            .PointStyles.Add lngSeries, lngStartPoint + i
            .PointStyles(lngSeries, lngStartPoint + i).Symbol.COLOR = vbWhite
            .PointStyles(lngSeries, lngStartPoint + i).Symbol.Shape = oc2dShapeNone
            .Data.y(lngSeries, lngStartPoint + i) = .Data.y(lngSeries, lngStartPoint) + i * dblAdd
        Next
    End With
End Sub

Function CheckRule(Rule As Integer, Optional Rule_N As Integer, Optional Rule_X As Integer, Optional Rule_M As Integer, Optional ByVal dblAvg As Double) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                       检查规则
    '    Rule                   =1规则1;=2规则2;...
    '    Rule_N                 N的值
    '    Rule_X                 X的值
    '    Rule_M                 M的值
    '返回                       =True违返规则;=false正常
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim IndexTmp As Integer
    Dim strTmp As String
    Dim SD As Double                            '标准差
    Dim BX As Double                            '靶值
    Dim ExceedSDCout As Integer                 '超过SD个数
    Dim ResCost As Double                       '临时记录上一数据的值
    Dim intTmp As Integer
    Dim strChkPoint() As String                 '失控点：格式为[序列号]-[序列内顺序号]
    Dim i As Long, aTmp() As String, j As Integer, N As Long
    Dim iState As Integer                       '结果状态：0－上升或在靶值上面、1－下降或在靶值下面
    
    ReDim strChkPoint(0) As String
    
    IndexTmp = Me.LivMain.SelectedItem.Index
    BX = Me.LivMain.ListItems(IndexTmp).SubItems(1)
    SD = Me.LivMain.ListItems(IndexTmp).SubItems(2)
    
    '规则1(1:N-XS N个连续结果超过X倍标准差)
    If Rule = 1 Then
        If Rule_N <= 0 Or Rule_X <= 0 Then
            Exit Function
        End If
        '单线
        If Me.ChkMultiLine.Value <> 1 Then
            With Me.ChartMain.ChartGroups(1)
                For i = 1 To .Data.LastPoint(8)
                    If (.Data(8, i) > BX + (SD * Rule_X) Or .Data(8, i) < BX - (SD * Rule_X)) And .Data(8, i) <> 1E+308 Then
                        If .PointStyles(8, i).Symbol.Shape <> oc2dShapeNone Then
                            ExceedSDCout = ExceedSDCout + 1
                            ReDim Preserve strChkPoint(ExceedSDCout) As String
                            strChkPoint(ExceedSDCout) = "8-" & i
                            If ExceedSDCout >= Rule_N Then
                                CheckRule = True
                            End If
                        End If
                    Else
                        If Not CheckRule Then ExceedSDCout = 0
                    End If
                Next
            End With
        Else
            '多线
            With Me.ChartMain.ChartGroups(1)
                For j = 8 To 10
                    For i = 1 To .Data.LastPoint(j)
                        If (.Data(j, i) > BX + (SD * Rule_X) Or .Data(j, i) < BX - (SD * Rule_X)) And .Data(j, i) <> 1E+308 Then
                            If .PointStyles(j, i).Symbol.Shape <> oc2dShapeNone Then
                                ExceedSDCout = ExceedSDCout + 1
                                ReDim Preserve strChkPoint(ExceedSDCout) As String
                                strChkPoint(ExceedSDCout) = j & "-" & i
                                If ExceedSDCout >= Rule_N Then
                                    CheckRule = True
                                End If
                            End If
                        Else
                            If Not CheckRule Then ExceedSDCout = 0
                        End If
                    Next
                Next
            End With
        End If
    End If
    
    '规则2:R-Xs 同一批结果之差超过X倍标准差.
    If Rule = 2 Then
        If Rule_X <= 0 Then
            Exit Function
        End If
        '单线
        If Me.ChkMultiLine.Value <> 1 Then
            With Me.ChartMain.ChartGroups(1)
                If (Abs(.Data.DataMax(8) - .Data.DataMin(8)) > Abs(SD * Rule_X)) Then
                    CheckRule = True
                    Exit Function
                End If
            End With
        Else
            With Me.ChartMain.ChartGroups(1)
                For j = 8 To 10
                    If (Abs(.Data.DataMax(j) - .Data.DataMin(j)) > Abs(SD * Rule_X)) Then
                        CheckRule = True
                        Exit Function
                    End If
                Next
            End With
        End If
    End If
    
    '3:N-T 连续N个结果上升或下降
    If Rule = 3 Then
        If Rule_N <= 0 Then
            Exit Function
        End If
        '单线
        If Me.ChkMultiLine.Value <> 1 Then
            With Me.ChartMain.ChartGroups(1)
                iState = -1
                For i = 2 To .Data.LastPoint(8)
                    If .Data(8, i) <> 1E+308 Then
                        If .PointStyles(8, i).Symbol.Shape <> oc2dShapeNone Then
                            If .Data(8, i) > .Data(8, i - 1) Then
                                '上升
                                If iState <> 0 And Not CheckRule Then ExceedSDCout = 0
                                
                                iState = 0
                                ExceedSDCout = ExceedSDCout + 1
                            ElseIf .Data(8, i) < .Data(8, i - 1) Then
                                '下降
                                If iState <> 1 And Not CheckRule Then ExceedSDCout = 0
                                
                                iState = 1
                                ExceedSDCout = ExceedSDCout + 1
                            Else
                                '相等
                                If Not CheckRule Then ExceedSDCout = 0
                                
                                iState = -1
                            End If
                            If ExceedSDCout > 0 Then
                                ReDim Preserve strChkPoint(ExceedSDCout) As String
                                strChkPoint(ExceedSDCout) = 8 & "-" & i
                                If ExceedSDCout >= Rule_N Then
                                    CheckRule = True
                                End If
                            End If
                        End If
                    End If
                Next
            End With
        Else
            '多线
            With Me.ChartMain.ChartGroups(1)
                For j = 8 To 10
                    iState = -1
                    For i = 2 To .Data.LastPoint(j)
                        If .Data(j, i) <> 1E+308 Then
                            If .PointStyles(j, i).Symbol.Shape <> oc2dShapeNone Then
                                If .Data(j, i) > .Data(j, i - 1) Then
                                    '上升
                                    If iState <> 0 And Not CheckRule Then ExceedSDCout = 0
                                    
                                    iState = 0
                                    ExceedSDCout = ExceedSDCout + 1
                                ElseIf .Data(j, i) < .Data(j, i - 1) Then
                                    '下降
                                    If iState <> 1 And Not CheckRule Then ExceedSDCout = 0
                                    
                                    iState = 1
                                    ExceedSDCout = ExceedSDCout + 1
                                Else
                                    '相等
                                    If Not CheckRule Then ExceedSDCout = 0
                                    
                                    iState = -1
                                End If
                                If ExceedSDCout > 0 Then
                                    ReDim Preserve strChkPoint(ExceedSDCout) As String
                                    strChkPoint(ExceedSDCout) = j & "-" & i
                                    If ExceedSDCout >= Rule_N Then
                                        CheckRule = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                Next
            End With
        End If

    End If
    
    '规则4:N-X 连续N个结果在一侧
    If Rule = 4 Then
        If Rule_N <= 0 Then
            Exit Function
        End If
        '单线
        If Me.ChkMultiLine.Value <> 1 Then
            With Me.ChartMain.ChartGroups(1)
                iState = -1
                For i = 1 To .Data.LastPoint(8)
                    If .Data(8, i) <> 1E+308 Then
                        If .PointStyles(8, i).Symbol.Shape <> oc2dShapeNone Then
                            If .Data(8, i) > BX Then
                                '上面
                                If iState <> 0 And Not CheckRule Then ExceedSDCout = 0
                                
                                iState = 0
                                ExceedSDCout = ExceedSDCout + 1
                            ElseIf .Data(8, i) < BX Then
                                '下面
                                If iState <> 1 And Not CheckRule Then ExceedSDCout = 0
                                
                                iState = 1
                                ExceedSDCout = ExceedSDCout + 1
                            Else
                                '相等
                                If Not CheckRule Then ExceedSDCout = 0
                                
                                iState = -1
                            End If
                            If ExceedSDCout > 0 Then
                                ReDim Preserve strChkPoint(ExceedSDCout) As String
                                strChkPoint(ExceedSDCout) = 8 & "-" & i
                                If ExceedSDCout >= Rule_N Then
                                    CheckRule = True
                                End If
                            End If
                        End If
                    End If
                Next
            End With
        Else
            With Me.ChartMain.ChartGroups(1)
                For j = 8 To 10
                    iState = -1
                    For i = 1 To .Data.LastPoint(j)
                        If .Data(j, i) <> 1E+308 Then
                            If .PointStyles(j, i).Symbol.Shape <> oc2dShapeNone Then
                                If .Data(j, i) > BX Then
                                    '上面
                                    If iState <> 0 And Not CheckRule Then ExceedSDCout = 0
                                    
                                    iState = 0
                                    ExceedSDCout = ExceedSDCout + 1
                                ElseIf .Data(j, i) < BX Then
                                    '下面
                                    If iState <> 1 And Not CheckRule Then ExceedSDCout = 0
                                    
                                    iState = 1
                                    ExceedSDCout = ExceedSDCout + 1
                                Else
                                    '相等
                                    If Not CheckRule Then ExceedSDCout = 0
                                    
                                    iState = -1
                                End If
                                If ExceedSDCout > 0 Then
                                    ReDim Preserve strChkPoint(ExceedSDCout) As String
                                    strChkPoint(ExceedSDCout) = j & "-" & i
                                    If ExceedSDCout >= Rule_N Then
                                        CheckRule = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                Next
            End With
        End If
    End If
    
    '规则5:(M of N)XS 连续N个结果中有M个结果超过X倍标准差
    If Rule = 5 Then
        If Rule_N < 0 Or Rule_X <= 0 Or Rule_M <= 0 Then
            Exit Function
        End If
        '单线
        If Me.ChkMultiLine.Value <> 1 Then
            With Me.ChartMain.ChartGroups(1)
                For i = 1 To .Data.LastPoint(8)
                    For N = 0 To Rule_N - 1
                        If i + N <= .Data.LastPoint(8) Then
                            If (.Data(8, i + N) > BX + (SD * Rule_X) Or .Data(8, i + N) < BX - (SD * Rule_X)) And .Data(8, i) <> 1E+308 Then
                                If .PointStyles(8, i).Symbol.Shape <> oc2dShapeNone Then
                                    ExceedSDCout = ExceedSDCout + 1
                                    ReDim Preserve strChkPoint(ExceedSDCout) As String
                                    strChkPoint(ExceedSDCout) = 8 & "-" & i + N
                                End If
                            End If
                            
                            If ExceedSDCout >= Rule_M Then
                                CheckRule = True
                            End If
                        End If
                    Next
                    If CheckRule Then
                        Exit For
                    Else
                        ExceedSDCout = 0
                    End If
                Next
            End With
        Else
            With Me.ChartMain.ChartGroups(1)
                For j = 8 To 10
                    For i = 1 To .Data.LastPoint(j)
                        For N = 0 To Rule_N - 1
                            If i + N <= .Data.LastPoint(j) Then
                                If (.Data(j, i + N) > BX + (SD * Rule_X) Or .Data(j, i + N) < BX - (SD * Rule_X)) And .Data(j, i) <> 1E+308 Then
                                    If .PointStyles(j, i).Symbol.Shape <> oc2dShapeNone Then
                                        ExceedSDCout = ExceedSDCout + 1
                                        ReDim Preserve strChkPoint(ExceedSDCout) As String
                                        strChkPoint(ExceedSDCout) = j & "-" & i + N
                                    End If
                                End If
                                
                                If ExceedSDCout >= Rule_M Then
                                    CheckRule = True
                                End If
                            End If
                        Next
                        If CheckRule Then
                            Exit For
                        Else
                            ExceedSDCout = 0
                        End If
                    Next
                    If CheckRule Then Exit For
                Next
            End With
        End If
    End If
    
    '将失控点标注为红色
    If CheckRule Then
        For i = 1 To UBound(strChkPoint)
            aTmp = Split(strChkPoint(i), "-")
            Me.ChartMain.ChartGroups(1).PointStyles(Val(aTmp(0)), Val(aTmp(1))).Symbol.COLOR = vbRed
            Me.ChartMain.ChartGroups(1).PointStyles(Val(aTmp(0)), Val(aTmp(1))).Symbol.Shape = oc2dShapeTriangle
        Next
    End If
End Function
Sub ShowCheckRule(ByVal dblAvg As Double)
    '''''''''''''''''''''''''''''''''''''''''''
    '功能               显示质控规则结果
    '''''''''''''''''''''''''''''''''''''''''''
    Dim intTmp As Integer
    Dim strTmp As String, strFoot As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    strTmp = Me.QualityRule
    
    Me.ChartMain.Footer.Location.Left = 10
    Me.ChartMain.Footer.Font.Bold = True
    Me.ChartMain.Footer.Adjust = oc2dAdjustLeft
    
    gstrSql = "select * from 检验质控规则 where id = [1] "
    Do Until Len(strTmp) = 0
        intTmp = InStr(strTmp, ",")
        If intTmp = 0 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTmp)
            strTmp = ""
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(strTmp, 1, intTmp - 1))
            strTmp = Mid(strTmp, intTmp + 1)
        End If
        
        If rsTmp.EOF <> True And rsTmp.BOF <> True Then
            If CheckRule(rsTmp("规则类型") + 1, rsTmp("N"), rsTmp("X"), rsTmp("M"), dblAvg) = True Then
                i = i + 1
                strFoot = strFoot & " " & i & ".违返规则：" & rsTmp("规则名称") & _
                    IIf(IsNull(rsTmp("说明")), "", "--" & rsTmp("说明")) & vbCrLf
            End If
        End If
        
        rsTmp.Close
    Loop
    Me.ChartMain.Footer.Text = strFoot
End Sub

Private Sub subPrint(bytMode As Byte)
    '''''''''''''''''''''''''''''''''''''''''''
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '''''''''''''''''''''''''''''''''''''''''''
    Dim objPrint As New zlPrintLvw
    
    If gstrUserName = "" Then Call GetUserInfo
    
    Select Case NowFocus
        Case 1
            If LivMain.SelectedItem Is Nothing Then Exit Sub
    
            If LivMain.ListItems.Count = 0 Then Exit Sub
            
            Set objPrint.Body.objData = Me.LivMain
        Case 2
            If LivData.SelectedItem Is Nothing Then Exit Sub
    
            If LivData.ListItems.Count = 0 Then Exit Sub
            
            Set objPrint.Body.objData = Me.LivData
    End Select
    
    objPrint.Title.Text = "质控查询"
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
        '处理按下按钮
    Select Case Button.Key
        
        Case "Quit"
            '退出
            mnuFileExit_Click
        Case "Print"
            '打印
            mnuFilePrint_Click
        Case "Preview"
            '预览
            mnuFilePreview_Click
        Case "Help"
            '帮助
            mnuHelpTopic_Click
        Case "规则"
            '规则
            mnuViewQualityRule_Click
    End Select
End Sub

Private Sub ChartMain_DblClick()
    If mLastSeries <> -1 Then
        With ChartMain.ChartGroups(1)
            .Data(mLastSeries, mLastPoint) = Val(InputBox("输入新值：", "修改", .Data(mLastSeries, mLastPoint)))
        End With
    End If
End Sub

Private Sub ChartMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objDataIndex As New Chart2DDataIndexResult
    With ChartMain.ChartGroups(1)
        Set objDataIndex = .CoordToDataIndexObject(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, oc2dFocusXY)
        If objDataIndex.Region = oc2dRegionInChartArea And objDataIndex.Distance < 10 Then
            If mLastSeries <> -1 Then
                If (mLastSeries = objDataIndex.Series And mLastPoint = objDataIndex.Point) Then
                    Exit Sub
                Else
                    .PointStyles(mLastSeries, mLastPoint).Symbol.Size = mLastSize
                End If
            End If
            .PointStyles.Add objDataIndex.Series, objDataIndex.Point
            mLastSeries = objDataIndex.Series
            mLastPoint = objDataIndex.Point
            mLastSize = .PointStyles(mLastSeries, mLastPoint).Symbol.Size
            .PointStyles(mLastSeries, mLastPoint).Symbol.Size = 10
        Else
            If mLastSeries <> -1 Then
                .PointStyles(mLastSeries, mLastPoint).Symbol.Size = mLastSize
                mLastSeries = -1
            End If
        End If
    End With
End Sub

