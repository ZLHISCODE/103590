VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManagePrice 
   AutoRedraw      =   -1  'True
   Caption         =   "门诊划价管理"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9675
   Icon            =   "frmManagePrice.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picVsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   7410
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1695
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4170
      Width           =   45
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   15
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9675
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4140
      Width           =   9675
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9675
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "划价"
               Key             =   "Price"
               Description     =   "划价"
               Object.ToolTipText     =   "进入划价窗口"
               Object.Tag             =   "划价"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modi"
               Description     =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Del"
               Description     =   "删除"
               Object.ToolTipText     =   "删除当前选择的划价单"
               Object.Tag             =   "删除"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅当前单据的内容"
               Object.Tag             =   "查阅"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "按设置条件重新筛选记录"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5844
      Width           =   9672
      _ExtentX        =   17066
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManagePrice.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11986
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
            AutoSize        =   2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
      Height          =   1665
      Left            =   7470
      TabIndex        =   2
      Top             =   4185
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   2937
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManagePrice.frx":115E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1665
      Left            =   0
      TabIndex        =   1
      Top             =   4185
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   2937
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManagePrice.frx":1478
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   6006
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManagePrice.frx":1792
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   60
      Top             =   450
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
            Picture         =   "frmManagePrice.frx":1AAC
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":1CC6
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":1EE0
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":20FA
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":2314
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":2A8E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":2CA8
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":2EC2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":30DC
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":32F6
            Key             =   "Adjust"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":3510
            Key             =   "Modi"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   645
      Top             =   450
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
            Picture         =   "frmManagePrice.frx":372A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":3944
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":3B5E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":3D78
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":3F92
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":470C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":4926
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":4B40
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":4D5A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":4F74
            Key             =   "Adjust"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":518E
            Key             =   "Modi"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Price 
         Caption         =   "门诊划价(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Price_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "修改单据(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit_Adjust 
         Caption         =   "调整时间(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEdit_Adjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "删除单据(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "打印划价通知单(&P)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit_Del_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "查阅单据(&V)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
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
         Begin VB.Menu mnuView_Tlb_1 
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
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "刷新方式(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后不要刷新数据(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后提示是否刷新(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后自动刷新数据(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB上的中联"
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
Attribute VB_Name = "frmManagePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsList As ADODB.Recordset  '单据列表
Private mblnNOMoved  As Boolean '筛选结果为收费单据时当前单据是否在后备表中.
Private mrsDetail As ADODB.Recordset
Private mrsMoney As ADODB.Recordset

Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    Operator As String
    PatientID As Long '门诊号 用于精确查找
    PatientName As String '姓名 用于模糊查找
    ChargeKind As String
    DeptID As Long
    str收费类别 As String
    int门诊标志 As Integer  '1-门诊;2-住院;3-门诊和住院 126174
End Type
Private SQLCondition As Type_SQLCondition

Private mstrFilter As String
Private mblnMax As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mbln收费 As Boolean

'消息相关对象变量
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
End Sub

Private Sub mnuEdit_Adjust_Click()
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo = "" Then
        MsgBox "当前没有单据可以调整！", vbInformation, gstrSysName
        Exit Sub
    End If

    On Error Resume Next
    Err.Clear
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 1
    frmCharge.mbytInState = 2
    frmCharge.mstrInNO = strNo
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改单据清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_Modi_Click()
    Dim strNo As String
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If strNo = "" Then
        MsgBox "当前没有单据可以修改！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrModiNO = ""
    
    On Error Resume Next
    Err.Clear
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mstrInNO = strNo
    frmCharge.mbytInFun = 1
    frmCharge.mbytInState = 0
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
    If gblnOK And gstrModiNO <> "" Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改单据清单内容,修改后的单据号为:[" & gstrModiNO & "],要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_Print_Click()
    Dim strNo As String
    
    If mbln收费 Then Exit Sub   '收过费了就没有划价单了
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo <> "" Then
        If MsgBox("确实要打印当前单据的划价通知单吗！", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me, "NO=" & strNo, 2)
        End If
    Else
        MsgBox "当前没有单据可以打印！", vbInformation, gstrSysName
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Dim blnPre As Boolean, intFrom As Integer
    
    blnPre = gbln药房单位
    intFrom = gint病人来源
        
    With frmSetExpence
        .mlngModul = mlngModul
        .mstrPrivs = mstrPrivs
        .mbytInFun = 1
        .mblnSetDrugStore = False
        .Show 1, Me
    End With
    
    '更改了药品单位参数,重新刷新
    If gbln药房单位 <> blnPre Or gint病人来源 <> intFrom Then
        If SQLCondition.Default Then SQLCondition.int门诊标志 = gint病人来源
        ShowBills mstrFilter
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo <> "" Then
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                    "NO=" & .TextMatrix(.Row, GetColNum("单据号")), _
                    "开单人=" & .TextMatrix(.Row, GetColNum("医生")))
        End With
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewFilter_Click()
    frmPriceFilter.mstrPrivs = mstrPrivs
    '病人来源
    If gint病人来源 = 1 Then
        frmPriceFilter.opt病人(0).Value = True
    ElseIf gint病人来源 = 2 Then
        frmPriceFilter.opt病人(1).Value = True
    End If
    
    frmPriceFilter.Show 1, Me
    If gblnOK Then
        
        With frmPriceFilter
            mbln收费 = .chk收费.Value = 1
            mstrFilter = .mstrFilter
            
            SQLCondition.Default = False
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.DeptID = 0
            If .cbo科室.ListIndex <> -1 Then
                SQLCondition.DeptID = .cbo科室.ItemData(.cbo科室.ListIndex)
            End If
            SQLCondition.PatientID = .mlngPrePatient
            SQLCondition.PatientName = UCase(.txt姓名.Text)
            SQLCondition.Operator = zlStr.NeedName(.cbo操作员.Text)
            SQLCondition.ChargeKind = zlStr.NeedName(.cbo费别.Text)
            SQLCondition.str收费类别 = "," & .mstr收费类别 & ","
            SQLCondition.int门诊标志 = IIf(.opt病人(0).Value, 0, IIf(.opt病人(1).Value, 1, 2)) + 1
        End With
        
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub

Private Sub mshList_EnterCell()
    Dim strNo As String
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If mshList.Row = 0 Or strNo = "" Then Exit Sub
    stbThis.Panels(2) = "共 " & mrsList.RecordCount & " 张单据"
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    Call ShowDetail(strNo)
    Call ShowMoney(strNo)
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled And mnuEdit_Del.Visible Then Call mnuEdit_Del_Click
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            '始终从当前行开始
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Public Function CheckBillDel(ByVal strPrivs As String, strNo As String, int记录性质 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查划价单是否允许删除
    '入参:
    '出参:
    '返回:允许删除,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-13 12:35:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, bln药品 As Boolean
    
     If InStr(1, mstrPrivs, ";药品划价删除;") > 0 And InStr(1, mstrPrivs, ";诊疗划价删除;") > 0 Then
        CheckBillDel = True: Exit Function
     End If
     '45774
     If InStr(1, mstrPrivs, ";药品划价删除;") = 0 And InStr(1, mstrPrivs, ";诊疗划价删除;") = 0 Then
        MsgBox "你不具体删除划价单的权限,请与管理员联系!", vbInformation, gstrSysName
        Exit Function
      End If
     bln药品 = InStr(1, mstrPrivs, ";药品划价删除;") > 0
     
    On Error GoTo errH
    
    strSQL = "Select Nvl(Count(ID),0) as 数目" & _
        " From 门诊费用记录 " & _
        " Where NO=[1] And 记录性质=[2] And 记录状态 =0    " & _
        IIf(bln药品 = False, "  And 收费类别 IN ('5','6','7')", "And Not 收费类别  in ('5','6','7')")
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNo, int记录性质)
    
    If Val(Nvl(rsTmp!数目)) = 0 Then
        CheckBillDel = True
        Exit Function
    End If
    If bln药品 Then
        MsgBox "注意:" & vbCrLf & "    划价单中包含了诊疗项目,你不具备删除诊疗项目权限,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    Else
        MsgBox "注意:" & vbCrLf & "    划价单中包含了药品项目,你不具备删除药品项目权限,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub mnuEdit_Del_Click()
    Dim strNo As String, strSQL As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If strNo = "" Then
        MsgBox "当前没有单据可以删除！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '检查删除权限
    If Not BillOperCheck(3, mshList.TextMatrix(mshList.Row, GetColNum("划价人")), _
        CDate(mshList.TextMatrix(mshList.Row, GetColNum("划价时间"))), "删除", , , 1) Then Exit Sub
    
    If HaveExecute(1, strNo, 1) Then
        MsgBox "该单据中包含已执行的内容,不允许删除！", vbInformation, gstrSysName
        Exit Sub
    End If
    If CheckBillDel(mstrPrivs, strNo, 1) = False Then Exit Sub
    
    If MsgBox("确实要将单据""" & strNo & """删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSQL = "zl_门诊划价记录_DELETE('" & strNo & "')"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改单据清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_Price_Click()
    On Error Resume Next
    Err.Clear
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 1
    frmCharge.mbytInState = 0
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_View_Click()
    Dim strNo As String, strDate As Date
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If strNo = "" Then
        MsgBox "当前没有单据可以查阅！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    strDate = mshList.TextMatrix(mshList.Row, GetColNum(IIf(mbln收费, "收费时间", "划价时间")))
    '显示单据内容
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 1
    frmCharge.mbytInState = 1
    frmCharge.mstrInNO = strNo
    frmCharge.mstrTime = strDate
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    ShowBills mstrFilter
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mshMoney_GotFocus()
    Call SetActiveList(mshMoney)
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        picVsc.Top = picVsc.Top + Y
        picVsc.Height = picVsc.Height - Y
        mshMoney.Top = mshMoney.Top + Y
        mshMoney.Height = mshMoney.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub picHsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub picVsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshDetail.Width + X < 1000 Or mshMoney.Width - X < 1000 Then Exit Sub
        picVsc.Left = picVsc.Left + X
        mshDetail.Width = mshDetail.Width + X
        mshMoney.Left = mshMoney.Left + X
        mshMoney.Width = mshMoney.Width - X
        Me.Refresh
    End If
End Sub

Private Sub picVsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '定位
            mnuViewGo_Click
        Case "Filter" '过滤
            mnuViewFilter_Click
        Case "View"
            mnuEdit_View_Click
        Case "Price"
            mnuEdit_Price_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshList.Row
    
    '表头
    If Not mbln收费 Then
        objOut.Title.Text = "门诊划价单据清单"
    Else
        objOut.Title.Text = "门诊收费单据清单"
    End If
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    With frmPriceFilter
        objRow.Add "时间：" & Format(.dtpBegin.Value, .dtpBegin.CustomFormat) & " 至 " & Format(.dtpEnd.Value, .dtpEnd.CustomFormat)
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mshList.Redraw = False
    Set objOut.Body = mshList
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.Col = 0: mshList.ColSel = mshList.COLS - 1
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub SetMenu(blnUsed As Boolean)
'功能：根据有无记录设置菜单可用状态
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEdit_Adjust.Enabled = blnUsed And Not mbln收费
    mnuEdit_Modi.Enabled = blnUsed And Not mbln收费
    tbr.Buttons("Modi").Enabled = blnUsed And Not mbln收费
    
    mnuEdit_Del.Enabled = blnUsed And Not mbln收费
    mnuEdit_Print.Enabled = blnUsed And Not mbln收费
    mnuEdit_View.Enabled = blnUsed And Not mbln收费
    tbr.Buttons("Del").Enabled = blnUsed And Not mbln收费
    tbr.Buttons("View").Enabled = blnUsed And Not mbln收费
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strSQL As String
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
        
    If gintDelPrice > 0 And InStr(mstrPrivs, "删除") > 0 Then
        If MsgBox("系统准备清除划价后超过 " & gintDelPrice & " 天未收费未发药的划价单,处理吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call zlCommFun.ShowFlash("正在清除划价单,请稍候 ...", Me)
            DoEvents
            
            strSQL = "zl_门诊划价记录_Clear(" & gintDelPrice & ")"
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            On Error GoTo 0
            
            Call zlCommFun.StopFlash
        End If
    End If
    Call RestoreWinState(Me, App.ProductName)
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("刷新方式", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    '权限设置
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    If InStr(mstrPrivs, ";划价;") = 0 Then
        mnuEdit_Price.Visible = False
        mnuEdit_Print.Visible = False
        mnuEdit_Price_.Visible = False
        tbr.Buttons("Price").Visible = False
    End If
    
    If InStr(mstrPrivs, ";修改;") = 0 Then
        mnuEdit_Modi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    If InStr(mstrPrivs, ";调整;") = 0 Then
        mnuEdit_Adjust.Visible = False
    End If
    If InStr(mstrPrivs, ";修改;") = 0 And InStr(mstrPrivs, ";调整;") = 0 Then
        mnuEdit_Adjust_.Visible = False
    End If
    
    If InStr(mstrPrivs, ";删除;") = 0 Then
        mnuEdit_Del.Visible = False
        mnuEdit_Del_.Visible = False
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("Del_").Visible = False
    End If
    
    '缺省过滤条件(当天内)
    mbln收费 = False
    mstrFilter = " And 登记时间 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And 划价人||''=[5]"
    frmPriceFilter.mblnDateMoved = False
    SQLCondition.Default = True
    SQLCondition.int门诊标志 = gint病人来源
    
    Call SetHeader
    Call SetDetail
    Call SetMoney
    Call SetMenu(False)
    
    stbThis.Panels(2).Text = "请刷新清单或重新设置过滤条件"
    
    '初始化消息处理对象模块
    Call zlMsgModuleInit
        
    Exit Sub
errH:
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Call zlCommFun.ShowFlash("正在清除划价单,请稍候 ...", Me)
        DoEvents
        Resume
    End If
    Call SaveErrLog
    Unload Me
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long
    Dim sngVsc As Single, sngHsc As Single

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    sngVsc = mshDetail.Height / (mshDetail.Height + mshList.Height)
    sngHsc = mshMoney.Width / (mshMoney.Width + mshDetail.Width)
    
    If mblnMax Then
        sngVsc = 0.3: sngHsc = 0.2
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    mshList.Left = Me.ScaleLeft
    mshList.Top = Me.ScaleTop + cbrH
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Left = 0
    picHsc.Width = mshList.Width
    
    mshDetail.Left = 0
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - picHsc.Height - mshList.Height
    mshDetail.Width = (Me.ScaleWidth - picVsc.Width) * (1 - sngHsc)
    
    picVsc.Top = mshDetail.Top
    picVsc.Left = mshDetail.Left + mshDetail.Width
    picVsc.Height = mshDetail.Height
    
    mshMoney.Top = mshDetail.Top
    mshMoney.Left = picVsc.Left + picVsc.Width
    mshMoney.Height = mshDetail.Height
    mshMoney.Width = Me.ScaleWidth - picVsc.Width - mshDetail.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    mstrFilter = ""
    Unload frmPriceFilter
    Unload frmPriceGo
    
    Call SaveWinState(Me, App.ProductName)
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "刷新方式", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
            Exit For
        End If
    Next
    '拆卸消息对象
    Call zlMsgModuleUnload
End Sub

Private Sub mnuViewGo_Click()
    frmPriceGo.mstrPrivs = mstrPrivs
    frmPriceGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmPriceGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With frmPriceGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("单据号")) = .txtNO.Text
            End If
            If .cbo操作员.ListIndex > 0 Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("划价人")) = zlStr.NeedName(.cbo操作员.Text)
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, GetColNum("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.COLS - 1
                        
            Call mshList_EnterCell
            mlngGo = i + 1
            
            stbThis.Panels(2).Text = "找到一条记录"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '按ESC取消
        If mblnGo = False Then
            stbThis.Panels(2).Text = "用户取消定位操作"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "已定位到清单尾部"
    Screen.MousePointer = 0
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshList.COLS - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    
    If Button = 1 And mshList.MousePointer = 99 Then
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshList.TextMatrix(1, GetColNum("单据号")) = "" Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    If Not mbln收费 Then
        strHead = "单据号,1,850|开单科室,1,850|医生,1,800|姓名,1,800|性别,1,500|年龄,1,500|应收金额,7,850|实收金额,7,850|划价人,1,800|划价时间,1,1850"
    Else
        strHead = "单据号,1,850|开单科室,1,850|医生,1,800|姓名,1,800|性别,1,500|年龄,1,500|应收金额,7,850|实收金额,7,850|划价人,1,800|收费时间,1,1850|收费人,1,800"
    End If
    
    With mshList
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        i = GetColNum("医生")
        If InStr(mstrPrivs, "医生查询") = 0 Then
            mshList.ColWidth(i) = 0
        ElseIf mshList.ColWidth(i) = 0 Then
            mshList.ColWidth(i) = 800
        End If
        
        .RowHeight(0) = 320
        
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .COLS - 1
        Call mshList_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub ShowBills(Optional ByVal strIF As String, Optional blnSort As Boolean)
'功能:按条件读取单据列表(过滤功能)
'参数:strIF=以"AND"开始的条件串
'     blnSort=不重新读取数据,仅重新显示已排序的内容
    Dim i As Long, strSQL As String, strTable As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        Call zlCommFun.ShowFlash("正在读取单据列表,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        
        If frmPriceFilter.mblnDateMoved Then
            '这种方式与条件分开写的效果是一样的,会用到两个表的索引
            strTable = zlGetFullFieldsTable("门诊费用记录", 2, "", True, "")
        Else
            strTable = "门诊费用记录"
        End If
        strIF = " Where 记录性质=1 And " & IIf(mbln收费, " 记录状态 IN(1,3)", "记录状态=0") & _
                " And 划价人 is Not NULL And 操作员姓名 IS " & IIf(mbln收费, "NOT", "") & " NULL " & strIF
         
        Select Case SQLCondition.int门诊标志
        Case 1 '门诊
            strIF = strIF & " And  门诊标志 in (1,4)"
        Case 2 '住院
            strIF = strIF & " And  门诊标志 =2"
        Case Else   '所有
        End Select
         
        strSQL = "Select * From " & strTable & " A " & strIF
        
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            If gblnUserIsClinic Then '113577，限制开单科室
                strSQL = strSQL & " And Not Exists(" & _
                        "Select 1 From (Select NO From " & strTable & " C " & strIF & _
                        " And Not Exists(Select 1 From 部门人员 D Where C.开单部门ID+0=D.部门ID And D.人员ID=[9]) Group by NO) E" & _
                        " Where A.NO=E.NO)"
            Else '限制执行科室
                strSQL = strSQL & " And Not Exists(" & _
                        "Select 1 From (Select NO From " & strTable & " C " & strIF & _
                        " And Not Exists(Select 1 From 部门人员 D Where C.执行部门ID+0=D.部门ID And D.人员ID=[9]) Group by NO) E" & _
                        " Where A.NO=E.NO)"
            End If
        End If
        
        strSQL = _
            "Select A.NO as 单据号,B.名称 as 开单科室,A.开单人 as 医生,Ltrim(A.姓名) as 姓名,A.性别,A.年龄," & _
            " To_Char(Sum(A.应收金额),'9999999" & gstrDec & "') as 应收金额," & _
            " To_Char(Sum(A.实收金额),'9999999" & gstrDec & "') as 实收金额,A.划价人," & _
            " To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as " & IIf(mbln收费, "收费时间,A.操作员姓名 as 收费人", "划价时间") & _
            " From (" & strSQL & ") A,部门表 B" & _
            " Where A.开单部门ID = B.ID" & _
            " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
            " Group by A.NO,B.名称,A.开单人,A.姓名,A.性别,A.年龄,A.划价人," & IIf(mbln收费, "A.操作员姓名,", "") & "A.登记时间" & _
            " Order by " & IIf(mbln收费, "收费时间", "划价时间") & " Desc,单据号 Desc"
        
        With SQLCondition
            If .Default Then .Operator = UserInfo.姓名
            Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Operator, .PatientName, .ChargeKind, .DeptID, _
                              UserInfo.ID, .str收费类别, .PatientID)
        End With
    End If
    
    mshList.Clear:   mshList.Rows = 2
    mshDetail.Clear: mshDetail.Rows = 2
    mshMoney.Clear:  mshMoney.Rows = 2
    
    If Not mbln收费 Then
        mshList.ForeColor = vbBlack:    mshDetail.ForeColor = vbBlack:   mshMoney.ForeColor = vbBlack
    Else
        mshList.ForeColor = &H808080:   mshDetail.ForeColor = &H808080:  mshMoney.ForeColor = &H808080
    End If
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "当前设置没有过滤出任何单据"
        Call SetMenu(False)
    Else
        Set mshList.DataSource = mrsList
        stbThis.Panels(2) = "共 " & mrsList.RecordCount & " 张单据"
        Call SetMenu(True)
    End If
    Call SetHeader
    Call SetDetail
    Call SetMoney
    
    mnuEdit_Del.Enabled = Not mrsList.EOF And Not mbln收费
    tbr.Buttons("Del").Enabled = Not mrsList.EOF And Not mbln收费
    mnuEdit_Modi.Enabled = Not mrsList.EOF And Not mbln收费
    tbr.Buttons("Modi").Enabled = Not mrsList.EOF And Not mbln收费
    mnuEdit_Adjust.Enabled = Not mrsList.EOF And Not mbln收费
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowDetail(Optional strNo As String, Optional blnSort As Boolean)
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        '如果之前的筛选选项,划价清单中如果包括收费记录,则要检查
        If frmPriceFilter.mblnDateMoved Then
            mblnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNo, , "1")
        Else
            mblnNOMoved = False
        End If
        strSQL = _
        " Select C.名称 as 类别,Nvl(E.名称,B.名称) as 名称," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "B.规格," & _
                IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
        "       To_Char(Avg(Nvl(A.付数,1)*A.数次)" & _
                IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ",'9999990.00000') as 数次, " & _
        "       A.费别,To_Char(Sum(A.标准单价)" & _
                IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
        "       To_Char(Sum(A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
        "       To_Char(Sum(A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
        "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型" & _
        " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,收费项目别名 E,药品规格 X" & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
        " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
        "       And A.记录性质=1 and A.记录状态 IN(0,1,3) And A.NO=[1]" & _
        "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
        " Group by Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称)," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称,", "") & " B.规格,A.计算单位,A.费别," & _
        "       D.名称,Nvl(A.费用类型,B.费用类型),X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1)" & _
        " Order by Nvl(A.价格父号,A.序号)"
        Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    End If
    
    mshDetail.Clear
    mshDetail.Rows = 2

    If Not mrsDetail.EOF Then Set mshDetail.DataSource = mrsDetail
    Call SetDetail
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "类别,1,750|名称,1,1800" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,2000", "") & "|规格,1,1000|单位,4,500|数次,7,850|费别,1,750|单价,7,850|应收金额,7,850|实收金额,7,850|执行科室,1,850|类型,1,850"
    
    With mshDetail
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        .RowHeight(0) = 320
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        'Call mshDetail_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub ShowMoney(Optional strNo As String, Optional blnSort As Boolean)
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        strSQL = _
            "Select " & IIf(gint分类合计 = 0, "A.收据费目", "B.名称") & " as 项目," & _
            " To_Char(Sum(A.实收金额),'9999999" & gstrDec & "') as 金额 " & _
            " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收入项目 B " & _
            " Where A.收入项目ID=B.ID AND A.记录性质=1" & _
            " And A.记录状态 IN(0,1,3) And A.NO=[1]" & _
            " Group by " & IIf(gint分类合计 = 0, "A.收据费目", "B.名称")
        Set mrsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    End If
    
    mshMoney.Clear
    mshMoney.Rows = 2
    
    If Not mrsMoney.EOF Then Set mshMoney.DataSource = mrsMoney
    Call SetMoney
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetMoney()
    Dim strHead As String
    Dim i As Long
    
    strHead = "项目,1,850|金额,7,850"
    With mshMoney
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshMoney, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        'Call mshMoney_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshDetail.MouseCol
    
    If Button = 1 And mshDetail.MousePointer = 99 Then
        If mshDetail.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshDetail.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsDetail Is Nothing Then Exit Sub
        
        Set mshDetail.DataSource = Nothing

        mrsDetail.Sort = mshDetail.TextMatrix(0, lngCol) & IIf(mshDetail.ColData(lngCol) = 0, "", " DESC")
        mshDetail.ColData(lngCol) = (mshDetail.ColData(lngCol) + 1) Mod 2
        
        Call ShowDetail(, True)
    End If
End Sub

Private Sub mshMoney_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshMoney.MouseRow = 0 Then
        mshMoney.MousePointer = 99
    Else
        mshMoney.MousePointer = 0
    End If
End Sub

Private Sub mshMoney_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshMoney.MouseCol
    
    If Button = 1 And mshMoney.MousePointer = 99 Then
        If mshMoney.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshMoney.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsMoney Is Nothing Then Exit Sub
        
        Set mshMoney.DataSource = Nothing

        mrsMoney.Sort = mshMoney.TextMatrix(0, lngCol) & IIf(mshMoney.ColData(lngCol) = 0, "", " DESC")
        mshMoney.ColData(lngCol) = (mshMoney.ColData(lngCol) + 1) Mod 2
        
        Call ShowMoney(, True)
    End If
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &H8000000D
        mshDetail.BackColorSel = &H8000000C
        mshMoney.BackColorSel = &H8000000C
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &H8000000C
        mshDetail.BackColorSel = &H8000000D
        mshMoney.BackColorSel = &H8000000C
    ElseIf obj Is mshMoney Then
        mshList.BackColorSel = &H8000000C
        mshDetail.BackColorSel = &H8000000C
        mshMoney.BackColorSel = &H8000000D
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


Private Function zlMsgModuleInit() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化消息模块
    '入参:lngModule -模块号
    '     strPivs-权限串
    '出参:objMsgModule-返回消息对象
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    zlMsgModuleInit = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlMsgModuleUnload() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:拆卸消息模块
    '入参:objMsgModule-消息对象
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    
    If mobjMsgModule Is Nothing Then Exit Function
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    zlMsgModuleUnload = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


