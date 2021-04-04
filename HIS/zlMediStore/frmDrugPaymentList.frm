VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDrugPaymentList 
   Caption         =   "药品付款事务"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   Icon            =   "frmDrugPaymentList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   5
      Top             =   2760
      Width           =   4815
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询范围:1999年8月12日至1999年9月12日"
         Height          =   180
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   3690
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   11775
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbTool"
      MinWidth1       =   6000
      MinHeight1      =   720
      Width1          =   6210
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "PrintView"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "Add"
               Description     =   "增加"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Payment"
                     Text            =   "付款单"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Imprest"
                     Text            =   "预付款单"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Description     =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Description     =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "审核"
               Key             =   "Verify"
               Description     =   "审核"
               Object.ToolTipText     =   "审核"
               Object.Tag             =   "审核"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "冲销"
               Key             =   "Strike"
               Description     =   "冲销"
               Object.ToolTipText     =   "冲销"
               Object.Tag             =   "冲销"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Search"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "Refresh"
               Description     =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmDrugPaymentList.frx":014A
         Begin VB.Timer LimitTime 
            Enabled         =   0   'False
            Interval        =   8000
            Left            =   6660
            Top             =   180
         End
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   4620
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugPaymentList.frx":0464
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11642
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   0
      Top             =   600
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
            Picture         =   "frmDrugPaymentList.frx":0CF8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":0F18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1138
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1354
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1574
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1794
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":19B0
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1BCC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1DE6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1F40
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":215C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   600
      Top             =   600
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
            Picture         =   "frmDrugPaymentList.frx":237C
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":259C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":27BC
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":29D8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":2BF8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":2E18
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":3034
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":3250
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":346A
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":35C4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":37E4
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshAddition 
      Height          =   2595
      Left            =   6810
      TabIndex        =   7
      Top             =   3000
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   4577
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorFixed  =   -2147483644
      GridColorFixed  =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image imgVsc_s 
      Height          =   2325
      Left            =   6480
      MousePointer    =   9  'Size W E
      Top             =   3210
      Width           =   120
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBillPrint 
         Caption         =   "单据打印(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileBillPreview 
         Caption         =   "单据预览(&L)"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "新增(&A)"
         Begin VB.Menu mnuEditAddPayment 
            Caption         =   "付款单(&P)"
         End
         Begin VB.Menu mnuEditAddImprest 
            Caption         =   "预付款单(&I)"
         End
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "审核(&C)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "冲销(&K)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "查看单据(&W)"
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
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
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
            Caption         =   "发送反馈(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmDrugPaymentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Private mstrFind As String
Private mlastRow As Long                '上次电击的行
Private mintPreCol As Integer           '前一次单据头的排序列
Private mintsort As Integer             '前一次单据头的排序
Private mintPreDetailCol As Integer     '前一次单据体的排序列
Private mintDetailsort As Integer       '前一次单据体的排序
Private mblnStartup As Boolean


Private Sub GetList(ByVal StrFind As String)
    Dim rsList As New Recordset
    Dim strUserPart As String
    
    On Error GoTo errHandle
    mlastRow = 0
    Call zlcommfun.ShowFlash("正在搜索药品付款记录,请稍候 ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    
    mshList.Redraw = False
    
    gstrSQL = "SELECT a.no,b.id, b.名称,nvl(预付款,0) as 预付款 , ltrim(to_char(SUM (a.金额),'9999999999999990.00')) AS 金额, a.填制人 AS 申请人," _
            & "TO_CHAR (min(a.填制日期), 'yyyy-MM-dd HH24:MI:SS') AS 申请日期, a.审核人," _
            & "TO_CHAR (min(a.审核日期), 'yyyy-MM-dd HH24:MI:SS') AS 审核日期, a.记录状态, a.摘要 " _
            & "FROM 药品付款记录 a, 药品供应商 b " _
            & "Where a.单位id = b.id " _
            & StrFind _
            & " GROUP BY a.no,b.id,b.名称,nvl(预付款,0),a.填制人,a.审核人,a.记录状态, a.摘要 " _
            & " ORDER BY a.no desc "
    
    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rsList = zldatabase.OpenSQLRecord(gstrSQL, "GetList")
    Call SQLTest
   
    Set mshList.Recordset = rsList
    With mshList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = True
            
            .TopRow = 1
            .rows = .rows - 99
            
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    SetListColWidth
    
    mshlist_EnterCell    '列出单据体
    
    SetStrikeColor
    mshList.Redraw = True
    Call zlcommfun.StopFlash
    Screen.MousePointer = vbDefault
    staThis.Panels(2).Text = "当前共有" & rsList.RecordCount & "张单据"
    rsList.Close
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim IntCol As Integer
    
    With mshList
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
            intStatus = .TextMatrix(intRow, .Cols - 2)
            If intStatus = 3 Then
                .Row = intRow
                For IntCol = 0 To .Cols - 1
                    .Col = IntCol
                    .CellForeColor = &H80000001
                Next
            End If
            If intStatus = 2 Then
                .Row = intRow
                For IntCol = 0 To .Cols - 1
                    .Col = IntCol
                    .CellForeColor = &HFF
                Next
            End If
        Next
    End With
                
End Sub

'表头列宽初始
Private Sub SetListColWidth()
    Dim IntCol As Integer
    
    With mshList
        .ColAlignment(4) = flexAlignRightCenter
        
        If mblnStartup = False Then
            For IntCol = 1 To .Cols - 1
                .ColWidth(IntCol) = 1500
            Next
        End If
        .ColWidth(1) = 0
        .ColWidth(1) = 0
        .ColWidth(3) = 0
        .ColWidth(.Cols - 2) = 0
    End With
End Sub


Private Sub SetDetailColWidth()
    Dim IntCol As Integer
    
    With mshDetail
        .ColAlignment(1) = flexAlignLeftCenter      '入库单号
        .ColAlignment(2) = flexAlignLeftCenter      '发票号
        .ColAlignment(3) = flexAlignRightCenter     '发票金额
        .ColAlignment(8) = flexAlignCenterCenter    '单位
        .ColAlignment(9) = flexAlignRightCenter     '数量
        .ColAlignment(10) = flexAlignRightCenter    '采购价
        .ColAlignment(11) = flexAlignRightCenter    '采购金额
           
        If mblnStartup = False Then
            For IntCol = 0 To .Cols - 1
                .ColWidth(IntCol) = 1000
            Next
            .ColWidth(4) = 2000
        End If
            
    End With
End Sub


'根据权限设置不同的显示项目
Private Sub SetVisable()
    '外购入库所有权限：参数设置、基本、所有库房、登记、修改、删除、验收、冲销

    If InStr(1, gstrprivs, "参数设置") = 0 Then
         mnuFileParameter.Visible = False
         mnuFileLine3.Visible = False                '相应的分割线
    End If
     
    If InStr(1, gstrprivs, "登记") = 0 Then
        mnuEditAdd.Visible = False
        tlbTool.Buttons("Add").Visible = False
    End If
    
    If InStr(1, gstrprivs, "修改") = 0 Then
        mnuEditModify.Visible = False
        tlbTool.Buttons("Modify").Visible = False
    End If
    
    If InStr(1, gstrprivs, "删除") = 0 Then
        mnuEditDel.Visible = False
        tlbTool.Buttons("Delete").Visible = False
         '对没有所有编辑权限时，把菜单和工具栏上的相应的分割线屏蔽。
        If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
            mnuEditLine1.Visible = False
            tlbTool.Buttons("EditSeparate").Visible = False
        End If
    End If
    
    If InStr(1, gstrprivs, "审核") = 0 Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    
    If InStr(1, gstrprivs, "冲销") = 0 Then
        mnuEditStrike.Visible = False
        tlbTool.Buttons("Strike").Visible = False
        
        If mnuEditVerify.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    If InStr(1, gstrprivs, "单据打印") = 0 Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
End Sub

Private Sub Form_Load()
    '恢复设置
    Dim strStart As String
    Dim strEnd As String
    Dim StrFind As String
    
    mblnStartup = False
    SetVisable  '根据权限设置不同的显示项目
    strStart = Format(DateAdd("m", -1, zldatabase.Currentdate), "yyyy-MM-dd")
    strEnd = Format(zldatabase.Currentdate, "yyyy-MM-dd")
    StrFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between To_Date('" & strStart & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & strEnd & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    mstrFind = StrFind
    
    lblRange.Caption = "查询范围:" & Format(DateAdd("m", -1, zldatabase.Currentdate), "yyyy年MM月dd日") & "至" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")
    
    GetList (mstrFind)  '列出单据头
    mblnStartup = True
    RestoreWinState Me, App.ProductName, Me.Caption
End Sub

Private Sub Form_Resize()
    '窗体位置设置
    
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If
    
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
        .Height = 720
    End With
    
    With picSeparate_s
        .Height = 300
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        .Width = imgVsc_s.Left   '- 10
    End With
    
    With mshAddition
        .Top = mshDetail.Top
        .Left = imgVsc_s.Left + imgVsc_s.Width  '+ 10
        .Width = mshList.Width - .Left
        .Height = mshDetail.Height
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, Me.Caption
End Sub

Private Sub imgVsc_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With imgVsc_s
            If .Left + x < 2000 Then Exit Sub
            If .Left + x > ScaleWidth - 2000 Then Exit Sub
            .Left = .Left + x
        End With
        
        Me.mshDetail.Width = Me.mshDetail.Width + x
        Me.mshAddition.Left = Me.mshAddition.Left + x
        Me.mshAddition.Width = Me.mshAddition.Width - x
    End If
End Sub


Private Sub mnuEditAddImprest_Click()
    Dim StrNo As String
    Dim BlnSuccess As Boolean
    
    StrNo = ""
    frmDrugImprestCard.ShowCard Me, StrNo, 1, , BlnSuccess
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditAddPayment_Click()
    Dim BlnSuccess  As Boolean
    
    FrmDrugPaymentCard.ShowCard Me, "", 1, 1, BlnSuccess
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditVerify_Click()
    '验收
    
    Dim StrNo As String
    Dim BlnSuccess As Boolean
    
    With mshList
        StrNo = .TextMatrix(.Row, 0)
        If .TextMatrix(.Row, 3) = 1 Then
            frmDrugImprestCard.ShowCard Me, StrNo, 3, .TextMatrix(.Row, .Cols - 2), BlnSuccess
        Else
            FrmDrugPaymentCard.ShowCard Me, StrNo, 3, .TextMatrix(.Row, .Cols - 2), BlnSuccess
        End If
        
    End With
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    '删除
    Dim strBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With mshList
        
        If .TextMatrix(.Row, 3) = 1 Then
            strTitle = "药品预付款单"
        Else
            strTitle = "药品付款单"
        End If
        On Error GoTo errHandle
        intRow = .Row
        strBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("你确实要删除单据号为“" & strBillNo & "”的" & strTitle & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_药品付款管理_delete('" & strBillNo & "')"
            
            If gstrSQL = "" Then Exit Sub
            Call SQLTest(App.Title, Me.Caption, gstrSQL)
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
            Call SQLTest
            intRecord = intRecord - 1
            mlastRow = 0
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                With mshDetail
                    .rows = 1
                    .rows = 2
                    .FixedRows = 1
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                End With
                SetEnable
                
            End If
                
            '.RowHeight(intRow) = 0
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
            mshlist_EnterCell
        End If
    End With
    staThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub mnuEditDisplay_Click()
    '查看单据
    
    Dim StrNo As String
    With mshList
        StrNo = .TextMatrix(.Row, 0)
        If .TextMatrix(.Row, 3) = 1 Then
            frmDrugImprestCard.ShowCard Me, StrNo, 4, .TextMatrix(.Row, .Cols - 2)
        Else
            FrmDrugPaymentCard.ShowCard Me, StrNo, 4, .TextMatrix(.Row, .Cols - 2)
        End If
    End With
    
End Sub

Private Sub mnuEditStrike_Click()
    '冲销
    With mshList
        If MsgBox("你确实要冲销单据号为“" & .TextMatrix(.Row, 0) & "”的单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If StrikeSave = True Then
                mnuViewRefresh_Click
            End If
        End If
    End With
End Sub

Private Function StrikeSave() As Boolean
    
    StrikeSave = False
    With mshList
        On Error GoTo errHandle
        gstrSQL = "zl_药品付款管理_STRIKE('" & .TextMatrix(.Row, 0) & "'," _
            & .TextMatrix(.Row, 4) & "," & .TextMatrix(.Row, 1) & ",'" & UserInfo.用户姓名 & "')"
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        Call SQLTest
        
    End With
    StrikeSave = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

    'MsgBox "存盘失败！", vbInformation, gstrSysName
End Function

Private Sub mnuEditModify_Click()
    '修改
    Dim StrNo As String
    Dim BlnSuccess As Boolean
    
    BlnSuccess = False
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        StrNo = .TextMatrix(.Row, 0)
        If .TextMatrix(.Row, 3) = 0 Then
            FrmDrugPaymentCard.ShowCard Me, StrNo, 2, mshList.TextMatrix(mshList.Row, mshList.Cols - 2), BlnSuccess
        Else
            frmDrugImprestCard.ShowCard Me, StrNo, 2, mshList.TextMatrix(mshList.Row, .Cols - 2), BlnSuccess
        End If
        If BlnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If IIf(IsNull(.TextMatrix(.Row, 3)), 0, .TextMatrix(.Row, 3)) = 1 Then
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_2", "zl8_bill_1320_2"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), 1
        Else
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_1", "zl8_bill_1320_1"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), 1
        End If
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If IIf(IsNull(.TextMatrix(.Row, 3)), 0, .TextMatrix(.Row, 3)) = 1 Then
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_2", "zl8_bill_1320_2"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), 2
        Else
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_1", "zl8_bill_1320_1"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), 2
        End If
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '输出到Excel
    mshList.Redraw = False
    subPrint 3
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnufileexit_Click()
    '退出
    Unload Me
    
End Sub

Private Sub mnuFileParameter_Click()
    '参数设置
    frm参数设置.设置参数 Me, Me.Caption
End Sub

Private Sub mnuFilePreView_Click()
    '打印预览
    mshList.Redraw = False
    subPrint 2
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    mshList.Redraw = False
    subPrint 1
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnuFilePrintSet_Click()
    '打印设置
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    '关于
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '帮助主题
'    ReportMan gcnOracle, Me
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewRefresh_Click()
    '刷新
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '查找
    
    Dim strStart As Date
    Dim strEnd As Date
    Dim strVerifyStart As Date
    Dim strVerifyEnd As Date
    Dim StrFind As String
    
    
    StrFind = FrmDrugPaymentSearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd)
    
    If StrFind <> "" Then
        mstrFind = StrFind
        GetList mstrFind
        If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:填制日期 " & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日") & "  审核日期 " & Format(strVerifyStart, "yyyy年MM月dd日") & "至" & Format(strVerifyEnd, "yyyy年MM月dd日")
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:填制日期 " & Format(strStart, "yyyy年MM月dd日") & "至" & Format(strEnd, "yyyy年MM月dd日")
        ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "查询范围:审核日期 " & Format(strVerifyStart, "yyyy年MM月dd日") & "至" & Format(strVerifyEnd, "yyyy年MM月dd日")
        End If
             
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '工具条索引
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            '取消所有的文本标签显示
            For intCount = 1 To .count
                .Item(intCount).Caption = ""
            Next
        Else
            '让所有的文本标签显示。说明：Tag中放的文本标签
            For intCount = 1 To .count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub

Private Sub mshDetail_Click()
    With mshDetail
         If .Row < 1 Or .TextMatrix(.Row, 0) = "" Then Exit Sub
         If .MouseRow = 0 Then
            DetailSort          '列排序
            Exit Sub
         End If
    End With
End Sub

Private Sub mshList_Click()
    With mshList
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            ListSort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshlist_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If mshList.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub mshlist_EnterCell()
    Dim rsDetail As New Recordset
    Dim strUnitName As String                       '单位名称:如门诊单位，住院单位等
    Dim str包装系数 As String
    Dim intUnit As Integer
    
    On Error GoTo errHandle
    If mlastRow = mshList.Row Then Exit Sub
    mlastRow = mshList.Row
    SetEnable
    
    intUnit = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Caption, "药品单位", "0")
    If glngSys \ 100 = 8 Then
        strUnitName = Choose(intUnit + 1, "c.药库单位", "c.售价单位")
        str包装系数 = Choose(intUnit + 1, "c.药库包装", "1")
    Else
        strUnitName = Choose(intUnit + 1, "c.药库单位", "c.门诊单位", "c.住院单位", "c.售价单位")
        str包装系数 = Choose(intUnit + 1, "c.药库包装", "c.门诊包装", "c.住院包装", "1")
    End If
    
    If mshList.Row >= 1 And LTrim(mshList.TextMatrix(mshList.Row, 0)) <> "" And IIf(mshList.TextMatrix(mshList.Row, 3) = "", 0, mshList.TextMatrix(mshList.Row, 3)) <> 1 And mshList.TextMatrix(mshList.Row, mshList.Cols - 2) <> "2" Then
        mshList.Col = 0
        mshList.ColSel = mshList.Cols - 1
        
        mshDetail.Redraw = False
        gstrSQL = "SELECT distinct b.审核日期 AS 入库日期, b.no, a.发票号, to_char(a.发票金额,'99999999999999990.00') as 发票金额 ," _
            & "'[' || c.编码 || ']' || DECODE (d.名称, NULL, e.通用名称, d.名称) AS 药品信息," _
            & "c.规格, b.产地, b.批号," & strUnitName & " AS 单位, to_char(b.实际数量/" & str包装系数 & ",'99999999999999990.00000')  AS 数量, to_char(b.成本价*" & str包装系数 & ",'99999999999999990.00000') as 采购价 ,to_char(b.成本金额,'99999999999999990.00') as 采购金额 " _
            & " FROM 药品应付记录 a, " _
               & "(SELECT * From 药品收发记录 WHERE 单据 = 1 and 供药单位id=" & mshList.TextMatrix(mshList.Row, 1) & " ) b," _
                & "药品目录 c," _
                & "药品别名 d," _
                & "药品信息 e " _
           & " Where a.收发id = b.id " _
           & "   AND b.药品id = c.药品id " _
           & "   AND c.药品id = d.药品id (+) " _
           & "   AND c.药名id = e.药名id " _
           & "   AND a.付款序号 = (SELECT DISTINCT 付款序号 FROM 药品付款记录 WHERE no='" & mshList.TextMatrix(mshList.Row, 0) & "' and  付款序号 is not null and 记录状态<>2 ) " _
          & "  ORDER BY a.发票号, b.no "
        
        With rsDetail
            Call SQLTest(App.Title, Me.Caption, gstrSQL)
            Set rsDetail = zldatabase.OpenSQLRecord(gstrSQL, "cmd产地_Click")
            Call SQLTest
            
            Set mshDetail.Recordset = rsDetail
            .Close
        End With
        With mshDetail
            If .rows = 1 Then
                .rows = .rows + 100
                .Row = 1
                .Redraw = True
                
                .TopRow = 1
                .rows = .rows - 99
            End If
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
        End With
        
        mshDetail.Redraw = True
    ElseIf LTrim(mshList.TextMatrix(mshList.Row, 0)) = "" Or mshList.TextMatrix(mshList.Row, 3) = "1" Or mshList.TextMatrix(mshList.Row, mshList.Cols - 2) = "2" Then
        With mshDetail
            .Cols = 12
            .rows = 2
            .Clear
    
            .TextMatrix(0, 0) = "入库日期"
            .TextMatrix(0, 1) = "入库单号"
            .TextMatrix(0, 2) = "发票号"
            .TextMatrix(0, 3) = "发票金额"
            .TextMatrix(0, 4) = "药品信息"
            .TextMatrix(0, 5) = "规格"
            .TextMatrix(0, 6) = "产地"
            .TextMatrix(0, 7) = "批号"
            .TextMatrix(0, 8) = "单位"
            .TextMatrix(0, 9) = "数量"
            .TextMatrix(0, 10) = "采购价"
            .TextMatrix(0, 11) = "采购金额"
            
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
        End With
    End If
    SetDetailColWidth
    
    If mshList.TextMatrix(mshList.Row, 3) = "1" Or mshList.TextMatrix(mshList.Row, mshList.Cols - 2) = "2" Then
        gstrSQL = "SELECT DECODE (预付款, 1, '是', 0, '否') AS 预付款," _
            & " TO_CHAR (金额, '99999999999999990.00') AS 金额, 结算方式, 结算号码, " _
            & " DECODE (预付款, 1, no, '') AS 相关预付款号 " _
            & " From 药品付款记录 " _
            & "WHERE no = '" & mshList.TextMatrix(mshList.Row, 0) _
            & "'  and 记录状态=" & mshList.TextMatrix(mshList.Row, mshList.Cols - 2) _
            & " ORDER BY 预付款,序号 "
    Else
        gstrSQL = "SELECT DECODE (预付款, 1, '是', 0, '否') AS 预付款," _
            & " TO_CHAR (金额, '99999999999999990.00') AS 金额, 结算方式, 结算号码, " _
            & " DECODE (预付款, 1, no, '') AS 相关预付款号 " _
            & " From 药品付款记录 " _
            & "WHERE 付款序号 = (SELECT DISTINCT 付款序号 FROM 药品付款记录 WHERE no='" & mshList.TextMatrix(mshList.Row, 0) & "' and  付款序号 is not null and 记录状态<>2 ) " _
            & "  and 记录状态=" & IIf(mshList.TextMatrix(mshList.Row, mshList.Cols - 2) = "", 1, mshList.TextMatrix(mshList.Row, mshList.Cols - 2)) _
            & "  and 预付款=0 "
        gstrSQL = gstrSQL & _
             " union " _
             & " SELECT DECODE (预付款, 1, '是', 0, '否') AS 预付款," _
            & " TO_CHAR (金额, '99999999999999990.00') AS 金额, 结算方式, 结算号码, " _
            & " DECODE (预付款, 1, no, '') AS 相关预付款号 " _
            & " From 药品付款记录 " _
            & "WHERE 付款序号 = (SELECT DISTINCT 付款序号 FROM 药品付款记录 WHERE no='" & mshList.TextMatrix(mshList.Row, 0) & "' and  付款序号 is not null and 记录状态<>2 ) " _
            & "  and 预付款=1 "
        gstrSQL = gstrSQL & _
            " union " _
            & " SELECT DECODE (预付款, 1, '是', 0, '否') AS 预付款," _
            & " TO_CHAR (金额, '99999999999999990.00') AS 金额, 结算方式, 结算号码, " _
            & " DECODE (预付款, 1, no, '') AS 相关预付款号 " _
            & " From 药品付款记录 " _
            & "WHERE 付款序号 = (SELECT DISTINCT 付款序号 FROM 药品付款记录 WHERE no='" & mshList.TextMatrix(mshList.Row, 0) & "' and  付款序号 is not null and 记录状态<>2 ) " _
            & "  and (记录状态=2) " _
            & "  and nvl(预付款,0)=0 "
            
        gstrSQL = " select * from (" & gstrSQL & ") order by 预付款 "
            
        '& "  and (记录状态=1 or 记录状态=3) "
            
    End If
    
    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rsDetail = zldatabase.OpenSQLRecord(gstrSQL, "cmd产地_Click")
    Call SQLTest
    
    Set mshAddition.Recordset = rsDetail
    rsDetail.Close
    With mshAddition
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = True
            
            .TopRow = 1
            .rows = .rows - 99
        End If
        
        If mblnStartup = False Then
            .ColWidth(0) = 800
            .ColWidth(1) = 1000
            .ColWidth(2) = 800
            .ColWidth(3) = 1000
            .ColWidth(4) = 1000
        End If
        
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        .ColAlignment(1) = flexAlignRightCenter
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshlist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditModify_Click
    End If
        
End Sub

Private Sub mshlist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    PopupMenu mnuEdit, 2
    
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '分割条设置
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Height = picSeparate_s.Top - .Top
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        mshAddition.Top = .Top
        mshAddition.Height = .Height
    End With
    
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Add"
            mnuEditAddPayment_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Strike"
            mnuEditStrike_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnufileexit_Click
        
    End Select
    
End Sub

'设置菜单和工具按钮的可用属性
Private Sub SetEnable()
    With mshList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
        
            
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditVerify.Visible = True Then
                mnuEditVerify.Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
            End If
            
            If mnuEditStrike.Visible = True Then
                mnuEditStrike.Enabled = False
                tlbTool.Buttons("Strike").Enabled = False
            End If
             
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
         Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            If .TextMatrix(.Row, .Cols - 4) = "" Then    '未审核单
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = True
                    tlbTool.Buttons("Verify").Enabled = True
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '审核单
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = True
                    tlbTool.Buttons("Strike").Enabled = True
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            Else   '2,3 冲销单
                If .TextMatrix(.Row, .Cols - 2) = 3 Then
                    .ToolTipText = "冲销单据的原单据"
                ElseIf .TextMatrix(.Row, .Cols - 2) = 2 Then
                    .ToolTipText = "冲销单据"
                End If
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
            End If
        End If
        
    End With
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    
    Set objPrint = New zlPrint1Grd
    
        
    objPrint.Title.Text = "药品付款单清册表"
    Set objPrint.Body = mshList
    
        
    'objRow.Add "打印人:" & UserInfo.用户姓名
    'objRow.Add "打印日期:" & Format(ZlDatabase.Currentdate, "yyyy-MM-dd")
    'objPrint.UnderAppRows.Add objRow
    
    objRow.Add "打印人：" & UserInfo.用户姓名
    objRow.Add "打印时间：" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Payment"
            mnuEditAddPayment_Click
        Case "Imprest"
            mnuEditAddImprest_Click
    End Select
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'对单据头列排序
Private Sub ListSort()
    Dim IntCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshList
        If .rows > 1 Then
            .Redraw = False
            IntCol = .MouseCol
            .Col = IntCol
            .ColSel = IntCol
            intTemp = .TextMatrix(.Row, 0)
                    
            If IntCol = 4 Then
                If IntCol = mintPreCol And mintsort = flexSortNumericDescending Then
                    .Sort = flexSortNumericAscending
                    mintsort = flexSortNumericAscending
                Else
                    .Sort = flexSortNumericDescending
                    mintsort = flexSortNumericDescending
                End If
            Else
                If IntCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                    .Sort = flexSortStringNoCaseAscending
                    mintsort = flexSortStringNoCaseAscending
                Else
                    .Sort = flexSortStringNoCaseDescending
                    mintsort = flexSortStringNoCaseDescending
                End If
            End If
                
            mintPreCol = IntCol
            .Row = FindRow(mshList, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

'对单据头列排序
Private Sub DetailSort()
    Dim IntCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshDetail
        If .rows > 1 Then
            .Redraw = False
            IntCol = .MouseCol
            .Col = IntCol
            .ColSel = IntCol
            intTemp = .TextMatrix(.Row, 1)
            
            Select Case IntCol
                Case 3, 9, 10, 11
                    If IntCol = mintPreDetailCol And mintDetailsort = flexSortNumericDescending Then
                       .Sort = flexSortNumericAscending
                       mintDetailsort = flexSortNumericAscending
                    Else
                       .Sort = flexSortNumericDescending
                       mintDetailsort = flexSortNumericDescending
                    End If
                    
                Case Else
                    If IntCol = mintPreDetailCol And mintDetailsort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       mintDetailsort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       mintDetailsort = flexSortStringNoCaseDescending
                    End If
            End Select
                
            mintPreDetailCol = IntCol
            .Row = FindRow(mshDetail, intTemp, 1)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

'寻找与某一列相等的行
Public Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal IntCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, IntCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, IntCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

