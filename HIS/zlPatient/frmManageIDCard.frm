VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmManageIDCard 
   AutoRedraw      =   -1  'True
   Caption         =   "就诊发卡管理"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8850
   Icon            =   "frmManageIDCard.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   4710
      Left            =   45
      TabIndex        =   3
      Top             =   780
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   8308
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
      MouseIcon       =   "frmManageIDCard.frx":0442
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8850
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   8730
         _ExtentX        =   15399
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
            NumButtons      =   12
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
               Caption         =   "发卡"
               Key             =   "IDCard"
               Description     =   "发卡"
               Object.ToolTipText     =   "进入发卡窗口"
               Object.Tag             =   "发卡"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退卡"
               Key             =   "Del"
               Description     =   "退卡"
               Object.ToolTipText     =   "对当前选中记录退卡"
               Object.Tag             =   "退卡"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅当前单据的内容"
               Object.Tag             =   "查阅"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "设置条件重新读取列表"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位在当前列表中满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   0
      Top             =   5490
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageIDCard.frx":075C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10530
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":0FF0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":120A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":1424
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":163E
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":1858
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":1FD2
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":21EC
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":2406
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":2620
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":283A
            Key             =   "Report"
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":2A54
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":2C6E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":2E88
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":30A2
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":32BC
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":3A36
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":3C50
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":3E6A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":4084
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":429E
            Key             =   "Report"
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
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileWorkReport 
         Caption         =   "打印缴款书(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_IDCard 
         Caption         =   "发卡(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "退卡(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "查阅(&V)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPass 
         Caption         =   "修改密码(&P)"
         Shortcut        =   {F4}
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
         Begin VB.Menu mnuViewTool_1 
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
      Begin VB.Menu mnuView_5 
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
      Begin VB.Menu mnuViewReFlash 
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
Attribute VB_Name = "frmManageIDCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''Option Explicit '要求变量声明
''''Private mrsList As ADODB.Recordset  '单据列表
''''Private mstrFilter As String
''''Private mblnCancel As Boolean
''''Private mblnGo As Boolean, mlngGo As Long
''''Private mlngCurRow As Long, mlngTopRow As Long
''''Private mstrPrivs As String
''''Private mlngModul As Long
''''Private mblnNOMoved As Boolean '显明细时记录当前选择的单据是否在在线数据表中,以其它操作时无需再判断
''''Private mcllFilterA As Collection
'''''by lesfeng 2010-1-11 性能优化
''''Private Sub InitFilter()
''''    '-----------------------------------------------------------------------------------------------------------
''''    '功能:初始化过滤条件
''''    '入参:
''''    '出参:
''''    '返回:
''''    '编制:lesfeng
''''    '日期:2010-01-11 16:10:40
''''    '-----------------------------------------------------------------------------------------------------------
''''    Set mcllFilterA = New Collection
''''    mcllFilterA.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "登记时间"
''''    mcllFilterA.Add Array("", ""), "单据号"
''''    mcllFilterA.Add Array("", ""), "票据号"
''''    mcllFilterA.Add "", "住院号"
''''    mcllFilterA.Add "", "姓名"
''''    mcllFilterA.Add "", "记录状态"
''''    mcllFilterA.Add "", "附加标志"
''''    mcllFilterA.Add "", "收款人"
''''    mstrFilter = ""
''''End Sub
''''
''''Private Sub Form_Activate()
''''    Call InitLocPar(mlngModul)
''''End Sub
''''
''''Private Sub mnuEditPass_Click()
''''    frmModiPass.Show 1, Me
''''End Sub
''''
''''Private Sub mnuFileLocalSet_Click()
''''    Call frmLocalSet.zlSetPara(Me, mstrPrivs, mlngModul)
''''    If glng磁卡ID > 0 Then
''''        If Not ExistBill(glng磁卡ID, 5) Then
''''            zldatabase.SetPara "共用就诊卡批次", 0, glngSys, mlngModul
''''            glng磁卡ID = 0
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub mnuFileWorkReport_Click()
''''    Call frmWorkTime.ShowMe(Me, 5)
''''End Sub
''''
''''Private Sub mnuReportItem_Click(Index As Integer)
''''    Dim strNO As String, strTmp As String
''''
''''    strNO = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
''''    If strNO <> "" Then
''''        With mshList
''''            If glngSys Like "8??" Then
''''                strTmp = "客户ID"
''''            Else
''''                strTmp = "病人ID"
''''            End If
''''            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
''''                    "NO=" & strNO, "就诊卡号=" & .TextMatrix(.Row, GetColNum("卡号")), _
''''                    strTmp & "=" & .TextMatrix(.Row, GetColNum(strTmp)))
''''        End With
''''    Else
''''        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
''''    End If
''''End Sub
''''
''''Private Sub mnuViewFilter_Click()
''''    frmIDCardFilter.Show 1, Me
''''    If gblnOK Then
''''        mstrFilter = frmIDCardFilter.mstrFilter
''''        'by lesfeng 2010-03-08 性能优化
''''        Set mcllFilterA = frmIDCardFilter.mcllFilter
''''        mblnCancel = (frmIDCardFilter.chkCancel.Value = Checked)
''''        mnuViewReFlash_Click
''''    End If
''''End Sub
''''
''''Private Sub mshList_DblClick()
''''    If mshList.MouseRow = 0 Then Exit Sub
''''    If mnuEdit_View.Enabled Then mnuEdit_View_Click
''''End Sub
''''
''''Private Sub mshList_EnterCell()
''''    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, 0) = "" Then Exit Sub
''''    mlngGo = mshList.Row
''''    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
''''
''''    If frmIDCardFilter.mblnDateMoved Then
''''        mblnNOMoved = zldatabase.NOMoved("住院费用记录", mshList.TextMatrix(mshList.Row, 0), , "5", Me.Caption)
''''    Else
''''        mblnNOMoved = False
''''    End If
''''
''''End Sub
''''
''''Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
''''    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled And mnuEdit_Del.Visible Then Call mnuEdit_Del_Click
''''End Sub
''''
''''Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    If Button = 2 Then PopupMenu mnuEdit, 2
''''End Sub
''''
''''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''''    Select Case KeyCode
''''        Case vbKeyF3
''''            '始终从当前行开始
''''            If mnuViewGo.Enabled Then Call SeekBill(False)
''''        Case vbKeyReturn
''''            If mnuEdit_View.Enabled Then mnuEdit_View_Click
''''        Case vbKeyEscape
''''            mblnGo = False
''''    End Select
''''End Sub
''''
''''Private Sub mnuEdit_Del_Click()
''''    If mshList.TextMatrix(mshList.Row, 0) = "" Then
''''        MsgBox "当前没有记录可以退卡！", vbExclamation, gstrSysName
''''        Exit Sub
''''    End If
''''
''''    '单据权限
''''    If Not BillOperCheck(8, mshList.TextMatrix(mshList.Row, GetColNum("发卡人")), _
''''        CDate(mshList.TextMatrix(mshList.Row, GetColNum("发卡时间"))), "退卡") Then Exit Sub
''''
''''    On Error Resume Next
''''    Err.Clear
''''
''''    '是否已转入后备数据表中
''''    If mblnNOMoved Then
''''        If Not ReturnMovedExes(mshList.TextMatrix(mshList.Row, 0), 5, Me.Caption) Then Exit Sub
''''        mblnNOMoved = False  '此时已转入在线数据表
''''    End If
''''
''''    frmIDCard.mbytInState = 2
''''    frmIDCard.mstrInNO = mshList.TextMatrix(mshList.Row, 0)
''''    frmIDCard.Show 1, Me
''''    If gblnOK Then Call mnuViewReFlash_Click
''''End Sub
''''
''''Private Sub mnuHelpTitle_Click()
''''ShowHelp App.ProductName, Me.hwnd, Me.Name
''''End Sub
''''
''''Private Sub mnuEdit_IDCard_Click()
''''    On Error Resume Next
''''    Err.Clear
''''
''''    frmIDCard.mbytInState = 0
''''    frmIDCard.Show 1, Me
''''    If gblnOK Then mnuViewReFlash_Click
''''End Sub
''''
''''Private Sub mnuEdit_View_Click()
''''    If mshList.TextMatrix(mshList.Row, 0) = "" Then
''''        MsgBox "当前没有记录可以查阅！", vbExclamation, gstrSysName
''''        Exit Sub
''''    End If
''''
''''    On Error Resume Next
''''    Err.Clear
''''    '显示单据内容
''''    frmIDCard.mbytInState = 1
''''    If mblnCancel Then frmIDCard.mblnViewCancel = True
''''    frmIDCard.mstrInNO = mshList.TextMatrix(mshList.Row, 0)
''''    frmIDCard.mblnNOMoved = mblnNOMoved
''''    frmIDCard.Show 1, Me
''''End Sub
''''
''''Private Sub mnuFile_Quit_Click()
''''    Unload Me
''''End Sub
''''
''''Private Sub mnuHelpAbout_Click()
''''    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
''''End Sub
''''
''''Private Sub mnuViewReFlash_Click()
''''    ShowBills mstrFilter
''''End Sub
''''
''''Private Sub mnuViewStatus_Click()
''''    mnuViewStatus.Checked = Not mnuViewStatus.Checked
''''    stbThis.Visible = Not stbThis.Visible
''''    Form_Resize
''''End Sub
''''
''''Private Sub mnuViewToolButton_Click()
''''    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
''''    cbr.Visible = Not cbr.Visible
''''    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
''''    Form_Resize
''''End Sub
''''
''''Private Sub mnuViewToolText_Click()
''''    Dim i As Integer
''''    mnuViewToolText.Checked = Not mnuViewToolText.Checked
''''    For i = 1 To tbr.Buttons.Count
''''        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
''''    Next
''''    cbr.Bands(1).MinHeight = tbr.ButtonHeight
''''    Form_Resize
''''End Sub
''''
''''Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
''''    Select Case Button.Key
''''        Case "Quit"
''''            mnuFile_Quit_Click
''''        Case "Go" '定位
''''            mnuViewGo_Click
''''        Case "Filter" '过滤
''''            mnuViewFilter_Click
''''        Case "View"
''''            mnuEdit_View_Click
''''        Case "IDCard"
''''            mnuEdit_IDCard_Click
''''        Case "Del"
''''            mnuEdit_Del_Click
''''        Case "Print"
''''            mnuFile_Print_Click
''''        Case "Preview"
''''            mnuFile_PreView_Click
''''        Case "Help"
''''            mnuHelpTitle_Click
''''    End Select
''''End Sub
''''
''''Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    If Button = 2 Then PopupMenu mnuViewTool, 2
''''End Sub
''''
''''Private Sub mnuFile_Excel_Click()
''''    Call OutputList(3)
''''End Sub
''''
''''Private Sub mnuFile_PreView_Click()
''''    Call OutputList(2)
''''End Sub
''''
''''Private Sub mnuFile_Print_Click()
''''    Call OutputList(1)
''''End Sub
''''
''''Private Sub mnuFile_PrintSet_Click()
''''    Call zlPrintSet
''''End Sub
''''
''''Private Sub OutputList(bytStyle As Byte)
'''''功能：输入出列表
'''''参数：bytStyle=1-打印,2-预览,3-输出到Excel
''''    Dim objOut As New zlPrint1Grd
''''    Dim objRow As New zlTabAppRow
''''    Dim bytR As Byte, intRow As Integer
''''
''''    intRow = mshList.Row
''''
''''    '表头
''''    If glngSys Like "8??" Then
''''        objOut.Title.Text = "会员卡发放单清单"
''''    Else
''''        objOut.Title.Text = "就诊卡发放单清单"
''''    End If
''''    objOut.Title.Font.Name = "楷体_GB2312"
''''    objOut.Title.Font.Size = 18
''''    objOut.Title.Font.Bold = True
''''
''''    '表项
''''    With frmIDCardFilter
''''        If IsNull(.dtpEnd.Value) Then
''''            objRow.Add "时间：" & Format(.dtpBegin.Value, "yyyy-MM-dd")
''''        Else
''''            objRow.Add "时间：" & Format(.dtpBegin.Value, "yyyy-MM-dd HH:MM") & " 至 " & Format(.dtpEnd.Value, "yyyy-MM-dd HH:MM")
''''        End If
''''        objRow.Add "性质：" & IIf(.chkCancel.Value = 1, "退卡记录", "发卡记录")
''''        objOut.UnderAppRows.Add objRow
''''    End With
''''
''''    Set objRow = New zlTabAppRow
''''    objRow.Add "打印人：" & UserInfo.姓名
''''    objRow.Add "打印日期：" & Format(zldatabase.Currentdate(), "yyyy年MM月dd日")
''''    objOut.BelowAppRows.Add objRow
''''
''''    '表体
''''    mshList.Redraw = False
''''    Set objOut.Body = mshList
''''
''''    '输出
''''    If bytStyle = 1 Then
''''        bytR = zlPrintAsk(objOut)
''''        Me.Refresh
''''        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
''''    Else
''''        zlPrintOrView1Grd objOut, bytStyle
''''    End If
''''
''''    mshList.Row = intRow
''''    mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
''''    mshList.Redraw = True
''''End Sub
''''
''''Private Sub mnuHelpWebHome_Click()
''''    zlHomePage hwnd
''''End Sub
''''
''''Private Sub mnuHelpWebMail_Click()
''''    zlMailTo hwnd
''''End Sub
''''
''''Private Sub SetHeader()
''''    Dim i As Integer
''''    With mshList
''''        .Redraw = False
''''        .Cols = 12
''''        .TextMatrix(0, 0) = "单据号"
''''        If mblnCancel Then
''''            .TextMatrix(0, 1) = "退卡时间"
''''        Else
''''            .TextMatrix(0, 1) = "发卡时间"
''''        End If
''''        .TextMatrix(0, 2) = "卡号"
''''        .TextMatrix(0, 3) = "类型"
''''        If glngSys Like "8??" Then
''''            .TextMatrix(0, 4) = "客户ID"
''''        Else
''''            .TextMatrix(0, 4) = "病人ID"
''''        End If
''''        .TextMatrix(0, 5) = "标识号"
''''        .TextMatrix(0, 6) = "姓名"
''''        .TextMatrix(0, 7) = "性别"
''''        .TextMatrix(0, 8) = "年龄"
''''        .TextMatrix(0, 9) = "金额"
''''        .TextMatrix(0, 10) = "记帐"
''''        If mblnCancel Then
''''            .TextMatrix(0, 11) = "退卡人"
''''        Else
''''            .TextMatrix(0, 11) = "发卡人"
''''        End If
''''
''''        .ColAlignment(0) = 4
''''        .ColAlignment(1) = 4
''''        .ColAlignment(2) = 1
''''        .ColAlignment(3) = 4
''''        .ColAlignment(4) = 1
''''        .ColAlignment(5) = 1
''''        .ColAlignment(6) = 1
''''        .ColAlignment(7) = 4
''''        .ColAlignment(8) = 4
''''        .ColAlignment(9) = 7
''''        .ColAlignment(10) = 4
''''        .ColAlignment(11) = 1
''''
''''        If Not Visible Then
''''            .ColWidth(0) = 850
''''            .ColWidth(1) = 1000
''''            .ColWidth(2) = 850
''''            .ColWidth(3) = 500
''''            .ColWidth(4) = 750
''''            If glngSys Like "8??" Then
''''                .ColWidth(5) = 0
''''            Else
''''                .ColWidth(5) = 750
''''            End If
''''            .ColWidth(6) = 800
''''            .ColWidth(7) = 500
''''            .ColWidth(8) = 500
''''            .ColWidth(9) = 850
''''            .ColWidth(10) = 500
''''            .ColWidth(11) = 800
''''        End If
''''
''''        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
''''
''''        .RowHeight(0) = 320
''''        For i = 0 To .Cols - 1
''''            .ColAlignmentFixed(i) = 4
''''        Next
''''        '恢复上次行
''''        If mlngCurRow = 0 Then mlngCurRow = 1
''''        If mlngTopRow = 0 Then mlngTopRow = 1
''''        If mlngCurRow <= .Rows - 1 Then
''''            .Row = mlngCurRow
''''        Else
''''            .Row = .Rows - 1
''''        End If
''''        If mlngTopRow <= .Rows - 1 Then
''''            .TopRow = mlngTopRow
''''        Else
''''            .TopRow = .Row
''''        End If
''''
''''         .Col = 0: .ColSel = .Cols - 1
''''        Call mshList_EnterCell
''''
''''        .Redraw = True
''''    End With
''''End Sub
''''
''''Private Sub ShowBills(Optional ByVal strIF As String, Optional blnSort As Boolean)
'''''功能:按条件读取单据列表(过滤功能)
'''''参数:strIF=以"AND"开始的条件串
''''    Dim strCard As String, i As Long
''''
''''    On Error GoTo errH
''''
''''    If Not blnSort Then
''''        Call zlCommFun.ShowFlash("正在读取单据列表,请稍候 ...", Me)
''''        DoEvents
''''        Me.Refresh
''''
''''        strIF = " Where 记录性质=5 " & strIF
''''        'by lesfeng 2010-03-08 性能优化
''''        If frmIDCardFilter.mblnDateMoved Then
''''            strIF = "" & _
''''            " Select NO,登记时间,实际票号,附加标志,病人id,标识号,姓名,性别,年龄,实收金额,记帐费用,操作员姓名 " & _
''''            " From 住院费用记录 " & strIF & _
''''            " UNION ALL " & _
''''            " Select NO,登记时间,实际票号,附加标志,病人id,标识号,姓名,性别,年龄,实收金额,记帐费用,操作员姓名 " & _
''''            " From H住院费用记录 " & strIF
''''        Else
''''            strIF = "Select NO,登记时间,实际票号,附加标志,病人id,标识号,姓名,性别,年龄,实收金额,记帐费用,操作员姓名 From 住院费用记录 " & strIF
''''        End If
''''
''''        strCard = "Decode(" & IIf(gblnShowCard, 1, 0) & ",1,A.实际票号,LPAD('*',Length(A.实际票号),'*')) as 卡号,"
''''        gstrSQL = _
''''        " Select A.NO as 单据号,To_Char(A.登记时间,'YYYY-MM-DD') as " & IIf(mblnCancel, "退卡", "发卡") & "时间," & strCard & _
''''        "           Decode(A.附加标志,1,'补卡',2,'换卡','发卡') as 类型,A.病人ID,A.标识号 as 住院号,A.姓名,A.性别,A.年龄," & _
''''        "           To_Char(" & IIf(mblnCancel, " - ", "") & "Sum(A.实收金额),'99990.00') as 金额," & _
''''        "           Decode(Nvl(A.记帐费用,0),0,NULL,'√') as 记帐," & _
''''        "           A.操作员姓名 as " & IIf(mblnCancel, "退卡人", "发卡人") & " " & _
''''        " From (" & strIF & ") A " & _
''''        " Group by A.NO,To_Char(A.登记时间,'YYYY-MM-DD'),A.实际票号,Decode(A.附加标志,1,'补卡',2,'换卡','发卡')," & _
''''        "           A.病人ID,A.标识号,A.姓名,A.性别,A.年龄,Decode(Nvl(A.记帐费用,0),0,NULL,'√'),A.操作员姓名" & _
''''        " Order by " & IIf(mblnCancel, "退卡", "发卡") & "时间 Desc,单据号 Desc"
''''        Set mrsList = New ADODB.Recordset
''''
''''        Set mrsList = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(mcllFilterA("登记时间")(0)), CDate(mcllFilterA("登记时间")(1)), _
''''        CStr(Val(mcllFilterA("单据号")(0))), CStr(Val(mcllFilterA("单据号")(1))), _
''''        CStr(Val(mcllFilterA("票据号")(0))), CStr(Val(mcllFilterA("票据号")(1))), CLng(Val(mcllFilterA("住院号"))), _
''''        CStr(mcllFilterA("姓名")), CLng(Val(mcllFilterA("记录状态"))), CLng(Val(mcllFilterA("附加标志"))), CStr(mcllFilterA("收款人")))
''''
''''    End If
''''
''''    mshList.Clear
''''    mshList.Rows = 2
''''
''''    mshList.ForeColor = IIf(mblnCancel, &HC0, ForeColor)
''''
''''    If mrsList.EOF Then
''''        Call SetHeader
''''        stbThis.Panels(2).Text = "当前设置没有过滤出任何单据"
''''        Call SetMenu(False)
''''    Else
''''        Set mshList.DataSource = mrsList
''''        Call SetHeader
''''        stbThis.Panels(2) = "共 " & mrsList.RecordCount & " 张单据"
''''        Call SetMenu(True)
''''    End If
''''
''''    mnuEdit_Del.Enabled = Not mblnCancel And Not mrsList.EOF
''''    tbr.Buttons("Del").Enabled = Not mblnCancel And Not mrsList.EOF
''''
''''    If Not blnSort Then Call zlCommFun.StopFlash
''''
''''    Me.Refresh
''''    Exit Sub
''''errH:
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''End Sub
''''
''''Private Sub SetMenu(blnUsed As Boolean)
'''''功能：根据有无记录设置菜单可用状态
''''    mnuFile_Print.Enabled = blnUsed
''''    mnuFile_Preview.Enabled = blnUsed
''''    mnuFile_Excel.Enabled = blnUsed
''''    tbr.Buttons("Print").Enabled = blnUsed
''''    tbr.Buttons("Preview").Enabled = blnUsed
''''
''''    mnuEdit_Del.Enabled = blnUsed
''''    mnuEdit_View.Enabled = blnUsed
''''    tbr.Buttons("Del").Enabled = blnUsed
''''    tbr.Buttons("View").Enabled = blnUsed
''''
''''    mnuViewGo.Enabled = blnUsed
''''    tbr.Buttons("Go").Enabled = blnUsed
''''End Sub
''''
''''Private Sub Form_Load()
''''    Dim curDate As Date
''''    'by lesfeng 2010-03-08 性能优化
''''    Call InitFilter
''''
''''    mstrPrivs = gstrPrivs
''''    mlngModul = glngModul
''''    Call zldatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
''''
''''    If glngSys Like "8??" Then Caption = "会员卡发放管理"
''''
''''    Call RestoreWinState(Me, App.ProductName)
''''
''''    If glng磁卡ID > 0 Then
''''        If Not ExistBill(glng磁卡ID, 5) Then
''''            zldatabase.SetPara "共用就诊卡批次", 0, glngSys, mlngModul
''''            glng磁卡ID = 0
''''        End If
''''    End If
''''
''''    '权限设置(改为不可见)
''''    If InStr(mstrPrivs, "发卡事务") = 0 Then
''''        mnuEdit_IDCard.Visible = False
''''        mnuEdit_Del.Visible = False
''''        tbr.Buttons("IDCard").Visible = False
''''        tbr.Buttons("Del").Visible = False
''''    End If
''''    If InStr(mstrPrivs, "修改密码") = 0 Then
''''        mnuEdit_1.Visible = False
''''        mnuEditPass.Visible = False
''''    End If
''''
''''    '缺省过滤条件
''''    curDate = zldatabase.Currentdate
''''    'by lesfeng 2010-03-08 性能优化
''''    mstrFilter = ""
''''    mstrFilter = mstrFilter & " And (登记时间  Between [1] And [2]) "
''''    mstrFilter = mstrFilter & " And 记录状态=[9]"
''''    mstrFilter = mstrFilter & " And 操作员姓名=[11]"
''''
''''    mcllFilterA.Remove "登记时间"
''''    mcllFilterA.Add Array(Format(DateAdd("d", -7, curDate), "yyyy-mm-dd") & " 00:00:00", Format(curDate, "yyyy-mm-dd") & " 23:59:59"), "登记时间"
''''    mcllFilterA.Remove "记录状态"
''''    mcllFilterA.Add "1", "记录状态"
''''    mcllFilterA.Remove "收款人"
''''    mcllFilterA.Add Trim(UserInfo.姓名), "收款人"
''''
''''    mblnCancel = False
''''
''''    Call SetHeader
''''    Call SetMenu(False)
''''
''''    stbThis.Panels(2).Text = "请刷新清单或重新设置过滤条件"
''''End Sub
''''
''''Private Sub Form_Resize()
''''    Dim cbrH As Long '工具条占用高度
''''    Dim staH As Long '状态栏占用高度
''''
''''    On Error Resume Next
''''
''''    If WindowState = 1 Then Exit Sub
''''
''''    mshList.MousePointer = 0
''''
''''    '靠齐控件宽度和高度
''''    cbrH = IIf(cbr.Visible, cbr.Height, 0)
''''    staH = IIf(stbThis.Visible, stbThis.Height, 0)
''''    With mshList
''''        .Left = Me.ScaleLeft
''''        .Top = Me.ScaleTop + cbrH
''''        .Width = Me.ScaleWidth
''''        .Height = Me.ScaleHeight - cbrH - staH
''''    End With
''''End Sub
''''
''''Private Sub Form_Unload(Cancel As Integer)
''''    mstrFilter = ""
''''    Unload frmIDCardFilter
''''    Unload frmIDCardFind
''''    Call SaveWinState(Me, App.ProductName)
''''End Sub
''''
''''Private Sub mnuViewGo_Click()
''''    If Not mblnCancel Then
''''        frmIDCardFind.lbl操作员.Caption = "发卡人"
''''    Else
''''        frmIDCardFind.lbl操作员.Caption = "退卡人"
''''    End If
''''    frmIDCardFind.Show 1, Me
''''    If gblnOK Then Call SeekBill(frmIDCardFind.optHead)
''''End Sub
''''
''''Private Sub SeekBill(blnHead As Boolean)
''''    Dim i As Long
''''    Dim blnFill As Boolean
''''
''''    Screen.MousePointer = 11
''''    mblnGo = True
''''    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
''''    Me.Refresh
''''
''''    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
''''        DoEvents
''''
''''        '比较条件
''''        blnFill = True
''''        With frmIDCardFind
''''            If .txtNO.Text <> "" Then
''''                blnFill = blnFill And mshList.TextMatrix(i, 0) = .txtNO.Text
''''            End If
''''            If .txtCard.Text <> "" Then
''''                blnFill = blnFill And mshList.TextMatrix(i, 2) = .txtCard.Text
''''            End If
''''            If .cbo操作员.ListIndex > 0 Then
''''                blnFill = blnFill And mshList.TextMatrix(i, 11) = NeedName(.cbo操作员.Text)
''''            End If
''''            If .txt姓名.Text <> "" Then
''''                blnFill = blnFill And UCase(mshList.TextMatrix(i, 6)) Like "*" & UCase(.txt姓名.Text) & "*"
''''            End If
''''            If IsNumeric(.txt住院号.Text) Then
''''                blnFill = blnFill And Val(mshList.TextMatrix(i, 5)) = Val(.txt住院号.Text)
''''            End If
''''        End With
''''
''''        '满足则退出
''''        If blnFill Then
''''            mlngGo = i + 1
''''            mshList.Row = i: mshList.TopRow = i
''''            mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
''''            stbThis.Panels(2).Text = "找到一条记录"
''''            Screen.MousePointer = 0: Exit Sub
''''        End If
''''
''''        '按ESC取消
''''        If mblnGo = False Then
''''            stbThis.Panels(2).Text = "用户取消定位操作"
''''            Screen.MousePointer = 0: Exit Sub
''''        End If
''''    Next
''''    mlngGo = 1
''''    stbThis.Panels(2).Text = "已定位到清单尾部"
''''    Screen.MousePointer = 0
''''End Sub
''''
''''Private Function GetColNum(strHead As String) As Integer
''''    Dim i As Integer
''''    For i = 0 To mshList.Cols - 1
''''        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
''''    Next
''''End Function
''''
''''Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    If mshList.MouseRow = 0 Then
''''        mshList.MousePointer = 99
''''    Else
''''        mshList.MousePointer = 0
''''    End If
''''End Sub
''''
''''Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    Dim lngCol As Long
''''
''''    lngCol = mshList.MouseCol
''''
''''    If Button = 1 And mshList.MousePointer = 99 Then
''''        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
''''        If mshList.TextMatrix(1, GetColNum("单据号")) = "" Then Exit Sub
''''
''''        Set mshList.DataSource = Nothing
''''        If mshList.TextMatrix(0, lngCol) = "客户ID" Then
''''            mrsList.Sort = "病人ID" & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
''''        Else
''''            mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
''''        End If
''''        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
''''
''''        Call ShowBills(, True)
''''    End If
''''End Sub
''''
''''Private Sub mnuHelpWebForum_Click()
''''    '-----------------------------------------------------------------------------
''''    '功能:链接到中联论坛
''''    '修改人:刘兴宏
''''    '修改日期:2006-12-11
''''    '-----------------------------------------------------------------------------
''''    Call zlWebForum(Me.hwnd)
''''End Sub
''''
