VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmKSSManagement 
   Caption         =   "抗菌药物授权"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKSSManagement.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptPati 
      Height          =   2400
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   7440
      _Version        =   589884
      _ExtentX        =   13123
      _ExtentY        =   4233
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.Frame fraType 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   600
      Width           =   2535
      Begin VB.OptionButton optOccasion 
         Caption         =   "住院"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   16
         Top             =   -10
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optOccasion 
         Caption         =   "门诊"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   15
         Top             =   -10
         Width           =   735
      End
      Begin VB.Label lblType 
         Caption         =   "使用场合"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.CheckBox chkIsShowAll 
      BackColor       =   &H8000000B&
      Caption         =   "只显示三级医师职务的人员"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   983
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox picGrant 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   2160
      ScaleHeight     =   2775
      ScaleWidth      =   7455
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton cmdMove 
         Height          =   495
         Index           =   0
         Left            =   3720
         Picture         =   "frmKSSManagement.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton cmdMove 
         Height          =   495
         Index           =   1
         Left            =   3240
         Picture         =   "frmKSSManagement.frx":711C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   375
      End
      Begin MSComctlLib.TreeView tvwSelect 
         Height          =   2535
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4471
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         ImageList       =   "ils16"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvwGrant 
         Height          =   2535
         Left            =   4200
         TabIndex        =   9
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4471
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         ImageList       =   "ils16"
         Appearance      =   1
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTemp 
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _cx             =   661
      _cy             =   661
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":79E6
            Key             =   "Male"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":E248
            Key             =   "feMale"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":14AAA
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":15044
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   7800
      MaxLength       =   30
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin XtremeSuiteControls.TaskPanel tplFunc 
      Height          =   5325
      Left            =   0
      TabIndex        =   0
      Top             =   1260
      Width           =   2100
      _Version        =   589884
      _ExtentX        =   3704
      _ExtentY        =   9393
      _StockProps     =   64
      Behaviour       =   1
      ItemLayout      =   2
      HotTrackStyle   =   3
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8010
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   635
      SimpleText      =   $"frmKSSManagement.frx":155DE
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmKSSManagement.frx":15625
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17092
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
   Begin MSComctlLib.ImageList img32 
      Left            =   840
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":15EB9
            Key             =   "Male"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":1C71B
            Key             =   "feMale"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":22F7D
            Key             =   "No"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":23857
            Key             =   "Yes"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":24131
            Key             =   "Pepoles"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1560
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":2A993
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":2AFDF
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":2B2FB
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":2B617
            Key             =   "Dept_No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":2B937
            Key             =   "Cert"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":2BA91
            Key             =   "Male"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKSSManagement.frx":322F3
            Key             =   "feMale"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMsgHave 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "已选人员列表"
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   983
      Visible         =   0   'False
      Width           =   1095
   End
   Begin XtremeSuiteControls.ShortcutCaption stcLabel 
      Height          =   300
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   7500
      _Version        =   589884
      _ExtentX        =   13229
      _ExtentY        =   529
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeSuiteControls.ShortcutCaption stcItem 
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   945
      Width           =   2100
      _Version        =   589884
      _ExtentX        =   3704
      _ExtentY        =   529
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeCommandBars.ImageManager ImgC 
      Left            =   2280
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmKSSManagement.frx":38B55
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   2880
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKSSManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'CommandBar辅助热键
Private Const FSHIFT = 4
Private Const FCONTROL = 8
Private Const FALT = 16

Private mstrPrivs As String
Private mlngModul As Long
Private mblnFirstLoad As Boolean '判断是否是第一次加载
Private mlngLastRunModule As Long
Private mobjShowCancel As CommandBarControl
Private mpicOld As IPictureDisp
Private mobjBar As CommandBar
Private mobjMenu As CommandBarPopup
Private mblnIsChange As Boolean
Private mblnIsFindFinish As Boolean
Private mrsFind As Recordset
Private mlngFindNum As Long
Private mlngCodeType As Long         '0-拼音,1-五笔
Private mstrUserDept As String
Private mintNoPrivs As Integer    '不可用的类别数，判断是否只有一个权限，如果是，则禁用调整权限。

'拖放
Private mMouseX As Long
Private mMouseY As Long
Private mblnIsUp As Boolean
Private mblnIsCheck As Boolean
Private mblnIsHaveCancle As Boolean

Private Enum mEnumPanel
    PanelItem_NotLimit = 1
    PanelItem_Limit = 2
    PanelItem_Special = 3
End Enum

Private Enum mEnumVsgPati
    col_选择 = 0
    COL_姓名 = 1
    COL_编号 = 2
    col_性别 = 3
    COL_专业职务 = 4
    COL_所属部门 = 5
    COL_授权人 = 6
    COL_授权日期 = 7
    COL_人员ID = 8
    col_记录状态 = 9
End Enum

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim lng路径ID As Long
    Dim objPopup As CommandBarPopup
    
    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    
    Case conMenu_Kss_Cancellation  '取消授权
        Call Cancellation
    Case conMenu_Kss_Adjustment    '调整权限
        Call Adjustment
    Case conMenu_Kss_Grant         '授权
        Call GrantTo
        mblnIsFindFinish = False
    Case conMenu_Edit_Untread     '取消
        Call CancleGrant
        mblnIsFindFinish = False
    Case conMenu_Edit_Save
        Call SaveGrant
    Case conMenu_View_Find '查找
        If Me.ActiveControl Is txtFind Then
            txtFind.SetFocus '有时需要定位一下
            If txtFind.Text <> "" Then
                'Call FuncFindPath
            End If
        Else
            txtFind.SetFocus
        End If
    Case conMenu_View_FindNext '查找下一个
        If txtFind.Text = "" Then
            txtFind.SetFocus
        Else
            Call txtFind_KeyPress(vbKeyReturn)
        End If
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
        cbsMain_Resize
    Case conMenu_Kss_ShowCancel  '显示取消授权的人员
        Control.Checked = Not Control.Checked
        If mlngLastRunModule <> 0 Then
            Call RunByModule(mlngLastRunModule & "")
        End If
    Case conMenu_View_Refresh '刷新
        If picGrant.Visible Then
            Call LoadPrss
        Else
            If mlngLastRunModule <> 0 Then Call RunByModule(mlngLastRunModule)
        End If
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '退出
        Unload Me
    End Select
End Sub

Private Sub CancleGrant()
'功能：取消保存
    Dim objItem As TaskPanelGroupItem

    If mblnIsChange Then
        If MsgBox("您已经进行了改动，确定要取消刚才的操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    picGrant.Visible = False
    rptPati.Visible = True
    mobjMenu.Enabled = True
    stcLabel.Caption = "人员列表"
    chkIsShowAll.Visible = False
    mblnIsChange = False
    lblMsgHave.Visible = False
    Call tvwGrant.Nodes.Clear
    For Each objItem In tplFunc.Groups(1).Items
        If objItem.Tag & "" <> "不可用" Then objItem.Enabled = True
    Next
End Sub

Private Sub SaveGrant()
'功能：保存授权信息
    Dim strSql As String
    Dim strPatiIDs As String
    Dim Node As Node
    Dim curDate As Date
    Dim lngNum As Long
    Dim objItem As TaskPanelGroupItem
    
    If tvwGrant.Nodes.Count = 0 Then Call CancleGrant: Exit Sub
    
    For Each Node In tvwGrant.Nodes
        If Not Node.Parent Is Nothing Then
            strPatiIDs = strPatiIDs & IIf(strPatiIDs = "", "", ",") & Mid(Node.Key, 2)
            lngNum = lngNum + 1
        End If
    Next
    
    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    
    strSql = "Zl_人员抗菌药物权限_Update('" & strPatiIDs & "'," & mlngLastRunModule & ",'" & UserInfo.姓名 & "',to_date('" & _
                curDate & "','YYYY-MM-DD HH24:MI:SS')," & IIf(optOccasion(0).Value, 1, 2) & ")"
        
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Call RunByModule(mlngLastRunModule & "")
    
    picGrant.Visible = False
    rptPati.Visible = True
    mobjMenu.Enabled = True
    stcLabel.Caption = "人员列表"
    chkIsShowAll.Visible = False
    lblMsgHave.Visible = False
    mblnIsChange = False
    Call tvwGrant.Nodes.Clear
    For Each objItem In tplFunc.Groups(1).Items
        If objItem.Tag & "" <> "不可用" Then objItem.Enabled = True
    Next
    MsgBox "本次授权成功，一共授权 " & lngNum & " 名人员。", vbInformation, Me.Caption
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GrantTo()
'功能：授权
    Dim objItem As TaskPanelGroupItem

    picGrant.Visible = True
    rptPati.Visible = False
    mobjMenu.Enabled = False
    For Each objItem In tplFunc.Groups(1).Items
        If Not objItem.Selected Then objItem.Enabled = False
    Next
    mblnIsChange = False
    stcLabel.Caption = "待选人员列表"
    chkIsShowAll.Visible = True
    lblMsgHave.Visible = True
    If mlngLastRunModule = 1 Then
        chkIsShowAll.Caption = "只显示住院医师"
    ElseIf mlngLastRunModule = 2 Then
        chkIsShowAll.Caption = "只显示主治医师"
    ElseIf mlngLastRunModule = 3 Then
        chkIsShowAll.Caption = "只显示(副)主任医师"
    End If
    Call cbsMain_Resize
    
    '加载数据
    Call LoadPrss
End Sub

Private Sub LoadPrss()
'功能：加载可授权的用户
    Dim strIsShowAll As String
    Dim rsTmp As Recordset
    Dim strTemp As String
    Dim strSql As String
    Dim i As Long, y As Long
    Dim blnIsInGrant As Boolean
    Dim Node As Node

    If chkIsShowAll.Value Then
        If mlngLastRunModule = 1 Then
            strTemp = "'医师'"
        ElseIf mlngLastRunModule = 2 Then
            strTemp = "'主治医师'"
        ElseIf mlngLastRunModule = 3 Then
            strTemp = "'主任医师', '副主任医师'"
        End If
        strIsShowAll = " And (a.专业技术职务 In (" & strTemp & ") And c.人员id Is Null)"
    End If

    strSql = "Select *" & vbNewLine & _
            "From (With Test As (Select Distinct a.Id, 'C' || a.Id As 人员id, a.姓名, a.性别, a.编号, 'D' || b.部门id As 树形部门id, b.部门id," & vbNewLine & _
                                "                Upper(zlSpellCode(a.姓名)) As 拼音简码, Upper(zlWbCode(a.姓名)) As 五笔简码" & vbNewLine & _
                                "From 人员表 A, 部门人员 B," & vbNewLine & _
                                "     (Select Distinct c.人员id, c.级别, c.记录状态, min(C.场合) as 场合,max(e.场合) as 场合2 " & vbNewLine & _
                                "       From 人员抗菌药物权限 C, 人员抗菌药物权限 E" & vbNewLine & _
                                "       Where c.人员id = e.人员id And c.级别 = e.级别 And c.记录状态 = 1 And e.记录状态 = 1 And c.场合 <= e.场合 Group By c.人员id, c.级别, c.记录状态) C, 人员性质说明 D " & vbNewLine & _
                                "Where c.人员id(+) = a.Id And a.Id = b.人员id And d.人员id = a.Id And b.缺省 = 1 And" & vbNewLine & _
                                "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And" & vbNewLine & _
                                "      (( c.级别 = 0) or (C.场合 <> [2] and C.级别 = [3] and (C.场合 + C.场合2  <> 3)) Or (d.人员性质 = '医生' And c.人员id Is Null)) "
                

    '判断是否有所有部门权限
    If InStr(mstrPrivs, ";所有部门;") = 0 Then
        strSql = strSql & " And Instr([1],','|| B.部门ID || ',')>0 "
    End If
    
    strSql = strSql & strIsShowAll & ")" & vbNewLine & _
            "Select *" & vbNewLine & _
            "       From (Select Distinct b.Id, 'D' || b.Id As 部门人员id, b.名称, '部门' 性别, b.编码, '' As 树形上级id,NULL 上级id, '' As 拼音简码, '' As 五笔简码" & vbNewLine & _
            "              From Test A, 部门表 B" & vbNewLine & _
            "              Where a.树形部门id = 'D' || b.Id And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null)" & vbNewLine & _
            "              Union All" & vbNewLine & _
            "              Select * From Test Order By 编码)" & vbNewLine & _
            "       Start With 上级id Is Null And 编码 <> '-'" & vbNewLine & _
            "       Connect By Prior 部门人员id = 树形上级id)"

    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "," & mstrUserDept & ",", IIf(optOccasion(0).Value, 1, 2), mlngLastRunModule)
    
    tvwSelect.Nodes.Clear
    Do Until rsTmp.EOF
        If rsTmp!性别 & "" = "部门" Then
            strTemp = "Dept"
        ElseIf rsTmp!性别 & "" = "男" Then
            strTemp = "Male"
        ElseIf rsTmp!性别 & "" = "女" Then
            strTemp = "feMale"
        Else
            strTemp = "Male"
        End If
        
        If IsNull(rsTmp("上级id")) Then
            tvwSelect.Nodes.Add , , "D" & rsTmp("id"), "【" & rsTmp("编码") & "】" & rsTmp("名称"), strTemp, strTemp
            tvwSelect.Nodes("D" & rsTmp("id")).Sorted = True
            tvwSelect.Nodes("D" & rsTmp("id")).Expanded = True
            tvwSelect.Nodes("D" & rsTmp("id")).ForeColor = &HFF0000
        Else
            '如果
            For i = 1 To tvwGrant.Nodes.Count
                If Not tvwGrant.Nodes(i).Parent Is Nothing Then
                     If tvwGrant.Nodes(i).Key = "C" & rsTmp("id") Then blnIsInGrant = True
                End If
            Next
            Set Node = tvwSelect.Nodes.Add("D" & rsTmp("上级id"), tvwChild, "C" & rsTmp("id"), "【" & rsTmp("编码") & "】" & rsTmp("名称"), strTemp, strTemp)
                tvwSelect.Nodes("C" & rsTmp("id")).Sorted = True
                Node.Tag = rsTmp!拼音简码 & "|" & rsTmp!五笔简码
            If blnIsInGrant Then
                Node.ForeColor = &H80000010
                Node.Checked = False
                blnIsInGrant = False
            End If
        End If
        
        rsTmp.MoveNext
    Loop
    Me.Refresh
    mlngFindNum = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Adjustment(Optional ByVal lngDragMode As Long)
'功能:调整授权
    Dim strSql As String
    Dim curDate As Date
    Dim strPatiIDs As String
    Dim i As Long
    Dim strMsg As String
    Dim strPatiName As String
    Dim strReturn As String   '返回的按钮信息
    Dim strCmds As String
    Dim strCmdsAll As String
    Dim lngMode As Long
    Dim blnIsCheck As Boolean
    Dim blnIsHaveCancle As Boolean
    Dim objItem As TaskPanelGroupItem
    
    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    With rptPati
        '判断是否有被选中的
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_选择).Checked Then
                    blnIsCheck = True
                    If rptPati.Rows(i).Record(col_记录状态).Value & "" = "0" Then
                        blnIsHaveCancle = True
                        Exit For
                    End If
                End If
            End If
        Next
        
        
        If blnIsCheck Then
            '确定多选的人员ID
            For i = 0 To rptPati.Rows.Count - 1
                If Not rptPati.Rows(i).GroupRow Then
                    If rptPati.Rows(i).Record(col_选择).Checked Then
                        strPatiIDs = strPatiIDs & IIf(strPatiIDs <> "", ",", "") & .Rows(i).Record(COL_人员ID).Value
                        strPatiName = strPatiName & IIf(strPatiName <> "", ",", "") & .Rows(i).Record(COL_姓名).Value
                    End If
                End If
            Next
        Else
            If .SelectedRows.Count = 0 Then
                MsgBox "请选中或勾选需要调整权限的人员。", vbInformation, Me.Caption
                Exit Sub
            End If
            '确定多选的人员ID
            For i = 0 To .SelectedRows.Count - 1
                If Not .SelectedRows(i).GroupRow Then
                    If .SelectedRows(i).Record(col_记录状态).Value & "" = "0" Then blnIsHaveCancle = True
                    strPatiIDs = strPatiIDs & IIf(strPatiIDs <> "", ",", "") & .SelectedRows(i).Record(COL_人员ID).Value
                    strPatiName = strPatiName & IIf(strPatiName <> "", ",", "") & .SelectedRows(i).Record(COL_姓名).Value
                End If
            Next
        End If
        
        If strPatiIDs = "" Then
            MsgBox "请选中或勾选需要调整权限的人员。", vbInformation, Me.Caption
            Exit Sub
        End If
        
        For Each objItem In tplFunc.Groups(1).Items
            If objItem.Enabled Then
                If objItem.Caption <> tplFunc.Groups(1).Items(mlngLastRunModule).Caption Then
                    strCmds = strCmds & IIf(strCmds = "", "", ",") & objItem.Caption
                End If
                strCmdsAll = strCmdsAll & IIf(strCmdsAll = "", "", ",") & objItem.Caption
            End If
        Next
        
        
        If InStr(strPatiName, ",") > 0 Then
            If lngDragMode <> 0 Then
                If lngDragMode = 1 Then
                    strReturn = "非限制使用"
                ElseIf lngDragMode = 2 Then
                    strReturn = "限制使用"
                ElseIf lngDragMode = 3 Then
                    strReturn = "特殊使用"
                End If
                strMsg = "您选择了多个人员，您确定要授予他们【" & strReturn & "】权限吗？"
                If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
            Else
                strMsg = "您选择了多个人员，请选择授需要授予他们的权限。"
                strReturn = zlCommFun.ShowMsgBox("调整权限", strMsg, IIf(blnIsHaveCancle, strCmdsAll, strCmds) & ",!取消(&C)", Me, vbInformation)
            End If
        Else
            If lngDragMode <> 0 Then
                If lngDragMode = 1 Then
                    strReturn = "非限制使用"
                ElseIf lngDragMode = 2 Then
                    strReturn = "限制使用"
                ElseIf lngDragMode = 3 Then
                    strReturn = "特殊使用"
                End If
                strMsg = "您选择了一个人员:【" & strPatiName & "】，您确定要授予他【" & strReturn & "】权限吗？"
                If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
            Else
                strMsg = "您选择了一个人员:【" & strPatiName & "】，请选择授需要授予他的权限。"
                strReturn = zlCommFun.ShowMsgBox("调整权限", strMsg, IIf(blnIsHaveCancle, strCmdsAll, strCmds) & ",!取消(&C)", Me, vbInformation)
            End If
        End If
        
        If strReturn = "" Then Exit Sub
        If strReturn = "非限制使用" Then
            lngMode = 1
        ElseIf strReturn = "限制使用" Then
            lngMode = 2
        ElseIf strReturn = "特殊使用" Then
            lngMode = 3
        Else
            Exit Sub
        End If
        
        strSql = "Zl_人员抗菌药物权限_Update('" & strPatiIDs & "'," & lngMode & ",'" & UserInfo.姓名 & "',to_date('" & _
                curDate & "','YYYY-MM-DD HH24:MI:SS')," & IIf(optOccasion(0).Value, 1, 2) & ")"
        
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        Call RunByModule(mlngLastRunModule & "")
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cancellation()
'功能:取消授权
    Dim strSql As String
    Dim curDate As Date
    Dim strPatiIDs As String
    Dim i As Long
    Dim strMsg As String
    Dim strPatiName As String
    Dim lngPatiID As Long   '用于定位到取消前用户定位的人员
    Dim blnIsCancel As Boolean   '判断选中的是否是已经取消的用户
    
    curDate = zlDatabase.Currentdate
    With rptPati
        '判断是否有被选中的
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_选择).Checked Then Exit For
            End If
        Next
        
        If i <= rptPati.Rows.Count - 1 Then
            '确定多选的人员ID
            For i = 0 To rptPati.Rows.Count - 1
                If Not rptPati.Rows(i).GroupRow Then
                    If rptPati.Rows(i).Record(col_选择).Checked Then
                        If .Rows(i).Record(col_记录状态).Value & "" = "0" Then
                            blnIsCancel = True
                        Else
                            strPatiIDs = strPatiIDs & IIf(strPatiIDs <> "", ",", "") & .Rows(i).Record(COL_人员ID).Value
                            strPatiName = strPatiName & IIf(strPatiName <> "", ",", "") & .Rows(i).Record(COL_姓名).Value
                        End If
                    End If
                End If
            Next
        Else
            If .SelectedRows.Count = 0 Then
                MsgBox "请选中或勾选需要取消权限的人员。", vbInformation, Me.Caption
                Exit Sub
            End If
            '确定多选的人员ID
            For i = 0 To .SelectedRows.Count - 1
                If Not .SelectedRows(i).GroupRow Then
                    If .SelectedRows(i).Record(col_记录状态).Value & "" = "0" Then
                        blnIsCancel = True
                    Else
                        strPatiIDs = strPatiIDs & IIf(strPatiIDs <> "", ",", "") & .SelectedRows(i).Record(COL_人员ID).Value
                        strPatiName = strPatiName & IIf(strPatiName <> "", ",", "") & .SelectedRows(i).Record(COL_姓名).Value
                    End If
                End If
            Next
        End If
        
        If strPatiIDs = "" Then
            If blnIsCancel = True Then
                MsgBox "该用户已经没有权限了，不能取消。", vbInformation, Me.Caption
            Else
                MsgBox "请选中或勾选需要取消权限的人员。", vbInformation, Me.Caption
            End If
            Exit Sub
        End If
        If InStr(strPatiName, ",") > 0 Then
            strMsg = "您确定要取消当前选择的一个或多个人员的权限吗？"
        Else
            strMsg = "您确定要取消【" & strPatiName & "】的抗生素使用权限吗？"
            lngPatiID = Val(strPatiIDs & "")
        End If
        
        strSql = "Zl_人员抗菌药物权限_Update('" & strPatiIDs & "',0,'" & UserInfo.姓名 & "',to_date('" & _
                curDate & "','YYYY-MM-DD HH24:MI:SS')," & IIf(optOccasion(0).Value, 1, 2) & ")"
        
        On Error GoTo errH
        If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Call RunByModule(mlngLastRunModule & "", lngPatiID)
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
'功能:记录表打印
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objReport As ReportControl
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    
    If rptPati.Visible = False Then Exit Sub
    If rptPati.Records.Count > 0 Then
        Set objReport = rptPati
        strSubhead = "【" & tplFunc.Groups(1).Items(mlngLastRunModule).Caption & "】人员表"
    Else
        Exit Sub
    End If
    
    '-------------------------------------------------
    If zlControl.RPTCopyToVSF(objReport, vsTemp) Is Nothing Then Exit Sub
    '调用打印部件处理
    
    Set objPrint.Body = Me.vsTemp
    objPrint.Title.Text = strSubhead
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人:" & UserInfo.姓名)
    Call objAppRow.Add("打印时间:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    lngBottom = lngBottom - IIf(stbThis.Visible, stbThis.Height, 0)
    stcItem.Top = lngTop
    With tplFunc
        .Top = lngTop + stcItem.Height
        .Height = lngBottom - .Top
    End With
    
    stcLabel.Top = stcItem.Top
    stcLabel.Width = lngRight - tplFunc.Width - 45
    
    With rptPati
        .Left = tplFunc.Left + tplFunc.Width + 45
        .Top = tplFunc.Top
        .Width = lngRight - tplFunc.Width - 45
        .Height = tplFunc.Height
    End With
    
    With picGrant
        .Left = rptPati.Left
        .Top = rptPati.Top
        .Width = rptPati.Width
        .Height = rptPati.Height
    End With
    
    chkIsShowAll.Move stcLabel.Left + 1400, stcLabel.Top
    lblMsgHave.Move picGrant.Left + tvwGrant.Left, stcLabel.Top
    
    Me.Refresh
End Sub

Private Sub SetControlVisible(ByRef Control As XtremeCommandBars.ICommandBarControl)
    '根据权限设置按钮可见状态
    
    Select Case Control.ID
        Case conMenu_Edit_Untread, conMenu_Edit_Save
            Control.Visible = picGrant.Visible
        Case conMenu_Kss_Adjustment, conMenu_Kss_Cancellation, conMenu_Kss_Grant
            Control.Visible = Not picGrant.Visible
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim rptRecord As ReportRecord
        
'    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    Select Case Control.ID
    
        Case conMenu_File_Preview
            Control.Enabled = Not picGrant.Visible
        Case conMenu_File_Print
            Control.Enabled = Not picGrant.Visible
        Case conMenu_File_Excel
            Control.Enabled = Not picGrant.Visible
        Case conMenu_Kss_ShowCancel
            Control.Enabled = Not picGrant.Visible
        Case conMenu_Kss_Grant   '授权
            Control.Enabled = mlngLastRunModule <> 0
        Case conMenu_Kss_Adjustment, conMenu_Kss_Cancellation  '调整权限,取消权限
            Control.Enabled = rptPati.SelectedRows.Count > 0
            If Control.Enabled = False Then
                For Each rptRecord In rptPati.Records
                    If rptRecord(col_选择).Checked Then Control.Enabled = True: Exit For
                Next
            End If
            If Control.ID = conMenu_Kss_Adjustment And Control.Enabled Then
                Control.Enabled = mintNoPrivs <> 2
            End If
        Case conMenu_View_ToolBar_Button '工具栏
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_FindNext '查找下一个
            Control.Visible = False
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub chkIsShowAll_Click()
    Call LoadPrss
End Sub

Private Function GetDelStr(ByRef objtvwFrm As TreeView, ByRef objtvwTo As TreeView, ByVal Index As Long, _
     ByRef strDel As String, ByVal blnIsSelect As Boolean, ByVal blnIsAllMove As Boolean) As String
'功能：将一个列表中的一项移到另外一个列表，并返回需要删除的字符串
'参数： objtvwFrm 需要移出的树表
'       objtvwTo 需要移进的树表
'       index移出列表需要移出的项的索引
'       strDel   需要删除的项的字符串
'       blnIsAllMove 是否是选择的父项
    Dim y As Long
    Dim Node As Node
    Dim lngDelParant As Long
    Dim blnIsExist As Boolean
    Dim NodeIsParant As Node
    
    If blnIsAllMove Then
        Set NodeIsParant = objtvwFrm.Nodes(Index)
    Else
        Set NodeIsParant = objtvwFrm.Nodes(Index).Parent
    End If
    If Not NodeIsParant Is Nothing Or blnIsAllMove Then
        '添加父项
        For y = 1 To objtvwTo.Nodes.Count
            If objtvwTo.Nodes(y).Key = NodeIsParant.Key Then blnIsExist = True: Exit For
        Next
        If Not blnIsExist Then
            If objtvwTo.Name = tvwSelect.Name Then
                objtvwTo.Nodes(NodeIsParant.Key).ForeColor = &H80000012
                objtvwTo.Nodes(NodeIsParant.Key).Checked = NodeIsParant.Checked
                objtvwTo.Nodes(NodeIsParant.Key).Expanded = NodeIsParant.Expanded
            Else
                Set Node = objtvwTo.Nodes.Add(, , NodeIsParant.Key, NodeIsParant.Text, NodeIsParant.Image, NodeIsParant.SelectedImage)
                Node.Expanded = NodeIsParant.Expanded
                Node.Checked = NodeIsParant.Checked
                Node.ForeColor = &HFF0000
            End If
        Else
            objtvwTo.Nodes(y).Checked = NodeIsParant.Checked
        End If
        blnIsExist = False
        '添加自己/添加子项
        If blnIsAllMove Then
            For y = NodeIsParant.Index + 1 To objtvwFrm.Nodes.Count
                If Not objtvwFrm.Nodes(y).Parent Is Nothing Then
                    If objtvwFrm.Nodes(y).Parent.Key = NodeIsParant.Key Then
                        If objtvwTo.Name = tvwSelect.Name Then
                            objtvwTo.Nodes(objtvwFrm.Nodes(y).Key).ForeColor = &H80000012
                            objtvwTo.Nodes(objtvwFrm.Nodes(y).Key).Checked = objtvwFrm.Nodes(Index).Checked
                        Else
                            Set Node = objtvwTo.Nodes.Add(NodeIsParant.Key, tvwChild, objtvwFrm.Nodes(y).Key, objtvwFrm.Nodes(y).Text, objtvwFrm.Nodes(y).Image, objtvwFrm.Nodes(y).SelectedImage)
                            Node.Checked = objtvwFrm.Nodes(Index).Checked
                        End If
                    End If
                End If
            Next
        Else
            If objtvwTo.Name = tvwSelect.Name Then
                objtvwTo.Nodes(objtvwFrm.Nodes(Index).Key).ForeColor = &H80000012
                objtvwTo.Nodes(objtvwFrm.Nodes(Index).Key).Checked = objtvwFrm.Nodes(Index).Checked
            Else
                Set Node = objtvwTo.Nodes.Add(NodeIsParant.Key, tvwChild, objtvwFrm.Nodes(Index).Key, objtvwFrm.Nodes(Index).Text, objtvwFrm.Nodes(Index).Image, objtvwFrm.Nodes(Index).SelectedImage)
                Node.Checked = objtvwFrm.Nodes(Index).Checked
            End If
        End If
        
        
        lngDelParant = 0
        If Not blnIsAllMove Then
            For y = NodeIsParant.Index + 1 To objtvwFrm.Nodes.Count
                If Not objtvwFrm.Nodes(y).Parent Is Nothing Then
                    If objtvwFrm.Nodes(y).Parent.Key = objtvwFrm.Nodes(Index).Parent.Key Then
                        If blnIsSelect Then
                            If blnIsSelect And objtvwFrm.Nodes(y).Key <> objtvwFrm.Nodes(Index).Key Then
                                lngDelParant = lngDelParant + 1
                            End If
                        Else
                            If Not objtvwFrm.Nodes(y).Checked Then
                                lngDelParant = lngDelParant + 1
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If lngDelParant = 0 Then
            '删除父项
            If InStr("," & strDel & ",", "," & NodeIsParant.Index & ",") <= 0 Then
                strDel = strDel & IIf(strDel = "", "", ",") & NodeIsParant.Index
            End If
        Else
            '删除自己
            If InStr("," & strDel & ",", "," & NodeIsParant.Index & ",") <= 0 Then
                strDel = strDel & IIf(strDel = "", "", ",") & objtvwFrm.Nodes(Index).Index
            End If
        End If
    End If
    GetDelStr = strDel
End Function

Private Sub cmdMove_Click(Index As Integer)
    Dim i As Long, y As Long
    Dim strDel As String
    
    If Index = 0 Then
        '授权
        For i = 1 To tvwSelect.Nodes.Count
            If tvwSelect.Nodes(i).Checked And tvwSelect.Nodes(i).ForeColor <> &H80000010 Then
                strDel = GetDelStr(tvwSelect, tvwGrant, i, strDel, False, False)
            End If
        Next
        If Not tvwSelect.SelectedItem Is Nothing Then
            If strDel = "" And tvwSelect.SelectedItem.ForeColor <> &H80000010 Then
                strDel = GetDelStr(tvwSelect, tvwGrant, tvwSelect.SelectedItem.Index, strDel, True, IIf(tvwSelect.SelectedItem.Parent Is Nothing, True, False))
            End If
        End If
        '删除
        For i = UBound(Split(strDel, ",")) To 0 Step -1
            'Call tvwSelect.Nodes.Remove(Val(Split(strDel, ",")(i)))
            If tvwSelect.Nodes(Val(Split(strDel, ",")(i))).Parent Is Nothing Then
                For y = tvwSelect.Nodes(Val(Split(strDel, ",")(i))).Index + 1 To tvwSelect.Nodes.Count
                    If tvwSelect.Nodes(y).Parent Is Nothing Then Exit For
                    tvwSelect.Nodes(y).ForeColor = &H80000010
                    tvwSelect.Nodes(y).Checked = False
                    tvwSelect.Nodes(Val(Split(strDel, ",")(i))).Checked = False
                Next
            Else
                tvwSelect.Nodes(Val(Split(strDel, ",")(i))).ForeColor = &H80000010
                tvwSelect.Nodes(Val(Split(strDel, ",")(i))).Checked = False
            End If
            mblnIsChange = True
        Next
    ElseIf Index = 1 Then
        '取消
        For i = 1 To tvwGrant.Nodes.Count
            If tvwGrant.Nodes(i).Checked Then
                strDel = GetDelStr(tvwGrant, tvwSelect, i, strDel, False, False)
            End If
        Next
        If strDel = "" And Not tvwGrant.SelectedItem Is Nothing Then
            strDel = GetDelStr(tvwGrant, tvwSelect, tvwGrant.SelectedItem.Index, strDel, True, IIf(tvwGrant.SelectedItem.Parent Is Nothing, True, False))
        End If
        '删除
        For i = UBound(Split(strDel, ",")) To 0 Step -1
            Call tvwGrant.Nodes.Remove(Val(Split(strDel, ",")(i)))
            mblnIsChange = True
        Next
    End If
End Sub

Private Sub Form_Activate()
    mblnFirstLoad = False
End Sub

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
    mblnIsUp = False     '拖动结束
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Dim tpGroupItem As TaskPanelGroupItem
    Dim strHead As String
    
    mstrPrivs = gstrPrivs
    If (InStr(mstrPrivs, ";非限制使用;") = 0 And InStr(mstrPrivs, ";限制使用;") = 0 And InStr(mstrPrivs, ";特殊使用;") = 0) Or InStr(mstrPrivs, ";基本;") = 0 Then
        If InStr(mstrPrivs, ";基本;") = 0 Then
            MsgBox "您没有抗菌药物授权的基本权限,请与管理员联系。", vbInformation, Me.Caption
        Else
            MsgBox "您没有任何一类抗生素授权的权限，请与管理员联系。", vbInformation, Me.Caption
        End If
        Unload Me
        Exit Sub
    End If
    mstrUserDept = GetUser科室IDs(False)
    mlngModul = glngModul
    mblnFirstLoad = True
    mlngFindNum = 0
    mblnIsFindFinish = False
    mlngCodeType = zlDatabase.GetPara("简码方式")
    
    'TaskPanel
    '----------------------------------------------------
    mintNoPrivs = 0
    Set tpGroup = tplFunc.Groups.Add(1, "权限分类")
    Set tpGroupItem = tpGroup.Items.Add(PanelItem_NotLimit, "非限制使用", xtpTaskItemTypeLink, PanelItem_NotLimit + 1)
    tpGroupItem.Selected = False
    If InStr(mstrPrivs, ";非限制使用;") = 0 Then
        tpGroupItem.Enabled = False: tpGroupItem.Tag = "不可用"
        mintNoPrivs = mintNoPrivs + 1
    End If
    Set tpGroupItem = tpGroup.Items.Add(PanelItem_Limit, "限制使用", xtpTaskItemTypeLink, PanelItem_Limit + 1)
    tpGroupItem.Selected = False
    If InStr(mstrPrivs, ";限制使用;") = 0 Then
        tpGroupItem.Enabled = False
        tpGroupItem.Tag = "不可用"
        mintNoPrivs = mintNoPrivs + 1
    End If
    Set tpGroupItem = tpGroup.Items.Add(PanelItem_Special, "特殊使用", xtpTaskItemTypeLink, PanelItem_Special + 1)
    tpGroupItem.Selected = False
    If InStr(mstrPrivs, ";特殊使用;") = 0 Then
        tpGroupItem.Enabled = False
        tpGroupItem.Tag = "不可用"
        mintNoPrivs = mintNoPrivs + 1
    End If
    
    tplFunc.SetMargins 1, 2, 0, 2, 2
    tplFunc.SelectItemOnFocus = True
    Call tplFunc.Icons.AddIcons(ImgC.Icons)
    tplFunc.SetIconSize 24, 24
    tpGroup.CaptionVisible = False
    tpGroup.Expanded = True
    stcItem.Caption = "权限分类"
    stcLabel.Caption = "人员列表"
    stcItem.Font.Size = 9
    stcLabel.Font.Size = 9
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    
    'ReportControl
    '-----------------------------------------------------
    Call InitPatiReportColumn
    
    Call RestoreWinState(Me, App.ProductName)
    
    mobjShowCancel.Checked = _
        GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示取消", 0)
End Sub

Private Sub InitPatiReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)或ItemIndex查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(col_选择, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.Editable = True
        objCol.Icon = img16.ListImages("unCheck").Index - 1
        Set objCol = .Columns.Add(COL_姓名, "姓名", 80, True)
            objCol.Groupable = False
        Set objCol = .Columns.Add(COL_编号, "编号", 60, True)
            objCol.Groupable = False
        Set objCol = .Columns.Add(col_性别, "性别", 32, True)
            objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_专业职务, "专业职务", 70, True)
        Set objCol = .Columns.Add(COL_所属部门, "所属部门", 80, True)
        Set objCol = .Columns.Add(COL_授权人, "授权人", 62, True)
        Set objCol = .Columns.Add(COL_授权日期, "授权日期", 125, True)
            objCol.Groupable = False
        Set objCol = .Columns.Add(COL_人员ID, "人员id", 0, False)
        Set objCol = .Columns.Add(col_记录状态, "记录状态", 0, False)
        For Each objCol In .Columns
            If objCol.Index <> col_选择 Then objCol.Editable = False
            If objCol.Width = 0 Then objCol.Visible = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "本分类下没有人员..."
            '.ShadeGroupHeadings = True
            
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = True
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = True
        .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(COL_所属部门)
        
    End With
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim lngCount As Long
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
            objControl.BeginGroup = True
    End With
    
    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Kss_Jurisdiction, "权限(&M)", -1, False)
    mobjMenu.ID = conMenu_Kss_Jurisdiction
    With mobjMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant, "授权(&G)")
        Set objControl = .Add(xtpControlButton, conMenu_Kss_Cancellation, "取消权限(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Kss_Adjustment, "调整权限(&A)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Kss_ShowCancel, "显示取消授权的人员(&C)")
            objControl.BeginGroup = True
        Set mobjShowCancel = objControl
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
            objControl.BeginGroup = True
    End With

    '主菜单右侧的查找
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, 0, "人员查找")
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = txtFind.hWnd
        objCustom.Flags = xtpFlagRightAlign
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set mobjBar = cbsMain.Add("工具栏", xtpBarTop)
    With mobjBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant, "授权(&G)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
        objControl.BeginGroup = True
        objControl.Visible = False
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        objControl.Visible = False
        Set objControl = .Add(xtpControlButton, conMenu_Kss_Cancellation, "取消权限(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Kss_Adjustment, "调整权限(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        Set objCustom = .Add(xtpControlCustom, conMenu_View_FindType, "场合")
            objCustom.Handle = fraType.hWnd
            objCustom.Flags = xtpFlagRightAlign
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找
        .Add 0, vbKeyF3, conMenu_View_FindNext '查找下一个
        
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With

    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnIsChange Then
        If MsgBox("您已经进行了改动，您确认要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    Call SaveWinState(Me, App.ProductName)
    mlngLastRunModule = 0

    If Not mobjShowCancel Is Nothing Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示取消", _
            IIf(mobjShowCancel.Checked, 1, 0)
    End If
    Set mobjShowCancel = Nothing
    Set mpicOld = Nothing
    Set mrsFind = Nothing
End Sub

Private Sub optOccasion_Click(Index As Integer)
    If mlngLastRunModule <> 0 Then Call RunByModule(mlngLastRunModule)
End Sub

Private Sub picGrant_Resize()
    With tvwSelect
        .Left = 0
        .Top = 0
        .Height = picGrant.Height
        .Width = picGrant.Width / 2 - 300
    End With
    
    With tvwGrant
        .Left = tvwSelect.Width + 600
        .Top = 0
        .Height = picGrant.Height
        .Width = picGrant.Width / 2 - 300
    End With
    
    cmdMove(0).Top = picGrant.Height / 2 - 1000
    cmdMove(0).Left = tvwSelect.Width + 100
    
    cmdMove(1).Top = picGrant.Height / 2
    cmdMove(1).Left = tvwSelect.Width + 100
    
End Sub

Private Sub rptPati_DragDrop(Source As Control, x As Single, y As Single)
    mblnIsUp = False     '拖动结束
End Sub

Private Sub rptPati_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Long
    
    If Button = 1 Then
        mMouseX = x
        mMouseY = y
        mblnIsUp = True
        '判断是否有被选中的
        mblnIsCheck = False
        mblnIsHaveCancle = False
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record(col_选择).Checked Then
                    mblnIsCheck = True
                    If rptPati.Rows(i).Record(col_记录状态).Value & "" = "0" Then
                        mblnIsHaveCancle = True
                        Exit For
                    End If
                End If
            End If
        Next
    End If
End Sub

Private Sub rptPati_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 1 Then
        With rptPati
            If mblnIsUp Then
                If .SelectedRows.Count = 0 Then Exit Sub
                If rptPati.SelectedRows.Count > 1 Or mblnIsCheck Then
                    Set Me.rptPati.DragIcon = img32.ListImages("Pepoles").Picture
                    '如果是拖动状态并且移动的距离长过150就出现图标
                    If mblnIsUp And (Abs(x - mMouseX) > 10 Or Abs(y - mMouseY) > 10) Then Me.rptPati.Drag 1
                ElseIf rptPati.SelectedRows.Count = 1 Then
                    If Not .SelectedRows(0).GroupRow Then
                        Set Me.rptPati.DragIcon = img32.ListImages(IIf(rptPati.SelectedRows(0).Record(col_性别).Value = "女", "feMale", "Male")).Picture
                        '如果是拖动状态并且移动的距离长过150就出现图标
                        If mblnIsUp And (Abs(x - mMouseX) > 10 Or Abs(y - mMouseY) > 10) Then Me.rptPati.Drag 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '如果点击表头的图片，就选中全部
    If Button = 1 Then
        If rptPati.HitTest(x, y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(x, y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = col_选择 Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptPati.Columns(col_选择).Icon = img16.ListImages("Check").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(col_选择).Checked = True
                        Next
                    Else
                        objColumn.Caption = ""
                        rptPati.Columns(col_选择).Icon = img16.ListImages("unCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(col_选择).Checked = False
                        Next
                    End If
                End If
            End If
        End If
    End If
    '还原状态栏
    stbThis.Panels(2).Text = stbThis.Panels(2).Tag
End Sub

Private Sub rptPati_SelectionChanged()
    '如果选中的行是取消权限的人员，则选中背景色设置为灰色
    With rptPati
        If .Visible = False Or .SelectedRows.Count = 0 Then Exit Sub
        If Not .SelectedRows(0).GroupRow Then
            If Val(.SelectedRows(0).Record(col_记录状态).Value & "") = 0 Then
                .PaintManager.HighlightBackColor = RGB(169, 210, 252)
                .PaintManager.HighlightForeColor = RGB(122, 123, 126)
            Else
                .PaintManager.HighlightBackColor = RGB(89, 169, 249)
                .PaintManager.HighlightForeColor = &H80000008
            End If
        End If
    End With
End Sub

Private Sub tplFunc_DragDrop(Source As Control, x As Single, y As Single)
    Dim strMsg As String
    Dim strSql As String
    Dim curDate As Date
    Dim lngGrade As Long
    Dim strMsgState As String
    
    curDate = zlDatabase.Currentdate
    If Source.Name = "rptPati" Then
        If mblnIsUp Then
            mblnIsUp = False     '拖动结束
            If rptPati.DragIcon = img32.ListImages("Yes").Picture Then
                If y > 0 And y < 500 Then
                    lngGrade = 1
                ElseIf y >= 500 And y < 1000 Then
                    lngGrade = 2
                ElseIf y >= 1000 And y < 1500 Then
                    lngGrade = 3
                End If
                '调整权限
                Call Adjustment(lngGrade)
            End If
        End If
    End If
    '还原状态栏
    stbThis.Panels(2).Text = stbThis.Panels(2).Tag
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tplFunc_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Dim i As Long
    Dim rowSelect As ReportRow
    Dim lngItem As Long
    
    If y < 500 Then
        lngItem = 1
    ElseIf y < 1000 Then
        lngItem = 2
    ElseIf y < 1500 Then
        lngItem = 3
    End If
    If Source.Name = "rptPati" Then
        If State = 0 Or State = 2 Then    '进入或者经过
            If State = 0 Then Set mpicOld = Source.DragIcon    '进入的时候记录进入的时候的图片，一边移出的时候还原
            If mblnIsUp Then
                If mblnIsCheck Then
                    If y >= 1500 Then
                        Set Source.DragIcon = img32.ListImages("No").Picture
                    Else
                        If mblnIsHaveCancle Then
                            Set Source.DragIcon = img32.ListImages(IIf(tplFunc.Groups(1).Items(lngItem).Enabled, "Yes", "No")).Picture
                        Else
                            If y < mlngLastRunModule * 500 And y > (mlngLastRunModule - 1) * 500 Then
                                Set Source.DragIcon = img32.ListImages("No").Picture
                            Else
                                Set Source.DragIcon = img32.ListImages(IIf(tplFunc.Groups(1).Items(lngItem).Enabled, "Yes", "No")).Picture
                            End If
                        End If
                    End If
                Else
                    If rptPati.SelectedRows.Count = 1 Then
                        If (y < mlngLastRunModule * 500 And y > (mlngLastRunModule - 1) * 500) And _
                        rptPati.SelectedRows(0).Record(col_记录状态).Value & "" <> "0" Or y >= 1500 Then
                            Set Source.DragIcon = img32.ListImages("No").Picture
                        Else
                            Set Source.DragIcon = img32.ListImages(IIf(tplFunc.Groups(1).Items(lngItem).Enabled, "Yes", "No")).Picture
                        End If
                    ElseIf rptPati.SelectedRows.Count > 1 Then
                        If y >= 1500 Then
                            Set Source.DragIcon = img32.ListImages("No").Picture
                        Else
                            For Each rowSelect In rptPati.SelectedRows
                                If Not rowSelect.GroupRow Then
                                    If rowSelect.Record(col_记录状态).Value & "" = "0" Then
                                        Set Source.DragIcon = img32.ListImages(IIf(tplFunc.Groups(1).Items(lngItem).Enabled, "Yes", "No")).Picture
                                        Exit For
                                    Else
                                        If y < mlngLastRunModule * 500 And y > (mlngLastRunModule - 1) * 500 Then
                                            Set Source.DragIcon = img32.ListImages("No").Picture
                                        Else
                                            Set Source.DragIcon = img32.ListImages(IIf(tplFunc.Groups(1).Items(lngItem).Enabled, "Yes", "No")).Picture
                                        End If
                                    End If
                                End If
                            Next
                             
                        End If
                    End If
                End If
            End If
            If Source.DragIcon = img32.ListImages("Yes").Picture Then
                If y < 500 Then
                    stbThis.Panels(2).Text = "放开鼠标左键可移入【非限制使用】分类。"
                ElseIf y < 1000 Then
                    stbThis.Panels(2).Text = "放开鼠标左键可移入【限制使用】分类。"
                ElseIf y < 1500 Then
                    stbThis.Panels(2).Text = "放开鼠标左键可移入【特殊使用】分类。"
                End If
            Else
                If y < 1500 Then
                    stbThis.Panels(2).Text = IIf(tplFunc.Groups(1).Items(lngItem).Enabled, "您选中的人员已经在这个分类了", "您没有这个分类的权限。")
                Else
                    stbThis.Panels(2).Text = "您可以将您选中的人员移入想调整的分类中。"
                End If
            End If
        ElseIf State = 1 Then '移出
            If Not mpicOld Is Nothing Then Set Source.DragIcon = mpicOld: Set mpicOld = Nothing
            '还原状态栏
            stbThis.Panels(2).Text = stbThis.Panels(2).Tag
        End If
    End If
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    If mlngLastRunModule <> Item.ID Then
        Call RunByModule("0" & Item.ID)
    End If
End Sub

Public Sub RunByModule(ByVal strModule As String, Optional ByVal lngPatiID As Long, Optional ByVal blnIsSelect As Boolean)
'功能：加载人员表格
'参数：lngPatiID 如果<>0则是刷新,刷新后定位到刷新前的哪一行，否则为第一次加载
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long, y As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objRow As ReportRow
    Dim strAllDept As String      '如果没有所有部门权限，则取所属科室的人员信息
    
    mlngLastRunModule = Val(strModule)
    
    On Error GoTo errH
    If Not blnIsSelect Then
        strSql = "Select a.人员id, b.姓名, b.编号, b.性别, b.专业技术职务, d.名称 As 部门名称, a.记录状态, a.操作人员, a.操作时间" & vbNewLine & _
                "From (Select 人员id, 记录状态, 操作人员, 操作时间,场合" & vbNewLine & _
                "       From 人员抗菌药物权限" & vbNewLine & _
                "       Where 级别 = [1] And 记录状态 = 1"
        
        If mobjShowCancel.Checked Then
            '确定是否显示取消授权的人员
            strSql = strSql & " Union All" & vbNewLine & _
                            "       Select b.人员id, b.记录状态, b.操作人员, b.操作时间,b.场合" & vbNewLine & _
                            "       From (Select a.人员id, a.记录状态, a.操作人员, a. 操作时间" & vbNewLine & _
                            "              From 人员抗菌药物权限 A" & vbNewLine & _
                            "              Where a.级别 = 0 And a.记录状态 = 1) A, 人员抗菌药物权限 B" & vbNewLine & _
                            "       Where a.人员id = b.人员id And b.级别 = [1] And" & vbNewLine & _
                            "             b.操作时间 = (Select Max(操作时间) From 人员抗菌药物权限 Where 人员id = a.人员id And 级别 <>0)"
    
        End If
        strSql = strSql & ") A, 人员表 B, 部门人员 C, 部门表 D Where a.人员id = b.Id And " & _
                        " c.人员id = a.人员id And c.部门id = d.Id And c.缺省 = 1 And a.场合=[3]"
        strAllDept = " And Instr([2],','||d.ID || ',')>0"
        If InStr(mstrPrivs, ";所有部门;") = 0 Then
            strSql = strSql & strAllDept
        End If
        
        strSql = strSql & " Order By d.名称"
        
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(strModule), "," & mstrUserDept & ",", IIf(optOccasion(0).Value, 1, 2))
        
        rptPati.Records.DeleteAll
        mblnIsFindFinish = False
        
        If rsTmp.RecordCount > 0 Then
            i = 0
            With rptPati
                Do While Not rsTmp.EOF
                    Set objRecord = .Records.Add()
                    
                    '选择按钮
                    Set objItem = objRecord.AddItem("")
                    objItem.HasCheckbox = True
                        objItem.Checked = False '缺省未选中
                    Set objItem = objRecord.AddItem(rsTmp!姓名 & "")
                        objItem.Icon = img16.ListImages.Item(IIf(rsTmp!性别 & "" = "女", "feMale", "Male")).Index - 1
                    Set objItem = objRecord.AddItem(rsTmp!编号 & "")
                    Set objItem = objRecord.AddItem(rsTmp!性别 & "")
                    Set objItem = objRecord.AddItem(rsTmp!专业技术职务 & "")
                    Set objItem = objRecord.AddItem(rsTmp!部门名称 & "")
                    Set objItem = objRecord.AddItem(rsTmp!操作人员 & "")
                    Set objItem = objRecord.AddItem(rsTmp!操作时间 & "")
                    Set objItem = objRecord.AddItem(rsTmp!人员ID & "")
                    Set objItem = objRecord.AddItem(rsTmp!记录状态 & "")
    
                    If rsTmp!记录状态 & "" = "0" Then
                        For y = COL_姓名 To col_记录状态
                            objRecord.Item(y).ForeColor = RGB(122, 123, 126)
                        Next
                    End If
                    
                    If rsTmp!记录状态 & "" <> "0" Then i = i + 1
                    rsTmp.MoveNext
                Loop
            End With
        End If
        rptPati.Populate
        stbThis.Panels(2).Text = "当前分类下有 " & i & " 名人员拥有该权限" & IIf(mobjShowCancel.Checked, "， " & rsTmp.RecordCount - i & " 名人员被取消了权限", "") & "。"
        stbThis.Panels(2).Tag = "当前分类下有 " & i & " 名人员拥有该权限" & IIf(mobjShowCancel.Checked, "， " & rsTmp.RecordCount - i & " 名人员被取消了权限", "") & "。"
    End If
    If lngPatiID <> 0 Then
        '如果是刷新，则定位到刷新前的那行
        For Each objRow In rptPati.Rows
            If Not objRow.GroupRow Then
                If Val(objRow.Record(COL_人员ID).Value & "") = lngPatiID And lngPatiID <> 0 Then
                    Set rptPati.FocusedRow = objRow
                    Exit For
                End If
            End If
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvwGrant_NodeCheck(ByVal Node As MSComctlLib.Node)
    Call NodeCheckMode(Node, tvwGrant)
End Sub

Private Sub tvwSelect_NodeCheck(ByVal Node As MSComctlLib.Node)
    Call NodeCheckMode(Node, tvwSelect)
End Sub

Private Sub NodeCheckMode(ByRef Node As MSComctlLib.Node, ByRef objtvwThis As TreeView)
'功能：让树表选中父节点，自动选中所有子节点，选中所有子节点，父节点也选中
    Dim i As Long
    Dim blnIsNothing As Boolean
    
    '如果是灰色的就退出
    If Node.ForeColor = &H80000010 Then Exit Sub
    If Node.Parent Is Nothing Then
        For i = Node.Index + 1 To objtvwThis.Nodes.Count
            If Not objtvwThis.Nodes(i).Parent Is Nothing And objtvwThis.Nodes(i).ForeColor <> &H80000010 Then
                If objtvwThis.Nodes(i).Parent.Key = Node.Key Then
                    objtvwThis.Nodes(i).Checked = Node.Checked
                End If
            End If
        Next
    Else
        For i = Node.Parent.Index + 1 To objtvwThis.Nodes.Count
            If Not objtvwThis.Nodes(i).Parent Is Nothing And objtvwThis.Nodes(i).ForeColor <> &H80000010 Then
                If objtvwThis.Nodes(i).Parent.Key = Node.Parent.Key Then
                    If Not objtvwThis.Nodes(i).Checked = Node.Checked Then blnIsNothing = True
                End If
            End If
        Next
        If blnIsNothing Then
            Node.Parent.Checked = False
        Else
            Node.Parent.Checked = Node.Checked
        End If
    End If
End Sub

Private Sub txtFind_Change()
    '值改变后算从新查找
    mblnIsFindFinish = False
    mlngFindNum = 0
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strSql As String
    Dim strMsg As String
    Dim Node As Node
    Dim i As Long
    Dim tpGroupItem As TaskPanelGroupItem
    Dim blnIsSelect As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If mblnIsFindFinish = False Then Set mrsFind = Nothing
    
    If picGrant.Visible Then
        '授权页面查询
        strMsg = UCase(Trim(txtFind.Text))
        If tvwSelect.Nodes.Count > 0 Then
            For Each Node In tvwSelect.Nodes
                If Not Node.Parent Is Nothing Then
                    If Mid(Node.Text, InStr(Node.Text, "【") + 1, InStr(Node.Text, "】") - 2) Like strMsg & "*" _
                        Or Mid(Node.Text, InStr(Node.Text, "】") + 1) Like IIf(gstrMatch = "%", "*", "") & strMsg & "*" _
                        Or Split(Node.Tag, "|")(mlngCodeType) Like IIf(gstrMatch = "%", "*", "") & strMsg & "*" Then
                        If i >= mlngFindNum Then
                            Node.Selected = True
                            mlngFindNum = mlngFindNum + 1
                            Exit For
                        Else
                            i = i + 1
                        End If
                    End If
                    
                    If Node.Index = tvwSelect.Nodes(tvwSelect.Nodes.Count).Index Then
                        If mlngFindNum = 0 And i = 0 Then
                            MsgBox "没有找到您查找的人员，或是您查找的人员已经有权限了。", vbInformation, Me.Caption
                        Else
                            MsgBox "已经查找到最后一个人员了。", vbInformation, Me.Caption
                            mlngFindNum = 0
                        End If
                    End If
                End If
            Next
        End If
    Else
        '人员查询
        If mblnIsFindFinish = False Then
            strSql = "Select a.人员id,a.级别,b.姓名,b.编号,a.取消" & vbNewLine & _
                    "From (Select 人员id,级别,级别 as 取消,场合" & vbNewLine & _
                    "       From 人员抗菌药物权限" & vbNewLine & _
                    "       Where 级别 <>0 And 记录状态 = 1 Union All" & vbNewLine & _
                    "       Select b.人员id,b.级别,a.级别 as 取消,B.场合" & vbNewLine & _
                    "       From (Select a.人员id,a.级别" & vbNewLine & _
                    "              From 人员抗菌药物权限 A" & vbNewLine & _
                    "              Where a.级别 = 0 And a.记录状态 = 1) A, 人员抗菌药物权限 B" & vbNewLine & _
                    "       Where a.人员id = b.人员id And b.级别 <>0 And" & vbNewLine & _
                    "             b.操作时间 = (Select Max(操作时间) From 人员抗菌药物权限 Where 人员id = a.人员id And 级别 <>0)) A, 人员表 B, 部门人员 C " & _
                    " Where a.人员id = b.Id And c.人员id = a.人员id And c.缺省 = 1 And A.场合=[4] " & _
                    IIf(mobjShowCancel.Checked, "", " And A.取消<>0")
            strMsg = UCase(Trim(txtFind.Text))
            
            '判断是否有所有部门的权限
            If InStr(mstrPrivs, ";所有部门;") = 0 Then
                strSql = strSql & " And Instr([3],','|| c.部门ID || ',')>0"
            End If
            
            If zlCommFun.IsCharChinese(strMsg) Then
                strSql = strSql & " And B.姓名 like [1]"
            ElseIf zlCommFun.IsCharAlpha(strMsg) Then
                strSql = strSql & " And (B.姓名 like [1] or UPPER(" & IIf(mlngCodeType = 1, "zlwbcode", "zlspellcode") & "(b.姓名)) like [1])"
            ElseIf IsNumeric(strMsg) Then
                strSql = strSql & " And (b.编号 like [2])"
            Else
                strSql = strSql & " And B.姓名 like [1]"
            End If
            '排除没有权限的类别
            For Each tpGroupItem In tplFunc.Groups(1).Items
                If tpGroupItem.Enabled = False Then
                    strSql = strSql & " And A.级别<>" & tpGroupItem.ID
                End If
            Next
            
            On Error GoTo errH
            Set mrsFind = zlDatabase.OpenSQLRecord(strSql, Me.Caption, gstrMatch & strMsg & "%", strMsg & "%", "," & mstrUserDept & ",", IIf(optOccasion(0).Value, 1, 2))
        End If
            
        If mrsFind.RecordCount > 0 Then
            If Not mrsFind.EOF Then
                '定位
                If tplFunc.Groups(1).Items(Val(mrsFind!级别 & "")).Selected Then blnIsSelect = True
                For Each tpGroupItem In tplFunc.Groups(1).Items
                    tpGroupItem.Selected = False
                Next
                tplFunc.Groups(1).Items(Val(mrsFind!级别 & "")).Selected = True
                Call RunByModule(Val(mrsFind!级别 & ""), Val(mrsFind!人员ID & ""), blnIsSelect)
                mrsFind.MoveNext
                mblnIsFindFinish = True
            Else
                MsgBox "已经查找到最后一个人员了。", vbInformation, Me.Caption
                mrsFind.MoveFirst
            End If
        Else
            MsgBox "没有找到您查找的人员，或是您查找的人员没有任何权限。", vbInformation, Me.Caption
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetUser科室IDs(Optional ByVal bln病区 As Boolean) As String
'功能：获取操作员所属的科室(本身所在科室+所属病区包含的科室),可能有多个
'参数：是否取所属病区下的科室
    Static rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, blnNew As Boolean
    
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    '没有强制限制临床,可能医技科室用
    If blnNew Then
        strSql = "Select 1 as 类别,部门ID From 部门人员 Where 人员ID=[1] Union" & _
                " Select Distinct 2 as 类别,B.科室ID From 部门人员 A,病区科室对应 B" & _
                " Where A.部门ID=B.病区ID And A.人员ID=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", UserInfo.ID)
    End If
    If bln病区 = False Then
        rsTmp.Filter = "类别 = 1"
    Else
        rsTmp.Filter = ""
    End If
    
    For i = 1 To rsTmp.RecordCount
        If InStr("," & GetUser科室IDs & ",", "," & rsTmp!部门ID & ",") = 0 Then
            GetUser科室IDs = GetUser科室IDs & "," & rsTmp!部门ID
        End If
        rsTmp.MoveNext
    Next
    GetUser科室IDs = Mid(GetUser科室IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
