VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmMedicalTeam 
   AutoRedraw      =   -1  'True
   Caption         =   "医疗小组管理"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10680
   Icon            =   "frmMedicalTeam.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList imgList 
      Left            =   6960
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalTeam.frx":06EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMember 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   4440
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   2
      Top             =   1680
      Width           =   3375
      Begin XtremeReportControl.ReportControl rpcMember 
         Height          =   1695
         Left            =   360
         TabIndex        =   3
         Top             =   120
         Width           =   2415
         _Version        =   589884
         _ExtentX        =   4260
         _ExtentY        =   2990
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
      End
   End
   Begin VB.PictureBox picTeam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   360
      ScaleHeight     =   2655
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
      Begin XtremeReportControl.ReportControl rpcTeam 
         Height          =   1695
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2415
         _Version        =   589884
         _ExtentX        =   4260
         _ExtentY        =   2990
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   900
      Left            =   5520
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
      WordWrap        =   -1  'True
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6585
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalTeam.frx":0C84
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15928
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
   Begin XtremeCommandBars.CommandBars cmbMain 
      Left            =   4800
      Top             =   960
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMedicalTeam.frx":1516
      Left            =   5520
      Top             =   960
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMedicalTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngMode As Long
Private mstrPrivs As String                              '权限串
Private mblnPauseTeam As Boolean                         '小组是否停用, True停用
Private mlngTeamID As Long                               '小组ID

Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Const FCONTROL = 8
Const VK_DELETE = &H2E
Const VK_F1 = &H70
Const VK_F5 = &H74

Const conPaneTeam = 1
Const conPaneMember = 2
Const conMenu_Label_Team = 9001
Const conMenu_Label_Member = 9002

Const conMenu_EditPopup = 3                     '编辑
Const conMenu_Edit_NewItem = 3001
Const conMenu_Edit_Modify = 3003
Const conMenu_Edit_Delete = 3004
Const conMenu_Edit_Pause = 3008
Const conMenu_Edit_Reuse = 3009
Const conMenu_Edit_CardBack = 3813
Const conMenu_Edit_CardCallBack = 3814
Const conMenu_Edit_Seat_Clear = 3534

Const conMenu_ViewPopup = 7                      '查看
Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Const conMenu_View_StatusBar = 702               '状态栏(&S)
Const conMenu_View_Refresh = 791                 '*刷新(&R)
Const conMenu_View_PauseTeam = 9003
Const conMenu_View_ToolBar = 701              '工具栏(&T)

Const conMenu_FilePopup = 1              '文件
Const conMenu_File_PrintSet = 101        '*打印设置(&S)…
Const conMenu_File_Preview = 102         '*预览(&V)
Const conMenu_File_Print = 103           '*打印(&P)
Const conMenu_File_Excel = 104           '输出到&Excel…
Const conMenu_File_Exit = 191            '*退出(&X)

Const conMenu_HelpPopup = 9                      '帮助
Const conMenu_Help_Help = 901                    '*帮助主题(&H)
Const conMenu_Help_Web = 902                     '&WEB上的中联
Const conMenu_Help_Web_Home = 9021               '中联主页(&H)
Const conMenu_Help_Web_Forum = 9023              '中联论坛(&F)
Const conMenu_Help_Web_Mail = 9022               '*发送反馈(&M)
Const conMenu_Help_About = 991                   '关于(&A)…


Private Enum mColTeam
    ID = 0: 停用: 科室: 小组名称: 说明
End Enum
Private Enum mColMember
    ID = 0: 姓名: 编号: 性别: 民族: 管理职务: 专业技术职务: 所属部门
End Enum

Private mbytTeamStatus As Byte
Private Property Let TeamStatus(ByVal bytVal As Byte)
'0-禁用状态, 1-启用状态, 2-停用状态
    Dim cbcTmp As CommandBarControl
    Dim blnTmp As Boolean
    
    Set cbcTmp = cmbMain.FindControl(, conMenu_EditPopup)
    For Each cbrControl In cbcTmp.CommandBar.Controls
        If cbrControl.ID = conMenu_Edit_Reuse Then
            If bytVal = 1 And InStr(mstrPrivs, "医疗小组编辑") > 0 Then
                cbrControl.Enabled = True
                cmbMain.FindControl(, conMenu_Edit_Reuse).Enabled = True
            Else
                cmbMain.FindControl(, conMenu_Edit_Reuse).Enabled = False
                cbrControl.Enabled = False
            End If
            Exit For
        End If
    Next
    For Each cbrControl In cbcTmp.CommandBar.Controls
        If cbrControl.ID = conMenu_Edit_Pause Then
            If bytVal = 2 And InStr(mstrPrivs, "医疗小组编辑") > 0 Then
                cbrControl.Enabled = True
                cmbMain.FindControl(, conMenu_Edit_Pause).Enabled = True
            Else
                cmbMain.FindControl(, conMenu_Edit_Pause).Enabled = False
                cbrControl.Enabled = False
            End If
            Exit For
        End If
    Next
    mbytTeamStatus = bytVal
End Property
Private Property Get TeamStatus() As Byte
    TeamStatus = mbytTeamStatus
End Property

Private Sub cmbMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim rpcRow As ReportRow
    Dim i As Long, lngRow As Long
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_Edit_NewItem
    '添加小组
        With frmMedicalTeamEdit
            .mstrPrivs = mstrPrivs
            .Status = 1
            If Not rpcTeam.FocusedRow Is Nothing Then
                For i = 0 To .cboDept.ListCount - 1
                    If .cboDept.ItemData(i) = rpcTeam.FocusedRow.Record.Item(0).Value Then
                        .cboDept.ListIndex = i
                        Exit For
                    End If
                Next
            End If
            .Show vbModal, Me
        End With
        Call RefreshToolbar
        RefreshRPCTeam mblnPauseTeam
        rpcTeam.SetFocus
    Case conMenu_Edit_Modify
    '编辑小组
        If rpcTeam.Rows.Count < 1 Then Exit Sub
        If rpcTeam.SelectedRows.Count < 1 Then Exit Sub
        strTmp = rpcTeam.FocusedRow.Record.Item(3).Value
        If InStr(strTmp, "【") > 0 Then Exit Sub
        With frmMedicalTeamEdit
            .mstrPrivs = mstrPrivs
            .Status = 2
            .TeamID = mlngTeamID
            For i = 0 To .cboDept.ListCount - 1
                If .cboDept.ItemData(i) = rpcTeam.FocusedRow.ParentRow.Record.Item(0).Value Then
                    .cboDept.ListIndex = i
                    Exit For
                End If
            Next
            .txtName = rpcTeam.FocusedRow.Record.Item(3).Value
            .txtExplain = rpcTeam.FocusedRow.Record.Item(4).Value
            .Show vbModal, Me
        End With
        Call RefreshToolbar
        RefreshRPCTeam mblnPauseTeam
        rpcTeam.SetFocus
    Case conMenu_Edit_Delete
    '删除小组
        If rpcTeam.Rows.Count < 1 Then Exit Sub
        If rpcTeam.SelectedRows.Count < 1 Then Exit Sub
        strTmp = rpcTeam.FocusedRow.Record.Item(3).Value
        If InStr(strTmp, "【") > 0 Then Exit Sub
        gstrSQL = "Select Count(*) rec From 医疗小组人员 Where 小组id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID)
        If rsTmp!rec = 0 Then
            rsTmp.Close
            If MsgBox("是否确认删除[" & rpcTeam.FocusedRow.Record.Item(3).Value & "]？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSQL = "Zl_临床医疗小组_DELETE(" & mlngTeamID & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                lngRow = rpcTeam.FocusedRow.Index
                RefreshRPCTeam mblnPauseTeam
                If lngRow >= rpcTeam.Rows.Count Then
                    If rpcTeam.Rows.Count > 0 Then
                        Set rpcTeam.FocusedRow = rpcTeam.Rows(rpcTeam.Rows.Count - 1)
                    End If
                Else
                    Set rpcTeam.FocusedRow = rpcTeam.Rows(lngRow)
                End If
                Call RefreshToolbar
            End If
        Else
            MsgBox "[" & rpcTeam.FocusedRow.Record.Item(3).Value & "]已经有成员应用！", vbInformation, gstrSysName
        End If
        rpcTeam.SetFocus
    Case conMenu_Edit_Reuse
    '启用
        Call TeamReusePause(True)
        Call RefreshToolbar
        rpcTeam.SetFocus
    Case conMenu_Edit_Pause
    '停用
        If rpcTeam.Records.Count < 1 Then Exit Sub
        Call TeamReusePause(False)
        Call RefreshToolbar
        rpcTeam.SetFocus
    Case conMenu_Edit_CardBack
    '增加成员
        If rpcTeam.Rows.Count < 1 Then Exit Sub
        For Each rpcRow In rpcTeam.Rows
            If Val(rpcRow.Record(2).Value) = mlngTeamID Then
                Set rpcTeam.FocusedRow = rpcRow
                Exit For
            End If
        Next
        strTmp = rpcTeam.FocusedRow.Record.Item(3).Value
        If InStr(strTmp, "【") > 0 Then Exit Sub
        With frmMedicalTeamMember
            .mstrPrivs = mstrPrivs
            .ShowMe Me, 1, rpcTeam.FocusedRow.ParentRow.Record.Item(0).Value, mlngTeamID
            If .mblnOK = True Then RefreshRPCMember mlngTeamID
        End With
        Call RefreshToolbar
        rpcMember.SetFocus
    Case conMenu_Edit_CardCallBack
    '转小组
        If rpcMember.SelectedRows.Count < 1 Then Exit Sub
        With frmMedicalTeamMember
            .mstrPrivs = mstrPrivs
            .ShowMe Me, 2, rpcTeam.FocusedRow.ParentRow.Record.Item(0).Value, mlngTeamID, rpcMember.FocusedRow.Record.Item(0).Value
            If .mblnOK = True Then RefreshRPCMember mlngTeamID
        End With
        Call RefreshToolbar
        rpcMember.SetFocus
    Case conMenu_Edit_Seat_Clear
    '移除
        If rpcMember.SelectedRows.Count < 1 Then Exit Sub
        If MsgBox("要移除【" & rpcMember.FocusedRow.Record.Item(2).Value & "】" & rpcMember.FocusedRow.Record.Item(1).Value & " 吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Dim strMess As String
            strMess = MedicalTeamPatients(mlngTeamID, Val(rpcMember.FocusedRow.Record.Item(0).Value))
            If strMess = "" Then
                gstrSQL = "select count(*) rec from 医疗小组人员 where 小组id=[1] and 人员id=[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, rpcMember.FocusedRow.Record.Item(0).Value)
                If rsTmp!rec = 0 Then
                    MsgBox "该医生已经被其他用户移除！", vbInformation, gstrSysName
                Else
                    gstrSQL = "Zl_医疗小组人员_Delete(" & mlngTeamID & "," & rpcMember.FocusedRow.Record.Item(0).Value & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
                lngRow = rpcMember.FocusedRow.Index
                RefreshRPCMember mlngTeamID
                If lngRow >= rpcMember.Rows.Count Then
                    If rpcMember.Rows.Count > 0 Then
                        Set rpcMember.FocusedRow = rpcMember.Rows(rpcMember.Rows.Count - 1)
                    End If
                Else
                    Set rpcMember.FocusedRow = rpcMember.Rows(lngRow)
                End If
            Else
                MsgBox "该医生当前有以下在院病人，" & vbNewLine & vbNewLine & strMess & vbNewLine & "不允许移除！", vbInformation, gstrSysName
            End If
        End If
        Call RefreshToolbar
        rpcMember.SetFocus
    Case conMenu_View_ToolBar_Button
        Control.Checked = Not Control.Checked
        Me.cmbMain(2).Visible = Control.Checked
        Me.cmbMain.RecalcLayout
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not Control.Checked
        For Each cbrControl In Me.cmbMain(2).Controls
            If cbrControl.Type = xtpControlButton Then
                cbrControl.Style = IIF(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            End If
        Next
        Me.cmbMain.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cmbMain.Options.LargeIcons = Not Me.cmbMain.Options.LargeIcons
        Me.cmbMain.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cmbMain.RecalcLayout
    Case conMenu_View_PauseTeam
    '控制显示停用小组
        Control.Checked = Not Control.Checked
        mblnPauseTeam = Control.Checked
        '更新rpcTeam控件
        RefreshRPCTeam mblnPauseTeam
        Call RefreshToolbar
    Case conMenu_View_Refresh
        RefreshRPCTeam mblnPauseTeam
        Call rpcTeam_SelectionChanged
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
    If Control.ID > 401 And Control.ID < 499 Then
        '执行自定义报表
        Call BillPrint_Custom(Control)
    End If
    End Select
    Exit Sub

errHandle:
    Call ERRCENTER
    Call SaveErrLog
End Sub

Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '打印自定义报表
    
    '默认参数：小组=小组id
    Dim strName As String
    
    strName = Split(Control.Parameter, ",")(1)
    
    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strName, Me, _
        "小组=" & IIF(mlngTeamID = 0, "", mlngTeamID))
End Sub
Private Sub cmbMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cmbMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
'    If Control.Type = xtpBarTypePopup Then
'        Select Case Control.Index
'        Case conMenu_EditPopup: Control.Visible = True
'        End Select
'    End If
'
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rpcMember.Records.Count <> 0) ' And mintEditState = 0)
'    Case conMenu_Edit_Save, conMenu_Edit_Untread
'        Control.Enabled = (mintEditState <> 0)
'    Case conMenu_Edit_NewItem
'        Control.Enabled = (InStr(1, mstrPrivs, "定义") > 0 And mintEditState = 0)
'    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_ApplyTo
'        Control.Enabled = (InStr(1, mstrPrivs, "定义") > 0 And mintEditState = 0)
'        If Control.Enabled Then Control.Enabled = mlngBillID <> 0
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cmbMain(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cmbMain(2).Controls(2).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cmbMain.Options.LargeIcons
    Case conMenu_View_StatusBar:      Control.Checked = Me.stbThis.Visible
    'Case conMenu_View_PauseTeam:      Control.Checked = Not Control.Checked
'    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case conPaneTeam: Item.Handle = picTeam.hwnd
        Case conPaneMember: Item.Handle = picMember.hwnd
    End Select
End Sub

Private Sub Form_Load()
    Dim cbcTmp As CommandBarControl
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    Call InitMenu
    Call InitToolBar
    Call InitDKP
    Call InitReportControl
    
    '添加自定义报表
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    RestoreWinState Me, App.ProductName
    
    '本地参数
    Set cbcTmp = cmbMain.FindControl(, conMenu_ViewPopup)
    For Each cbrControl In cbcTmp.CommandBar.Controls
        If cbrControl.Type = xtpControlButton Then
            If cbrControl.ID = conMenu_View_PauseTeam Then
                cbrControl.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示停用", 0)) = 1)
                mblnPauseTeam = cbrControl.Checked
                Exit For
            End If
        End If
    Next
    
    RefreshRPCTeam mblnPauseTeam
    Call RefreshToolbar
    
End Sub

Private Sub InitMenu()
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cmbMain.VisualTheme = xtpThemeOffice2003
    Set cmbMain.Icons = zlCommFun.GetPubIcons
    With cmbMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cmbMain.EnableCustomization False

    '-----------------------------------------------------
    '菜单定义
    Me.cmbMain.ActiveMenuBar.Title = "菜单"
    Me.cmbMain.ActiveMenuBar.EnableDocking xtpFlagAlignTop
    Set cbrMenuBar = Me.cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        If InStr(mstrPrivs, "医疗小组编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        If InStr(mstrPrivs, "医疗小组编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        If InStr(mstrPrivs, "医疗小组编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用(&S)"): cbrControl.BeginGroup = True
        If InStr(mstrPrivs, "医疗小组编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Pause, "停用(&T)")
        If InStr(mstrPrivs, "医疗小组编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "添加(&N)"): cbrControl.BeginGroup = True
        If InStr(mstrPrivs, "组成人员编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardCallBack, "转小组(&R)")
        If InStr(mstrPrivs, "组成人员编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Seat_Clear, "移除(&E)")
        If InStr(mstrPrivs, "组成人员编辑") = 0 Then cbrControl.Enabled = False
        'Set cbrControl = .Add(xtpControlButton, xxx, "设为小组组长(&O)")
        'If InStr(mstrPrivs, "组成人员编辑") = 0 Then cbrControl.Enabled = False
    End With

    Set cbrMenuBar = Me.cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        Dim cbrChild As CommandBarControl
        Set cbrChild = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
        Set cbrChild = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
        Set cbrChild = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_PauseTeam, "显示停用小组(&V)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    End With
    
    Set cbrMenuBar = Me.cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的中联")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cmbMain.KeyBindings
        '.Add FCONTROL, Asc("S"), conMenu_Edit_Reuse
        '.Add FCONTROL, Asc("T"), conMenu_Edit_Pause
        .Add FCONTROL, Asc("A"), conMenu_Edit_CardBack
        .Add FCONTROL, VK_DELETE, conMenu_Edit_Seat_Clear
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cmbMain.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
End Sub

Private Sub InitToolBar()
    Set cbrToolBar = Me.cmbMain.Add("小组工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    'cbrToolBar.EnableDocking xtpFlagAlignTop
    With cbrToolBar.Controls
        'Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        'Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlLabel, conMenu_Label_Team, "小组：")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增")
        If InStr(mstrPrivs, "医疗小组编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        If InStr(mstrPrivs, "医疗小组编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        If InStr(mstrPrivs, "医疗小组编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用"): cbrControl.BeginGroup = True
        If InStr(mstrPrivs, "医疗小组编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Pause, "停用")
        If InStr(mstrPrivs, "医疗小组编辑") = 0 Then cbrControl.Enabled = False
        
        Set cbrControl = .Add(xtpControlLabel, conMenu_Label_Member, "成员："): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "添加")
        If InStr(mstrPrivs, "组成人员编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardCallBack, "转小组")
        If InStr(mstrPrivs, "组成人员编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Seat_Clear, "移除")
        If InStr(mstrPrivs, "组成人员编辑") = 0 Then cbrControl.Enabled = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub InitDKP()
    Dim panTeam As Pane, panMember As Pane
    
    With dkpMain
        Set panMember = .CreatePane(conPaneMember, 500, 100, DockLeftOf)
        panMember.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        panMember.Title = "小组成员列表"

        Set panTeam = .CreatePane(conPaneTeam, 300, 100, DockLeftOf)
        panTeam.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        panTeam.Title = "小组列表"

        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        '.Options.ThemedFloatingFrames = True
        If Not cmbMain Is Nothing Then .SetCommandBars Me.cmbMain
    End With
End Sub

Private Sub InitReportControl()
    Dim rpcCol As ReportColumn
    With Me.rpcTeam
        Set rpcCol = .Columns.Add(mColTeam.ID, "ID", 0, False)
        rpcCol.Visible = False
        Set rpcCol = .Columns.Add(mColTeam.停用, "停用", 0, False)
        rpcCol.Visible = False
        Set rpcCol = .Columns.Add(mColTeam.科室, "科室", 150, False)
        rpcCol.Visible = False
        Set rpcCol = .Columns.Add(mColTeam.小组名称, "小组名称", 150, False)
        rpcCol.TreeColumn = True
        Set rpcCol = .Columns.Add(mColTeam.说明, "说明", 300, False)
        'rpcCol.AutoSize = True
        For Each rpcCol In .Columns
            rpcCol.Editable = False
            rpcCol.Groupable = False
            rpcCol.Sortable = False
            rpcCol.Resizable = True
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = rpcTeam.PaintManager.BackColor
            .NoItemsText = "没有可显示的小组..."
            .VerticalGridStyle = xtpGridSolid
        End With
        .PreviewMode = False
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .AllowColumnResize = True
        .SetImageList imgList
    End With
    
    With Me.rpcMember
        Set rpcCol = .Columns.Add(mColMember.ID, "ID", 0, False)
        Set rpcCol = .Columns.Add(mColMember.姓名, "姓名", 80, False)
        Set rpcCol = .Columns.Add(mColMember.编号, "编号", 60, False)
        Set rpcCol = .Columns.Add(mColMember.性别, "性别", 30, False)
        Set rpcCol = .Columns.Add(mColMember.民族, "民族", 30, False)
        Set rpcCol = .Columns.Add(mColMember.管理职务, "管理职务", 80, False)
        Set rpcCol = .Columns.Add(mColMember.专业技术职务, "专业技术职务", 90, False)
        Set rpcCol = .Columns.Add(mColMember.所属部门, "所属部门", 500, False)
        'rpcCol.AutoSize = True
        For Each rpcCol In .Columns
            rpcCol.Editable = False
            rpcCol.Groupable = False
            rpcCol.Resizable = True
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoItemsText = "没有可显示的成员..."
            .VerticalGridStyle = xtpGridSolid
        End With
        .PreviewMode = False
        .AutoColumnSizing = False
        .AllowColumnRemove = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim cbrTmp As CommandBarControl
    Set cbrTmp = cmbMain.FindControl(, conMenu_ViewPopup)
    If cbrTmp Is Nothing Then Exit Sub
    
    SaveWinState Me, App.ProductName
    
    For Each cbrControl In cbrTmp.CommandBar.Controls
        If cbrControl.Type = xtpControlButton Then
            If cbrControl.ID = conMenu_View_PauseTeam Then
                SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示停用", IIF(cbrControl.Checked, 1, 0)
                Exit For
            End If
        End If
    Next
    mlngTeamID = 0
End Sub

Private Sub picMember_Resize()
    With rpcMember
        .Top = 0
        .Left = 0
        .Width = picMember.Width
        .Height = picMember.Height
    End With
End Sub

Private Sub picTeam_Resize()
    With rpcTeam
        .Top = 0
        .Left = 0
        .Width = picTeam.Width
        .Height = picTeam.Height
    End With
End Sub

Private Sub FillReportControl(ByVal rsData As ADODB.Recordset, ByVal bytNO As Byte)
    Dim rpcVal As ReportControl
    Dim rpcRec As ReportRecord, rpcRecChild As ReportRecord
    Dim rpcRecItem As ReportRecordItem
    Dim lngId As Long
    Dim strDept As String, strNo As String
    
    If bytNO = 0 Then
        Set rpcVal = rpcTeam
    Else
        Set rpcVal = rpcMember
    End If
    
    rpcVal.Records.DeleteAll
    rpcVal.Populate
    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub
    
    With rsData
        .MoveFirst
        Do While Not .EOF
            If bytNO = 0 Then
                If lngId <> !ID Then
                    Set rpcRec = rpcVal.Records.Add
                    Set rpcRecItem = rpcRec.AddItem(CStr(!ID))
                    rpcRec.AddItem CStr(!ID)
                    rpcRec.AddItem ""
                    rpcRec.AddItem CStr("【" & !编码 & "】" & !科室名称)
                    rpcRec.AddItem ""
                End If
                
                Set rpcRecChild = rpcRec.Childs.Add
                Set rpcRecItem = rpcRecChild.AddItem(CStr(!ID))
                rpcRecChild.AddItem CStr("" & !停用)
                rpcRecChild.AddItem CStr("" & !小组ID)
                'rpcRecChild.AddItem CStr("" & !名称)
                Set rpcRecItem = rpcRecChild.AddItem(CStr("" & !名称))
                If !停用 = 1 Then
                    rpcRecItem.ForeColor = vbRed
                    rpcRecItem.Icon = 0
                End If
                rpcRecChild.AddItem CStr("" & !说明)
                rpcRec.Expanded = True
                lngId = !ID
            Else
                strNo = !编号
                .MoveNext
                If .EOF Then
                    .MovePrevious
                    strDept = strDept & CStr("" & !所属部门) & ";"
                    Set rpcRec = rpcVal.Records.Add
                    Set rpcRecItem = rpcRec.AddItem(CStr(!ID))
                    rpcRec.AddItem CStr("" & !姓名)
                    rpcRec.AddItem CStr("" & !编号)
                    rpcRec.AddItem CStr("" & !性别)
                    rpcRec.AddItem CStr("" & !民族)
                    rpcRec.AddItem CStr("" & !管理职务)
                    rpcRec.AddItem CStr("" & !专业技术职务)
                    rpcRec.AddItem CStr(Left(strDept, Len(strDept) - 1))
                    strDept = ""
                ElseIf strNo <> !编号 Then
                    .MovePrevious
                    strDept = strDept & CStr("" & !所属部门) & ";"
                    Set rpcRec = rpcVal.Records.Add
                    Set rpcRecItem = rpcRec.AddItem(CStr(!ID))
                    rpcRec.AddItem CStr("" & !姓名)
                    rpcRec.AddItem CStr("" & !编号)
                    rpcRec.AddItem CStr("" & !性别)
                    rpcRec.AddItem CStr("" & !民族)
                    rpcRec.AddItem CStr("" & !管理职务)
                    rpcRec.AddItem CStr("" & !专业技术职务)
                    rpcRec.AddItem CStr(Left(strDept, Len(strDept) - 1))
                    strDept = ""
                Else
                    .MovePrevious
                    strDept = strDept & CStr("" & !所属部门) & ";"
                End If
                
            End If
            '工具栏控制
            .MoveNext
        Loop
    End With
    rpcVal.Populate
    
    If rpcVal.Rows.Count >= 1 Then
        rpcVal.SelectedRows.Add rpcVal.Rows(0)
        Set rpcVal.FocusedRow = rpcVal.Rows(0)
    End If
    
End Sub

Private Sub rpcMember_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim cbpPopupMenu As CommandBarPopup
    Dim cbcControl As CommandBarControl
    If Button = vbRightButton Then
        Set cbpPopupMenu = cmbMain.FindControl(, conMenu_EditPopup)
        If Not cbpPopupMenu Is Nothing Then
            For Each cbcControl In cbpPopupMenu.CommandBar.Controls
                If cbcControl.ID = conMenu_Edit_CardBack Or cbcControl.ID = conMenu_Edit_CardCallBack _
                Or cbcControl.ID = conMenu_Edit_Seat_Clear Then
                    cbcControl.Visible = True
                Else
                    cbcControl.Visible = False
                End If
            Next
            cbpPopupMenu.CommandBar.ShowPopup
            For Each cbcControl In cbpPopupMenu.CommandBar.Controls
                cbcControl.Visible = True
            Next
        End If
    End If
End Sub

Private Sub rpcMember_SelectionChanged()
    Call RefreshToolbar
End Sub

Private Sub rpcTeam_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
'    If Row.Record.Item(1).Value = 1 And InStr(Row.Record.Item(2).Value, "【") = 0 Then
'        Dim fntTmp As New StdFont
'        fntTmp.Strikethrough = True
'        Set Metrics.Font = fntTmp
'    End If
End Sub

Private Sub rpcTeam_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim cbpPopupMenu As CommandBarPopup
    Dim cbcControl As CommandBarControl
    If Button = vbRightButton Then
        Set cbpPopupMenu = cmbMain.FindControl(, conMenu_EditPopup)
        If Not cbpPopupMenu Is Nothing Then
            For Each cbcControl In cbpPopupMenu.CommandBar.Controls
                cbcControl.Visible = False
                If cbcControl.ID = conMenu_Edit_NewItem Or cbcControl.ID = conMenu_Edit_Modify _
                Or cbcControl.ID = conMenu_Edit_Delete Or cbcControl.ID = conMenu_Edit_Reuse _
                Or cbcControl.ID = conMenu_Edit_Pause Then
                    cbcControl.Visible = True
                End If
            Next
            cbpPopupMenu.CommandBar.ShowPopup
            For Each cbcControl In cbpPopupMenu.CommandBar.Controls
                cbcControl.Visible = True
            Next
        End If
    End If
End Sub

Private Sub rpcTeam_SelectionChanged()
    Dim recCurrent As ReportRecord
    Dim cmcTmp As CommandBarControl
    
    If rpcTeam.FocusedRow Is Nothing Then Exit Sub
    Set recCurrent = rpcTeam.FocusedRow.Record
    
    '医疗小组未发生选择变化，不刷新医疗小组成员列表。
    If Val(rpcTeam.Tag) = recCurrent.Item(2).Value Then Exit Sub
    
    If InStr(recCurrent.Item(3).Value, "【") = 0 Then
        mlngTeamID = Val(recCurrent.Item(2).Value)
        RefreshRPCMember mlngTeamID
        If Val(recCurrent.Item(1).Value) = 0 Then
            '停用
            TeamStatus = 2
        Else
            TeamStatus = 1
        End If
        stbThis.Panels(2).Text = "本医疗小组有" & rpcMember.Records.Count & "位小组成员"
    Else
        rpcMember.Records.DeleteAll
        rpcMember.Populate
        TeamStatus = 0
        mlngTeamID = 0
        stbThis.Panels(2).Text = "本科室有" & recCurrent.Childs.Count & "个医疗小组"
    End If
    
    rpcTeam.Tag = mlngTeamID
    
    Call RefreshToolbar
End Sub

Private Sub RefreshRPCTeam(ByVal blnPauseTeam As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim cbcTmp As CommandBarControl
    Dim rpcRow As ReportRow
    
    On Error GoTo errHandle
    strTmp = "Select A.ID, A.名称 科室名称, A.编码, B.ID 小组id, B.名称, B.说明, Case When B.撤档时间 >= To_Date('3000-1-1', 'yyyy-mm-dd') Then 0 Else 1 End 停用 " & vbNewLine & _
             "From 部门表 A, 临床医疗小组 B, 部门性质说明 C " & _
             IIF(InStr(mstrPrivs, "所有科室") = 0, ", 部门人员 D ", "") & vbNewLine & _
             "Where A.ID = B.科室id And B.科室id = C.部门id And substr(B.名称,1,1)<>'-' And C.工作性质 = '临床' " & _
             "  and C.服务对象 in (2,3) And A.撤档时间 >= To_Date('3000-1-1', 'yyyy-mm-dd') " & vbNewLine & _
             IIF(InStr(mstrPrivs, "所有科室") = 0, " And A.ID=D.部门ID and D.人员ID=[1] ", "")
    '不显示停用小组
    If Not blnPauseTeam Then
        strTmp = strTmp & " And B.撤档时间 >= To_Date('3000-1-1', 'yyyy-mm-dd') " & vbNewLine
    End If
    strTmp = strTmp & "Order By A.编码, B.名称 "
    '所有科室
    If InStr(mstrPrivs, "所有科室") = 0 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, glngUserId)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption)
    End If
    
    FillReportControl rsTmp, 0
    
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        If mlngTeamID = 0 Then mlngTeamID = rsTmp!小组ID
        For Each rpcRow In rpcTeam.Rows
            If Val(rpcRow.Record(2).Value) = mlngTeamID Then
                Set rpcTeam.FocusedRow = rpcRow
                Exit For
            End If
        Next
        If rpcTeam.FocusedRow Is Nothing Then Exit Sub
        RefreshRPCMember rpcTeam.FocusedRow.Record.Item(2).Value 'rpcTeam.Rows(1).Record.Item(1).Value
        If Val(rpcTeam.FocusedRow.Record.Item(1).Value) = 0 Then
            TeamStatus = 2
        Else
            TeamStatus = 1
        End If
        stbThis.Panels(2).Text = "本医疗小组有" & rpcMember.Records.Count & "位小组成员"
    End If
    Exit Sub
errHandle:
    If ERRCENTER() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub RefreshRPCMember(ByVal lngTeamID As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
'    strTmp = "Select A.ID, A.姓名, A.编号, A.性别, A.民族, B.所属部门, A.管理职务, A.专业技术职务" & vbNewLine & _
'             "From 人员表 A," & vbNewLine & _
'             "     (Select A.人员id, Wmsys.Wm_Concat(B.名称) 所属部门" & vbNewLine & _
'             "       From 部门人员 A, 部门表 B" & vbNewLine & _
'             "       Where A.部门id = B.ID And B.撤档时间 >= To_Date('3000/01/01', 'yyyy-mm-dd')" & vbNewLine & _
'             "       Group By A.人员id) B, 医疗小组人员 C" & vbNewLine & _
'             "Where A.ID = B.人员id And A.ID = C.人员id and C.小组ID=[1]"
    On Error GoTo errHandle
    strTmp = "Select A.ID, A.姓名, A.编号, A.性别, A.民族, C.名称 所属部门, A.管理职务, A.专业技术职务" & vbNewLine & _
             "From 人员表 A, 部门人员 B, 部门表 C, 医疗小组人员 D" & vbNewLine & _
             "Where A.ID = B.人员id And B.部门id = C.ID And A.ID = D.人员id And D.小组ID=[1]" & vbNewLine & _
             " And C.撤档时间 >= To_Date('3000/01/01', 'yyyy-mm-dd')" & vbNewLine & _
             "Order By A.ID, C.名称"
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, lngTeamID)
    FillReportControl rsTmp, 1
    Exit Sub
errHandle:
    If ERRCENTER() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rpcMember.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vsfPrint, Me.rpcMember) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vsfPrint
    objPrint.Title.Text = "目录"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        
        '标题行复制
        .Row = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = rptCol.Width * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol
        
        '数据行复制
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    zlReportToVSFlexGrid = False
End Function

Private Function TeamReusePause(ByVal blnReuse As Boolean) As Boolean
    Dim lngRow As Long
    Dim rpcRow As ReportRow
    
    On Error GoTo errHandle
    For Each rpcRow In rpcTeam.Rows
        If Val(rpcRow.Record(2).Value) = mlngTeamID Then
            Set rpcTeam.FocusedRow = rpcRow
            Exit For
        End If
    Next
    If rpcTeam.FocusedRow Is Nothing Then
        MsgBox "请选择医疗小组！", vbInformation, gstrSysName
        Exit Function
    End If
    lngRow = rpcTeam.FocusedRow.Index
    If blnReuse Then
        Call zlDatabase.ExecuteProcedure("zl_临床医疗小组_Reuse(" & rpcTeam.FocusedRow.Record.Item(2).Value & ")", Me.Caption)
    Else
        If rpcMember.Records.Count > 0 Then
            MsgBox "[" & rpcTeam.FocusedRow.Record.Item(3).Value & "]存在小组成员，不允许停用操作！", vbInformation, gstrSysName
            Exit Function
        End If
        Call zlDatabase.ExecuteProcedure("zl_临床医疗小组_Stop(" & rpcTeam.FocusedRow.Record.Item(2).Value & ")", Me.Caption)
    End If
    RefreshRPCTeam mblnPauseTeam
    
    If lngRow >= rpcTeam.Rows.Count Then
        If rpcTeam.Rows.Count > 0 Then
            Set rpcTeam.FocusedRow = rpcTeam.Rows(rpcTeam.Rows.Count - 1)
        End If
    Else
        Set rpcTeam.FocusedRow = rpcTeam.Rows(lngRow)
    End If

    TeamReusePause = True
    Exit Function
errHandle:
    If ERRCENTER() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RefreshToolbar()
    Dim cbpPopupMenu As CommandBarPopup
    Dim cbcControl As CommandBarControl
    
    Set cbpPopupMenu = cmbMain.FindControl(, conMenu_EditPopup)
    
    cmbMain.FindControl(, conMenu_Edit_NewItem).Enabled = InStr(mstrPrivs, "医疗小组编辑") <> 0
    If rpcTeam.FocusedRow Is Nothing Then
        For Each cbcControl In cbpPopupMenu.CommandBar.Controls
            Select Case cbcControl.ID
                Case conMenu_Edit_NewItem
                    cbcControl.Enabled = InStr(mstrPrivs, "医疗小组编辑") <> 0
                Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_CardBack, conMenu_Edit_CardCallBack, conMenu_Edit_Seat_Clear _
                    , conMenu_Edit_Reuse, conMenu_Edit_Pause
                    cbcControl.Enabled = False
            End Select
        Next
        cmbMain.FindControl(, conMenu_Edit_Modify).Enabled = False
        cmbMain.FindControl(, conMenu_Edit_Delete).Enabled = False
        cmbMain.FindControl(, conMenu_Edit_Reuse).Enabled = False
        cmbMain.FindControl(, conMenu_Edit_Pause).Enabled = False
        cmbMain.FindControl(, conMenu_Edit_CardBack).Enabled = False
        cmbMain.FindControl(, conMenu_Edit_CardCallBack).Enabled = False
        cmbMain.FindControl(, conMenu_Edit_Seat_Clear).Enabled = False
    ElseIf InStr(rpcTeam.FocusedRow.Record.Item(3).Value, "【") = 0 Then
        For Each cbcControl In cbpPopupMenu.CommandBar.Controls
            Select Case cbcControl.ID
                Case conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete
                    cbcControl.Enabled = True And rpcTeam.FocusedRow.Record.Item(1).Value = "0" And InStr(mstrPrivs, "医疗小组编辑") <> 0
                Case conMenu_Edit_CardBack
                    cbcControl.Enabled = True And rpcTeam.FocusedRow.Record.Item(1).Value = "0" And InStr(mstrPrivs, "组成人员编辑") <> 0
                Case conMenu_Edit_CardCallBack, conMenu_Edit_Seat_Clear
                    cbcControl.Enabled = rpcMember.Rows.Count > 0 And rpcTeam.FocusedRow.Record.Item(1).Value = "0" And InStr(mstrPrivs, "组成人员编辑") <> 0
            End Select
        Next
        cmbMain.FindControl(, conMenu_Edit_Modify).Enabled = True And rpcTeam.FocusedRow.Record.Item(1).Value = "0" And InStr(mstrPrivs, "医疗小组编辑") <> 0
        cmbMain.FindControl(, conMenu_Edit_Delete).Enabled = True And rpcTeam.FocusedRow.Record.Item(1).Value = "0" And InStr(mstrPrivs, "医疗小组编辑") <> 0
        cmbMain.FindControl(, conMenu_Edit_CardBack).Enabled = True And rpcTeam.FocusedRow.Record.Item(1).Value = "0" And InStr(mstrPrivs, "组成人员编辑") <> 0
        cmbMain.FindControl(, conMenu_Edit_CardCallBack).Enabled = rpcMember.Rows.Count > 0 And rpcTeam.FocusedRow.Record.Item(1).Value = "0" And InStr(mstrPrivs, "组成人员编辑") <> 0
        cmbMain.FindControl(, conMenu_Edit_Seat_Clear).Enabled = rpcMember.Rows.Count > 0 And rpcTeam.FocusedRow.Record.Item(1).Value = "0" And InStr(mstrPrivs, "组成人员编辑") <> 0
    Else
        For Each cbcControl In cbpPopupMenu.CommandBar.Controls
            Select Case cbcControl.ID
                Case conMenu_Edit_NewItem
                    cbcControl.Enabled = True And InStr(mstrPrivs, "医疗小组编辑") <> 0
                Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_CardBack, conMenu_Edit_CardCallBack, conMenu_Edit_Seat_Clear
                    cbcControl.Enabled = False
            End Select
        Next
        cmbMain.FindControl(, conMenu_Edit_Modify).Enabled = False
        cmbMain.FindControl(, conMenu_Edit_Delete).Enabled = False
        cmbMain.FindControl(, conMenu_Edit_CardBack).Enabled = False
        cmbMain.FindControl(, conMenu_Edit_CardCallBack).Enabled = False
        cmbMain.FindControl(, conMenu_Edit_Seat_Clear).Enabled = False
    End If

End Sub
