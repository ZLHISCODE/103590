VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmProcEdit 
   Caption         =   "编辑过程"
   ClientHeight    =   8016
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13128
   Icon            =   "frmProcEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8016
   ScaleWidth      =   13128
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   1
      Left            =   1680
      ScaleHeight     =   4056
      ScaleWidth      =   9648
      TabIndex        =   1
      Top             =   3960
      Width           =   9645
      Begin VB.PictureBox picEdit 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   3645
         Left            =   1200
         ScaleHeight     =   3648
         ScaleWidth      =   8616
         TabIndex        =   5
         Top             =   360
         Width           =   8610
         Begin XtremeSyntaxEdit.SyntaxEdit synProcEdit 
            Height          =   2295
            Left            =   360
            TabIndex        =   8
            Top             =   120
            Width           =   2445
            _Version        =   983043
            _ExtentX        =   4313
            _ExtentY        =   4048
            _StockProps     =   84
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            EnableSyntaxColorization=   -1  'True
            ShowLineNumbers =   -1  'True
            ShowSelectionMargin=   -1  'True
            ShowScrollBarVert=   -1  'True
            ShowScrollBarHorz=   -1  'True
            EnableVirtualSpace=   0   'False
            EnableAutoIndent=   -1  'True
            ShowWhiteSpace  =   0   'False
            ShowCollapsibleNodes=   -1  'True
            AutoCompleteWndWidth=   160
         End
         Begin VB.PictureBox picBase 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   3615
            Left            =   3120
            ScaleHeight     =   3588
            ScaleWidth      =   5124
            TabIndex        =   9
            Top             =   0
            Width           =   5145
            Begin VB.TextBox txtLocRow 
               Height          =   300
               Left            =   1125
               TabIndex        =   11
               Top             =   105
               Width           =   1530
            End
            Begin VB.CommandButton cmdProcName 
               Height          =   275
               Left            =   4630
               Picture         =   "frmProcEdit.frx":6852
               Style           =   1  'Graphical
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   898
               Width           =   275
            End
            Begin VB.ComboBox cboProcType 
               Height          =   300
               Left            =   1125
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   495
               Width           =   1530
            End
            Begin VB.ComboBox cboOwner 
               Height          =   300
               Left            =   3525
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   495
               Width           =   1380
            End
            Begin VB.TextBox txtNote 
               Height          =   1260
               Left            =   1125
               MultiLine       =   -1  'True
               TabIndex        =   20
               Top             =   1275
               Width           =   3780
            End
            Begin VB.ComboBox cboProcName 
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   1125
               TabIndex        =   17
               Text            =   "cboProcName"
               Top             =   885
               Width           =   3780
            End
            Begin VB.Label lblNotic 
               AutoSize        =   -1  'True
               Caption         =   "说明：过程编辑区支持快捷键CTRL+A(全选)、CTRL+Z(撤销)、CTRL+C(复制)、CTRL+V(粘贴)、CTRL+F(查找)、CTRL+H(替换)、CTRL+G(定位行)"
               ForeColor       =   &H002222B2&
               Height          =   540
               Left            =   300
               TabIndex        =   22
               Top             =   2760
               Width           =   4815
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblLocRow 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "定位行号(&G)"
               Height          =   180
               Left            =   120
               TabIndex        =   10
               Top             =   165
               Width           =   990
            End
            Begin VB.Label lblProcType 
               AutoSize        =   -1  'True
               Caption         =   "过程类型(&T)"
               Height          =   180
               Left            =   120
               TabIndex        =   12
               Top             =   555
               Width           =   990
            End
            Begin VB.Label lblOwner 
               AutoSize        =   -1  'True
               Caption         =   "所有者(&O)"
               Height          =   180
               Left            =   2685
               TabIndex        =   14
               Top             =   555
               Width           =   810
            End
            Begin VB.Label lblNote 
               AutoSize        =   -1  'True
               Caption         =   "过程说明(&P)"
               Height          =   180
               Left            =   120
               TabIndex        =   19
               Top             =   1335
               Width           =   990
            End
            Begin VB.Label lblProcName 
               AutoSize        =   -1  'True
               Caption         =   "过程名称(&N)"
               Height          =   180
               Left            =   120
               TabIndex        =   16
               Top             =   945
               Width           =   990
            End
         End
      End
      Begin XtremeSuiteControls.TabControl TbcBase 
         Height          =   975
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   120
         Width           =   1935
         _Version        =   589884
         _ExtentX        =   3413
         _ExtentY        =   1720
         _StockProps     =   64
      End
   End
   Begin VB.Frame fraHSplit 
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   21
      Top             =   4800
      Width           =   9615
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2970
      Index           =   0
      Left            =   240
      ScaleHeight     =   2976
      ScaleWidth      =   9696
      TabIndex        =   0
      Top             =   1560
      Width           =   9690
      Begin VB.PictureBox picLast 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   2280
         ScaleHeight     =   2592
         ScaleWidth      =   3576
         TabIndex        =   4
         Top             =   240
         Width           =   3570
         Begin XtremeSyntaxEdit.SyntaxEdit synLastProc 
            Height          =   2295
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   2445
            _Version        =   983043
            _ExtentX        =   4313
            _ExtentY        =   4048
            _StockProps     =   84
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ReadOnly        =   -1  'True
            EnableSyntaxColorization=   -1  'True
            ShowLineNumbers =   -1  'True
            ShowSelectionMargin=   -1  'True
            ShowScrollBarVert=   -1  'True
            ShowScrollBarHorz=   -1  'True
            EnableVirtualSpace=   0   'False
            EnableAutoIndent=   -1  'True
            ShowWhiteSpace  =   0   'False
            ShowCollapsibleNodes=   -1  'True
            AutoCompleteWndWidth=   160
         End
         Begin SHDocVwCtl.WebBrowser wbrCompare 
            Height          =   1935
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   1935
            ExtentX         =   3413
            ExtentY         =   3413
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
      End
      Begin XtremeSuiteControls.TabControl TbcBase 
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   3413
         _StockProps     =   64
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   7320
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmProcEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==模块变量
'==============================================================
Private mobjMain As Object '父窗体
Private mlngKey As Long '存储过程ID
Private mblnOk As Boolean '是否确定退出
Private mptType As ProcType '存储过程类型
Private mblnChange As Boolean '是否数据修改
Private mintState As Integer '存储过程状态
Private mrsProcedure As ADODB.Recordset '过程名清单
Private mstrProcName As String '存储过程名称
Private mblnLoad As Boolean
Private Enum PaneEnum
    PE_历史变动 = 0
    PE_当前过程 = 1
End Enum
Private mfrmProcedureOwnerCon As frmProcOwnerConn



'==============================================================
'==公共接口
'==============================================================
Public Function ShowMe(ByVal objMain As Object, ByVal lngKey As Long, Optional ByVal ptType As ProcType) As Boolean
'参数：ptType=存储过程管理主界面，选择的过程类型
    Set mobjMain = objMain
    mlngKey = lngKey
    mptType = ptType
    mblnOk = False
    mblnLoad = False
    Me.Show 1, objMain
    ShowMe = mblnOk
End Function

'==============================================================
'==控件事件
'==============================================================
Private Sub cboProcName_Click()
    Dim strOwner As String
    If Trim(cboProcName.Text) <> "" And cboProcName.Tag = "" Then
        synProcEdit.Text = gclsBase.GetProgram(Trim(cboProcName.Text), strOwner)
        If strOwner <> "" Then
            cboOwner.Text = strOwner
        End If
    End If
End Sub

Private Sub cboProcName_KeyPress(KeyAscii As Integer)
    If mptType <> ProcType.用户过程 And mlngKey = 0 Then
        Call SendMessage(cboProcName.hwnd, CB_SHOWDROPDOWN, 1, 0)
    End If
End Sub

Private Sub cboProcType_Click()
    Select Case cboProcType.ItemData(cboProcType.ListIndex)
        Case ProcType.变动过程, ProcType.空白过程
            If mlngKey = 0 Then
                lblOwner.Visible = False: cboOwner.Visible = False
                LoadProcNames
            Else
                cboOwner.Locked = True: cboProcName.Locked = True
            End If
        Case ProcType.用户过程
            If mlngKey = 0 Then
                lblOwner.Visible = True: cboOwner.Visible = True
                If mptType <> 用户过程 Then cboProcName.Clear
                cboProcName.Text = "ZLUSER_"
            Else
                cboOwner.Locked = True: cboProcName.Locked = True
            End If
    End Select
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    Case conMenu_Edit_SaveExit
        If ValidData Then
            If SaveProcData() Then
                mblnOk = True
                Unload Me
            End If
        End If
    Case conMenu_Edit_Save
        If ValidData Then
            If SaveProcData(True) Then
                mblnOk = True
                mblnChange = False
            End If
        End If
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case conMenu_Edit_Save
            '用户过程不能暂存
            Control.Enabled = mblnChange And mptType <> ProcType.用户过程 Or mlngKey = 0
        Case conMenu_Edit_SaveExit
            Control.Enabled = mblnChange Or mintState = ProcState.调整中 Or mlngKey = 0
    End Select
End Sub

Private Sub cmdProcName_Click()
    Dim objText As TextStream
    Dim strSQL As String
    Dim objSQL As clsSQLInfo
    Dim objScript As New clsRunScript
    
    On Error GoTo errH
    cdg.DialogTitle = "选择导入过程文件"
    cdg.Filter = "自定义过程文件|*.Sql"
    cdg.flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdg.InitDir = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Path", "Import", App.Path & "\Import\ImportProcdure")
    cdg.FileName = ""
    cdg.MaxFileSize = 32767
    cdg.CancelError = True
    On Error Resume Next
    cdg.ShowOpen
    If err.Number = 0 Then
        On Error GoTo errH
        Me.Refresh
        If cdg.FileTitle <> "" Then
            Set objText = gobjFile.OpenTextFile(cdg.FileName, ForAppending)
            objText.WriteLine "/" '保证存在存储过程结束符
            objText.Close
            If objScript.OpenFile(cdg.FileName) Then
                Do While Not objScript.EOF
                    If objScript.SQLInfo.Block Then
                        If objScript.SQLInfo.BlockType Like "*PROCEDURE*" Or objScript.SQLInfo.BlockType Like "*FUNCTION*" Then
                            Set objSQL = New clsSQLInfo
                            Call objSQL.CopySQL(objScript.SQLInfo)
                            Exit Do
                        End If
                    End If
                    objScript.ReadNextSQL
                Loop
            Else
                Exit Sub
            End If
            If objSQL Is Nothing Then
                MsgBox "选择文件中未发现有效的存储过程！", vbInformation, Me.Caption
                Exit Sub
            End If
            If LoadSQLInfo(objSQL) Then
                mblnChange = True
            End If
        End If
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyG And Shift = vbCtrlMask Then
        Call gclsBase.LocationObj(txtLocRow)
    End If
End Sub

Private Sub Form_Load()
    '清空数据,由于窗体缓存导致的错误
    fraHSplit.Tag = ""
    wbrCompare.Navigate ("")
    synLastProc.Text = ""
    synProcEdit.Text = ""
    txtNote.Text = "": picPane(PE_历史变动).Tag = ""
    cboProcName.Clear: cboProcName.Text = "": cboProcName.Tag = "": cboOwner.Clear
    lblOwner.Visible = True: cboOwner.Visible = True
    cboProcType.Locked = False: cboProcName.Locked = False: cboOwner.Locked = False
    wbrCompare.Visible = True: synLastProc.Visible = True
    fraHSplit.Visible = False: picPane(PE_历史变动).Visible = False
    If Not mblnLoad Then
        Call InitCommandBar
    End If
    Call InitTbc
    Call InitSQLArea
    Call FillData
    Call Form_Resize
    mblnChange = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not mblnOk And mblnChange Then
        If MsgBox("过程已经发生改变，直接退出将丢失修改。确认退出吗？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim dbRate As Double, lngTotal As Long
    Dim lngLeft As Long, lngRight As Long, lngTop As Long, lngBottom As Long
    
    On Error Resume Next
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If picPane(PE_历史变动).Tag = "隐藏" Then
        picPane(PE_当前过程).Move lngLeft, lngTop, lngRight - lngLeft - 30, lngBottom - lngTop - 30
        fraHSplit.Visible = False
        picPane(PE_历史变动).Visible = False
    Else
        fraHSplit.Visible = True
        picPane(PE_历史变动).Visible = True
        If fraHSplit.Tag = "" Then '没有拖动就按默认比例分屏
            lngTotal = picPane(PE_历史变动).Height + picPane(PE_当前过程).Height
            dbRate = picPane(PE_历史变动).Height / lngTotal
        Else
            lngTotal = lngBottom - lngTop - 30 - fraHSplit.Height - 30
            dbRate = (fraHSplit.Top) / lngTotal
        End If
        
        If dbRate < 0.1 Then
            dbRate = 0.1
        ElseIf dbRate > 0.9 Then
            dbRate = 0.9
        End If
        lngTotal = lngBottom - lngTop - 30 - fraHSplit.Height - 30
        picPane(PE_历史变动).Move lngLeft, lngTop + 30, lngRight - lngLeft, lngTotal * dbRate
        fraHSplit.Move lngLeft, picPane(PE_历史变动).Top + picPane(PE_历史变动).Height + 15, lngRight - lngLeft, fraHSplit.Height
        picPane(PE_当前过程).Move lngLeft, fraHSplit.Top + fraHSplit.Height + 15, lngRight - lngLeft, lngBottom - fraHSplit.Top - fraHSplit.Height - 15
        fraHSplit.Tag = ""
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mrsProcedure Is Nothing) Then Set mrsProcedure = Nothing
    If gobjFile.FolderExists(App.Path & "\Reports") Then Call gobjFile.DeleteFolder(App.Path & "\Reports")
    If Not (mfrmProcedureOwnerCon Is Nothing) Then Unload mfrmProcedureOwnerCon
End Sub

Private Sub fraHSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then fraHSplit.Top = fraHSplit.Top + y
End Sub

Private Sub fraHSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraHSplit.Tag = "拖动"
    Call Form_Resize
End Sub

Private Sub picEdit_Resize()
    On Error Resume Next
    picBase.Move picEdit.ScaleWidth - 30 - picBase.Width, 15, picBase.Width, picEdit.ScaleHeight - 30
    synProcEdit.Move 15, 15, picBase.Left - 15, picEdit.ScaleHeight - 30
End Sub

Private Sub picLast_Resize()
    On Error Resume Next
    wbrCompare.Move 15, 0, picLast.ScaleWidth - 30, picLast.ScaleHeight
    synLastProc.Move 15, 15, picLast.ScaleWidth - 30, picLast.ScaleHeight - 30
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    TbcBase(Index).Move 0, 0, picPane(Index).ScaleWidth, picPane(Index).ScaleHeight
    picPane(Index).Refresh
End Sub

Private Sub synLastProc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        synLastProc.SelectAll 'Ctrl+A
    ElseIf KeyCode = vbKeyC And Shift = vbCtrlMask Then
        synLastProc.Copy
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        synProcEdit.ShowFindReplaceDialog (False)
    End If
End Sub

Private Sub synProcEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        synProcEdit.SelectAll 'Ctrl+A
    ElseIf KeyCode = vbKeyZ And Shift = vbCtrlMask Then
        synProcEdit.UnDo
    ElseIf KeyCode = vbKeyC And Shift = vbCtrlMask Then
        synProcEdit.Copy
    ElseIf KeyCode = vbKeyV And Shift = vbCtrlMask Then
        synProcEdit.Paste
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        synProcEdit.ShowFindReplaceDialog (False)
    ElseIf KeyCode = vbKeyH And Shift = vbCtrlMask Then
        synProcEdit.ShowFindReplaceDialog (True)
    ElseIf KeyCode = vbKeyS And Shift = vbCtrlMask Then
        synProcEdit.ShowFindReplaceDialog (True)
    End If
End Sub

Private Sub txtLocRow_GotFocus()
    Call gclsBase.TxtSelAll(txtLocRow)
End Sub

Private Sub txtLocRow_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    
    If KeyAscii = vbKeyReturn Then
        lngRow = Val(txtLocRow.Text)
        If lngRow = 0 Then lngRow = 1
        If synProcEdit.RowsCount < lngRow Then
            synProcEdit.CurrPos.Row = synProcEdit.RowsCount
        Else
            synProcEdit.CurrPos.Row = lngRow
        End If
        Call gclsBase.LocationObj(txtLocRow)
    Else
        If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) <= 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub synProcEdit_TextChanged(ByVal nRowFrom As Long, ByVal nRowTo As Long, ByVal nActions As Long)
    mblnChange = True
End Sub
'==============================================================
'==私有方法
'==============================================================

Private Sub InitCommandBar()
    '******************************************************************************************************************
    '功能：初始菜单工具栏
    '参数：无
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
'    cbsMain.DeleteAll
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_SaveExit, "完成(&S)")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Save, "暂存(&C)")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出(&X)")
    mblnLoad = True
End Sub

Private Sub InitTbc()
    With TbcBase(PE_历史变动).PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .BoldSelected = True
        .ClientFrame = xtpTabFrameSingleLine
        .ShowIcons = True
        .DisableLunaColors = False
        .Position = xtpTabPositionTop
        .Appearance = xtpTabAppearanceVisio
        .Color = xtpTabColorOffice2003
        .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
        .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
        TbcBase(PE_历史变动).RemoveAll
        TbcBase(PE_历史变动).InsertItem(0, "上次过程", picLast.hwnd, 1).Tag = "上次过程"
    End With
    
    With TbcBase(PE_当前过程).PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .BoldSelected = True
        .ClientFrame = xtpTabFrameSingleLine
        .ShowIcons = True
        .DisableLunaColors = False
        .Position = xtpTabPositionTop
        .Appearance = xtpTabAppearanceVisio
        .Color = xtpTabColorOffice2003
        .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
        .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
        TbcBase(PE_当前过程).RemoveAll
        TbcBase(PE_当前过程).InsertItem(0, "本次过程", picEdit.hwnd, 1).Tag = "本次过程"
    End With
End Sub

Private Sub InitSQLArea()
    Dim strPath As String, strColor As String
    '语法控件颜色方案
    synLastProc.Font.name = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontName", "Fixedsys")
    synLastProc.Font.Size = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontSize", 12)
    synLastProc.Font.Underline = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontUnderline", 0)
    synLastProc.Font.Italic = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontItalic", 0)
    synLastProc.Font.Bold = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontBold", 0)
    synLastProc.Font.Strikethrough = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontStrikethru", 0)
    synLastProc.BorderStyle = xtpBorderClientEdge
    
    synProcEdit.Font.name = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontName", "Fixedsys")
    synProcEdit.Font.Size = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontSize", 12)
    synProcEdit.Font.Underline = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontUnderline", 0)
    synProcEdit.Font.Italic = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontItalic", 0)
    synProcEdit.Font.Bold = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontBold", 0)
    synProcEdit.Font.Strikethrough = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontStrikethru", 0)
    synProcEdit.BorderStyle = xtpBorderClientEdge
    
    '设置控件的显示颜色方案为：SQL
    If Not gblnInIDE Then '增加多环境支持
        strPath = App.Path & "\PUBLIC\_sql.schclass"
    Else
        strPath = gobjFile.GetParentFolderName(GetSetting("ZLSOFT", "公共全局", "程序路径")) & "\PUBLIC\_sql.schclass"
    End If
    If Not gobjFile.FileExists(strPath) Then
        strPath = "C:\Appsoft\PUBLIC\_sql.schclass"
    End If
    If gobjFile.FileExists(strPath) Then
        strColor = ReadFileToString(strPath)
    Else
        strColor = ""
    End If
    synLastProc.SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
    synLastProc.SyntaxScheme = strColor
    
    synProcEdit.SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
    synProcEdit.SyntaxScheme = strColor
    
End Sub

Private Sub FillData()
    '获取字段长度
    cboProcName.Tag = gclsBase.GetMaxLength("zlProcedure", "名称")
    txtNote.MaxLength = gclsBase.GetMaxLength("zlProcedure", "说明")
    '添加存储过程类型
    Call LoadProcType
    '加载所有者
    Call LoadOwner
    '加载上次过程或对比信息
    Call LoadProcInfo
    cboProcName.Tag = IIf(synProcEdit.Text <> "", "不加载", "")
    cboProcType.ListIndex = mptType - 1
    cboProcName.Text = mstrProcName
    If mlngKey <> 0 Then
        cboProcName.AddItem mstrProcName
        cboProcName.ListIndex = 0 '加载数据库过程源码
        cboProcType.Locked = True
        cboProcName.Locked = True
    End If
End Sub

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTMp As String, strCurProc As String
    Dim objCurProc As New clsSQLInfo, objSouce As New clsSQLInfo
    Dim arrTmp As Variant, strSQL As String
    Dim cnOracle As ADODB.Connection
    Dim strPassword As String, strError As String
    Dim rsTmp As ADODB.Recordset, rsCur As ADODB.Recordset
    '过程名称验证
    If gclsBase.StrIsValid(cboProcName.Text, Val(cboProcName.Tag)) = False Then
        gclsBase.LocationObj cboProcName
        Exit Function
    End If
    If Trim(cboProcName.Text) = "" Then
        MsgBox "过程名称不能为空值，必须输入！", vbInformation + vbOKOnly, "中联软件"
        gclsBase.LocationObj cboProcName
        Exit Function
    End If
    '所有者验证
    If cboOwner.ListIndex = -1 Then
        MsgBox "请指定过程所有者!", vbInformation + vbOKOnly, "中联软件"
        gclsBase.LocationObj cboOwner
        Exit Function
    ElseIf cboOwner.ItemData(cboOwner.ListIndex) = 0 Then
        MsgBox "请指定过程所有者！", vbInformation + vbOKOnly, "中联软件"
        gclsBase.LocationObj cboOwner
        Exit Function
    End If
    '过程类型验证
    If cboProcType.ListIndex = -1 Then
        MsgBox "请指定过程类型!", vbInformation + vbOKOnly, "中联软件"
        gclsBase.LocationObj cboProcType
        Exit Function
    End If
    '过程说明验证
    If gclsBase.StrIsValid(txtNote.Text, txtNote.MaxLength) = False Then
        gclsBase.LocationObj txtNote
        Exit Function
    End If
    If mlngKey = 0 Then
        '验证过程名称是否匹配
        If Trim(txtNote.Text) = "" Then
            MsgBox "用户过程的过程说明不能为空！", vbInformation + vbOKOnly, "中联软件"
            gclsBase.LocationObj txtNote
            Exit Function
        End If
    End If
    strTMp = gclsBase.GetProgram(Trim(cboProcName.Text), , True)
    If mptType <> ProcType.用户过程 And mlngKey = 0 Then
        If strTMp = "" Then
            MsgBox "该过程不是" & IIf(mptType = ProcType.变动过程, "变动过程", "空白过程") & "！", vbInformation + vbOKOnly, "中联软件"
            Exit Function
        End If
    End If
    strCurProc = GetCurrentProctext(True)
    If Not objCurProc.LoadSQL(strCurProc, vbCrLf) Or Not objCurProc.Block Then
        MsgBox "无法解析编辑区域的存储过程，请格式化后重新保存！", vbInformation + vbOKOnly, "中联软件"
        Exit Function
    End If
    Set rsCur = objCurProc.AnsySQL()
    If rsCur Is Nothing Then
        MsgBox "无法解析编辑区域的存储过程，请格式化后重新保存！", vbInformation + vbOKOnly, "中联软件"
        Exit Function
    End If
    
    If strTMp <> "" Then
        If Not objSouce.LoadSQL(strTMp & vbCrLf & "/", vbCrLf) Then
            MsgBox "无法解析数据库中该存储过程，请格式化并编译后重试！", vbInformation + vbOKOnly, "中联软件"
            Exit Function
        End If
        Set rsTmp = objSouce.AnsySQL
        If rsTmp Is Nothing Then
            MsgBox "无法解析数据库中该存储过程，请格式化并编译后重试！", vbInformation + vbOKOnly, "中联软件"
            Exit Function
        End If
        If mptType <> ProcType.用户过程 Then
            strError = CompareProcPars(rsTmp, rsCur)
            If strError <> "" Then
                MsgBox "变动过程或空白过程不允许改过程参数以及过程名称或返回值！差异信息如下：" & strError, vbInformation + vbOKOnly, "中联软件"
                Exit Function
            End If
        End If
    End If
    rsCur.Filter = "位置=-1"
    If Trim(cboProcName.Text) <> rsCur!名称 Then
        MsgBox "编辑区域过程名称不匹配！", vbInformation + vbOKOnly, "中联软件"
        Exit Function
    End If
    '判断当前登录用户是否与过程所有者匹配
    strSQL = "Select User From Dual"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
    If Trim(cboOwner.Text) <> rsTmp!User And Not CollectionHave(gcolOwnerConn, "K" & cboOwner.Text) Then
        If mfrmProcedureOwnerCon Is Nothing Then Set mfrmProcedureOwnerCon = New frmProcOwnerConn
        If mfrmProcedureOwnerCon.ShowDialog(Me, cboOwner.Text, strPassword) Then
            
            Set cnOracle = gobjRegister.GetConnection(gstrServer, cboOwner.Text, strPassword, True, OraOLEDB, "", False)
            If cnOracle.State = adStateClosed Then
                Exit Function
            End If
            Call SetSQLTrace(gstrServer, cboOwner.Text, cnOracle)
            gcolOwnerConn.Add cnOracle, "K" & cboOwner.Text
        Else
            Exit Function
        End If
    End If
    ValidData = True
End Function

Private Function SaveProcData(Optional ByVal bln暂存 As Boolean) As Boolean
'功能：保存存储过程数据
'参数：bln暂存-是否是暂存数据
    Dim lngKey As Long
    Dim arrSQL() As Variant
    Dim objSQL As New clsSQLInfo
    Dim strTMp As String
    
    On Error GoTo errH
    If mlngKey = 0 Then
        lngKey = gclsBase.GetNextId("zlProcedure")
        If Not bln暂存 Then
            strTMp = gclsBase.GetProgram(cboProcName.Text)
        End If
    Else '修改
        lngKey = mlngKey
    End If
    Call gclsBase.AddItem(arrSQL, "Zl_Zlprocedure_Update(" & lngKey & "," & mptType & ",'" & cboProcName.Text & "'," & IIf(bln暂存, ProcState.调整中, ProcState.已调整) & ",'" & txtNote.Text & "','" & cboOwner.Text & "')")
    Call gclsBase.GetProcSQL(lngKey, ProcTextType.本次自定过程, GetCurrentProctext, arrSQL)
    If strTMp <> "" Then
        Call gclsBase.GetProcSQL(lngKey, ProcTextType.本次标准过程, strTMp, arrSQL)
    End If
    SaveProcData = gclsBase.ExecuteProcedureBeach(gcnOracle, arrSQL, "保存存储过程")
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub LoadProcInfo()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strProcText As String, pttType As ProcTextType
    Dim i As Long
    
    If mlngKey <> 0 Then
        '获取存储过程代码
        strSQL = "Select a.Id, a.类型, a.名称,Upper(a.所有者) 所有者, 说明,状态,性质, 序号, b.内容" & vbNewLine & _
                        "From Zlprocedure a, Zlproceduretext b" & vbNewLine & _
                        "Where a.Id = b.过程id(+) And a.Id = [1]" & vbNewLine & _
                        "Order By b.性质, b.序号"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mlngKey)
        If rsTmp.EOF Then
            mlngKey = 0
        Else
            mptType = Val(rsTmp!类型 & "")
            mstrProcName = rsTmp!名称 & ""
            mintState = Nvl(rsTmp!状态, 1)
            txtNote.Text = rsTmp!说明 & ""
            For i = 0 To cboOwner.ListCount
                If cboOwner.List(i) = rsTmp!所有者 & "" Then
                    cboOwner.ListIndex = i: Exit For
                End If
            Next
        End If
    End If
    If mlngKey <> 0 And mptType <> ProcType.用户过程 Then '用户过程不需要比较，不需要加载上次过程
        wbrCompare.Visible = True: synLastProc.Visible = False
        '将本次自定过程对应的上次标准过程与本次标准过程进行比较
        Call DealWithTmpFolder(True)
        '生成ProcTextType.上次标准过程
        rsTmp.Filter = "性质=" & ProcTextType.上次标准过程: rsTmp.Sort = "序号"
        Call CreateProcText(rsTmp, App.Path & "\Standard\" & mstrProcName & ".sql")
        '生成ProcTextType.本次标准过程
        rsTmp.Filter = "性质=" & ProcTextType.本次标准过程: rsTmp.Sort = "序号"
        Call CreateProcText(rsTmp, App.Path & "\NewStandard\" & mstrProcName & ".sql")
        '生成ProcTextType.本次自定过程
        rsTmp.Filter = "性质=" & ProcTextType.本次自定过程: rsTmp.Sort = "序号"
        Call CreateProcText(rsTmp, App.Path & "\ThisProcedure\" & mstrProcName & ".sql")
        '上次标准与本次标准对比，存在对比文件是因此两次不同，不存在是因为相同或者两者不同时存在
        If gobjFile.FileExists(App.Path & "\Standard\" & mstrProcName & ".sql") Then
            Call CompareFolder(App.Path & "\Standard", App.Path & "\NewStandard", App.Path & "\Reports")
        End If
        If gobjFile.FileExists(App.Path & "\Reports\" & mstrProcName & ".sql.htm") Then
            TbcBase(PE_历史变动).Item(0).Caption = "上次标准过程(左) 与 本次标准过程(右)差异对比"
            Call wbrCompare.Navigate(App.Path & "\Reports\" & mstrProcName & ".sql.htm")
        Else
            If gobjFile.FileExists(App.Path & "\NewStandard\" & mstrProcName & ".sql") Then
                Call CompareFolder(App.Path & "\NewStandard", App.Path & "\ThisProcedure", App.Path & "\Reports")
            End If
            If gobjFile.FileExists(App.Path & "\Reports\" & mstrProcName & ".sql.htm") Then
                TbcBase(PE_历史变动).Item(0).Caption = "本次标准过程(左) 与 本次自定过程(右)差异对比"
                Call wbrCompare.Navigate(App.Path & "\Reports\" & mstrProcName & ".sql.htm")
            Else
                If gobjFile.FileExists(App.Path & "\Standard\" & mstrProcName & ".sql") Then
                    pttType = ProcTextType.上次标准过程
                ElseIf gobjFile.FileExists(App.Path & "\NewStandard\" & mstrProcName & ".sql") Then
                    pttType = ProcTextType.本次标准过程
                End If
                If pttType <> 0 Then
                    wbrCompare.Visible = False: synLastProc.Visible = True
                    '生成上次过程
                    rsTmp.Filter = "性质=" & pttType
                    rsTmp.Sort = "序号"
                    strProcText = ""
                    If Not rsTmp.EOF Then
                        Do While Not rsTmp.EOF
                            strProcText = strProcText & rsTmp!内容 & ""
                            rsTmp.MoveNext
                        Loop
                        synLastProc.Text = strProcText
                    End If
                    If strProcText = "" Then picPane(PE_历史变动).Tag = "隐藏"
                Else
                    picPane(PE_历史变动).Tag = "隐藏"
                End If
            End If
        End If
        '生成本次过程
        rsTmp.Filter = "性质=" & ProcTextType.本次自定过程
        rsTmp.Sort = "序号"
        strProcText = ""
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                strProcText = strProcText & rsTmp!内容 & ""
                rsTmp.MoveNext
            Loop
            synProcEdit.Text = strProcText
        End If
    Else
        wbrCompare.Visible = False: synLastProc.Visible = False
        picPane(PE_历史变动).Tag = "隐藏"
    End If
End Sub

Private Sub CreateProcText(ByRef rsProc As ADODB.Recordset, ByVal strFile As String)
'      blnAdjustName=是否去掉过程的名称所带的双引号
    Dim objText As TextStream, strProcText As String
    Dim strName As String
    
    If Not rsProc.EOF Then
        strName = rsProc!名称 & ""
        Do While Not rsProc.EOF
            If rsProc!序号 = 1 Then
                '名称带双引号，则去掉
                If UCase(rsProc!内容) Like "*" & """" & UCase(strName) & """" & "*" Then
                    strProcText = strProcText & Replace(UCase(rsProc!内容), """" & UCase(strName) & """", strName)
                Else
                    strProcText = strProcText & rsProc!内容
                End If
            Else
                strProcText = strProcText & rsProc!内容
            End If
            rsProc.MoveNext
        Loop
        Set objText = gobjFile.CreateTextFile(strFile)
        objText.Write strProcText
        objText.Close
    End If
End Sub

Private Sub LoadProcNames()
'功能：加载存储过程
    Dim strSQL As String
    
    cboProcName.Clear
    If mrsProcedure Is Nothing Then
        strSQL = "Select Object_Name,Owner" & vbNewLine & _
                        "From All_Objects a" & vbNewLine & _
                        "Where a.Owner In (Select Distinct 所有者 From Zlsystems) And a.Object_Type In ('PROCEDURE', 'FUNCTION') And" & vbNewLine & _
                        "      a.Object_Name Not In (Select Upper(名称) 名称 From Zlprocedure) And a.Object_Name Not Like 'ZL%_UPGRADECHECK'" & vbNewLine & _
                        "Order By Object_Name"
        Set mrsProcedure = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取过程清单")
    End If
    mrsProcedure.Filter = "": mrsProcedure.Sort = "Object_Name"
    If mrsProcedure.RecordCount > 0 Then mrsProcedure.MoveFirst
    Do While Not mrsProcedure.EOF
        cboProcName.AddItem mrsProcedure!Object_Name
        mrsProcedure.MoveNext
    Loop
'    If mrsProcedure.RecordCount > 0 Then cboProcName.ListIndex = 0
End Sub

Private Sub LoadOwner()
'功能：加载所有者
    Dim strSQL As String, rsTmp As ADODB.Recordset
    '获取所有者
    strSQL = "Select Distinct Upper(所有者)  所有者 from zlSystems a"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取所有者")
    With cboOwner
        If Not rsTmp.EOF Then
            .AddItem "--所有者--"
            .ItemData(.NewIndex) = 0
            Do While Not rsTmp.EOF
                .AddItem rsTmp!所有者 & ""
                .ItemData(.NewIndex) = .NewIndex
                rsTmp.MoveNext
            Loop
            .ListIndex = -1
        End If
    End With
End Sub

Private Sub LoadProcType()
'功能：加载过程类型
    With cboProcType
        .Clear
        .AddItem "1-变动过程"
        .ItemData(.NewIndex) = 1
        .AddItem "2-空白过程"
        .ItemData(.NewIndex) = 2
        .AddItem "3-用户过程"
        .ItemData(.NewIndex) = 3
        .ListIndex = -1
    End With
End Sub

Private Sub DealWithTmpFolder(Optional ByVal blnCreate As Boolean)
'功能：处理临时目录
    '转换为大写的脚本
    If gobjFile.FolderExists(App.Path & "\Standard") Then Call gobjFile.DeleteFolder(App.Path & "\Standard", True)
    If gobjFile.FolderExists(App.Path & "\NewStandard") Then Call gobjFile.DeleteFolder(App.Path & "\NewStandard", True)
    If gobjFile.FolderExists(App.Path & "\ThisProcedure") Then Call gobjFile.DeleteFolder(App.Path & "\ThisProcedure", True)
    If gobjFile.FolderExists(App.Path & "\Reports") Then Call gobjFile.DeleteFolder(App.Path & "\Reports", True)
    If blnCreate Then
        Call gobjFile.CreateFolder(App.Path & "\Standard")
        Call gobjFile.CreateFolder(App.Path & "\NewStandard")
        Call gobjFile.CreateFolder(App.Path & "\ThisProcedure")
        Call gobjFile.CreateFolder(App.Path & "\Reports")
    End If
End Sub

Private Function LoadSQLInfo(ByVal objSQL As clsSQLInfo) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    If mlngKey <> 0 Then
        If UCase(objSQL.BlockName) <> UCase(cboProcName.Text) Then
            MsgBox "存储过程名称不匹配，选择的过程为""" & objSQL.BlockName & """！", vbInformation, Me.Caption
            Exit Function
        End If
    End If

    strSQL = "Select a.Id From Zlprocedure a Where Upper(a.名称)= [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "读取过程", UCase(objSQL.BlockName))
    If Not rsTmp.EOF Then
        mlngKey = Val(rsTmp!Id & "")
        Call Form_Load
    Else
        cboProcName.Locked = False: cboOwner.Locked = False
        mrsProcedure.Filter = "Object_Name='" & UCase(objSQL.BlockName) & "'"
        If mrsProcedure.EOF And mptType <> ProcType.用户过程 Then
            MsgBox "数据库中不存在存储过程或函数""" & objSQL.BlockName & """", vbInformation, Me.Caption
            Exit Function
        End If
        cboProcName.Text = objSQL.BlockName
        If Not mrsProcedure.EOF Then
            cboOwner.Text = mrsProcedure!Owner & ""
            cboProcName.Locked = True
            cboOwner.Locked = True
        End If
    End If
    synProcEdit.Text = objSQL.SQL
    LoadSQLInfo = True
End Function

Private Function GetCurrentProctext(Optional ByVal blnAppEnd As Boolean) As String
'功能：获取当前编辑区的过程内容
    Dim strSQL As String, blnHaveEnd As Boolean
    Dim i As Long

    For i = 1 To synProcEdit.RowsCount
        strSQL = strSQL & IIf(strSQL = "", "", vbCrLf) & synProcEdit.RowText(i)
        If blnAppEnd Then
            If TrimComment(TrimEx(synProcEdit.RowText(i))) = "/" Then
                blnHaveEnd = True
            End If
        End If
    Next
    '没有语句结束符，则自动增加一个。
    If Not blnHaveEnd And blnAppEnd Then
        strSQL = strSQL & IIf(strSQL = "", "", vbCrLf) & "/"
    End If
    GetCurrentProctext = strSQL
End Function

Private Function CompareProcPars(ByVal rsLeft As ADODB.Recordset, ByVal rsRigth As ADODB.Recordset) As String
'功能：对存储过程进行比较，返回比较结果。若无差异返回空
'参数：strLeftInfo=左边存储过程的参数信息
'      strRightInfo=右边过程的存储过程信息
'返回：差异信息，无差异不返回。
    Dim rsCom As ADODB.Recordset, strErr As String, intIndex As Integer
    Dim strSQL As String, rsDataType As ADODB.Recordset, strTMp As String
    Dim arrTmp As Variant
    
    
    On Error GoTo errH
    If gobjFile.FileExists("C:\rsLeft.xml") Then Call gobjFile.DeleteFile("C:\rsLeft.xml", True)
    If gobjFile.FileExists("C:\rsRigth.xml") Then Call gobjFile.DeleteFile("C:\rsRigth.xml", True)
    If gobjFile.FileExists("C:\rsCom.xml") Then Call gobjFile.DeleteFile("C:\rsCom.xml", True)
    rsLeft.Save "C:\rsLeft.xml", adPersistXML
    rsRigth.Save "C:\rsRigth.xml", adPersistXML
    '-1删除，0-不变,1-新增,2-更新
    Set rsCom = GetCompareRec(rsLeft, rsRigth, "位置", "方向,类型描述,类型,默认值", "名称,位置")
    rsCom.Save "C:\rsCom.xml", adPersistXML
    With rsCom
        '查看名称，类型是否改变
        .Filter = "MainKey='-1'"
        If !State = 2 Or !名称 <> !名称_New Then
            intIndex = intIndex + 1
            strErr = strErr & vbNewLine & intIndex & "-过程类型或名称差异:【" & !类型 & "】" & !名称 & " <---> 【" & !类型_New & "】" & !名称_New
        End If
        .Filter = "State<>0"
        If .RecordCount = 0 Then Exit Function '无差异，则退出
        '比较返回值
        .Filter = "MainKey='0' And State <> 0"
        If .RecordCount > 0 Then
            intIndex = intIndex + 1
            strErr = strErr & vbNewLine & intIndex & "-返回值类型差异:" & IIf(!State = 1, "无返回类型", !类型) & " <---> " & IIf(!State = -1, "无返回类型", !类型_New)
        End If
        '比较参数,优先处理参数缺失或者新增
        .Filter = "MainKey<>'0' And MainKey<>'-1' And State=-1" '缺失参数比较
        .Sort = "MainKey"
        Do While Not .EOF
            intIndex = intIndex + 1
            strErr = strErr & vbNewLine & intIndex & "-第" & !MainKey & "位参数缺失:名称:" & !名称 & " 入出方向:" & !方向 & " 参数类型:" & !类型描述 & !类型 & IIf(!默认值 & "" = "", "", " 默认值:" & !默认值)
            .MoveNext
        Loop
        .Filter = "MainKey<>'0' And MainKey<>'-1' And State=1" '添加参数
        .Sort = "MainKey"
        Do While Not .EOF
            intIndex = intIndex + 1
            strErr = strErr & vbNewLine & intIndex & "-新增第" & !MainKey & "位参数:名称:" & !名称_New & " 入出方向:" & !方向_New & " 参数类型:" & !类型描述_New & !类型_New & IIf(!默认值_New & "" = "", "", " 默认值:" & !默认值_New)
            .MoveNext
        Loop
        '参数类型变更
        '先处理动态类型的参数
        .Filter = "(MainKey<>'0' And MainKey<>'-1' And State=2 And 类型描述<>'') OR (MainKey<>'0' And MainKey<>'-1' And State=2 And 类型描述_New<>'')"
        .Sort = "MainKey"
        Do While Not .EOF
            If !类型描述 & "" <> "" Then
                If InStr(strSQL & ";", ";" & !类型描述 & ";") = 0 Then
                    strSQL = strSQL & ";" & !类型描述
                End If
            End If
            If !类型描述_New & "" <> "" Then
                If InStr(strSQL & ";", ";" & !类型描述_New & ";") = 0 Then
                    strSQL = strSQL & ";" & !类型描述_New
                End If
            End If
            .MoveNext
        Loop
        If strSQL <> "" Then
            strSQL = Mid(strSQL, 2)
            strSQL = "Select a.C1 || '.' || a.C2 Key, b.Data_Type" & vbNewLine & _
                    "From (Select C1, C2 From Table(f_Str2list2('" & strSQL & "', ';', '.'))) a, All_Tab_Columns b" & vbNewLine & _
                    "Where a.C1 = b.Table_Name(+) And a.C2 = b.Column_Name(+)"
            Set rsDataType = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
            .Sort = "MainKey"
            Do While Not .EOF
                If !类型描述 & "" <> "" Then
                    rsDataType.Filter = "Key='" & !类型描述 & "'"
                    .Update "类型", rsDataType!DATA_TYPE & ""
                End If
                If !类型描述_New & "" <> "" Then
                    rsDataType.Filter = "Key='" & !类型描述_New & "'"
                    .Update "类型_New", rsDataType!DATA_TYPE & ""
                End If
                .MoveNext
            Loop
        End If
        '调整比较结果
        .Filter = "MainKey<>'0' And MainKey<>'-1' And State=2"
        .Sort = "MainKey"
        Do While Not .EOF
            '类型描述存在差异或者类型差异，则对二者进行
            If !类型 <> "" And !类型_New <> "" Then
                If !类型 & "" = !类型_New & "" Then
                    strTMp = !DifInfo & ""
                    If InStr("," & strTMp & ",", ",类型描述,") > 0 Then
                        strTMp = Replace("," & strTMp, ",类型描述", "")
                    End If
                    If InStr("," & strTMp & ",", ",类型,") > 0 Then
                        strTMp = Replace("," & strTMp, ",类型", "")
                    End If
                    .Update "DifInfo", strTMp
                End If
            End If
            If !DifInfo & "" = "" Then
                .Update "State", 0
            End If
            .MoveNext
        Loop
        .Filter = "MainKey<>'0' And MainKey<>'-1' And State=2"
        .Sort = "MainKey"
        Do While Not .EOF
            intIndex = intIndex + 1
            strErr = strErr & vbNewLine & intIndex & "-第" & !MainKey & "位参数存在差异(忽略名称差异):" & vbNewLine & _
                     "名称:" & !名称 & " 入出方向:" & !方向 & " 参数类型:" & IIf(!类型描述 = "", !类型, !类型描述 & "%TYPE(" & IIf(!类型 = "", "无法获取类型", !类型) & ")") & IIf(!默认值 & "" = "", "", " 默认值:" & !默认值) & vbNewLine & _
                     "名称:" & !名称_New & " 入出方向:" & !方向_New & " 参数类型:" & IIf(!类型描述_New = "", !类型_New, !类型描述_New & "%TYPE(" & IIf(!类型_New = "", "无法获取类型", !类型_New) & ")") & IIf(!默认值_New & "" = "", "", " 默认值:" & !默认值_New)
            .MoveNext
        Loop
    End With
    CompareProcPars = strErr
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function
