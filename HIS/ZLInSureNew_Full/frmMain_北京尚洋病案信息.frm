VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMain_北京尚洋病案信息 
   Caption         =   "职工普通门诊报销"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13905
   Icon            =   "frmMain_北京尚洋病案信息.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   13905
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7560
      TabIndex        =   2
      ToolTipText     =   "快捷键：F3"
      Top             =   0
      Width           =   1320
   End
   Begin VB.PictureBox picPTMZ 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   6570
      Left            =   1080
      ScaleHeight     =   6570
      ScaleWidth      =   11445
      TabIndex        =   0
      Top             =   420
      Width           =   11445
      Begin VSFlex8Ctl.VSFlexGrid vsfPTMZ 
         Height          =   4695
         Left            =   105
         TabIndex        =   1
         Top             =   135
         Width           =   10635
         _cx             =   18759
         _cy             =   8281
         Appearance      =   2
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmMain_北京尚洋病案信息.frx":6852
         ScrollTrack     =   -1  'True
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
         ExplorerBar     =   7
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
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1770
      Top             =   405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_北京尚洋病案信息.frx":6A6D
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_北京尚洋病案信息.frx":78BF
            Key             =   "RootSel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   9855
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
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
            Object.Width           =   19659
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain_北京尚洋病案信息"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long      '锁定控件，不刷新
Private mstrPrivs               As String               '权限串
Private mobjFindKey             As CommandBarPopup      '查询
Private mstrFindKey             As String               '查询串
Private mlngModule              As Long                 '模块号
Private mstrSaveKey             As String               '保存的上次的分类选择关键字
Private mRsPTMZ                 As ADODB.Recordset      '数据集
Private mRsPTMZMX               As ADODB.Recordset      '数据集
Private mRsPTMZBX               As ADODB.Recordset      '数据集
Private mRsPTMZBXMX             As ADODB.Recordset      '数据集
Private mstrSortID              As String               '排序定位
Private mcbrPopupBar            As CommandBar           '弹出窗口
Private mintInsure              As Integer              '险类
Dim cbrPopupItem                As CommandBarControl    '弹出项

'打印模式
Private Enum gzlPrintModeS
    zlPrint = 1         '打印
    zlView = 2          '查看
    zlExcel = 3         '输出到Excel
End Enum
Private mzlPrintModeS           As gzlPrintModeS        '打印

Private Const mstrPTMZ = "select A.RESIDENCE_NO As ID,A.Up As 是否上传,A.UpMan AS 上传人,A.UpDateTime As 上传时间,A.StickID As 病人ID,A.CnName As 姓名,A.Sex As 性别,A.IDENTITY_NUMBER As 身份证号,B.医保号,A.RESIDENCE_NO As 住院号," & vbNewLine & _
                        "A.MEDICAL_RECORD_NO As 病案号,A.ADMISSION_DATE 入院日期,A.CONTACT_PERSON AS 联系人,A.CONTACT_PHONE AS 联系电话,A.CONTACT_ADDRESS As 联系地址" & vbNewLine & _
                        "from 长治病案信息 A,保险帐户 B" & vbNewLine & _
                        "Where A.Stickid=B.病人ID"
                          
Public Property Let intinsure(ByVal vNewValue As String)
    mintInsure = vNewValue
End Property

'==============================================================================
'=功能： 初始菜单工具栏
'==============================================================================
Private Sub InitCommandBar()
    Dim objMenu         As CommandBarPopup
    Dim objBar          As CommandBar
    Dim objExtendedBar  As CommandBar
    Dim objPopup        As CommandBarPopup
    Dim objControl      As CommandBarControl
    Dim cbrCustom       As CommandBarControlCustom
    
    On Error GoTo ErrH

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '编辑
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "新增登记(&N)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "修改登记(&E)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除登记(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Send, "上传病案(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SendBack, "撤销上传(&J)")
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Find, "过滤(&F)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)

    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrSysName)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, gstrSysName & "主页(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, gstrSysName & "论坛(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&E)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)
    
    '主菜单右侧的查找
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16

    mstrFindKey = Trim(GetPara("定位依据"))
    If InStr("住院号,医保号", mstrFindKey) = 0 Then mstrFindKey = "住院号"

    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.流水号", , , "流水号")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.医保号", , , "医保号")

    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = txtLocation.hwnd
    cbrCustom.flags = xtpFlagRightAlign

    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "前一条")
    objControl.flags = xtpFlagRightAlign
    objControl.Style = xtpButtonIcon

    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "后一条")
    objControl.flags = xtpFlagRightAlign
    objControl.Style = xtpButtonIcon
    
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "新增", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Send, "上传病案")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SendBack, "撤销上传")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        
        .Add FCONTROL, vbKeyF, conMenu_View_Find            '查找
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '新增
        .Add FCONTROL, vbKeyE, conMenu_Edit_Modify          '修改
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete       '删除
        .Add 0, vbKeyF3, conMenu_View_Location              '定位
        .Add 0, vbKeyF4, conMenu_View_Option                '选择定位依据
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      '前一条
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '后一条
    End With
    '------------------------------------------------------------------------------------------------------------------
    '弹出菜单分类
    
    Set mcbrPopupBar = cbsMain.Add("弹出项目菜单", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)", True)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "修改(&E)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Send, "上传病案(&S)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SendBack, "撤销上传(&J)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    SaveFlexState vsfPTMZ, Me.Name
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name, "窗口", Me.WindowState
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name, "LEFT", Me.Left
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name, "TOP", Me.Top
End Sub

Private Sub picDetail_Resize()
'    tbcPage.Move picDetail.Left + 15, tbcPage.Top + 15
End Sub

'==============================================================================
'=定位得到焦点选中
'==============================================================================
Private Sub txtLocation_GotFocus()
    On Error GoTo ErrH
    Call zlControl.TxtSelAll(txtLocation)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=快速定位
'==============================================================================
Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    
    On Error GoTo ErrH
    
    lngRow = 0
    If txtLocation.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        '读取大于当前行的记录数据
        For lngLoop = vsfPTMZ.Row + 1 To vsfPTMZ.Rows - 1
            If InStr(UCase(vsfPTMZ.TextMatrix(lngLoop, vsfPTMZ.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '读取小于当前行的记录数据
        If lngRow = 0 Then
            For lngLoop = 0 To vsfPTMZ.Row
                If InStr(UCase(vsfPTMZ.TextMatrix(lngLoop, vsfPTMZ.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If vsfPTMZ.Rows > 1 And lngRow >= 1 Then vsfPTMZ.Row = lngRow
        vsfPTMZ.ShowCell lngRow, vsfPTMZ.ColIndex(mstrFindKey)
        Call LocationObj(txtLocation)
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 排序后定位记录 vsfPTMZ
'==============================================================================
Private Sub vsfPTMZ_AfterSort(ByVal COL As Long, Order As Integer)
    Dim lngRow      As Long
    On Error GoTo ErrH
'    vsfSetRow vsfPTMZ, mstrSortID, "病种ID"
    lngRow = vsfPTMZ.FindRow(mstrSortID, -1, vsfPTMZ.ColIndex("ID"), False, True)
    If lngRow > 0 Then vsfPTMZ.Row = lngRow
    vsfPTMZ.ShowCell lngRow, 1
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 排序前记录病种ID vsfPTMZ
'==============================================================================
Private Sub vsfPTMZ_BeforeSort(ByVal COL As Long, Order As Integer)
    On Error GoTo ErrH
    mstrSortID = "" & vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 右键菜单 vsfAuditItem
'==============================================================================
Private Sub vsfPTMZ_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ErrH

    Select Case Button
        Case 2          '弹出菜单处理
        
            Call SendLMouseButton(vsfPTMZ.hwnd, x, y)

            mcbrPopupBar.ShowPopup
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 控件初始化
'==============================================================================
Private Sub InitControl()
    
    On Error GoTo ErrH
    
    Call InitCommandBar
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picPTMZ_Resize()
    On Error Resume Next
    vsfPTMZ.Move 15, 15, picPTMZ.Width - 30, picPTMZ.Height - 30
End Sub

'==============================================================================
'=功能： 位置设置
'==============================================================================
Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnNewCancel        As Boolean
    On Error GoTo ErrH
    
    Select Case Control.ID
        Case conMenu_Edit_NewItem                       '增加项目
            Call NewPTMZ
        Case conMenu_Edit_Modify                        '修改项目
            Call EditPTMZ
        Case conMenu_Edit_Delete                        '删除项目
            Call DeletePTMZ
        Case conMenu_Edit_Send                          '交易
            Call UpdateCenter
        Case conMenu_Edit_SendBack                      '撤销结算
            Call CancelUpdate
        Case conMenu_View_Find                          '搜索查找
            Call FindPTMZ
        Case conMenu_File_Preview   '预览
            mzlPrintModeS = zlView
            Call ItemPrint
        Case conMenu_File_Print   '打印
            mzlPrintModeS = zlPrint
            Call ItemPrint
        Case conMenu_File_Excel '输出到&Excel
            mzlPrintModeS = zlExcel
            Call ItemPrint
        Case conMenu_View_Forward
            Call ForwardPTMZ
        Case conMenu_View_Backward
            Call BackwardPTMZ
        Case conMenu_View_Option
            mobjFindKey.Execute
        Case conMenu_View_LocationItem
            mstrFindKey = Control.Parameter
            mobjFindKey.Caption = mstrFindKey
            cbsMain.RecalcLayout
        Case conMenu_View_Location
            LocationObj txtLocation
        Case conMenu_View_Refresh               '刷新
            Call RefreshPTMZ
        Case Else
            If Control.ID > 400 And Control.ID < 500 Then
                Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
            Else
                 '与业务无关的功能，公共的功能
                Call CommandBarExecutePublic(Control, Me, vsfPTMZ, "职工普通门诊报销")
            End If
    End Select
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
'
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo ErrH

    With vsfPTMZ
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel
                Control.Enabled = ((.Rows > 1) And IsPrivs(mstrPrivs, "上传"))
            Case conMenu_EditPopup
                Control.Visible = IsPrivs(mstrPrivs, "上传")
            Case conMenu_Edit_NewItem                    '增加项目
                Control.Visible = IsPrivs(mstrPrivs, "上传")
            Case conMenu_Edit_Modify                        '修改项目
                Control.Visible = IsPrivs(mstrPrivs, "上传")
                Control.Enabled = ((.Rows > 1) And IsPrivs(mstrPrivs, "上传"))
            Case conMenu_Edit_Delete                  '删除
                Control.Visible = IsPrivs(mstrPrivs, "上传")
                Control.Enabled = ((.Rows > 1) And IsPrivs(mstrPrivs, "上传"))
                If .Rows > 1 Then
                    Control.Enabled = .TextMatrix(.Row, .ColIndex("是否上传")) <> "1"
                End If
            Case conMenu_Edit_Send
                Control.Visible = IsPrivs(mstrPrivs, "上传")
                Control.Enabled = ((.Rows > 1) And IsPrivs(mstrPrivs, "上传"))
                If .Rows > 1 Then
                    Control.Enabled = .TextMatrix(.Row, .ColIndex("是否上传")) <> "1"
                End If
            Case conMenu_Edit_SendBack                      '撤销结算
                Control.Visible = IsPrivs(mstrPrivs, "上传")
                Control.Enabled = ((.Rows > 1) And IsPrivs(mstrPrivs, "上传"))
                If .Rows > 1 Then
                    Control.Enabled = .TextMatrix(.Row, .ColIndex("是否上传")) = "1"
                End If
            Case conMenu_View_Refresh
                Control.Visible = IsPrivs(mstrPrivs, "上传")
            Case conMenu_View_Forward
                Control.Enabled = .Row > 1
            Case conMenu_View_Backward
                Control.Enabled = .Row + 1 < .Rows
            Case conMenu_View_Find, conMenu_View_Refresh
                Control.Enabled = True
            Case conMenu_View_LocationItem, conMenu_View_LocationItem, conMenu_View_LocationItem
                If InStr(Control.Caption, mstrFindKey) > 0 Then
                    Control.Checked = True
                Else
                    Control.Checked = False
                End If

            Case Else
                Call CommandBarUpdatePublic(Control, Me)
        End Select
    End With
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 打印 ItemPrint
'==============================================================================
Private Sub ItemPrint()
    On Error GoTo ErrH
    subPrint (mzlPrintModeS)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    Dim lngLoop         As Long
    Dim objControl      As Object
    Dim objPrint        As New zlPrint1Grd
    Dim objAppRow       As zlTabAppRow
    
    If vsfPTMZ Is Nothing Then Exit Sub
    LockWindowUpdate vsfPTMZ.hwnd
    vsfPTMZ.ColHidden(vsfPTMZ.ColIndex("图标")) = True
    Call SearchPrintData(vsfPTMZ, frmPubResource.msfPrint)
    vsfPTMZ.ColHidden(vsfPTMZ.ColIndex("图标")) = False
    LockWindowUpdate 0
    '调用打印部件处理
    Set objPrint.Body = frmPubResource.msfPrint
    objPrint.Title.Text = Me.Caption
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人：" & UserInfo.姓名)
    Call objAppRow.Add("打印时间：" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub Form_Load()
On Error GoTo ErrH
    mstrPrivs = gstrPrivs
    Call InitControl
    If GetPersonSet Then
        '使用个性化设置【调已保存的格式】
        RestoreWinState Me, App.ProductName
        RestoreFlexState vsfPTMZ, Me.Name
        Me.WindowState = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name, "窗口", 0)
        If Me.WindowState = 0 Then
            Me.Left = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name, "LEFT", Me.Left)
            Me.Top = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name, "TOP", Me.Top)
        End If
    End If
    gstrSQL = mstrPTMZ
    gstrSQL = gstrSQL & vbCrLf & " And nvl(A.Up,0)=0 "
    
    '加载数据
    Call DataLoadPTMZ(gstrSQL)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrH
    picPTMZ.Move Me.ScaleLeft, Me.ScaleTop + 800, Me.ScaleWidth, Me.ScaleHeight - stbThis.Height - 800
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DataLoadPTMZ(sSql As String)
    Dim strField        As String
    Dim strFieldWIDth   As String
    Dim varField        As Variant
    Dim varFieldWIDth   As Variant
    Dim i               As Integer
On Error GoTo ErrH

    Set mRsPTMZ = zlDatabase.OpenSQLRecord(sSql, Me.Caption)
    Set vsfPTMZ.DataSource = mRsPTMZ
    '使用个性化设置【调已保存的格式】
    If GetPersonSet Then
        With vsfPTMZ
            strField = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name & "\VSFlexGrID", .Name & "名称", "")
            strFieldWIDth = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & Me.Name & "\VSFlexGrID", .Name & "宽度", "")
            varField = Split(strField, ",")
            varFieldWIDth = Split(strFieldWIDth, ",")
            For i = 0 To UBound(varField)
                If varField(i) <> "" And Val(varFieldWIDth(i)) <> 0 Then
                    If .ColIndex(varField(i)) <> -1 Then
                         .ColPosition(.ColIndex(varField(i))) = i
                         .ColWidth(i) = Val(varFieldWIDth(i))
                    End If
                End If
            Next
        End With
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'新增
Private Sub NewPTMZ()
    Dim str登记id       As String
    Dim strWhere        As String
On Error GoTo ErrH
    With frmMain_北京尚洋病案信息编辑
        .Show vbModal
        str登记id = .HospitalNumber
    End With
    Set frmMain_北京尚洋病案信息编辑 = Nothing
    If str登记id = "" Then Exit Sub '点击取消
    '刷新明细数据
    gstrSQL = mstrPTMZ
    strWhere = "And 登记ID='" & str登记id & "'"
    Call DataLoadPTMZ(strWhere)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'修改
Private Sub EditPTMZ()
    Dim str登记id       As String
    Dim strWhere        As String
On Error GoTo ErrH
    With frmMain_北京尚洋病案信息编辑
        str登记id = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
        .HospitalNumber = str登记id
        .UpdateCenter = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("是否上传")) = "1"
        .Show vbModal
        str登记id = .HospitalNumber
    End With
    Set frmMain_北京尚洋病案信息编辑 = Nothing
    If str登记id = "" Then Exit Sub '点击取消
    '刷新明细数据
    Call DataLoadPTMZ(mRsPTMZ.Source)
    '重新定位
    vsfSetRow vsfPTMZ, str登记id, "ID"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'删除
Private Sub DeletePTMZ()
    Dim str登记id       As String
    Dim strWhere        As String
On Error GoTo ErrH
    str登记id = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
    If MsgBox("确认删除住院号【" & str登记id & "】的病案信息吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    
    gstrSQL = "zl_长治病案信息_Delete('" & str登记id & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    '刷新明细数据
    Call DataLoadPTMZ(mRsPTMZ.Source)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'刷新
Private Sub RefreshPTMZ()
    Dim str登记id       As String
    Dim strErrMsg       As String
    Dim strWhere        As String
On Error GoTo ErrH
    If vsfPTMZ.Row > 1 Then
        str登记id = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
    End If
    '刷新明细数据
    Call DataLoadPTMZ(mRsPTMZ.Source)
    '重新定位
    vsfSetRow vsfPTMZ, str登记id, "ID"
    Exit Sub
ErrH:
    strErrMsg = Me.Name & "|" & Me.Caption & "|RefreshPTMZ:" & vbCrLf & Err.Description
    MsgBox strErrMsg, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

'前一条
Private Sub ForwardPTMZ()
    Dim strErrMsg       As String
On Error GoTo ErrH
    With vsfPTMZ
        If .Row > 1 Then
            .Row = .Row - 1
            .ShowCell .Row, .COL
        End If
    End With
    Exit Sub
ErrH:
    strErrMsg = Me.Name & "|" & Me.Caption & "|ForwardPTMZ:" & vbCrLf & Err.Description
    MsgBox strErrMsg, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

'后一条
Private Sub BackwardPTMZ()
    Dim strErrMsg       As String
On Error GoTo ErrH
    With vsfPTMZ
        If .Row < .Rows - 1 Then
            .Row = .Row + 1
            .ShowCell .Row, .COL
        End If
    End With
    Exit Sub
ErrH:
    strErrMsg = Me.Name & "|" & Me.Caption & "|BackwardPTMZ:" & vbCrLf & Err.Description
    MsgBox strErrMsg, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

'搜索查找
Private Sub FindPTMZ()
    Dim str登记id       As String
    Dim strErrMsg       As String
    Dim strWhere        As String
On Error GoTo ErrH
    '刷新明细数据
    With frmMain_北京尚洋病案信息过滤
        .Show vbModal
        strWhere = .strWhere
    End With
    Set frmMain_北京尚洋病案信息过滤 = Nothing
    If strWhere = "" Then Exit Sub

    Call DataLoadPTMZ(mstrPTMZ & strWhere)
    '重新定位
    vsfSetRow vsfPTMZ, str登记id, "ID"
    Exit Sub
ErrH:
    strErrMsg = Me.Name & "|" & Me.Caption & "|FindPTMZ:" & vbCrLf & Err.Description
    MsgBox strErrMsg, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

'上传病案到中心
Private Sub UpdateCenter()
    Dim str登记id As String
    Dim rsTmp As ADODB.Recordset
    Dim cnTest As ADODB.Connection
    Dim strServer As String
    Dim strUser As String
    Dim strPwd As String
    Dim str参数值 As String
    Dim strWhere As String
On Error GoTo ErrH
    str登记id = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
    If MsgBox("确认上传住院号【" & str登记id & "】的病案信息吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    '连接病案服务器
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_北京尚洋)
    Do Until rsTmp.EOF
        str参数值 = IIf(IsNull(rsTmp("参数值")), "", rsTmp("参数值"))
        Select Case rsTmp("参数名")
            Case "病案用户名"
                strUser = str参数值
            Case "病案用户密码"
                strPwd = str参数值
            Case "病案服务器"
                strServer = str参数值
        End Select
        rsTmp.MoveNext
    Loop
    
    Set cnTest = New ADODB.Connection
    If cnTest.State = adStateOpen Then cnTest.Close
    cnTest.ConnectionString = "Provider=MSDAORA.1;Password=" & Trim(strPwd) & ";User ID=" & strUser & ";Data Source=" & strServer & ";Persist Security Info=True"
    cnTest.CursorLocation = adUseClient
    cnTest.Open
    If Err <> 0 Then
        MsgBox "病案服务器连接失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    '上传
    gstrSQL = "Select * From 长治病案信息 where RESIDENCE_NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str登记id)
    
   
'    With rsTmp
'        gstrSQL = "" & _
'        "INSERT INTO VIEW_MEDICAL_RECORD_INFO" & vbNewLine & _
'        "    (" & vbNewLine & _
'        "      HOSPITAL_NUMBER,RESIDENCE_NO,IN_COUNT,MEDICAL_RECORD_NO,MARITAL_STATUS,STATUS,BIRTH_ADDRESS,IDENTITY_NUMBER,UNIT_NAME,UNIT_ADDRESS," & vbNewLine & _
'        "      UNIT_PHONE,UNIT_ZIPCODE,REGISTER_ADDRESS,REGISTER_ZIPCODE,CONTACT_PERSON,RELATIONSHIP,CONTACT_ADDRESS,CONTACT_PHONE,ADMISSION_DATE,ADMISSION_DEPT," & vbNewLine & _
'        "      IN_DEPT_ZONE,DEPT_TRANSFERED_TO,DISCHARGE_DATE,DISCHARGE_DEPT,OUT_DEPT_ZONE,PAT_ADM_CONDITION,DIAGNOSIS_DATE,ALERGY_DRUGS,HBSAG,HCV_AB,HIV_AB," & vbNewLine & _
'        "      CLINIC_INHOSPITAL,IN_OUT,BEFORE_AFTER_TREATMENT,CLINIC_PATHOLOGY,EMIT_PATHOLOGY,EMER_TREAT_TIMES,ESC_EMER_TIMES,DIRECTOR,DIRECTOR_DOCTOR,ATTENDING_DOCTOR," & vbNewLine & _
'        "      INHOSPITAL_DOCTOR,REFRESH_DOCTOR,GRADUATE_DOCTOR,INTERM,CODE_NAME,MEDICAL_RECORD_MASS,CONTROL_DOCTOR,CONTROL_NURSE,BAL_DATE,BODY_EXAMINE_FLAG," & vbNewLine & _
'        "      FIRST_FLAG,FOLLOW_FLAG,FOLLOW_TERM,TEACH_MR_FLAG,BLOOD_TYPE,RH,BLOOD_TRAN_REACT_FLAG,ERYTHROCYTE,HEMOBLAST,PLASM," & vbNewLine & _
'        "      BLOOD,OTHER_BLOOD,HANDLE,HANDLE_DATE,IN_DIAGNOSIS_CODE,IN_DIAGNOSIS_NAME,IN_DIAGNOSIS_DATE,OUT_DIAGNOSIS_CODE1,OUT_DIAGNOSIS_NAME1,OUT_DIAGNOSIS_DATE1," & vbNewLine & _
'        "      TREAT_RESULT1,OUT_DIAGNOSIS_CODE2,OUT_DIAGNOSIS_NAME2,OUT_DIAGNOSIS_DATE2,TREAT_RESULT2,OUT_DIAGNOSIS_CODE3,OUT_DIAGNOSIS_NAME3,OUT_DIAGNOSIS_DATE3,TREAT_RESULT3,OPERATION_CODE1," & vbNewLine & _
'        "      OPERATION_NAME1,WOUND_GRADE1,HEAL1,OPERATING_DATE1,ANAESTHESIA_METHOD1,OPERATION_CODE2,OPERATION_NAME2,WOUND_GRADE2,HEAL2,OPERATING_DATE2," & vbNewLine & _
'        "      ANAESTHESIA_METHOD2,OPERATION_CODE3,OPERATION_NAME3,WOUND_GRADE3,HEAL3,OPERATING_DATE3,ANAESTHESIA_METHOD3" & vbNewLine & _
'        "      )"
'        gstrSQL = gstrSQL & vbNewLine & _
'        "VALUES" & vbNewLine & _
'        "  (" & vbNewLine & _
'        "   '" & !HOSPITAL_NUMBER & "' , '" & !RESIDENCE_NO & "' , '" & !IN_COUNT & "' , '" & !MEDICAL_RECORD_NO & "' , '" & !MARITAL_STATUS & "' , '" & !Status & "' , '" & !BIRTH_ADDRESS & "' , '" & !IDENTITY_NUMBER & "' , '" & !UNIT_NAME & "' , '" & !UNIT_ADDRESS & "' ," & vbNewLine & _
'        "   '" & !UNIT_PHONE & "' , '" & !UNIT_ZIPCODE & "' , '" & !REGISTER_ADDRESS & "' , '" & !REGISTER_ZIPCODE & "' , '" & !CONTACT_PERSON & "' , '" & !RELATIONSHIP & "' , '" & !CONTACT_ADDRESS & "' , '" & !CONTACT_PHONE & "' , to_date('" & !ADMISSION_DATE & "','yyyy-dd-mm hh24:mi:ss'), '" & !ADMISSION_DEPT & "' ," & vbNewLine & _
'        "   '" & !IN_DEPT_ZONE & "' , '" & !DEPT_TRANSFERED_TO & "' ,to_date( '" & !DISCHARGE_DATE & "','yyyy-dd-mm hh24:mi:ss'), '" & !DISCHARGE_DEPT & "' , '" & !OUT_DEPT_ZONE & "' , '" & !PAT_ADM_CONDITION & "' ,to_date( '" & !DIAGNOSIS_DATE & "','yyyy-dd-mm hh24:mi:ss') , '" & !ALERGY_DRUGS & "' , '" & !HBSAG & "' , '" & !HCV_AB & "' , '" & !HIV_AB & "' ," & vbNewLine & _
'        "   '" & !CLINIC_INHOSPITAL & "' , '" & !IN_OUT & "' , '" & !BEFORE_AFTER_TREATMENT & "' , '" & !CLINIC_PATHOLOGY & "' , '" & !EMIT_PATHOLOGY & "' , '" & !EMER_TREAT_TIMES & "' , '" & !ESC_EMER_TIMES & "' , '" & !DIRECTOR & "' , '" & !DIRECTOR_DOCTOR & "' , '" & !ATTENDING_DOCTOR & "' ," & vbNewLine & _
'        "   '" & !INHOSPITAL_DOCTOR & "' , '" & !REFRESH_DOCTOR & "' , '" & !GRADUATE_DOCTOR & "' , '" & !INTERM & "' , '" & !CODE_NAME & "' , '" & !MEDICAL_RECORD_MASS & "' , '" & !CONTROL_DOCTOR & "' , '" & !CONTROL_NURSE & "' , to_date('" & !BAL_DATE & "','yyyy-dd-mm hh24:mi:ss') , '" & !BODY_EXAMINE_FLAG & "' ," & vbNewLine & _
'        "   '" & !FIRST_FLAG & "' , '" & !FOLLOW_FLAG & "' , '" & !FOLLOW_TERM & "' , '" & !TEACH_MR_FLAG & "' , '" & !BLOOD_TYPE & "' , '" & !RH & "' , '" & !BLOOD_TRAN_REACT_FLAG & "' , '" & !ERYTHROCYTE & "' , '" & !HEMOBLAST & "' , '" & !PLASM & "' ," & vbNewLine & _
'        "   '" & !BLOOD & "' , '" & !OTHER_BLOOD & "' , '" & !Handle & "' , to_date('" & !HANDLE_DATE & "','yyyy-dd-mm hh24:mi:ss') , '" & !IN_DIAGNOSIS_CODE & "' , '" & !IN_DIAGNOSIS_NAME & "' ,to_date( '" & !IN_DIAGNOSIS_DATE & "','yyyy-dd-mm hh24:mi:ss') , '" & !OUT_DIAGNOSIS_CODE1 & "' , '" & !OUT_DIAGNOSIS_NAME1 & "' ,to_date( '" & !OUT_DIAGNOSIS_DATE1 & "' ,'yyyy-dd-mm hh24:mi:ss')," & vbNewLine & _
'        "   '" & !TREAT_RESULT1 & "' , '" & !OUT_DIAGNOSIS_CODE2 & "' , '" & !OUT_DIAGNOSIS_NAME2 & "' ,to_date( '" & !OUT_DIAGNOSIS_DATE2 & "','yyyy-dd-mm hh24:mi:ss') , '" & !TREAT_RESULT2 & "' , '" & !OUT_DIAGNOSIS_CODE3 & "' , '" & !OUT_DIAGNOSIS_NAME3 & "' , to_date('" & !OUT_DIAGNOSIS_DATE3 & "','yyyy-dd-mm hh24:mi:ss') , '" & !TREAT_RESULT3 & "' , '" & !OPERATION_CODE1 & "' ," & vbNewLine & _
'        "   '" & !OPERATION_NAME1 & "' , '" & !WOUND_GRADE1 & "' , '" & !HEAL1 & "' , to_date('" & !OPERATING_DATE1 & "','yyyy-dd-mm hh24:mi:ss') , '" & !ANAESTHESIA_METHOD1 & "' , '" & !OPERATION_CODE2 & "' , '" & !OPERATION_NAME2 & "' , '" & !WOUND_GRADE2 & "' , '" & !HEAL2 & "' , to_date('" & !OPERATING_DATE2 & "','yyyy-dd-mm hh24:mi:ss') ," & vbNewLine & _
'        "   '" & !ANAESTHESIA_METHOD2 & "' , '" & !OPERATION_CODE3 & "' , '" & !OPERATION_NAME3 & "' , '" & !WOUND_GRADE3 & "' , '" & !HEAL3 & "' , to_date('" & !OPERATING_DATE3 & "','yyyy-dd-mm hh24:mi:ss'), '" & !ANAESTHESIA_METHOD3 & "'" & vbNewLine & _
'        "  )"
'    End With

With rsTmp
        gstrSQL = "" & _
        "INSERT INTO VIEW_MEDICAL_RECORD_INFO" & vbNewLine & _
        "    (" & vbNewLine & _
        "      HOSPITAL_NUMBER,RESIDENCE_NO,IN_COUNT,MEDICAL_RECORD_NO,MARITAL_STATUS,STATUS,BIRTH_ADDRESS,IDENTITY_NUMBER,UNIT_NAME,UNIT_ADDRESS," & vbNewLine & _
        "      UNIT_PHONE,UNIT_ZIPCODE,REGISTER_ADDRESS,REGISTER_ZIPCODE,CONTACT_PERSON,RELATIONSHIP,CONTACT_ADDRESS,CONTACT_PHONE,ADMISSION_DATE,ADMISSION_DEPT," & vbNewLine & _
        "      IN_DEPT_ZONE,DEPT_TRANSFERED_TO,DISCHARGE_DATE,DISCHARGE_DEPT,OUT_DEPT_ZONE,PAT_ADM_CONDITION,DIAGNOSIS_DATE,ALERGY_DRUGS,HBSAG,HCV_AB,HIV_AB," & vbNewLine & _
        "      CLINIC_INHOSPITAL,IN_OUT,BEFORE_AFTER_TREATMENT,CLINIC_PATHOLOGY,EMIT_PATHOLOGY,EMER_TREAT_TIMES,ESC_EMER_TIMES,DIRECTOR,DIRECTOR_DOCTOR,ATTENDING_DOCTOR," & vbNewLine & _
        "      INHOSPITAL_DOCTOR,REFRESH_DOCTOR,GRADUATE_DOCTOR,INTERM,CODE_NAME,MEDICAL_RECORD_MASS,CONTROL_DOCTOR,CONTROL_NURSE,BAL_DATE,BODY_EXAMINE_FLAG," & vbNewLine & _
        "      FIRST_FLAG,FOLLOW_FLAG,FOLLOW_TERM,TEACH_MR_FLAG,BLOOD_TYPE,RH,BLOOD_TRAN_REACT_FLAG,ERYTHROCYTE,HEMOBLAST,PLASM," & vbNewLine & _
        "      BLOOD,OTHER_BLOOD,HANDLE,HANDLE_DATE,IN_DIAGNOSIS_CODE,IN_DIAGNOSIS_NAME,IN_DIAGNOSIS_DATE,OUT_DIAGNOSIS_CODE1,OUT_DIAGNOSIS_NAME1,OUT_DIAGNOSIS_DATE1," & vbNewLine & _
        "      TREAT_RESULT1,OUT_DIAGNOSIS_CODE2,OUT_DIAGNOSIS_NAME2,OUT_DIAGNOSIS_DATE2,TREAT_RESULT2,OUT_DIAGNOSIS_CODE3,OUT_DIAGNOSIS_NAME3,OUT_DIAGNOSIS_DATE3,TREAT_RESULT3,OPERATION_CODE1," & vbNewLine & _
        "      OPERATION_NAME1,WOUND_GRADE1,HEAL1,OPERATING_DATE1,ANAESTHESIA_METHOD1,OPERATION_CODE2,OPERATION_NAME2,WOUND_GRADE2,HEAL2,OPERATING_DATE2," & vbNewLine & _
        "      ANAESTHESIA_METHOD2,OPERATION_CODE3,OPERATION_NAME3,WOUND_GRADE3,HEAL3,OPERATING_DATE3,ANAESTHESIA_METHOD3" & vbNewLine & _
        "      )"
        gstrSQL = gstrSQL & vbNewLine & _
        "VALUES" & vbNewLine & _
        "  (" & vbNewLine & _
        "   '" & !HOSPITAL_NUMBER & "' , '" & !RESIDENCE_NO & "' , '" & !IN_COUNT & "' , '" & !MEDICAL_RECORD_NO & "' , '" & !MARITAL_STATUS & "' , '" & !Status & "' , '" & !BIRTH_ADDRESS & "' , '" & !IDENTITY_NUMBER & "' , '" & !UNIT_NAME & "' , '" & !UNIT_ADDRESS & "' ," & vbNewLine & _
        "   '" & !UNIT_PHONE & "' , '" & !UNIT_ZIPCODE & "' , '" & !REGISTER_ADDRESS & "' , '" & !REGISTER_ZIPCODE & "' , '" & !CONTACT_PERSON & "' , '" & !RELATIONSHIP & "' , '" & !CONTACT_ADDRESS & "' , '" & !CONTACT_PHONE & "' , to_date('" & !ADMISSION_DATE & "','yyyy-mm-dd hh24:mi:ss'), '" & !ADMISSION_DEPT & "' ," & vbNewLine & _
        "   '" & !IN_DEPT_ZONE & "' , '" & !DEPT_TRANSFERED_TO & "' ,to_date( '" & !DISCHARGE_DATE & "','yyyy-mm-dd hh24:mi:ss'), '" & !DISCHARGE_DEPT & "' , '" & !OUT_DEPT_ZONE & "' , '" & !PAT_ADM_CONDITION & "' ,to_date( '" & !DIAGNOSIS_DATE & "','yyyy-mm-dd hh24:mi:ss') , '" & !ALERGY_DRUGS & "' , '" & !HBSAG & "' , '" & !HCV_AB & "' , '" & !HIV_AB & "' ," & vbNewLine & _
        "   '" & !CLINIC_INHOSPITAL & "' , '" & !IN_OUT & "' , '" & !BEFORE_AFTER_TREATMENT & "' , '" & !CLINIC_PATHOLOGY & "' , '" & !EMIT_PATHOLOGY & "' , '" & !EMER_TREAT_TIMES & "' , '" & !ESC_EMER_TIMES & "' , '" & !DIRECTOR & "' , '" & !DIRECTOR_DOCTOR & "' , '" & !ATTENDING_DOCTOR & "' ," & vbNewLine & _
        "   '" & !INHOSPITAL_DOCTOR & "' , '" & !REFRESH_DOCTOR & "' , '" & !GRADUATE_DOCTOR & "' , '" & !INTERM & "' , '" & !CODE_NAME & "' , '" & !MEDICAL_RECORD_MASS & "' , '" & !CONTROL_DOCTOR & "' , '" & !CONTROL_NURSE & "' , to_date('" & !BAL_DATE & "','yyyy-mm-dd hh24:mi:ss') , '" & !BODY_EXAMINE_FLAG & "' ," & vbNewLine & _
        "   '" & !FIRST_FLAG & "' , '" & !FOLLOW_FLAG & "' , '" & !FOLLOW_TERM & "' , '" & !TEACH_MR_FLAG & "' , '" & !BLOOD_TYPE & "' , '" & !RH & "' , '" & !BLOOD_TRAN_REACT_FLAG & "' , '" & !ERYTHROCYTE & "' , '" & !HEMOBLAST & "' , '" & !PLASM & "' ," & vbNewLine & _
        "   '" & !BLOOD & "' , '" & !OTHER_BLOOD & "' , '" & !Handle & "' , to_date('" & !HANDLE_DATE & "','yyyy-mm-dd hh24:mi:ss') , '" & !IN_DIAGNOSIS_CODE & "' , '" & !IN_DIAGNOSIS_NAME & "' ,to_date( '" & !IN_DIAGNOSIS_DATE & "','yyyy-mm-dd hh24:mi:ss') , '" & !OUT_DIAGNOSIS_CODE1 & "' , '" & !OUT_DIAGNOSIS_NAME1 & "' ,to_date( '" & !OUT_DIAGNOSIS_DATE1 & "' ,'yyyy-mm-dd hh24:mi:ss')," & vbNewLine & _
        "   '" & !TREAT_RESULT1 & "' , '" & !OUT_DIAGNOSIS_CODE2 & "' , '" & !OUT_DIAGNOSIS_NAME2 & "' ,to_date( '" & !OUT_DIAGNOSIS_DATE2 & "','yyyy-mm-dd hh24:mi:ss') , '" & !TREAT_RESULT2 & "' , '" & !OUT_DIAGNOSIS_CODE3 & "' , '" & !OUT_DIAGNOSIS_NAME3 & "' , to_date('" & !OUT_DIAGNOSIS_DATE3 & "','yyyy-mm-dd hh24:mi:ss') , '" & !TREAT_RESULT3 & "' , '" & !OPERATION_CODE1 & "' ," & vbNewLine & _
        "   '" & !OPERATION_NAME1 & "' , '" & !WOUND_GRADE1 & "' , '" & !HEAL1 & "' , to_date('" & !OPERATING_DATE1 & "','yyyy-mm-dd hh24:mi:ss') , '" & !ANAESTHESIA_METHOD1 & "' , '" & !OPERATION_CODE2 & "' , '" & !OPERATION_NAME2 & "' , '" & !WOUND_GRADE2 & "' , '" & !HEAL2 & "' , to_date('" & !OPERATING_DATE2 & "','yyyy-mm-dd hh24:mi:ss') ," & vbNewLine & _
        "   '" & !ANAESTHESIA_METHOD2 & "' , '" & !OPERATION_CODE3 & "' , '" & !OPERATION_NAME3 & "' , '" & !WOUND_GRADE3 & "' , '" & !HEAL3 & "' , to_date('" & !OPERATING_DATE3 & "','yyyy-mm-dd hh24:mi:ss'), '" & !ANAESTHESIA_METHOD3 & "'" & vbNewLine & _
        "  )"
    End With
 

    cnTest.Execute gstrSQL
    '更新本地标识
    gstrSQL = "zl_长治病案信息_UpServer('" & str登记id & "','1','" & UserInfo.姓名 & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    '刷新明细数据
    Call DataLoadPTMZ(mRsPTMZ.Source)
    '重新定位
    vsfSetRow vsfPTMZ, str登记id, "ID"
    Exit Sub
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

Private Sub CancelUpdate()
    Dim str登记id As String
    Dim rsTmp As ADODB.Recordset
    Dim cnTest As ADODB.Connection
    Dim strServer As String
    Dim strUser As String
    Dim strPwd As String
    Dim str参数值 As String
    Dim strWhere As String
On Error GoTo ErrH
    str登记id = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
    If MsgBox("确认撤销上传住院号【" & str登记id & "】的病案信息吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    '连接病案服务器
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_北京尚洋)
    Do Until rsTmp.EOF
        str参数值 = IIf(IsNull(rsTmp("参数值")), "", rsTmp("参数值"))
        Select Case rsTmp("参数名")
            Case "病案用户名"
                strUser = str参数值
            Case "病案用户密码"
                strPwd = str参数值
            Case "病案服务器"
                strServer = str参数值
        End Select
        rsTmp.MoveNext
    Loop
    
    Set cnTest = New ADODB.Connection
    If cnTest.State = adStateOpen Then cnTest.Close
    cnTest.ConnectionString = "Provider=MSDAORA.1;Password=" & Trim(strPwd) & ";User ID=" & strUser & ";Data Source=" & strServer & ";Persist Security Info=True"
    cnTest.CursorLocation = adUseClient
    cnTest.Open
    If Err <> 0 Then
        MsgBox "病案服务器连接失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    '上传
    gstrSQL = "Select * From VIEW_MEDICAL_RECORD_INFO where RESIDENCE_NO='" & str登记id & "'"
    Set rsTmp = cnTest.Execute(gstrSQL)
    If Not ChkRsState(rsTmp) Then
        MsgBox "中心住院号【" & str登记id & "】未撤销！本地不能撤销上传！", vbCritical, gstrSysName
        Exit Sub
    End If
    
    '更新本地标识
    gstrSQL = "zl_长治病案信息_UpServer('" & str登记id & "','0','" & UserInfo.姓名 & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    '刷新明细数据
    Call DataLoadPTMZ(mRsPTMZ.Source)
    '重新定位
    vsfSetRow vsfPTMZ, str登记id, "ID"
    Exit Sub
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub
