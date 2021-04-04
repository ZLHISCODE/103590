VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmLabMBReagent 
   Caption         =   "酶标试剂设置"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11655
   Icon            =   "frmLabMBReagent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   4290
      Left            =   90
      ScaleHeight     =   4290
      ScaleWidth      =   8895
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   660
      Width           =   8895
      Begin VB.Frame fraTool 
         Caption         =   "过滤条件"
         Height          =   975
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   8850
         Begin VB.TextBox txt项目 
            Height          =   300
            Left            =   4725
            TabIndex        =   11
            Top             =   255
            Width           =   1950
         End
         Begin VB.TextBox txt厂商 
            Height          =   300
            Left            =   1035
            TabIndex        =   8
            Top             =   570
            Width           =   5640
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   300
            Left            =   1035
            TabIndex        =   5
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   73400323
            CurrentDate     =   40553
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   2520
            TabIndex        =   7
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   73400323
            CurrentDate     =   40553
         End
         Begin VB.Label lbl项目 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "测试项目"
            Height          =   180
            Left            =   3930
            TabIndex        =   10
            Top             =   285
            Width           =   720
         End
         Begin VB.Label lbl厂商 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "试剂厂商 "
            Height          =   180
            Left            =   240
            TabIndex        =   9
            Top             =   630
            Width           =   810
         End
         Begin VB.Label lbl效期 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "试剂效期               － "
            Height          =   180
            Left            =   240
            TabIndex        =   6
            Top             =   285
            Width           =   2340
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   3765
         Left            =   705
         TabIndex        =   2
         Top             =   1320
         Width           =   7875
         _cx             =   13891
         _cy             =   6641
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
         BackColorFixed  =   15790320
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.TextBox txtEdit 
            Height          =   375
            Left            =   5460
            TabIndex        =   3
            Top             =   2340
            Visible         =   0   'False
            Width           =   1125
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   635
      SimpleText      =   $"frmLabMBReagent.frx":000C
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabMBReagent.frx":0053
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15478
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   135
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgICON 
      Bindings        =   "frmLabMBReagent.frx":08E7
      Left            =   525
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmLabMBReagent.frx":08FB
   End
End
Attribute VB_Name = "frmLabMBReagent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' API declares
Private Type POINTAPI
        x As Long
        Y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'公共部份菜单ID定义:*表示有图标
'*********************************************************************
Private Const mconMenu_FilePopup = 1    '文件
Private Const mconMenu_ManagePopup = 2  '管理
Private Const mconMenu_EditPopup = 3    '编辑
Private Const mconMenu_ReportPopup = 4  '报表
Private Const mconMenu_ViewPopup = 7    '查看
Private Const mconMenu_ToolPopup = 8     '工具
Private Const mconMenu_HelpPopup = 9    '帮助

'文件菜单
Private Const mconMenu_File_Open = 100              '*打开(&O)…
Private Const mconMenu_File_PrintSet = 101          '*打印设置(&S)…
Private Const mconMenu_File_Preview = 102           '*预览(&V)
Private Const mconMenu_File_Print = 103             '*打印(&P)
Private Const mconMenu_File_Excel = 104             '输出到&Excel…

Private Const mconMenu_File_MedRecSetup = 1051      '打印设置(&S)
Private Const mconMenu_File_MedRecPreview = 1052    '打印预览(&P)

Private Const mconMenu_File_RowPrint = 121        '记录打印(&R)

Private Const mconMenu_File_Exit = 191            '*退出(&X)

'编辑

Private Const mconMenu_Manage_Append = 3001     '*增加(&Y)
Private Const mconMenu_Manage_Delete = 3004     '*删除(&D)
Private Const mconMenu_Manage_Modify = 3003       '*修改(&M)

Private Const mconMenu_Manage_Stop = 3503     '*保存(&C)
Private Const mconMenu_Manage_Cancle = 3014   '取消

'查看菜单
Private Const mconMenu_View_ToolBar = 701              '工具栏(&T)
Private Const mconMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Private Const mconMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Private Const mconMenu_View_ToolBar_Size = 7013           '大图标(&B)
Private Const mconMenu_View_StatusBar = 702            '状态栏(&S)
Private Const mconMenu_View_Append = 703               '附加信息(&A)
Private Const mconMenu_View_Expend = 711               '展开/折叠组(&X)
Private Const mconMenu_View_Expend_CurCollapse = 7111     '折叠当前组(&C)
Private Const mconMenu_View_Expend_CurExpend = 7112       '展开当前组(&E)
Private Const mconMenu_View_Expend_AllCollapse = 7113     '折叠所有组(&L)
Private Const mconMenu_View_Expend_AllExpend = 7114       '展开所有组(&X)
Private Const mconMenu_View_Find = 721                 '*查找(&F)
Private Const mconMenu_View_FindNext = 722             '继续查找(&N)
Private Const mconMenu_View_FindType = 723             '查找方式(&Y)
Private Const mconMenu_View_Filter = 731               '*数据过滤(&I),子窗体的过滤功能
Private Const mconMenu_View_Notify = 732               '*医嘱提醒(&B)
Private Const mconMenu_View_Busy = 733                 '诊室忙(&M)
Private Const mconMenu_View_Hide = 741                 '*隐藏(&H)
Private Const mconMenu_View_Show = 742                 '*显示(&S)
Private Const mconMenu_View_Backward = 743             '*后退(&B)
Private Const mconMenu_View_Forward = 744              '*前进(&F)
Private Const mconMenu_View_Option = 781               '选项(&O)
Private Const mconMenu_View_Refresh = 791              '*刷新(&R)
Private Const mconMenu_View_Jump = 792                 '跳转(&J)

'帮助菜单
Private Const mconMenu_Help_Help = 901        '*帮助主题(&H)
Private Const mconMenu_Help_Web = 902         '&WEB上的中联
Private Const mconMenu_Help_Web_Home = 9021       '中联主页(&H)
Private Const mconMenu_Help_Web_Forum = 9023   '中联论坛(&F)
Private Const mconMenu_Help_Web_Mail = 9022       '*发送反馈(&M)
Private Const mconMenu_Help_About = 991       '关于(&A)…

'其它常量定义
'*********************************************************************
'CommandBar固有常量定义
Private Const mXTP_ID_WINDOW_LIST = 35000 '窗体列表
Private Const mXTP_ID_TOOLBARLIST = 59392 '工具栏列表
Private Const mID_INDICATOR_CAPS = 59137 '状态栏（大写）
Private Const mID_INDICATOR_NUM = 59138 '状态栏（数字）
Private Const mID_INDICATOR_SCRL = 59139 '状态栏（滚动）

'CommandBar辅助热键
Private Const mFSHIFT = 4
Private Const mFCONTROL = 8
Private Const mFALT = 16

Private mlngModul As Long, mstrPrivs As String

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, objControl As CommandBarControl
    
    Select Case Control.ID
    Case mconMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case mconMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case mconMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case mconMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case mconMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hWnd)
    Case mconMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case mconMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hWnd)
    Case mconMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case mconMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case mconMenu_File_Exit                             '退出
        Unload Me
        
    Case mconMenu_Manage_Modify                         '修改
        vfgList.SelectionMode = flexSelectionFree
        vfgList.Editable = flexEDKbdMouse
         
    Case mconMenu_Manage_Stop                           '保存

        If SaveData = True Then
            vfgList.Editable = flexEDNone
            vfgList.SelectionMode = flexSelectionByRow
            Call LoadData
        End If
    Case mconMenu_Manage_Cancle                         '取消
        vfgList.Editable = flexEDNone
        vfgList.SelectionMode = flexSelectionByRow
        Call LoadData
    Case mconMenu_Manage_Append                         '增加
        With vfgList
            .SelectionMode = flexSelectionFree
            .Editable = flexEDKbdMouse
            .Rows = .Rows + 1

            .TextMatrix(.Rows - 1, .ColIndex("试剂厂商")) = Trim(.TextMatrix(.Rows - 2, .ColIndex("试剂厂商")))
            .TextMatrix(.Rows - 1, .ColIndex("试剂效期")) = Format(Now + 365, "yyyy-MM-dd")
            .Select .Rows - 1, .ColIndex("试剂批号")
            
        End With
    Case mconMenu_Manage_Delete '删除
    
        If vfgList.Row > 0 Then
            vfgList.RemoveItem (vfgList.Row)
            vfgList.SelectionMode = flexSelectionFree
            vfgList.Editable = flexEDKbdMouse
        End If
    Case mconMenu_View_Refresh '刷新
        Call LoadData
    End Select

End Sub

Private Function SaveData() As Boolean
    '保存数据
    Dim i As Integer, iRow As Integer
    Dim strSQL() As String, str批号 As String, str效期 As String
    Dim blnRollBack As Boolean
    On Error GoTo ErrHandle
    
    
    With vfgList
        
        If .Rows > 1 Then
            ReDim strSQL(.Rows - 1)
        End If
        For i = 1 To .Rows - 1
            str批号 = Replace(Trim(.TextMatrix(i, .ColIndex("试剂批号"))), "'", "")
'            If str批号 = "" Then
'                MsgBox "试剂批号不能为空！", vbInformation, gstrSysName
'                Exit Function
'            End If
            str效期 = Replace(Trim(.TextMatrix(i, .ColIndex("试剂效期"))), "'", "")
            If str效期 = "" Then
                
                MsgBox "试剂效期不能为空！", vbInformation, gstrSysName
                Exit Function
            Else
                If IsDate(str效期) = False Then
                    MsgBox "试剂效期不正确，请按YYYY-MM-DD格式填写！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '检查编码是否有重复
            For iRow = 1 To .Rows - 1
                If i <> iRow Then
                    If str批号 = Replace(Trim(.TextMatrix(iRow, .ColIndex("试剂批号"))), "'", "") Then
                        MsgBox "试剂批号[" & str批号 & "]重复，请调整！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            Next
            If str批号 <> "" Then
                strSQL(i) = "Zl_检验酶标试剂_Edit(0,'" & str批号 & "',To_Date('" & str效期 & "','YYYY-MM-DD'),'" & Replace(Trim(.TextMatrix(i, .ColIndex("试剂厂商"))), "'", "") & "'," & _
                            "'" & Replace(Trim(.TextMatrix(i, .ColIndex("测试方法"))), "'", "") & "'," & _
                            IIf(.TextMatrix(i, .ColIndex("测试项目")) = "", "Null", .TextMatrix(i, .ColIndex("项目ID"))) & ")"
            End If
        Next
        
        If vfgList.Rows >= 2 Then
            gstrSql = "Zl_检验酶标试剂_Edit(1)"
            gcnOracle.BeginTrans
            blnRollBack = True
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            For i = LBound(strSQL) To UBound(strSQL)
                If strSQL(i) Like "Zl_检验酶标试剂_Edit*" Then
                    Call zlDatabase.ExecuteProcedure(strSQL(i), Me.Caption)
                End If
            Next
            gcnOracle.CommitTrans
            SaveData = True
        Else
            gstrSql = "Zl_检验酶标试剂_Edit(1)"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            SaveData = True
        End If
    End With
    Exit Function
ErrHandle:
    If blnRollBack = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With Me.picMain
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
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
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
        
    '-------------------------------
    Case mconMenu_Manage_Stop, mconMenu_Manage_Cancle '保存,取消
        Control.Enabled = vfgList.Editable = flexEDKbdMouse
    Case mconMenu_Manage_Modify, mconMenu_Manage_Append, mconMenu_Manage_Delete, _
         mconMenu_View_Refresh  '修改，增加,删除,刷新
        Control.Enabled = vfgList.Editable = flexEDNone
    End Select
End Sub

Private Sub Form_Load()

    '菜单工具栏
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
    cbsMain.Icons = imgICON.Icons
    dtpStart.Value = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    dtpStart.Value = CDate("3000-01-01")
    txt厂商 = ""
    txt项目 = ""
    Call initMenus
    Call LoadData
    
End Sub

Private Sub initVfgList()
    Dim strHead As String
    '1 左对齐 4 居中 7 右对齐
    strHead = "试剂批号,1200,4;试剂效期,1000,1;试剂厂商,2800,1;测试方法,2500,1;测试项目,2500,1;项目Id,0,1;Modify,0,1"
    Call SetVsFlexGridHead(strHead, vfgList)
    
End Sub

Private Sub SetVsFlexGridHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid)
    '功能：初始vsFlexGrid
    '           有一固定行，初始化后，只有一行记录，无固定列。
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'vsGrid:    要初始化的控件

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
         
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
        Next
        
        '固定行文字居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
    End With
End Sub

Private Sub LoadData()
    Dim strSQL As String
    Dim rsTmp As New adodb.Recordset
    Dim dateS As Date, dateE As Date
    Dim strWhere As String, strItem As String, strSang As String
    
    dateS = CDate(Format(dtpStart.Value, "yyyy-MM-dd"))
    dateE = CDate(Format(dtpEnd.Value, "yyyy-MM-dd"))
    
    If dateE < dateS Then
        dtpStart = dateE
        dtpEnd = dateS
        dateS = CDate(Format(dtpStart.Value, "yyyy-MM-dd"))
        dateE = CDate(Format(dtpEnd.Value, "yyyy-MM-dd"))
    End If
    strWhere = ""
    If Trim(txt项目) <> "" Then
        strWhere = strWhere & " And B.名称 Like [3] "
        strItem = "%" & DelInvalidChar(Trim(txt项目)) & "%"
    End If
    If Trim(txt厂商) <> "" Then
        strWhere = strWhere & " And A.试剂厂商 Like [4] "
        strSang = "%" & DelInvalidChar(Trim(txt厂商)) & "%"
    End If
    
    strSQL = "Select A.试剂批号, A.试剂效期, A.试剂厂商, A.测试方法, A.测试项目id, B.名称" & vbNewLine & _
            "From 检验酶标试剂 A, 诊疗项目目录 B" & vbNewLine & _
            "Where A.测试项目id = B.ID(+) And A.试剂效期 between [1] and [2] " & vbNewLine & _
             strWhere & _
            "Order By A.试剂效期 Desc"

    With vfgList
        .Clear
        Call initVfgList
        '.ColComboList(.ColIndex("颜色")) = "..."
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dateS, dateE, strItem, strSang)
        Do Until rsTmp.EOF
             
            .TextMatrix(.Rows - 1, .ColIndex("试剂批号")) = zlCommFun.Nvl(rsTmp.Fields("试剂批号"))
            .TextMatrix(.Rows - 1, .ColIndex("试剂效期")) = zlCommFun.Nvl(rsTmp.Fields("试剂效期"))
            .TextMatrix(.Rows - 1, .ColIndex("试剂厂商")) = zlCommFun.Nvl(rsTmp.Fields("试剂厂商"))
            .TextMatrix(.Rows - 1, .ColIndex("测试方法")) = zlCommFun.Nvl(rsTmp.Fields("测试方法"))
            .TextMatrix(.Rows - 1, .ColIndex("测试项目")) = zlCommFun.Nvl(rsTmp.Fields("名称"))
            .TextMatrix(.Rows - 1, .ColIndex("项目ID")) = zlCommFun.Nvl(rsTmp.Fields("测试项目Id"))
            
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        
        If .Rows > 0 Then
            .Rows = .Rows - 1
        End If
        '行选择
        .SelectionMode = flexSelectionByRow
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub picMain_Resize()
    On Error Resume Next
    With Me.fraTool
        .Left = picMain.ScaleLeft
        .Top = picMain.ScaleTop
        .Width = picMain.ScaleWidth
    End With
    With Me.vfgList
        .Left = picMain.ScaleLeft
        .Top = Me.fraTool.Top + Me.fraTool.Height
        .Width = picMain.ScaleWidth
        .Height = picMain.ScaleHeight - Me.fraTool.Height
    End With

End Sub

Private Sub vfgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vfgList.TextMatrix(Row, vfgList.ColIndex("Modify")) = "Update"
End Sub

Private Sub vfgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfgList.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    
End Sub

Private Sub vfgList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If vfgList.Editable = flexEDNone Then Exit Sub
    If NewCol = vfgList.ColIndex("测试项目") Then
        vfgList.ComboList = "..."
    Else
        vfgList.ComboList = ""
    End If
End Sub

Private Sub vfgList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call vfgListButtonClick(Row, Col)
End Sub

Private Sub vfgList_DblClick()

'    Dim pt As POINTAPI
'
'    With vfgList
'        If .Editable = flexEDNone Then Exit Sub
'
'        If .MouseCol = .ColIndex("颜色") Then
'            pt.x = .ColPos(.MouseCol) \ Screen.TwipsPerPixelX
'            pt.y = (.RowPos(.MouseRow) + .RowHeight(.MouseRow)) \ Screen.TwipsPerPixelY
'            ClientToScreen .hWnd, pt
'
'            frmSelColor.lblRow = .MouseRow
'            frmSelColor.lblCol = .MouseCol
'            frmSelColor.lngColor = .Cell(flexcpFloodColor, .MouseRow, .MouseCol)
'            frmSelColor.Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
'            frmSelColor.Show vbModal, Me
'
'        End If
'    End With
End Sub

Private Sub initMenus()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "文件(&F)", -1, False) '固有
    objMenu.ID = mconMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_File_PrintSet, "打印设置(&S)…") '固有
        Set objControl = .Add(xtpControlButton, mconMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, mconMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, mconMenu_File_Excel, "输出到&Excel…")

        Set objControl = .Add(xtpControlButton, mconMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ManagePopup, "编辑(&E)", -1, False)
    objMenu.ID = mconMenu_ManagePopup
    With objMenu.CommandBar.Controls
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Stop, "保存(&S)")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Cancle, "取消(&C)")
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Append, "增加(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Modify, "修改(&M)")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Delete, "删除(&D)")
        
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = mconMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, mconMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, mconMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, mconMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, mconMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, mconMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_View_StatusBar, "状态栏(&S)") '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "帮助(&H)", -1, False) '固有
    objMenu.ID = mconMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_Help_Help, "帮助主题(&H)") '固有
        
        Set objPopup = .Add(xtpControlButtonPopup, mconMenu_Help_Web, "&WEB上的" & gstrProductName) '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, mconMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False '固有
            .Add xtpControlButton, mconMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False '固有
            .Add xtpControlButton, mconMenu_Help_Web_Mail, "发送反馈(&M)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, mconMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True '固有
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
       ' Set objControl = .Add(xtpControlButton, mconMenu_File_Preview, "预览") '固有
       ' Set objControl = .Add(xtpControlButton, mconMenu_File_Print, "打印") '固有

        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Stop, "保存"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Cancle, "取消"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_View_Refresh, "刷新")
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Append, "增加"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Modify, "修改")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Delete, "删除")

        Set objControl = .Add(xtpControlButton, mconMenu_Help_Help, "帮助"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, mconMenu_File_Exit, "退出") '固有
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, mconMenu_File_Print '打印
        .Add 0, vbKeyF1, mconMenu_Help_Help '帮助
    End With

    '设置一些公共的不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
        .AddHiddenCommand mconMenu_File_PrintSet '打印设置
        .AddHiddenCommand mconMenu_File_Excel '输出到Excel
    End With

    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub vfgList_EnterCell()
    On Error GoTo ErrHandle
    With vfgList
    
        If .Col = .ColIndex("测试项目") And .Row > 0 Then
            If txtEdit.Tag = "False" Then
                txtEdit.Left = .CellLeft
                txtEdit.Top = .CellTop
                txtEdit.Height = .CellHeight - 12
                txtEdit.Width = .CellWidth - 12
                txtEdit.Tag = "True"
            End If
        Else
            txtEdit.Tag = "False"
        End If
        
        Dim blnCancle As Boolean
        Call vfgList_BeforeEdit(.Row, .Col, blnCancle)
        If Not blnCancle Then
            Call .CellBorder(.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    With vfgList
    If (.Col = .ColIndex("测试项目")) And KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    Else
        If .Col = .ColIndex("测试项目") And vfgList.ComboList = "..." Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                txtEdit.Text = .EditText
                Call vfgList_CellButtonClick(.Row, .Col)
                txtEdit.Tag = False
                txtEdit.Visible = False
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End If
 End With
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vfgList
            If Col = .ColIndex("测试项目") Then
                txtEdit.Text = .EditText
                .EditText = ""
                Call vfgListButtonClick(Row, Col)
                txtEdit.Tag = False
                txtEdit.Visible = False
            
            ElseIf Col + 1 > .Cols - 3 Then
                If Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                End If
                If .TextMatrix(Row + 1, .ColIndex("试剂厂商")) = "" Then .TextMatrix(Row + 1, .ColIndex("试剂厂商")) = Trim(.TextMatrix(Row, .ColIndex("试剂厂商")))
                If .TextMatrix(Row + 1, .ColIndex("试剂效期")) = "" Then .TextMatrix(Row + 1, .ColIndex("试剂效期")) = Format(Now + 365, "yyyy-MM-dd")
                
                .Select Row + 1, .ColIndex("试剂效期")
            Else
                .Select Row, Col + 1
            End If
        End With
    End If
End Sub

Private Sub vfgList_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = cbsMain.ActiveMenuBar.FindControl(, mconMenu_ManagePopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub
Private Sub vfgListButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New adodb.Recordset
    Dim strSQL   As String, strInput As String
    Dim vRect As RECT, blnCanel As Boolean
    Dim i As Integer
    On Error GoTo ErrHandle
    
    If Col = vfgList.ColIndex("测试项目") Then
        '提取材料
        '--------------------------------------------------------------------------------------
            strInput = UCase(Trim(txtEdit))
            If InStr(strInput, " ") > 0 Then
                strInput = Trim(Split(strInput, " ")(0))
            End If
            If strInput = "" Then
                strSQL = "Select C.ID, C.编码, C.名称, A.缩写" & vbNewLine & _
                    "From 检验项目 A, 检验报告项目 B, 诊疗项目目录 C" & vbNewLine & _
                    "Where A.诊治项目id = B.报告项目id And B.诊疗项目id = C.ID And 项目类别 = 4 And C.组合项目 = 0"
            Else
                strSQL = "Select C.ID, C.编码, C.名称, A.缩写" & vbNewLine & _
                    "From 检验项目 A, 检验报告项目 B, 诊疗项目目录 C" & vbNewLine & _
                    "Where A.诊治项目id = B.报告项目id And B.诊疗项目id = C.ID And 项目类别 = 4 And C.组合项目 = 0 " & vbNewLine & _
                    " And (C.编码 like '%" & UCase(strInput) & "%' or C.名称 like '%" & UCase(strInput) & "%' or 缩写 Like '%" & UCase(strInput) & "%')"
            End If

            vRect = GetControlRect(txtEdit.hWnd)
            Set rsTmp = New adodb.Recordset
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "选择项目", False, "", "", False, False, True, _
                                                 vRect.Left, vRect.Top, txtEdit.Height, blnCanel, True, True)
            If Not blnCanel And rsTmp.State <> 0 Then
                If Not rsTmp.EOF Then
                    With vfgList
                        .EditText = Trim(zlCommFun.Nvl(rsTmp.Fields("名称")))
                        .TextMatrix(.Row, .ColIndex("测试项目")) = Trim(zlCommFun.Nvl(rsTmp.Fields("名称")))
                        .TextMatrix(.Row, .ColIndex("项目ID")) = zlCommFun.Nvl(rsTmp.Fields("ID"), "")
                    End With
                End If
                Set rsTmp = Nothing
            Else
                With vfgList
                    .EditText = ""
                    .TextMatrix(.Row, .ColIndex("测试项目")) = ""
                End With
            End If
            txtEdit = ""
    End If
    Call zlCommFun.PressKey(vbKeyRight)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

