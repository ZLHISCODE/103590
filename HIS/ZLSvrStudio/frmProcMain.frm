VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmProcMain 
   Caption         =   "自定义过程管理"
   ClientHeight    =   6996
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13176
   Icon            =   "frmProcMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6996
   ScaleWidth      =   13176
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picMain 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4755
      Left            =   0
      ScaleHeight     =   4752
      ScaleWidth      =   12612
      TabIndex        =   0
      Top             =   600
      Width           =   12612
      Begin VB.PictureBox picHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEBD7&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Left            =   15
         ScaleHeight     =   1800
         ScaleWidth      =   12540
         TabIndex        =   1
         Top             =   15
         Width           =   12540
         Begin VB.Frame fraSplit 
            Height          =   45
            Left            =   0
            TabIndex        =   28
            Top             =   1360
            Width           =   12840
         End
         Begin VB.PictureBox picStep 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   350
            Index           =   1
            Left            =   7285
            ScaleHeight     =   324
            ScaleWidth      =   936
            TabIndex        =   26
            Top             =   720
            Width           =   960
            Begin VB.Label lblStep 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "过程调整"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   27
               Top             =   84
               Width           =   720
            End
         End
         Begin VB.PictureBox picStep 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   350
            Index           =   0
            Left            =   1775
            ScaleHeight     =   324
            ScaleWidth      =   1296
            TabIndex        =   24
            Top             =   720
            Width           =   1320
            Begin VB.Label lblStep 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "应用系统升级"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   85
               Width           =   1080
            End
         End
         Begin VB.CommandButton cmdFun 
            Caption         =   "生成(&G)"
            Height          =   350
            Index           =   3
            Left            =   8880
            TabIndex        =   20
            Top             =   720
            Width           =   1100
         End
         Begin VB.CommandButton cmdFun 
            Caption         =   "检查(&J)"
            Height          =   350
            Index           =   2
            Left            =   5555
            TabIndex        =   17
            Top             =   720
            Width           =   1100
         End
         Begin VB.CommandButton cmdFun 
            Caption         =   "升级完成(&C)"
            Height          =   350
            Index           =   1
            Left            =   3725
            TabIndex        =   14
            Top             =   720
            Width           =   1200
         End
         Begin VB.CommandButton cmdFun 
            Caption         =   "收集(&S)"
            Height          =   350
            Index           =   0
            Left            =   45
            TabIndex        =   11
            Top             =   720
            Width           =   1100
         End
         Begin VB.ComboBox cboProcState 
            Height          =   276
            Left            =   6840
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1476
            Width           =   1410
         End
         Begin VB.TextBox txtLocation 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   8916
            TabIndex        =   5
            ToolTipText     =   "请直接按回车键进行过滤"
            Top             =   1488
            Width           =   1695
         End
         Begin VB.OptionButton optType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "用户过程"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   3240
            TabIndex        =   4
            Top             =   1536
            Width           =   1305
         End
         Begin VB.OptionButton optType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "空白过程"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   1872
            TabIndex        =   3
            Top             =   1536
            Width           =   1305
         End
         Begin VB.OptionButton optType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "变动过程"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   528
            TabIndex        =   2
            Top             =   1536
            Width           =   1305
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "---＞"
            Height          =   180
            Index           =   3
            Left            =   6736
            TabIndex        =   23
            Top             =   805
            Width           =   468
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   1560
            Width           =   360
         End
         Begin VB.Label lblWarn 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   $"frmProcMain.frx":6852
            ForeColor       =   &H002222B2&
            Height          =   540
            Left            =   48
            TabIndex        =   21
            Top             =   120
            Width           =   6348
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "---＞"
            Height          =   180
            Index           =   4
            Left            =   8326
            TabIndex        =   19
            Top             =   805
            Width           =   468
         End
         Begin VB.Label lblResult 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "待调整过程清单"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   5555
            TabIndex        =   18
            Top             =   1120
            Width           =   1260
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "---＞"
            Height          =   180
            Index           =   2
            Left            =   5006
            TabIndex        =   16
            Top             =   805
            Width           =   468
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "---＞"
            Height          =   180
            Index           =   1
            Left            =   3176
            TabIndex        =   15
            Top             =   805
            Width           =   468
         End
         Begin VB.Label lblResult 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "各类过程清单"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   48
            TabIndex        =   13
            Top             =   1120
            Width           =   1080
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "---＞"
            Height          =   180
            Index           =   0
            Left            =   1226
            TabIndex        =   12
            Top             =   805
            Width           =   468
         End
         Begin VB.Label lblProcState 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "状态"
            Height          =   180
            Left            =   6360
            TabIndex        =   10
            Top             =   1536
            Width           =   360
         End
         Begin VB.Label lblLocation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "定位："
            Height          =   180
            Left            =   8388
            TabIndex        =   6
            Top             =   1536
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1752
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   7332
         _cx             =   12938
         _cy             =   3096
         Appearance      =   1
         BorderStyle     =   0
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmProcMain.frx":6904
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
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   6636
      Width           =   13176
      _ExtentX        =   23241
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "frmProcMain.frx":699B
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20355
            MinWidth        =   8819
            Text            =   "当前共有待调整0个；调整中0个"
            TextSave        =   "当前共有待调整0个；调整中0个"
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
   Begin MSComctlLib.ImageList imgEdit 
      Left            =   1080
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":722F
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":77C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":7D63
            Key             =   "签名"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":80B5
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":E917
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15179
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15641
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFlow 
      Left            =   1800
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15B09
            Key             =   "node"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15C50
            Key             =   "currnode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15D9F
            Key             =   "multnode"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15F21
            Key             =   "currmultnode"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":160E7
            Key             =   "arrow"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":1656A
            Key             =   "arrowlate"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":169E5
            Key             =   "arrow_Branch"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":16E05
            Key             =   "arrowlate_Branch"
         EndProperty
      EndProperty
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
Attribute VB_Name = "frmProcMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'===变量
'==============================================================
Private mfrmProgramEdit As frmProcEdit
Private mfrmBuildScript As frmProcBuildScript
Private mfrmProcedureRelating As frmProcRelating
Private mfrmCollectUpdate As frmProcCollectUpdate
Private mintProcType As Integer
Private mblnReading As Boolean
Private mlngProcID As Long '当前行ID
Private mobjMain As Object
Private mobjTip As clsTipSwap  '悬浮提示
Private Enum OptProcType
    OPT_变动过程 = 0
    OPT_空白过程 = 1
    OPT_用户过程 = 2
End Enum

Private Enum ProcCol
    PC_序号 = 0
    PC_选择 = 1
    PC_过程 = 2
    PC_状态 = 3
    PC_说明 = 4
End Enum


Private Enum cmdFun
    CF_收集 = 0
    CF_升级完成 = 1
    CF_检查 = 2
    CF_生成 = 3
End Enum

'==============================================================
'==公共接口
'==============================================================
Public Function ShowMe(ByVal objParent As Object)
    Me.Show 1, objParent
End Function

'==============================================================
'==控件事件
'==============================================================
Private Sub cboProcState_Click()
    If cboProcState.Tag = "" Then
        If IsSelData Then
            mlngProcID = Val(vsfMain.RowData(vsfMain.Row))
        Else
            mlngProcID = 0
        End If
    End If
    '读取加载数据
    Call RefreshData
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngLoop As Long
    Dim objControl As CommandBarControl
    Dim strSQL As String
    
    On Error GoTo errHand
    
    Dim rs As ADODB.Recordset
    Select Case Control.Id
    Case conMenu_File_PrintSet
        '打印设置
        Call zlPrintSet
    Case conMenu_File_Preview
        '预览
        PrintProcs 2
    Case conMenu_File_Print
        '打印
        PrintProcs 1
    Case conMenu_File_Excel
        '输出到Excel
        PrintProcs 3
    Case conMenu_Edit_NewItem
        If mfrmProgramEdit Is Nothing Then Set mfrmProgramEdit = New frmProcEdit
        Call mfrmProgramEdit.ShowMe(Me, 0, mintProcType)
    Case conMenu_Edit_Modify
        If mfrmProgramEdit Is Nothing Then Set mfrmProgramEdit = New frmProcEdit
        If vsfMain.RowData(vsfMain.Row) > 0 Then
            If mfrmProgramEdit.ShowMe(Me, vsfMain.RowData(vsfMain.Row), mintProcType) Then
                Call RefreshData
            End If
        End If
    Case conMenu_Edit_Disuse
        If MsgBox("您确认升级完成了吗？" & vbCrLf & "此操作会将本次升级前的过程记录更为上次过程记录！", vbOKCancel + vbInformation + vbDefaultButton2, "中联软件") = vbOK Then
            gcnOracle.Execute "Zl_Zlproceduretext_Move()"
            Call RefreshData
        End If
    Case conMenu_Edit_Audit
        If mfrmCollectUpdate Is Nothing Then Set mfrmCollectUpdate = New frmProcCollectUpdate
        If mfrmCollectUpdate.ShowMe(Me, 1) Then Call RefreshData
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Change_PaitNote
        Set rs = gclsBase.GetProcByState(1, 2)
        If rs.BOF = False Then
            MsgBox "检测到有过程还未调整完成，请先进行调整后再生成。", vbInformation + vbOKOnly, "中联软件"
            Exit Sub
        End If
        If mfrmBuildScript Is Nothing Then
            Set mfrmBuildScript = New frmProcBuildScript
        End If
        Call mfrmBuildScript.ShowMe(Me)
    Case conMenu_Edit_Delete
        Call FunDeleteProc
    Case conMenu_Edit_Untread
        Call FunRestoreProc
    Case conMenu_Edit_Word
        If mfrmCollectUpdate Is Nothing Then Set mfrmCollectUpdate = New frmProcCollectUpdate
        If mfrmCollectUpdate.ShowMe(Me) Then Call RefreshData
    Case conMenu_Edit_Confirm '确认调整
        Call FunConfirmProc
    Case conMenu_File_Exit
        Unload Me
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
        For lngLoop = 2 To cbsMain.Count
            cbsMain(lngLoop).Visible = Not cbsMain(lngLoop).Visible
        Next
        cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text      '按钮文字
        For lngLoop = 2 To cbsMain.Count
            For Each objControl In cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size      '大图标
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
    Case conMenu_View_StatusBar         '状态栏
        stbThis.Visible = Not stbThis.Visible
        cbsMain.RecalcLayout
    Case conMenu_Help_Help              '帮助主题
'        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((ParamInfo.系统号) / 100))
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me)
    End Select
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngRight As Long, lngTop As Long, lngBottom As Long
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If stbThis.Visible Then lngBottom = lngBottom - stbThis.Height
    picMain.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnSelData As Boolean
    With vsfMain
        Select Case Control.Id
            Case conMenu_Edit_Delete
                Control.Visible = Not optType(OPT_空白过程).value
                Control.Enabled = IsSelData() And Control.Visible
            Case conMenu_Edit_Modify
                Control.Enabled = IsSelData() And Control.Visible
            Case conMenu_Edit_Untread
                Control.Visible = (optType(OPT_变动过程).value Or optType(OPT_空白过程).value)
                Control.Enabled = IsSelData() And Control.Visible
            Case conMenu_Edit_Confirm '确认调整
                Control.Visible = optType(OPT_用户过程).value
                Control.Enabled = IsSelData() And Control.Visible
            Case conMenu_View_ToolBar_Button            '工具栏
                If cbsMain.Count >= 2 Then
                    Control.Checked = cbsMain(2).Visible
                End If
            Case conMenu_View_ToolBar_Text              '图标文字
                If cbsMain.Count >= 2 Then
                    Control.Checked = Not (cbsMain(2).Controls(1).Style = xtpButtonIcon)
                End If
            Case conMenu_View_ToolBar_Size              '大图标
                Control.Checked = cbsMain.Options.LargeIcons
            Case conMenu_View_StatusBar                 '状态栏
                Control.Checked = stbThis.Visible
        End Select
    End With
End Sub

Private Sub cmdFun_Click(Index As Integer)
    Dim objControl  As CommandBarControl
    Set objControl = cbsMain.FindControl(xtpControlButton, Decode(Index, CF_收集, conMenu_Edit_Word, CF_升级完成, conMenu_Edit_Disuse, CF_检查, conMenu_Edit_Audit, CF_生成, conMenu_Manage_Change_PaitNote))
    
    If Not objControl Is Nothing Then
        Call cbsMain_Execute(objControl)
    End If
End Sub

Private Sub Form_Load()
    '应用OEM图标
    Call ApplyOEM(stbThis)
    '初始化菜单
    Call InitCommandBar
    '默认展示变动过程
    optType(OPT_变动过程).value = True
    Call OptType_Click(OPT_变动过程)
    '读取加载数据
    Call RefreshData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mfrmBuildScript Is Nothing) Then Unload mfrmBuildScript
    If Not (mfrmCollectUpdate Is Nothing) Then Unload mfrmCollectUpdate
    If Not (mfrmProcedureRelating Is Nothing) Then Unload mfrmProcedureRelating
    If Not (mfrmProgramEdit Is Nothing) Then Unload mfrmProgramEdit
End Sub

Private Sub OptType_Click(Index As Integer)
    Dim arrTmp As Variant, strTMp As String, i As Integer
    Dim strOldType As String, intIndex As Integer
    If IsSelData Then
        mlngProcID = Val(vsfMain.RowData(vsfMain.Row))
    Else
        mlngProcID = 0
    End If
    mintProcType = (Index + 1)
    strTMp = "全部,-1,待检查,0,待调整,1,调整中,2,已调整,3,无变化,4"
    strOldType = cboProcState.Text: intIndex = -1
    arrTmp = Split(strTMp, ",")
    cboProcState.Clear
    For i = LBound(arrTmp) To UBound(arrTmp) Step 2
        cboProcState.AddItem arrTmp(i)
        cboProcState.ItemData(cboProcState.NewIndex) = Val(arrTmp(i + 1))
        If intIndex = -1 Then
            If arrTmp(i) = strOldType Then intIndex = cboProcState.NewIndex
        End If
    Next
    If intIndex = -1 Then intIndex = 0
    cboProcState.Tag = "刷新" '标识不需要重新获取当前行ID
    cboProcState.ListIndex = intIndex
    cboProcState.Tag = ""
End Sub

Private Sub optType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'    Call ShowTips(Index)
End Sub

Private Sub picHeader_Resize()
    On Error Resume Next
    txtLocation.Move picHeader.ScaleWidth - txtLocation.Width - 75
    lblLocation.Move txtLocation.Left - lblLocation.Width - 30
    cboProcState.Move lblLocation.Left - cboProcState.Width - 60
    lblProcState.Move cboProcState.Left - lblProcState.Width - 30
    fraSplit.Width = picHeader.ScaleWidth - fraSplit.Left
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    picHeader.Move 15, 15, picMain.ScaleWidth - 30
    vsfMain.Move 15, picHeader.Top + picHeader.Height + 15, picMain.ScaleWidth - 30, picMain.ScaleHeight - (picHeader.Top + picHeader.Height + 15) - 15
End Sub

Private Sub txtLocation_GotFocus()
    Call gclsBase.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim intCol As Integer
    
    If KeyAscii = vbKeyReturn Then
        intCol = vsfMain.ColIndex("过程")
        lngRow = vsfMain.FindRow(UCase(txtLocation.Text), intCol, 2, vsfMain.Row + 1)
        If lngRow = -1 Then
            lngRow = vsfMain.FindRow(UCase(txtLocation.Text), intCol, 2)
        End If
        If lngRow > 0 And vsfMain.Row <> lngRow Then
            vsfMain.Row = lngRow
            vsfMain.ShowCell vsfMain.Row, vsfMain.Col
        End If
        Call gclsBase.LocationObj(txtLocation)
    End If
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfMain
        .Redraw = False
        If .Rows - 1 > 0 Then
            .Cell(flexcpForeColor, .FixedRows, PC_序号, .Rows - 1, PC_序号) = Color.深灰色
            .Cell(flexcpFontBold, .FixedRows, PC_序号, .Rows - 1, PC_序号) = False
            .Cell(flexcpFontBold, NewRow, PC_序号, NewRow, PC_序号) = True
            .Cell(flexcpForeColor, NewRow, PC_序号, NewRow, PC_序号) = Color.兰色
        End If
        .Redraw = True
    End With
End Sub

Private Sub vsfMain_AfterSort(ByVal Col As Long, Order As Integer)
    Call SetSerial
End Sub

Private Sub vsfMain_BeforeSort(ByVal Col As Long, Order As Integer)
    If Col = PC_选择 Then
        Call SelRow
        Order = flexSortNone
    End If
End Sub

Private Sub vsfMain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <= PC_选择 Then
        Cancel = True
    End If
End Sub

Private Sub vsfMain_Click()
    If vsfMain.Col = PC_选择 Then
        vsfMain.ExplorerBar = flexExNone
    Else
        vsfMain.ExplorerBar = flexExSort
    End If
    If vsfMain.Col = PC_选择 Then
        Call SelRow(vsfMain.MouseRow)
    End If
End Sub

Private Sub vsfMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim cbrPopupBar As CommandBar
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call gclsBase.SendLMouseButton(vsfMain.hwnd, x, y)
        Set cbrPopupBar = gclsBase.CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        cbrPopupBar.ShowPopup
    End Select
End Sub

Private Sub vsfMain_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
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
    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.Id = conMenu_FilePopup
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "打印预览(&V)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    '------------------------------------------------------------------------------------------------------------------
    '编辑
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.Id = conMenu_EditPopup
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Word, "收集登记(&S)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "新建登记(&N)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Disuse, "升级完成(&C)", True)
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "差异检查(&J)", True)
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "修改过程(&M)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除过程(&D)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "恢复过程(&R)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Confirm, "确认调整(&T)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Change_PaitNote, "生成脚本(&G)", True)
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.Id = conMenu_ViewPopup
    Set objPopup = gclsBase.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.Id = conMenu_HelpPopup
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objPopup = gclsBase.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrWebSustainer)
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, gstrWebSustainer & "主页(&H)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, gstrWebSustainer & "论坛(&F)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
'    objBar.SetIconSize 16, 16
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Word, "收集")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "新建")
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Audit, "检查", True)
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改", True)
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Untread, "恢复")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Confirm, "确认调整(&T)")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Manage_Change_PaitNote, "生成", True)
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Disuse, "升级完成(&C)", True)
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
    
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新增
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改
        .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除
        
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
        
    End With
End Sub

Private Sub RefreshData()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngRow As Long, lngCurRow As Long
    Dim lng待调整 As Long, lng调整中 As Long
    Dim intState As Integer
    
    intState = cboProcState.ItemData(cboProcState.ListIndex)
    
    '清空原有数据
    strSQL = "Select Id, Decode(类型, 1, '标准过程', 2, '空白过程', 3, '用户过程') As 类型, 名称 As 过程," & vbNewLine & _
            "       Decode(状态,0,'待检查', 1, '待调整', 2, '调整中', 3, '已调整', 4, '无变化') As 状态描述, 状态, 说明, 修改人员, 修改时间, 上次修改人员," & vbNewLine & _
            "       上次修改时间" & vbNewLine & _
            "From Zlprocedure" & vbNewLine & _
            "Where 类型 = [1]" & IIf(intState = -1, "", " And 状态=[2] ")
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取过程清单", mintProcType, intState)
    With vsfMain
        .Redraw = flexRDNone
        .Cell(flexcpPicture, 0, PC_选择) = imgEdit.ListImages("UnCheck").Picture
        .Cell(flexcpPictureAlignment, 0, PC_选择) = flexAlignCenterCenter
        .Rows = vsfMain.FixedRows
        .Tag = ""
        .RowData(0) = rsTmp.RecordCount
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .TextMatrix(lngRow, PC_序号) = lngRow
            .TextMatrix(lngRow, PC_过程) = rsTmp!过程 & ""
            .TextMatrix(lngRow, PC_状态) = rsTmp!状态描述 & ""
            .Cell(flexcpData, lngRow, PC_状态) = rsTmp!状态
            .TextMatrix(lngRow, PC_说明) = rsTmp!说明 & ""
            .RowData(lngRow) = Val(rsTmp!Id & "")
            If .RowData(lngRow) = mlngProcID Then
                lngCurRow = lngRow
            End If
            If .TextMatrix(lngRow, PC_状态) = "待调整" Then
                .Cell(flexcpForeColor, lngRow, PC_状态) = vbRed
                lng待调整 = lng待调整 + 1
            ElseIf .TextMatrix(lngRow, PC_状态) = "调整中" Then
                .Cell(flexcpForeColor, lngRow, PC_状态) = vbBlue
                lng调整中 = lng调整中 + 1
            Else
                .Cell(flexcpForeColor, lngRow, PC_状态) = &H80000008
            End If
            rsTmp.MoveNext
        Loop
        If .Rows <> vsfMain.FixedRows Then
            vsfMain.Row = vsfMain.FixedRows
            If lngCurRow = 0 Then lngCurRow = vsfMain.FixedRows
        End If
        Call vsfMain_AfterRowColChange(-1, -1, lngCurRow, lngCurRow)
        .Redraw = flexRDDirect
    End With
    stbThis.Panels(2).Text = "当前共有待调整 " & lng待调整 & " 个,调整中 " & lng调整中 & " 个。"
End Sub

Private Sub SelRow(Optional ByVal lngRow As Long)
'功能：批量选择vsDetailParas，或取消选择
'          lngRow=0-选择或取消选择所有行，>0选择或取消选择指定行
    Dim blnSel As Boolean, i As Long
    
    With vsfMain
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        If lngRow = 0 Then
            blnSel = Val(.ColData(PC_选择)) = 0
            .Cell(flexcpPicture, lngRow, PC_选择) = imgEdit.ListImages(IIf(blnSel, "AllCheck", "UnCheck")).Picture
            .ColData(PC_选择) = IIf(blnSel, 1, 0) '标记图标状态
            For i = .FixedRows To .Rows - 1
                If Val(.RowData(i)) <> 0 Then
                    .TextMatrix(i, PC_选择) = IIf(blnSel, -1, 0)
                End If
            Next
            If blnSel Then
                .Tag = Val(.RowData(0))
            Else
                .Tag = 0
            End If
        Else
            If Val(.RowData(lngRow)) <> 0 Then
                blnSel = Val(.TextMatrix(lngRow, PC_选择)) = 0
                .TextMatrix(lngRow, PC_选择) = IIf(blnSel, -1, 0)
                .Tag = (Val(.Tag) + IIf(blnSel, 1, -1))
                If Val(.Tag) = 0 Then '所有的都未选择，则将图标更新为批量未勾选
                    .Cell(flexcpPicture, 0, PC_选择) = imgEdit.ListImages("UnCheck").Picture
                    .ColData(PC_选择) = 0
                ElseIf Val(.Tag) = Val(.RowData(0)) Then '所有的都选择，则将图标更新为批量勾选
                    .Cell(flexcpPicture, 0, PC_选择) = imgEdit.ListImages("AllCheck").Picture
                    .ColData(PC_选择) = 1
                End If
            End If
        End If
    End With
End Sub

Private Sub SetSerial()
'功能：生成行序号
    Dim i As Long
    With vsfMain
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, PC_序号) = i
        Next
        If .Rows - 1 > 0 Then
            .Cell(flexcpForeColor, .FixedRows, PC_序号, .Rows - 1, PC_序号) = Color.深灰色
            .Cell(flexcpFontBold, .FixedRows, PC_序号, .Rows - 1, PC_序号) = False
        End If
        If .Row > 0 Then
            .Cell(flexcpFontBold, .Row, PC_序号, .Row, PC_序号) = True
            .Cell(flexcpForeColor, .Row, PC_序号, .Row, PC_序号) = Color.兰色
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FunDeleteProc()
'功能：删除用户存储过程
    Dim i As Long, lngCout As Long
    Dim strProcIDs As String, strProcMsg As String
    Dim intSel As Integer
    Dim arrTmp As Variant, strSQL As String, rsTmp As ADODB.Recordset
    Dim blnOperate As Boolean
    
    On Error GoTo errH
    With vsfMain
        intSel = -1
        If Val(.Tag) = 0 Then  '没有选择项，查看当前行
            If .Row > 0 Then
                strProcIDs = .RowData(.Row)
                strProcMsg = "删除过程：" & .TextMatrix(.Row, PC_过程)
            End If
        ElseIf Val(.Tag) = .RowData(0) Then
            strProcIDs = "*" & mintProcType
            strProcMsg = "删除所有" & Decode(mintProcType, ProcType.变动过程, "变动过程", ProcType.空白过程, "空白过程", "用户过程") & "(共计" & Val(.Tag) & "个)"
        ElseIf Val(.Tag) > .RowData(0) * 0.9 Then '90%的被选择，则采取反向处理
            strProcIDs = "-" & mintProcType
            intSel = 0
        End If
        If strProcMsg = "" Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, PC_选择)) = intSel Then
                    strProcIDs = strProcIDs & "," & .RowData(i)
                End If
                If lngCout < 5 Then
                    If Val(.TextMatrix(i, PC_选择)) = -1 Then
                        lngCout = lngCout + 1
                        strProcMsg = strProcMsg & vbNewLine & .TextMatrix(i, PC_过程)
                    End If
                End If
            Next
            If strProcIDs Like ",*" Then
                strProcIDs = Mid(strProcIDs, 2)
            End If
            strProcMsg = "删除如下过程：" & strProcMsg & vbNewLine & _
                                IIf(lngCout = Val(.Tag), "", "... ..." & vbNewLine & "(共计" & Val(.Tag) & "个)")

        End If
        If MsgBox("确定" & strProcMsg & "吗?", vbInformation + vbOKCancel, "中联软件") = vbOK Then
            If mfrmProcedureRelating Is Nothing Then Set mfrmProcedureRelating = New frmProcRelating
            If Not mfrmProcedureRelating.CheckRelation(Me, strProcIDs) Then Exit Sub
            strSQL = GetProcSQL(strProcIDs)
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "删除存储过程", mintProcType, strProcIDs)
            blnOperate = True: lngCout = rsTmp.RecordCount
            Call ShowFlash("正在删除过程，请稍候！", 0)
            For i = 1 To rsTmp.RecordCount
                Call ShowFlash("正在删除过程【" & rsTmp!名称 & "】", i / lngCout)
                strSQL = "Zl_Zlprocedure_Delete(" & rsTmp!Id & ")"
                Call ExecuteProcedure(strSQL, "删除存储过程")
                rsTmp.MoveNext
            Next
            Call ShowFlash("")
        Else
            Exit Sub
        End If
        blnOperate = False
    End With
    Call RefreshData
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    If blnOperate Then RefreshData
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FunRestoreProc()
'功能：讲变动过程或空白过程恢复为本次标准存储过程
    Dim i As Long, lngCout As Long, rsTmp As ADODB.Recordset
    Dim strProcIDs As String, strProcMsg As String
    Dim intSel As Integer, strPreID As String, strPreName As String
    Dim strSQL As String, strProcText As String
    Dim blnOperate As Boolean
    Dim lngTotal As Long
    
    On Error GoTo errH
    With vsfMain
        intSel = -1
        If Val(.Tag) = 0 Then  '没有选择项，查看当前行
            If .Row > 0 Then
                strProcIDs = .RowData(.Row)
                strProcMsg = "恢复过程：" & .TextMatrix(.Row, PC_过程)
            End If
        ElseIf Val(.Tag) = .RowData(0) Then
            strProcIDs = "*" & mintProcType
            strProcMsg = "恢复所有" & Decode(mintProcType, ProcType.变动过程, "变动过程", ProcType.空白过程, "空白过程", "用户过程") & "(共计" & Val(.Tag) & "个)"
        ElseIf Val(.Tag) > .RowData(0) * 0.9 Then '90%的被选择，则采取反向处理
            strProcIDs = "-" & mintProcType
            intSel = 0
        End If
        If strProcMsg = "" Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, PC_选择)) = intSel Then
                    strProcIDs = strProcIDs & "," & .RowData(i)
                End If
                If lngCout < 5 Then
                    If Val(.TextMatrix(i, PC_选择)) = -1 Then
                        lngCout = lngCout + 1
                        strProcMsg = strProcMsg & vbNewLine & .TextMatrix(.Row, PC_过程)
                    End If
                End If
            Next
            If strProcIDs Like ",*" Then
                strProcIDs = Mid(strProcIDs, 2)
            End If
            strProcMsg = "确定恢复如下过程吗?过程如下：" & strProcMsg & vbNewLine & _
                                IIf(lngCout = Val(.Tag), "", "... ..." & vbNewLine & "(共计" & Val(.Tag) & "个)")

        End If
        If mintProcType <> ProcType.空白过程 Then
            strProcMsg = "操作执行后，在本次升级中，将不再对该过程进行管理并且将数据库中该过程对象恢复为升级之后的标准过程。" & vbNewLine & strProcMsg
        Else
            strProcMsg = "操作会将数据库中该过程对象恢复为升级之后标准过程。" & vbNewLine & strProcMsg
        End If
        If MsgBox(strProcMsg, vbInformation + vbOKCancel, "中联软件") = vbOK Then
            strSQL = "Select a.Id, a.名称, b.序号, b.内容" & vbNewLine & _
                        "From (" & GetProcSQL(strProcIDs) & ") a, Zlproceduretext b" & vbNewLine & _
                        "Where a.Id = b.过程id  And b.性质 = [3]" & vbNewLine & _
                        "Order By a.Id, b.序号"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "恢复存储过程", mintProcType, strProcIDs, ProcTextType.本次标准过程)
            strPreID = "": strSQL = "": strProcText = ""
            blnOperate = True
            Call ShowFlash("正在恢复过程，请稍候！", 0, Me, True)
            lngTotal = Val(.Tag): lngCout = 0
            If lngTotal = 0 Then lngTotal = 1
            Do While Not rsTmp.EOF
                If rsTmp!Id & "" <> strPreID Then
                    If strPreID <> "" Then
                        lngCout = lngCout + 1
                        strProcText = strProcText
                        Call ShowFlash("正在恢复过程【" & strPreName & "】", lngCout / lngTotal)
                        Call gcnOldOra.Execute(strProcText)
                        If mintProcType <> ProcType.空白过程 Then
                            strSQL = "Zl_Zlprocedure_Delete(" & strPreID & ")"
                            Call ExecuteProcedure(strSQL, "删除存储过程")
                        End If
                    End If
                    strPreID = rsTmp!Id
                    strProcText = rsTmp!内容
                    strPreName = rsTmp!名称
                Else
                    strProcText = strProcText & rsTmp!内容
                End If
                rsTmp.MoveNext
            Loop
            If strPreID <> "" Then
                strProcText = strProcText
                Call ShowFlash("正在恢复过程【" & strPreName & "】", 100)
                Call gcnOldOra.Execute(strProcText)
                If mintProcType <> ProcType.空白过程 Then
                    strSQL = "Zl_Zlprocedure_Delete(" & strPreID & ")"
                    Call ExecuteProcedure(strSQL, "删除存储过程")
                End If
            End If
            ShowFlash ("")
        Else
            Exit Sub
        End If
        blnOperate = False
    End With
    Call RefreshData
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    ShowFlash ("")
    If blnOperate Then RefreshData
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FunConfirmProc()
'功能：确认已经调整过程
'功能：讲变动过程或空白过程恢复为本次标准存储过程
    Dim i As Long, lngCout As Long, rsTmp As ADODB.Recordset
    Dim strProcIDs As String, strProcMsg As String
    Dim intSel As Integer
    Dim strSQL As String
    Dim blnOperate As Boolean
    Dim lngTotal As Long
    
    On Error GoTo errH
    With vsfMain
        intSel = -1
        If Val(.Tag) = 0 Then  '没有选择项，查看当前行
            If .Row > 0 Then
                strProcIDs = .RowData(.Row)
                strProcMsg = "确认已经调整如下过程：" & .TextMatrix(.Row, PC_过程)
            End If
        ElseIf Val(.Tag) = .RowData(0) Then
            strProcIDs = "*" & mintProcType
            strProcMsg = "确认已经调整所有用户过程(共计" & Val(.Tag) & "个)"
        ElseIf Val(.Tag) > .RowData(0) * 0.9 Then '90%的被选择，则采取反向处理
            strProcIDs = "-" & mintProcType
            intSel = 0
        End If
        If strProcMsg = "" Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, PC_选择)) = intSel Then
                    strProcIDs = strProcIDs & "," & .RowData(i)
                End If
                If lngCout < 5 Then
                    If Val(.TextMatrix(i, PC_选择)) = -1 Then
                        lngCout = lngCout + 1
                        strProcMsg = strProcMsg & vbNewLine & .TextMatrix(.Row, PC_过程)
                    End If
                End If
            Next
            If strProcIDs Like ",*" Then
                strProcIDs = Mid(strProcIDs, 2)
            End If
            strProcMsg = "确认已经调整如下过程吗?过程如下：" & strProcMsg & vbNewLine & _
                                IIf(lngCout = Val(.Tag), "", "... ..." & vbNewLine & "(共计" & Val(.Tag) & "个)")

        End If
        If MsgBox(strProcMsg, vbInformation + vbOKCancel, "中联软件") = vbOK Then
            strSQL = GetProcSQL(strProcIDs)
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "确认调整存储过程", mintProcType, strProcIDs)
            blnOperate = True
            Call ShowFlash("正在调整过程，请稍候！", 0, Me, True)
            lngTotal = Val(.Tag): lngCout = 0
            If lngTotal = 0 Then lngTotal = 1
            Do While Not rsTmp.EOF
                Call ShowFlash("正在调整过程【" & rsTmp!名称 & "】", lngCout / lngTotal)
                strSQL = "Zl_Zlprocedure_Confirm(" & rsTmp!Id & ")"
                Call ExecuteProcedure(strSQL, "调整存储过程")
                rsTmp.MoveNext
            Loop
            ShowFlash ("")
        Else
            Exit Sub
        End If
        blnOperate = False
    End With
    Call RefreshData
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    ShowFlash ("")
    If blnOperate Then RefreshData
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function GetProcSQL(strIDs As String) As String
'根据ID串获取准备删除或恢复的存储过程
'strIDs=ID条件，*类型 表示该类型的所有对象。-类型,ID1...:表示该类型去掉特定的ID,ID1,...：表示只获取这些ID
    Dim strProcs As String, intType As Integer, strTMp As String
    Dim lngPos As String, i As Integer

    '获取本次检查的存储过程
    If strIDs Like "[*]*" Then
        strProcs = "Select Id, Upper(名称) 名称, Upper(所有者) 所有者 From Zlprocedure Where 类型 = [1]"
        intType = Val(Mid(strIDs, 2))
    ElseIf strIDs Like "-*" Then
        lngPos = InStr(strIDs, ",")
        strTMp = Mid(strIDs, 1, lngPos - 1)
        strIDs = Mid(strIDs, lngPos + 1)
        intType = Val(Mid(strTMp, 2))
        strProcs = "Select Id, Upper(名称) 名称, Upper(所有者) 所有者" & vbNewLine & _
                    "From Zlprocedure a, Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) b" & vbNewLine & _
                    "Where 类型 = [1] And a.Id = b.Column_Value(+) And b.Column_Value Is Null"
    Else
        strProcs = "Select Id, Upper(名称) 名称, Upper(所有者) 所有者" & vbNewLine & _
                        "From Zlprocedure" & vbNewLine & _
                        "Where Id In (Select Column_Value From Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)))"
    End If
    GetProcSQL = strProcs
End Function

Private Function IsSelData() As Boolean
    If vsfMain.Row >= vsfMain.FixedRows Then
        IsSelData = vsfMain.RowData(vsfMain.Row) <> 0
    Else
        IsSelData = False
    End If
End Function

Private Sub PrintProcs(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    
    '表头
    objOut.Title.Text = "过程管理：" & Decode(mintProcType, ProcType.变动过程, "变动过程", ProcType.空白过程, "空白过程", "用户过程") & "(" & cboProcState.Text & ")"
    objOut.Title.Font.name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    Set objRow = New zlTabAppRow
    objRow.Add "时间：" & Format(CurrentDate(), "yyyy-MM-dd HH:mm:ss")
    objOut.UnderAppRows.Add objRow
    
    '表体
    Set objOut.Body = vsfMain
    '输出
    vsfMain.Redraw = False
    lngRow = vsfMain.Row: lngCol = vsfMain.Col
        
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    vsfMain.Row = lngRow: vsfMain.Col = lngCol
    vsfMain.Redraw = True
End Sub

