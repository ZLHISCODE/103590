VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmAppUpgrade 
   BackColor       =   &H80000005&
   Caption         =   "系统升迁管理"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmAppUpgrade.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   10155
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Index           =   0
      Left            =   0
      ScaleHeight     =   5775
      ScaleWidth      =   9735
      TabIndex        =   5
      Top             =   600
      Width           =   9732
      Begin VB.CommandButton cmdkillProcess 
         Caption         =   "中断客户端连接的进程定义(&P)"
         Height          =   350
         Left            =   6315
         TabIndex        =   32
         Top             =   1320
         Width           =   2790
      End
      Begin VB.CommandButton cmdSelALl 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   120
         TabIndex        =   30
         Top             =   5160
         Width           =   1100
      End
      Begin VB.CommandButton cmdNotSel 
         Caption         =   "全消(&R)"
         Height          =   350
         Left            =   1200
         TabIndex        =   29
         Top             =   5160
         Width           =   1100
      End
      Begin MSComctlLib.ProgressBar pgbGetUpInfo 
         Height          =   180
         Left            =   4200
         TabIndex        =   28
         Top             =   2760
         Visible         =   0   'False
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Frame fraUpMode 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   960
         TabIndex        =   24
         Top             =   2700
         Width           =   2175
         Begin VB.OptionButton optUpMode 
            BackColor       =   &H80000005&
            Caption         =   "提前升迁"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   27
            Top             =   60
            Width           =   1215
         End
         Begin VB.OptionButton optUpMode 
            BackColor       =   &H80000005&
            Caption         =   "常规升迁"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   26
            Top             =   60
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdExec 
         Caption         =   "执行(&E)…"
         Height          =   350
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5160
         Width           =   1100
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   2
         Left            =   0
         TabIndex        =   22
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   3
         Left            =   1020
         TabIndex        =   12
         Top             =   2115
         Width           =   5940
      End
      Begin VB.CommandButton cmdFunc 
         Appearance      =   0  'Flat
         Caption         =   "禁用并中断客户端"
         Height          =   350
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   1660
      End
      Begin VB.CommandButton cmdFunc 
         Appearance      =   0  'Flat
         Caption         =   "锁定用户帐号"
         Height          =   350
         Index           =   1
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   1300
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   1
         Left            =   900
         TabIndex        =   8
         Top             =   1155
         Width           =   5940
      End
      Begin VB.CommandButton cmdFunc 
         Appearance      =   0  'Flat
         Caption         =   "禁用触发器"
         Height          =   350
         Index           =   3
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   1100
      End
      Begin VB.CommandButton cmdFunc 
         Appearance      =   0  'Flat
         Caption         =   "禁用后台作业"
         Height          =   350
         Index           =   2
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   1300
      End
      Begin VSFlex8Ctl.VSFlexGrid vsSysSel 
         Height          =   1812
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   9372
         _cx             =   16531
         _cy             =   3196
         Appearance      =   3
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   8
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAppUpgrade.frx":04F9
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
      Begin MSComctlLib.ProgressBar pgbState 
         Height          =   180
         Left            =   960
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComDlg.CommonDialog cdgPub 
         Left            =   3600
         Top             =   1980
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblConfigureFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "升迁不仅需要升级脚本，还要求升迁的目标系统的安装配置文件，以获取对应的管理工具脚本。"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3000
         Width           =   7575
      End
      Begin VB.Label lblUpMode 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "升迁模式"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正在检查"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblExplain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAppUpgrade.frx":05E1
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   10080
      End
      Begin VB.Label lblMainPath 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "系统安装目录：C:\Appsoft"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   2370
         Width           =   2160
      End
      Begin VB.Label lblSel 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "更改…"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   2460
         TabIndex        =   14
         Top             =   2370
         Width           =   540
      End
      Begin VB.Label lblUpgrade 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "升迁执行"
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblManage 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "升迁准备"
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   720
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5052
      Index           =   1
      Left            =   3660
      ScaleHeight     =   5055
      ScaleWidth      =   9615
      TabIndex        =   1
      Top             =   1515
      Width           =   9612
      Begin VB.ComboBox cboSys 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   4560
      End
      Begin VSFlex8Ctl.VSFlexGrid vsUpLog 
         Height          =   3708
         Left            =   120
         TabIndex        =   2
         Top             =   828
         Width           =   9372
         _cx             =   16531
         _cy             =   6540
         Appearance      =   0
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
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAppUpgrade.frx":0686
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
      Begin VB.Label lblSys 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应用系统"
         Height          =   180
         Left            =   165
         TabIndex        =   4
         Top             =   360
         Width           =   720
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   6240
      Left            =   0
      TabIndex        =   20
      Top             =   480
      Width           =   9780
      _Version        =   589884
      _ExtentX        =   17251
      _ExtentY        =   11007
      _StockProps     =   64
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "系统升迁管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   0
      Top             =   105
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   0
      Picture         =   "frmAppUpgrade.frx":0761
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmAppUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum SysSelCol
    Col_Sel = 0
    Col_编号 = 1
    Col_名称 = 2
    Col_配置文件 = 3
    Col_当前版本 = 4
    Col_目标版本 = 5
    Col_检查结果 = 6
End Enum

Private Enum SysUpCol
    Col_序号 = 0
    Col_升迁时间 = 1
    Col_原始版本 = 2
    Col_预期目标 = 3
    Col_结果版本 = 4
    Col_升迁结果 = 5
    Col_提前执行 = 6
End Enum

Private Enum FuncCtrlType
    FCT_客户端 = 0
    FCT_用户账户 = 1
    FCT_后台作业 = 2
    FCT_触发器 = 3
End Enum

Private mrsSysInfo As ADODB.Recordset
Private mrsSysUpFiles As ADODB.Recordset
Private mrsMainPath As ADODB.Recordset

Private mstrSysJobs As String  '手工禁用的系统调度
Private mblnLoadSysFiles As Boolean '是否已经加载ZLSysFiles中的配置文件
Private mblnLastUpInfo As Boolean '是否获取上次升迁历史
'Private mstrMaxUpVer As String '各个系统单独升级时不考虑版本对应时的最大版本。
Private mobjOprateLog As TextStream
'===========================================================================
'==公共接口
'===========================================================================
Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
End Sub

'===========================================================================
'==事件
'===========================================================================
Private Sub cboSys_Click()
    Call LoadData(1)
End Sub

Private Sub cmdExec_Click()
    Dim objfrmUpSys As frmAppUpgradeNew
    Dim strRunModule As String
    
    '系统信息记录集
    Call RecToLog(mrsSysInfo, "系统编号", "原始系统系统记录集")
    If VilidateUpgrade Then
        '业务高峰期检测
        If optUpMode(1).value Then
            If Not CheckRushHours("0102", "提前升迁") Then
                Exit Sub
            End If
        End If
        If MsgBox("请检查当前勾选升级的系统是否符合要求，是否确认升级？", vbInformation + vbYesNo, gstrSysName) = vbYes Then
            If Not CheckCBOPars Then Exit Sub
            Set objfrmUpSys = New frmAppUpgradeNew '用来清除模块变量
            Call objfrmUpSys.ShowMe(frmMDIMain, mrsSysInfo, mrsSysUpFiles, optUpMode(1).value, strRunModule)
            If strRunModule <> "" Then
                Unload Me
                Call frmMDIMain.RunByModule(strRunModule)
                Exit Sub
            End If
            vsSysSel.Tag = ""
            Call LoadSystems
            Call LoadData(0)
            Call VilidateUpgrade
        End If
    Else
        MsgBox "请选择你要升级的应用系统", vbInformation, gstrSysName
    End If
    Call RecToLog(mrsSysInfo, "系统编号", "验证过的系统系统记录集")
End Sub

Private Sub cmdFunc_Click(Index As Integer)
    If cmdFunc(Index).ToolTipText <> "" Then
        MsgBox cmdFunc(Index).ToolTipText, vbInformation, gstrSysName
        Exit Sub
    End If
'    WriteTraceLog
    Call FuncControlExec(Index)
End Sub

Private Sub cmdkillProcess_Click()
    frmKillProcessManage.ShowMe ("0102")
End Sub

Private Sub cmdNotSel_Click()
    Call SetSelBeach
End Sub

Private Sub cmdSelAll_Click()
    Call SetSelBeach(True)
End Sub

Private Sub Form_Activate()
    If tbPage.Item(0).Selected Then
        Call VilidateUpgrade
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF5 Then '刷新界面
        Call tbPage_SelectedChanged(tbPage.Item(IIf(tbPage.Item(0).Selected, 0, 1)))
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errH
    '测试
    WriteTraceLog String(80, "/")
    WriteTraceLog String(4, "/") & "服务器：" & gstrServer
    WriteTraceLog String(4, "/") & "时间：" & Format(CurrentDate, "yyyy-MM-dd HH:MM:SS")
    WriteTraceLog String(80, "/")
    '初始化变量
    tbPage.Tag = "未加载"
    '初始化界面
    tbPage.InsertItem 0, "升迁操作", picMain(0).hwnd, 0
    tbPage.InsertItem 1, "升迁历史", picMain(1).hwnd, 0
    tbPage.Tag = ""
    Call LoadSystems
    Call tbPage_SelectedChanged(tbPage.Item(0))
    cmdkillProcess.Visible = CheckAndAdjustMustTable("zlkillprocess")
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub Form_Resize()
    Dim i As Long
    On Error Resume Next
    tbPage.Height = Me.ScaleHeight - tbPage.Top + 15
    tbPage.Width = Me.ScaleWidth - tbPage.Left + 15
    For i = 0 To 1
        picMain(i).Left = 0
        picMain(i).Width = tbPage.Width - 60
        picMain(i).Height = tbPage.Height - picMain(i).Top
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsSysInfo = Nothing
    Set mrsSysUpFiles = Nothing
    mstrSysJobs = ""
    Set mrsMainPath = Nothing
    mblnLoadSysFiles = False
    mblnLastUpInfo = False
End Sub

Private Sub lblSel_Click()
    Dim strFolderName As String
    Dim strOldPath As String
    strFolderName = lblMainPath.Tag
    
    strFolderName = OpenFolder(Me, "选择系统安装目录")
    If strFolderName = "" Then Exit Sub
    If lblMainPath.Tag <> strFolderName Then
        lblMainPath.Tag = "": lblMainPath.Caption = "系统安装目录："
        Call GetAllSetup(strFolderName)
        Call optUpMode_Click(IIf(optUpMode(0).value, 0, 1))
    End If
End Sub

Private Sub optUpMode_Click(Index As Integer)
    With vsSysSel
        .Cell(flexcpText, .FixedRows, Col_目标版本, .Rows - 1, Col_检查结果) = ""
        .Cell(flexcpForeColor, .FixedRows, Col_目标版本, .Rows - 1, Col_检查结果) = &H80000008
        .Cell(flexcpChecked, .FixedRows, Col_Sel, .Rows - 1, Col_Sel) = True
        Call VilidateUpgrade
    End With
End Sub

Private Sub picMain_Resize(Index As Integer)
    Dim sngWidth As Long '最小宽度
    Dim i As Long
    
    If Index = 1 Then
        pgbState.Visible = False
        lblState.Visible = False
    End If
    On Error Resume Next
    sngWidth = picMain(0).ScaleWidth - 200
    If Index = 1 Then
        cboSys.Width = sngWidth - cboSys.Left - 300
        vsUpLog.Width = sngWidth - vsUpLog.Left - 300
        vsUpLog.Height = picMain(0).ScaleHeight - vsUpLog.Top - 100
    Else
        vsSysSel.Width = sngWidth - vsUpLog.Left - 90
        If vsSysSel.Top + vsSysSel.Rows * vsSysSel.RowHeightMin + cmdSelALl.Height + 200 < picMain(0).ScaleHeight Then
            vsSysSel.Height = vsSysSel.Rows * vsSysSel.RowHeightMin + 30
        Else
            vsSysSel.Height = IIf(vsSysSel.Rows < 13, vsSysSel.Rows, 12) * vsSysSel.RowHeightMin + 30
        End If
        lblExplain.Width = vsSysSel.Width
        lblExplain.Refresh
        '系统控制标签的位置设置
        Call SetCtrlPosOnLine(True, -1, lblExplain, 60, lblManage, 60, cmdFunc(FCT_客户端), 90, lblState, 60, lblUpgrade, 60, lblMainPath, 90, lblUpMode, 90, lblConfigureFile, 30, vsSysSel)

        fraSplit(0).Left = -30: fraSplit(0).Width = lblManage.Left - fraSplit(0).Left
        Call SetCtrlPosOnLine(False, 0, lblManage, -1 * (lblManage.Width + fraSplit(0).Width), fraSplit(0), lblManage.Width, fraSplit(1))
        fraSplit(1).Width = picMain(0).ScaleWidth - fraSplit(1).Left + 100

        Call SetCtrlPosOnLine(False, 0, cmdFunc(FCT_客户端), 120, cmdFunc(FCT_用户账户), 120, cmdFunc(FCT_后台作业), 120, cmdFunc(FCT_触发器), 720, cmdkillProcess)
        Call SetCtrlPosOnLine(False, 0, lblState, 0, pgbState)

        fraSplit(2).Left = -30: fraSplit(2).Width = lblUpgrade.Left - fraSplit(2).Left
        Call SetCtrlPosOnLine(False, 0, lblUpgrade, -1 * (lblUpgrade.Width + fraSplit(2).Width), fraSplit(2), lblUpgrade.Width, fraSplit(3))
        fraSplit(3).Width = picMain(0).ScaleWidth - fraSplit(3).Left + 100

        Call SetCtrlPosOnLine(False, 0, lblMainPath, 120, lblSel)
        Call SetCtrlPosOnLine(False, 0, lblUpMode, 120, fraUpMode, 120, pgbGetUpInfo)
        Call SetCtrlPosOnLine(True, 1, vsSysSel, 90, cmdExec)
        Call SetCtrlPosOnLine(True, -1, vsSysSel, 90, cmdSelALl)
        Call SetCtrlPosOnLine(False, 0, cmdSelALl, 60, cmdNotSel)
    End If
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If tbPage.Tag = "" Then
        Call LoadData(Item.Index)
        picMain_Resize (Item.Index)
    End If
End Sub

Private Sub vsSysSel_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsSysSel
        .ComboList = ""
        .FocusRect = flexFocusLight
         .ToolTipText = ""
        If NewCol = Col_配置文件 Then
             .ComboList = "..."
             .FocusRect = flexFocusSolid
        End If
    End With
End Sub

Private Sub vsSysSel_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strFile As String
    Dim strMainPath As String

    If Col = Col_配置文件 Then
        With cdgPub
            .DialogTitle = "选择应用安装配置文件"
            If Trim(vsSysSel.TextMatrix(Row, Col_编号)) = "" Then
                .Filter = "服务器工具脚本(zlServer.Sql)|zlServer.Sql"
            Else
                .Filter = "应用安装配置文件(zlSetup.ini)|zlSetup.ini"
                .Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
            End If
            strFile = IIf(Mid(vsSysSel.TextMatrix(Row, Col), 1, 1) = "$", lblMainPath.Tag & Mid(vsSysSel.TextMatrix(Row, Col), 2), vsSysSel.TextMatrix(Row, Col))
            If gobjFile.FileExists(strFile) Then
                .InitDir = gobjFile.GetParentFolderName(strFile)
                .Filename = gobjFile.GetFileName(strFile)
            Else
                .InitDir = "": .Filename = ""
                If vsSysSel.Cell(flexcpData, Row, Col) <> "" Then
                    If gobjFile.FolderExists(gobjFile.GetParentFolderName(vsSysSel.Cell(flexcpData, Row, Col))) Then
                        .InitDir = gobjFile.GetParentFolderName(vsSysSel.Cell(flexcpData, Row, Col))
                    End If
                End If
            End If
            On Error Resume Next
            .CancelError = True
            .ShowOpen
            err.Clear: On Error GoTo errH
            If .Filename = gobjFile.GetFileName(strFile) Then .Filename = ""
            If .Filename <> "" And .Filename <> "zlSetup.ini" And .Filename <> "zlServer.Sql" Then
                If .Filename <> vsSysSel.Cell(flexcpData, Row, Col) Then
                    '配置文件改变，检查配置文件
                    If CheckInitFile(Val(vsSysSel.TextMatrix(Row, Col_编号)), .Filename) Then
                        vsSysSel.TextMatrix(Row, Col) = .Filename
                         vsSysSel.Cell(flexcpData, Row, Col) = .Filename
                        Call ReSetMainPath(Row)
                        Call VilidateUpgrade(Row)
                    End If
                End If
            End If
            On Error GoTo 0
        End With
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub vsSysSel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (vsSysSel.MouseCol = Col_目标版本 Or vsSysSel.MouseCol = Col_检查结果) And vsSysSel.MouseRow >= vsSysSel.FixedRows Then
        If vsSysSel.TextMatrix(vsSysSel.MouseRow, Col_检查结果) <> "" Then
            vsSysSel.ToolTipText = vsSysSel.TextMatrix(vsSysSel.MouseRow, Col_检查结果)
        Else
            vsSysSel.ToolTipText = ""
        End If
    Else
        vsSysSel.ToolTipText = ""
    End If
End Sub

Private Sub vsSysSel_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = Col_Sel And Row > vsSysSel.FixedRows Or Col = Col_配置文件) Then Cancel = True
End Sub

Private Sub vsUpLog_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsUpLog
        If NewRow >= .FixedRows Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, NewCol)
        End If
    End With
End Sub

'===========================================================
'方法
'===========================================================
Private Sub LoadSystems()
'功能：加载应用系统
'参数：intPageIndex=0，升迁页系统添加，intPageIndex=1,升迁历史页系统添加
    Dim strSQL As String, rsSys As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strVer As String
    Dim i As Long
    On Error GoTo errH
    '获取管理工具版本号
    strVer = GetToolsVersion
    '增加共享号排序，主要是将主系统排在前面
    strSQL = "Select 编号 系统编号, 名称 系统名称, 版本号 系统版本号, 所有者 系统所有者, 共享号, 正常安装 From Zlsystems where Upper(所有者)=[1] Order by Nvl(共享号,0),编号"
    Set rsSys = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "读取安装系统", gstrUserName)
    With rsSys
        '添加管理工具历史记录查看。
        cboSys.Clear
        cboSys.addItem String(5, " ") & RPAD("服务器管理工具", 18) & " v" & VerPAD(strVer)
        cboSys.ItemData(cboSys.NewIndex) = -1
        Do While Not .EOF
            If Val(Split(!系统版本号, ".")(0)) > 9 Then
                    cboSys.addItem Lpad(!系统编号, 4) & "-" & RPAD(!系统名称 & "", 18) & " v" & VerPAD(!系统版本号 & "")
                    cboSys.ItemData(cboSys.NewIndex) = !系统编号
                    If cboSys.ListIndex = -1 And UCase(!系统所有者 & "") = UCase(gstrUserName) Then
                        cboSys.ListIndex = cboSys.NewIndex
                    End If
            End If
            .MoveNext
        Loop
        If cboSys.ListIndex = -1 Then cboSys.ListIndex = 0
    End With
    If rsSys.RecordCount <> 0 Then rsSys.MoveFirst
    '填写已安装系统清单
    With vsSysSel
        '目标版本，最终版本为系统单独升级时的本次升迁目标以及最终目标
        Set mrsSysInfo = CopyNewRec(rsSys, True, "系统编号,系统名称,系统版本号,系统所有者,共享号,正常安装", Array("Sort", adInteger, 2, 0, "升级", adInteger, 1, 0, "配置文件", adVarChar, 2000, Empty, _
                                                                                       "目标版本", adVarChar, 20, Empty, "提前目标版本", adVarChar, 20, Empty, "最终版本", adVarChar, 20, Empty, _
                                                                                        "升迁结果", adInteger, 1, 0, "中止信息", adVarChar, 2000, Empty, "可升级", adInteger, 1, 0, "检查结果", adVarChar, 2000, Empty, _
                                                                                        "提前升迁结果", adInteger, 1, 0, "提前中止信息", adVarChar, 2000, Empty, "可提前升级", adInteger, 1, 0, "提前检查结果", adVarChar, 2000, Empty))
        .Rows = .FixedRows
        '获取管理工具版本号
        strVer = GetToolsVersion
        mrsSysInfo.AddNew Array("系统编号", "系统名称", "系统版本号", "系统所有者", "共享号", "正常安装", "Sort", "配置文件", "可升级", "可提前升级", "升级"), _
                                        Array(0, "管理工具", strVer, "ZLTOOLS", Null, 1, .Rows, Null, 1, 1, 1)
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, Col_Sel) = IIf(strVer & "" = "", 0, 1)
        .TextMatrix(.Rows - 1, Col_编号) = ""
        .TextMatrix(.Rows - 1, Col_名称) = "服务器管理工具"
        .TextMatrix(.Rows - 1, Col_当前版本) = VerPAD(strVer & "")
        .TextMatrix(.Rows - 1, Col_检查结果) = ""
        .Cell(flexcpForeColor, .Rows - 1, Col_Sel, .Rows - 1, .Cols - 1) = IIf(strVer & "" = "", vbRed, vbBlue)
        Do While Not rsSys.EOF
            If Val(Split(rsSys!系统版本号, ".")(0)) > 9 Then
                mrsSysInfo.AddNew Array("系统编号", "系统名称", "系统版本号", "系统所有者", "共享号", "正常安装", "Sort", "配置文件", "可升级", "可提前升级", "升级"), _
                                                Array(rsSys!系统编号, rsSys!系统名称, rsSys!系统版本号, rsSys!系统所有者, rsSys!共享号, rsSys!正常安装, .Rows, Null, 1, 1, 1)
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, Col_Sel) = 1
                .TextMatrix(.Rows - 1, Col_编号) = rsSys!系统编号 & ""
                .Cell(flexcpData, .Rows - 1, Col_编号) = Val(rsSys!共享号 & "")
                .TextMatrix(.Rows - 1, Col_名称) = rsSys!系统名称 & ""
                .TextMatrix(.Rows - 1, Col_当前版本) = VerPAD(rsSys!系统版本号 & "")
                .TextMatrix(.Rows - 1, Col_检查结果) = ""
            End If
            rsSys.MoveNext
        Loop
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, Col_名称) <> 0 Then
                mrsSysInfo.Filter = "系统编号=" & .Cell(flexcpData, i, Col_名称)
                .RowData(i) = Val(mrsSysInfo!序号 & "")
            End If
        Next
        Call GetLastUpgrade
    End With
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub GetLastUpgrade()
'功能：获取上次升迁信息
    Dim rsUpgrade As ADODB.Recordset
    Dim strSQL As String, strFilter As String
    Dim lngSys As Long
    Dim i As Long
    
    On Error GoTo errH
    mblnLastUpInfo = False
    '检查ZLUPGRADE表及其字段”提前执行“
    If Not CheckAndAdjustMustTable("ZLUPGRADE", "提前执行", True) Then
        Exit Sub
    End If
    If cboSys.ListCount > 1 Then
        '检查表ZLBAKSPACES
        If Not CheckAndAdjustMustTable("ZLBAKSPACES", , True) Then
            Exit Sub
        End If
        '检查表ZLBAKTABLES
        If Not CheckAndAdjustMustTable("ZLBAKTABLES", , True) Then
            Exit Sub
        End If
    End If
    mblnLastUpInfo = True
    '获取所有系统上次升迁以及上次提前升迁信息
    strSQL = "Select Nvl(系统,0) 系统编号 , 提前执行, 中止语句, 升迁结果, 结果版本" & vbNewLine & _
                    "From (Select 系统, 提前执行, 升迁时间, 中止语句, 升迁结果, 结果版本, Max(升迁时间) Over(Partition By 系统, Decode(提前执行, Null, -1, 0)) 当前时间" & vbNewLine & _
                    "       From Zlupgrade) a" & vbNewLine & _
                    "Where A.升迁时间 = A.当前时间" & vbNewLine & _
                    "Order By 系统"
    Set rsUpgrade = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取上次升迁信息")
    
    For i = vsSysSel.FixedRows To vsSysSel.Rows - 1
        lngSys = Val(vsSysSel.TextMatrix(i, Col_编号))
        strFilter = "系统编号 = " & lngSys
        mrsSysInfo.Filter = strFilter
        '系统上次执行升迁信息
        rsUpgrade.Filter = strFilter & " And  提前执行=Null"
        If Not rsUpgrade.EOF Then
            mrsSysInfo.Update Array("升迁结果", "中止信息"), Array(rsUpgrade!升迁结果, FormatUpgradeBreak(lngSys, rsUpgrade!结果版本 & "", rsUpgrade!中止语句 & ""))
            '系统最近一次正常升级不成功，不能进行提前执行
            If Val(rsUpgrade!升迁结果 & "") = 1 Then
                mrsSysInfo.Update Array("可提前升级", "提前检查结果"), Array(0, "系统最近一次正常升级不成功，不能进行提前执行！")
            End If
        Else
            mrsSysInfo.Update Array("升迁结果", "中止信息"), Array(0, FormatUpgradeBreak(lngSys, mrsSysInfo!系统版本号 & ""))
        End If
        '系统上次执行提前升迁信息
        rsUpgrade.Filter = strFilter & " And 提前执行<>Null"
        If Not rsUpgrade.EOF Then
            mrsSysInfo.Update Array("提前升迁结果", "提前中止信息"), Array(rsUpgrade!升迁结果, FormatUpgradeBreak(lngSys, rsUpgrade!结果版本 & "", rsUpgrade!中止语句 & ""))
        Else
            mrsSysInfo.Update Array("提前升迁结果", "提前中止信息"), Array(0, FormatUpgradeBreak(lngSys, mrsSysInfo!系统版本号 & ""))
        End If
    Next
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub LoadData(ByVal intPageIdx As Integer)
'功能：数据加载
'    intPageIdx=页面索引，1-升迁页面,0-日志界面
    Dim blnUp As Boolean
    
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strCtrlInfo As String
    Dim arrInfo As Variant, arrTmp As Variant, i As Long
    Dim lngSys As Long
    Dim intIndex As Long, strCaption As String, strTag As String
    
    On Error GoTo errH
    If intPageIdx = 1 Then
        lngSys = cboSys.ItemData(cboSys.ListIndex)
        If lngSys = Val(cboSys.Tag) Then Exit Sub
        cboSys.Tag = lngSys
        strSQL = "Select * From zlUpgrade Where " & IIf(lngSys = -1, "系统 Is Null ", "系统=[1] ") & " Order by 升迁时间"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取升迁历史", lngSys)
        With vsUpLog
            .Rows = 1
            On Error Resume Next
            Do While Not rsTmp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, Col_序号) = .Rows - 1
                .TextMatrix(.Rows - 1, Col_升迁时间) = Format(rsTmp!升迁时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(.Rows - 1, Col_原始版本) = VerPAD(rsTmp!原始版本 & "")
                .TextMatrix(.Rows - 1, Col_预期目标) = VerPAD(rsTmp!目标版本 & "")
                .TextMatrix(.Rows - 1, Col_结果版本) = VerPAD(rsTmp!结果版本 & "")
                .TextMatrix(.Rows - 1, Col_升迁结果) = IIf(Nvl(rsTmp!升迁结果, 0) = 0, "正常完成", "中途中止")
                '可能没有提前执行这一列
                .TextMatrix(.Rows - 1, Col_提前执行) = rsTmp!提前执行 & ""
                If rsTmp!提前执行 & "" <> "" Then
                    .TextMatrix(.Rows - 1, Col_提前执行) = "√"
                End If
                If Nvl(rsTmp!升迁结果, 0) <> 0 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                End If
                rsTmp.MoveNext
            Loop
            err.Clear: On Error GoTo errH
            If .Rows > 1 Then
                .Row = .Rows - 1
                .ShowCell .Row, .Col
            End If
        End With
    Else
        '仅有管理工具情况:
        '1、没有安装其他应用系统
        '2、安装了系统，但是以其他用户登录如SYS或其他用户登录，非系统所有者
        '此时需要做如下处理
        If vsSysSel.Tag = "" Then
            lblMainPath.Tag = App.Path
            lblMainPath.Caption = "系统安装目录：" & App.Path
            Call GetAllSetup
            vsSysSel.Tag = "已经加载"
        End If
        '结构验证,程序已经缓存对象第一次检查的结果，因此不用担心性能问题，每个内置的对象检查已经缓存检查结果
        '检查Zlupgradeconfig表
        If Not CheckAndAdjustMustTable("Zlupgradeconfig", , True) Then
            For i = 0 To 3
                cmdFunc(i).ToolTipText = "表Zlupgradeconfig缺失，无法执行该功能！"
            Next
            Exit Sub
        End If
        '检查Zlclients表与“系统升级禁用”字段
        If Not CheckAndAdjustMustTable("Zlclients", "系统升级禁用", True) Then
            cmdFunc(FCT_客户端).ToolTipText = "表Zlclients缺失或者表Zlclients的字段""系统升级禁用""缺失，无法执行该功能！"
        End If
        '检查ZLAutoJob表与“系统升级停用”字段
        If Not CheckAndAdjustMustTable("ZLAutoJobs", "系统升级停用", True) Then
            cmdFunc(FCT_后台作业).ToolTipText = "表ZLAutoJob缺失或者表ZLAutoJob的字段""系统升级停用""缺失，无法执行该功能！"
        End If
        If cboSys.ListCount > 1 Then
            '检查上机人员表表与“系统升级锁定”字段
            If Not CheckAndAdjustMustTable("上机人员表", "系统升级锁定", True, gstrUserName) Then
                cmdFunc(FCT_用户账户).ToolTipText = "表上机人员表缺失或者表上机人员表的字段""系统升级锁定""缺失，无法执行该功能！"
            End If
        Else
            cmdFunc(FCT_用户账户).ToolTipText = "请以应用所有者登录管理工具执行该功能！"
        End If
        '检查ZLTriggers表
        If Not CheckAndAdjustMustTable("ZLTriggers", , True) Then
            cmdFunc(FCT_触发器).ToolTipText = "表ZLTriggers缺失，无法执行该功能！"
        End If
        '获取系统升迁准备操作按钮状态并设置按钮状态
        strSQL = "Select 项目, 内容" & vbNewLine & _
                        "From Zlupgradeconfig" & vbNewLine & _
                        "Where 项目 In ('客户端状态', '用户状态', '后台作业状态', '触发器状态', '禁用的系统调度')"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取系统升迁准备操作按钮状态")
        Do While Not rsTmp.EOF
            intIndex = -1
            Select Case rsTmp!项目 & ""
                Case "客户端状态"
                    strCaption = IIf(Val(rsTmp!内容 & "") = 1, "禁用并中断客户端", "启用客户端"): intIndex = FCT_客户端
                Case "用户状态"
                    strCaption = IIf(Val(rsTmp!内容 & "") = 1, "锁定用户帐号", "恢复用户帐号"): intIndex = FCT_用户账户
                Case "后台作业状态"
                    strCaption = IIf(Val(rsTmp!内容 & "") = 1, "禁用后台作业", "启用后台作业"): intIndex = FCT_后台作业
                Case "触发器状态"
                    strCaption = IIf(Val(rsTmp!内容 & "") = 1, "禁用触发器", "启用触发器"): intIndex = FCT_触发器
                Case "禁用的系统调度"
                    mstrSysJobs = rsTmp!内容 & ""
            End Select
            strTag = Val(rsTmp!内容 & "") & ""
            If intIndex <> -1 Then
                cmdFunc(intIndex).Caption = strCaption: cmdFunc(intIndex).Tag = strTag
            End If
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub ReSetMainPath(Optional ByVal lngRow As Long = -1)
'功能：主路径没有被使用则自动重置主路径，使用主路径的路径自动修改为简写模式
'        :lngRow=当前修改行
    Dim blnRest As Boolean '是否重置路径
    Dim i As Long, lngTmpRow As Long
    Dim strMainPath As String, strMainPathOld As String
    Dim blnNotHave As Boolean
    
    On Error GoTo errH
    With vsSysSel
        blnRest = True
        If lblMainPath.Tag <> "" Then
            If lngRow >= .FixedRows Then
                If .TextMatrix(lngRow, Col_配置文件) = "" Then lngRow = -1
            End If
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, Col_配置文件) <> "" Then
                    If UCase(Mid(.TextMatrix(i, Col_配置文件), 1, Len(lblMainPath.Tag) + 1)) = UCase(lblMainPath.Tag) & "\" Then
                        .TextMatrix(i, Col_配置文件) = "$" & Mid(.TextMatrix(i, Col_配置文件), Len(lblMainPath.Tag) + 1)
                        blnRest = False
                    ElseIf Mid(.TextMatrix(i, Col_配置文件), 1, 1) = "$" Then
                        blnRest = False
                    End If
                    If lngTmpRow = 0 Then lngTmpRow = i
                End If
            Next
        End If
        If blnRest Then
            On Error Resume Next
            If lngRow >= lngTmpRow Then
                strMainPath = gobjFile.GetFile(.Cell(flexcpData, lngRow, Col_配置文件)).ParentFolder.ParentFolder.ParentFolder
            Else
                strMainPath = gobjFile.GetFile(.Cell(flexcpData, lngTmpRow, Col_配置文件)).ParentFolder.ParentFolder
            End If
            If err.Number <> 0 Then
                err.Clear
            End If
            On Error GoTo errH
            If strMainPath <> "" Then
                '更改主路径
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, Col_配置文件) <> "" Then
                        If UCase(Mid(.TextMatrix(i, Col_配置文件), 1, Len(strMainPath) + 1)) = UCase(strMainPath) & "\" Then '应用程序安装路径在用，则不改变
                            .TextMatrix(i, Col_配置文件) = "$" & Mid(.TextMatrix(i, Col_配置文件), Len(strMainPath) + 1)
                        End If
                    End If
                Next
                lblMainPath.Tag = strMainPath
                lblMainPath.Caption = "系统安装目录：" & strMainPath
            End If
        End If
    End With
    Call SetCtrlPosOnLine(False, 0, lblMainPath, 120, lblSel)
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub GetAllSetup(Optional ByVal strMainPath As String)
'功能：获取ZLSOFT软件各个系统安装配置文件
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim rsMainPath As ADODB.Recordset
    Dim strPath As String
    Dim strFile As String
    Dim i As Integer, blnAdd As Boolean
    
    On Error GoTo errH
    '清空上次内容
    vsSysSel.Cell(flexcpText, vsSysSel.FixedRows, Col_配置文件, vsSysSel.Rows - 1, Col_当前版本 - 1) = ""
    vsSysSel.Cell(flexcpData, vsSysSel.FixedRows, Col_配置文件, vsSysSel.Rows - 1, Col_当前版本 - 1) = ""
    vsSysSel.Cell(flexcpText, vsSysSel.FixedRows, Col_当前版本 + 1, vsSysSel.Rows - 1, vsSysSel.Cols - 1) = ""
    vsSysSel.Cell(flexcpData, vsSysSel.FixedRows, Col_当前版本 + 1, vsSysSel.Rows - 1, vsSysSel.Cols - 1) = ""
    mblnLoadSysFiles = False
    '获取安装配置文件与候选主目录
    If mrsMainPath Is Nothing Or strMainPath <> "" Then
        Set mrsMainPath = CopyNewRec(Nothing, True, , Array("序号", adInteger, 3, 0, "系统编号", adInteger, 5, 0, "路径", adVarChar, 2000, Empty))
        On Error Resume Next
        '0、若执定主目录，则该主目录优先
        If strMainPath <> "" Then
            mrsMainPath.AddNew Array("序号", "系统编号", "路径"), Array(1, 0, UCase(strMainPath))
        End If
        '1、优先通过通过注册表确定,注册表优先是由于系统最近安装系统总会产生注册信息
        strPath = GetSetting("ZLSOFT", "公共全局", "程序路径")
        strPath = gobjFile.GetFile(strPath).ParentFolder
        If err.Number = 0 Then
            mrsMainPath.Filter = "路径='" & UCase(strPath) & "'"
            If mrsMainPath.EOF Then mrsMainPath.AddNew Array("序号", "系统编号", "路径"), Array(2, 0, UCase(strPath))
        Else
            err.Clear
        End If
        '通过系统目录读取
        strPath = gobjFile.GetFolder(Mid(gobjFile.GetSpecialFolder(WindowsFolder), 1, 1) & ":\APPSOFT")
        If err.Number = 0 Then
            mrsMainPath.Filter = "路径='" & UCase(strPath) & "'"
            If mrsMainPath.EOF Then mrsMainPath.AddNew Array("序号", "系统编号", "路径"), Array(3, 0, UCase(strPath))
        Else
            err.Clear
        End If
        '2、任意10版本系统的安装配置文件确定
        '3、通过任意10版本系统的升级配置文件确定
        strSQL = "Select A.系统 系统编号, A.操作, A.文件名 From Zlsysfiles a Where  A.操作 in(1,2) Order By 系统,操作"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取系统升级安装的配置文件")
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!操作 & "") = 1 Then
                strPath = gobjFile.GetFile(rsTmp!文件名 & "").ParentFolder.ParentFolder.ParentFolder
                strFile = rsTmp!文件名 & ""
            Else
                strPath = gobjFile.GetFile(rsTmp!文件名 & "").ParentFolder.ParentFolder.ParentFolder.ParentFolder
                strFile = gobjFile.GetFile(rsTmp!文件名 & "").ParentFolder.ParentFolder.ParentFolder & "\应用脚本\ZLSETUP.INI"
            End If
            If err.Number = 0 Then
                mrsMainPath.Filter = "路径='" & UCase(strPath) & "' And 系统编号=0"
                If mrsMainPath.EOF Then mrsMainPath.AddNew Array("序号", "系统编号", "路径"), Array(i + 3, 0, UCase(strPath))
                If Not gobjFile.FileExists(strFile) Then strFile = ""
            Else
                err.Clear
                strFile = ""
            End If
            If strFile <> "" Then
                mrsMainPath.Filter = "路径='" & UCase(strFile) & "' And 系统编号=" & rsTmp!系统
                If mrsMainPath.EOF Then mrsMainPath.AddNew Array("序号", "系统编号", "路径"), Array(i + 4, rsTmp!系统编号, UCase(strFile))
            End If
            rsTmp.MoveNext
        Next
    End If
    mrsMainPath.Filter = "系统编号<>0"
    mblnLoadSysFiles = mrsMainPath.RecordCount = 0 '没有读取到ZLSysFiles，则默认已经加载
    mrsMainPath.Filter = "系统编号=0"
    mrsMainPath.Sort = "序号,路径"
    If mrsMainPath.RecordCount <> 0 Then
        blnAdd = strMainPath = ""
        For i = 0 To mrsMainPath.RecordCount - 1
            If mrsMainPath!路径 & "" <> "" Then
                If GetSetupInit(mrsMainPath!路径 & "", blnAdd) Then Exit For
                If blnAdd Then blnAdd = Not blnAdd
            End If
            mrsMainPath.MoveNext
        Next
        '对主路径进行特殊字符标识，没有使用主路径，则自动更换
        Call ReSetMainPath
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function GetSetupInit(Optional ByVal strMainPath As String, Optional ByRef blnAdd As Boolean) As Boolean
'功能：获取各个系统的安装配置文件
'参数：strMainPath="",通过系统文件ZLSysFiles读取文件，<>""，通过+路径读取文件
'           blnAdd=是否只读取未读取的系统的配置文件
    Dim strPath As String, lngCurSys As Long
    Dim strFile As String
    Dim blnGet As Boolean, blnAllGet As Boolean, blnToolsGet As Boolean, blnSysFileGet As Boolean
    Dim strTmp As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long
    
    With vsSysSel
        '自动读取时，主动按上次保存的ZLSysFiles读取
        If blnAdd And Not mblnLoadSysFiles Then Call LoadSysFiles
        '自动读取或主目录读取
        blnAllGet = True
        For i = .FixedRows To .Rows - 1
            lngCurSys = Val(.TextMatrix(i, Col_编号))
            If blnAdd And .TextMatrix(i, Col_配置文件) = "" Or Not blnAdd Then
                If lngCurSys = 0 Then
                    strTmp = "\TOOLS\ZLSERVER.SQL"
                    strFile = strMainPath & strTmp
                Else
                    strTmp = "\" & Decode(lngCurSys \ 100, 1, "ZLHIS10", 3, "ZLMEDREC10", 4, "ZLMATERIAL10", _
                                                                                6, "ZLDEVICE10", 21, "ZLPEIS10", 22, "ZLBLOOD10", _
                                                                                23, "ZLINFECT10", 24, "ZLOPER10", _
                                                                                25, "ZLLIS10", 26, "ZLPSS10", 27, "ZLHEC10") & "\应用脚本\ZLSETUP.INI"
                    strFile = strMainPath & strTmp
                End If
                If gobjFile.FileExists(strFile) Then
                    If CheckInitFile(lngCurSys, strFile, True) Then
                        .Cell(flexcpData, i, Col_配置文件) = gobjFile.GetFile(strFile).Path
                        .TextMatrix(i, Col_配置文件) = .Cell(flexcpData, i, Col_配置文件)
                        blnGet = True
                    End If
                End If
                If .TextMatrix(i, Col_配置文件) = "" Then blnAllGet = False
            End If
            '是否获取了管理工具配置文件
            If .TextMatrix(i, Col_配置文件) <> "" And lngCurSys = 0 Then
                blnToolsGet = True
            End If
        Next
        '手工指定主目录后加载ZLSYsFiles中的配置文件
        If Not blnAdd And Not mblnLoadSysFiles Then
            blnSysFileGet = LoadSysFiles
            blnAllGet = blnSysFileGet And blnToolsGet
        End If
        If Not blnAdd And Not blnGet Then
            MsgBox "在主目录" & strMainPath & "下未找到任何系统安装配置文件，系统将自动读取安装配置文件。"
        Else
            '设置主目录
            If blnGet And lblMainPath.Tag = "" Then
                lblMainPath.Tag = gobjFile.GetFolder(strMainPath).Path
                lblMainPath.Caption = "系统安装目录：" & lblMainPath.Tag
            End If
        End If
    End With

    GetSetupInit = blnAllGet
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Function

Private Function LoadSysFiles() As Boolean
'功能：加载ZLSysFiles中的记录的安装配置文件
    Dim blnAllGet As Boolean, i As Long
    Dim lngCurSys As Long
    
    On Error GoTo errH
    With vsSysSel
        blnAllGet = True
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Col_配置文件) = "" Then
                lngCurSys = Val(.TextMatrix(i, Col_编号))
                If lngCurSys <> 0 Then
                    mrsMainPath.Filter = "系统编号=" & lngCurSys
                    mrsMainPath.Sort = "序号"
                    Do While Not mrsMainPath.EOF
                        If gobjFile.FileExists(mrsMainPath!路径 & "") Then
                            If CheckInitFile(lngCurSys, mrsMainPath!路径 & "", True) Then
                                .Cell(flexcpData, i, Col_配置文件) = gobjFile.GetFile(mrsMainPath!路径 & "").Path
                                .TextMatrix(i, Col_配置文件) = gobjFile.GetFile(mrsMainPath!路径 & "").Path
                                Exit Do
                            End If
                        End If
                        mrsMainPath.MoveNext
                    Loop
                    If .TextMatrix(i, Col_配置文件) = "" Then blnAllGet = False
                End If
            End If
        Next
    End With
    mblnLoadSysFiles = True
    LoadSysFiles = blnAllGet
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub FuncControlExec(ByVal fctTmp As FuncCtrlType)
'功能：执行升级前置功能
'参数：fctTmp=功能类型
    Dim strInfoName As String
    Dim blnLock As Boolean '是否锁定对象
    Dim strLog As String
    Dim rsSQL As ADODB.Recordset
    Dim strStep As String
    Dim cnDBA As ADODB.Connection
    '获取操作日志对象
    '随用随创建，使用完后关闭，防止一台电脑开多个管理工具相互占用日志文件
    strLog = GetLogPath(LT_系统控制)
    If mobjOprateLog Is Nothing Then
        If gobjFile.FileExists(strLog) Then
            Set mobjOprateLog = gobjFile.OpenTextFile(strLog, ForAppending)
        Else
            Set mobjOprateLog = gobjFile.CreateTextFile(strLog, True)
        End If
    End If
    On Error GoTo errH
    '获取信息名
    strInfoName = Decode(fctTmp, FCT_客户端, "客户端状态", FCT_用户账户, "用户状态", FCT_后台作业, "后台作业状态", FCT_触发器, "触发器状态")
    '按钮状态验证，防止重复处理，若状态未进过验证，则提示是否继续处理
    If Not FuncControlStateVali(fctTmp, strInfoName) Then Exit Sub
    blnLock = Val(cmdFunc(fctTmp).Tag) = 1 '当前是启用状态，则禁用对象
    '获取执行SQL
    Set rsSQL = FuncControlSQL(fctTmp, blnLock, strInfoName)
    rsSQL.Filter = "DBA=1"
    If Not rsSQL.EOF Then
        Set cnDBA = GetConnection("System")
        If cnDBA Is Nothing Then Exit Sub
    End If
    mobjOprateLog.WriteLine String(120, "=")
    mobjOprateLog.WriteLine strInfoName & IIf(blnLock, " 启用-->禁用", " 禁用-->启用")
    mobjOprateLog.WriteLine "执行时间：" & Format(CurrentDate, "yyyy-MM-dd HH:mm:ss")
    mobjOprateLog.WriteLine String(120, "=")
    '执行SQL
    rsSQL.Filter = ""
    rsSQL.Sort = "序号"
    On Error Resume Next
    Do While Not rsSQL.EOF
        Call SetProgressState(Val(rsSQL!进度 & ""), rsSQL!步骤 & "")
        mobjOprateLog.WriteLine "执行SQL:" & rsSQL!内容
        If rsSQL!DBA = 1 Then
            cnDBA.Execute rsSQL!内容
        Else
            gcnOracle.Execute rsSQL!内容
        End If
        If err.Number <> 0 Then
            mobjOprateLog.WriteLine "执行结果:" & err.Description
            err.Clear
        Else
            mobjOprateLog.WriteLine "执行结果:执行成功" & err.Description
        End If
        rsSQL.MoveNext
    Loop
    On Error GoTo errH
    '处理标记，防止再次处理以及并发,并刷新界面，重置按钮状态
    Call LoadData(0)
    '处理完成
    Call SetProgressState(101)
     '随用随创建，使用完后关闭，防止一台电脑开多个管理工具相互占用日志文件
    mobjOprateLog.Close
    Set mobjOprateLog = Nothing
    Me.Refresh
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function FuncControlStateVali(ByVal fctTmp As FuncCtrlType, ByVal strInfoName As String) As Boolean
'功能：功能按钮状态验证
'参数：fctTmp =控件索引
'         strInfoName=空间使用的信息名
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String, blnCando As Boolean
    '设置进度
    On Error GoTo errH
    strSQL = "Select 项目, 内容 From Zlupgradeconfig Where 项目 =[1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName, strInfoName)
    blnCando = True
    If rsTmp.EOF Then
        If Val(cmdFunc(fctTmp).Tag) <> 1 Then
            blnCando = False
        End If
        On Error Resume Next
        strSQL = "Insert Into ZLTOOLS.zlUpgradeConfig(项目,内容) values('" & strInfoName & "',1)"
        gcnOracle.Execute strSQL
        If err.Number <> 0 Then err.Clear
        On Error GoTo errH
    Else
        If Val(rsTmp!内容 & "") <> Val(cmdFunc(fctTmp).Tag) Then
            blnCando = False
        End If
    End If
    If Not blnCando Then
        MsgBox "功能""" & cmdFunc(fctTmp).Caption & """已经执行！", vbInformation, gstrSysName
        Call LoadData(0) '刷新界面
        Exit Function
    End If
    FuncControlStateVali = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Function FuncControlSQL(ByVal fctTmp As FuncCtrlType, ByVal blnLock As Boolean, ByVal strInfoName As String) As ADODB.Recordset
'功能：功能所需执行的步骤以及SQL
'参数：fctTmp =控件索引
'         blnLock=是否禁用对象
    Dim strSQL As String
    Dim strAdjustSQL As String
    Dim rsObjectSQL As ADODB.Recordset, rsSysScheduler As ADODB.Recordset
    Dim rsSQLStep As ADODB.Recordset '修正SQL以及步骤
    Dim strTmp As String, rsTmp As ADODB.Recordset
    Dim lngSort As Long, lngAll As Long
    Dim arrTmp As Variant, i As Long
  
  
    On Error GoTo errH
    '序号=SQL执行顺序，以及步骤顺序
    '类型=0：SQL,1:步骤以及进度
    Set rsSQLStep = CopyNewRec(Nothing, True, , Array("序号", adInteger, 5, 0, "步骤", adVarChar, 100, Empty, "进度", adDouble, 5, 0, "内容", adVarChar, 2000, Empty, "DBA", adInteger, Empty, Empty))
    Select Case fctTmp
        Case FCT_客户端
            '加载SQL以及进度
            lngAll = lngAll + 2
            If blnLock Then
                Call KillSessions
            Else
                '启用客户端
                strAdjustSQL = "Update Zlclients Set 禁止使用 = 0, 系统升级禁用 = Null Where Nvl(系统升级禁用, 0) = 1"
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, IIf(blnLock, "禁用客户端", "启用客户端"), (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
        Case FCT_用户账户
            If blnLock Then
                '标记锁定的用户
                strAdjustSQL = "Update " & gstrUserName & ".上机人员表 b" & vbNewLine & _
                                    "Set 系统升级锁定 = 1" & vbNewLine & _
                                    "Where Exists (Select 1 From Dba_Users a Where Account_Status = 'OPEN' And A.Username = Upper(B.用户名)) And Upper(B.用户名)<>'" & UCase(gstrUserName) & "'"
                '锁定用户账户
                strSQL = "Select 'alter user ' || 用户名 || ' account lock '  SQL" & vbNewLine & _
                            "From " & gstrUserName & ".上机人员表 b" & vbNewLine & _
                            "Where Exists (Select 1 From Dba_Users a Where Account_Status = 'OPEN' And A.Username = Upper(B.用户名)) And Upper(B.用户名)<>[1] " & vbNewLine & _
                            "Order By 用户名"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName, UCase(gstrUserName))
            Else
                '解锁用户账户
                strSQL = "Select 'alter user ' || 用户名 || ' account unlock ' SQL From " & gstrUserName & ".上机人员表 Where Nvl(系统升级锁定, 0)  = 1"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                '取消用户锁定标记
                strAdjustSQL = "Update " & gstrUserName & ".上机人员表 b Set 系统升级锁定 = Null Where Nvl(系统升级锁定, 0) = 1"
            End If
            '加载SQL以及进度
            lngAll = rsObjectSQL.RecordCount + 2
            '标记锁定的用户
            If blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, "标记锁定的用户", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
            '锁定或解锁用户
            Do While Not rsObjectSQL.EOF
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, IIf(blnLock, "锁定用户账户", "解锁用户账户"), (lngSort * 100 / lngAll), rsObjectSQL!SQL, 0)
                rsObjectSQL.MoveNext
            Loop
            '取消用户锁定标记
            If Not blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, "取消用户锁定标记", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
        Case FCT_后台作业
            If blnLock Then
                '标记停用的后台作业
                strAdjustSQL = "Update Zlautojobs a" & vbNewLine & _
                                        "Set A.系统升级停用 = 1" & vbNewLine & _
                                        "Where Exists (Select 1 From All_Jobs b Where B.Broken = 'N' And B.Job = A.作业号)"
                '停用后台作业
                strSQL = "Select 'Dbms_Job.Broken(' || 作业号 || ', True)' SQL" & vbNewLine & _
                                "From Zlautojobs a, All_Jobs b" & vbNewLine & _
                                "Where B.Job = A.作业号 And B.Broken = 'N'"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                '停用系统调度
                strSQL = "Select Job_Name, 'Call dbms_scheduler.disable(' || Chr(39) || A.Owner|| '.'||Job_Name || Chr(39) || ')'  SQL" & vbNewLine & _
                            "From Dba_Scheduler_Jobs a" & vbNewLine & _
                            "Where A.Job_Name In ('GATHER_STATS_JOB', 'AUTO_SPACE_ADVISOR_JOB') And Enabled = 'TRUE'"
                Set rsSysScheduler = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                If rsSysScheduler.RecordCount <> 0 Then lngAll = 1: strTmp = ""
                lngAll = lngAll + rsObjectSQL.RecordCount + rsSysScheduler.RecordCount
            Else
                '启用后台作业
                strSQL = "Select 'Dbms_Job.Broken(' || 作业号 || ', False)' Sql From Zlautojobs Where 系统升级停用 = 1"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                lngAll = rsObjectSQL.RecordCount
                '取消停用标记
                strAdjustSQL = "Update Zlautojobs Set 系统升级停用 = Null Where Nvl(系统升级停用, 0) = 1"
                '启用系统调度
                strSQL = "select 内容 from zlUpgradeConfig where 项目='禁用的系统调度'"
                Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                strTmp = rsTmp!内容 & ""
                If strTmp <> "" Then
                    lngAll = lngAll + 1
                    arrTmp = Split(strTmp, ",")
                    strTmp = ""
                    For i = LBound(arrTmp) To UBound(arrTmp)
                        strTmp = strTmp & ",'" & arrTmp(i) & "'"
                    Next
                    strTmp = Mid(strTmp, 2)
                    strSQL = "Select Job_Name, 'Call dbms_scheduler.enable(' || Chr(39) || A.Owner|| '.'||Job_Name || Chr(39) || ')' SQL" & vbNewLine & _
                            "From Dba_Scheduler_Jobs a" & vbNewLine & _
                            "Where A.Job_Name In (" & strTmp & ")"
                    Set rsSysScheduler = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                    lngAll = lngAll + rsSysScheduler.RecordCount
                End If
            End If
            '加载SQL以及进度
            lngAll = lngAll + 1
            '标记停用的后台作业
            If blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, "标记停用的后台作业", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
            '停用或启用后台作业
            Do While Not rsObjectSQL.EOF
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, IIf(blnLock, "停用后台作业", "启用后台作业"), (lngSort * 100 / lngAll), rsObjectSQL!SQL, 0)
                rsObjectSQL.MoveNext
            Loop
            '取消停用标记
            If Not blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, "取消停用标记", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
            '停用或启用系统调度
            If Not rsSysScheduler Is Nothing Then
                Do While Not rsSysScheduler.EOF
                    lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, IIf(blnLock, "停用系统调度", "启用系统调度"), (lngSort * 100 / lngAll), rsSysScheduler!SQL, 1)
                    If blnLock Then
                        strTmp = strTmp & "," & rsSysScheduler!Job_Name
                    End If
                    rsSysScheduler.MoveNext
                Loop
            End If
            If blnLock And strTmp <> "" Then strTmp = Mid(strTmp, 2)
            '标记停用系统调度或取消系统调度停用标记
            If strTmp <> "" Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, IIf(blnLock, "标记停用系统调度", "取消系统调度停用标记"), (lngSort * 100 / lngAll), "Update Zlupgradeconfig Set 内容 = " & IIf(blnLock, "'" & strTmp & "'", "Null") & " Where 项目 = '禁用的系统调度'", 0)
            End If
        Case FCT_触发器
            If blnLock Then
                '标记禁用的触发器
                strAdjustSQL = "Insert Into Zltriggers" & vbNewLine & _
                                        "  (所有者, 名称)" & vbNewLine & _
                                        "  Select Owner, Trigger_Name" & vbNewLine & _
                                        "  From Dba_Triggers a" & vbNewLine & _
                                        "  Where A.Owner In (Select Upper(所有者)" & vbNewLine & _
                                        "                    From Zlbakspaces" & vbNewLine & _
                                        "                    Union" & vbNewLine & _
                                        "                    Select Upper(所有者)" & vbNewLine & _
                                        "                    From Zlsystems" & vbNewLine & _
                                        "                    Union" & vbNewLine & _
                                        "                    Select 'ZLTOOLS' From Dual) And A.Status = 'ENABLED'"
                '禁用触发器
                strSQL = "Select 'alter trigger ' || Owner || '.' || Trigger_Name || ' disable' SQL" & vbNewLine & _
                                "From Dba_Triggers a" & vbNewLine & _
                                "Where A.Owner In (Select Upper(所有者)" & vbNewLine & _
                                "                  From Zlbakspaces" & vbNewLine & _
                                "                  Union" & vbNewLine & _
                                "                  Select Upper(所有者)" & vbNewLine & _
                                "                  From Zlsystems" & vbNewLine & _
                                "                  Union" & vbNewLine & _
                                "                  Select 'ZLTOOLS' From Dual) And A.Status = 'ENABLED'"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
            Else
                '启用触发器
                strSQL = "Select 'alter trigger ' || 所有者 || '.' || 名称 || ' disable'  SQL From Zltriggers"
                Set rsObjectSQL = gclsBase.OpenSQLRecord(gcnOracle, strSQL, gstrSysName)
                '删除触发器禁用标志
                strAdjustSQL = "Delete ZLTOOLS.ZLTriggers"
            End If
            '加载SQL以及进度
            lngAll = rsObjectSQL.RecordCount + 2
            '标记禁用的触发器
            If blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, "标记禁用的触发器", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
            '启用或禁用触发器
            Do While Not rsObjectSQL.EOF
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, IIf(blnLock, "停用后台作业", "启用后台作业"), (lngSort * 100 / lngAll), rsObjectSQL!SQL, 1)
                rsObjectSQL.MoveNext
            Loop
            '删除触发器禁用标志
            If Not blnLock Then
                lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, "删除触发器禁用标志", (lngSort * 100 / lngAll), strAdjustSQL, 0)
            End If
    End Select
    '记录执行情况
    lngSort = lngSort + 1: rsSQLStep.AddNew Array("序号", "步骤", "进度", "内容", "DBA"), Array(lngSort, "执行标记", (lngSort * 100 / lngAll), "Update Zlupgradeconfig Set 内容 = " & IIf(blnLock, 0, 1) & " Where 项目 = '" & strInfoName & "'", 0)
    Set FuncControlSQL = rsSQLStep
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub SetProgressState(ByVal lngValue As Long, Optional ByVal strStep As String)
'功能：设置进度条
'参数：lngValue=进度，-1：隐藏进度条   >=0：显示进度条，并展示进度
'          strStep=当前步骤名
    If lngValue = -1 Then
        pgbState.Visible = False: lblState.Visible = False: pgbState.value = 0
    Else
        If lngValue > 100 Then
            pgbState.Visible = False
            lblState.Caption = "处理完毕！"
            Sleep (1000)
            lblState.Visible = False
        Else
            pgbState.value = lngValue
            If lblState.Caption <> strStep Then
                lblState.Caption = strStep
                Call SetCtrlPosOnLine(False, 0, lblState, 0, pgbState)
            End If
            pgbState.Visible = True: lblState.Visible = True
        End If
    End If
End Sub

Private Function VilidateUpgrade(Optional ByVal lngRow As Long) As Boolean
    Dim i As Long, strMaxVer As String, strCurMaxVer As String
    Dim strMaxTools As String
    Dim blnUp As Boolean
    Dim strFilter As String, strFilterSys As String
    Dim lngBegin  As Long, lngEnd As Long
    
    On Error GoTo errH
    pgbGetUpInfo.value = 0: pgbGetUpInfo.Visible = True
    With vsSysSel
        If lngRow > .FixedRows Then
            strFilterSys = "系统编号=" & Val(.TextMatrix(i, Col_编号))
            lngBegin = lngRow: lngEnd = lngRow
        Else '管理工具读取时，对所有的系统进行刷新
            lngBegin = .FixedRows: lngEnd = .Rows - 1
        End If
        If lngRow <= .FixedRows Or mrsSysUpFiles Is Nothing Then '获取升迁文件,初始化
            Set mrsSysUpFiles = GetUpgradeFiles(Nothing, -1, "", "")
        Else '清空当前系统的数据
            Call RecDelete(mrsSysUpFiles, strFilterSys)
        End If
        '清空上次升迁检查信息
        '"升迁结果", "中止信息","提前升迁结果", "提前中止信息"不清空
        Call RecUpdate(mrsSysInfo, strFilterSys, "升级", 1, "目标版本", Null, "最终版本", Null, "可升级", 1, "检查结果", Null, "可提前升级", 1, "提前检查结果", Null)
        '上次常规升迁未正常完成的不能提前执行
        Call RecUpdate(mrsSysInfo, strFilterSys & IIf(strFilterSys <> "", " And ", "") & "升迁结果=1", "可提前升级", 0, "提前检查结果", "系统最近一次正常升级不成功，不能进行提前执行！")
        .Cell(flexcpText, lngBegin, Col_目标版本, lngEnd, Col_检查结果) = ""
        .Cell(flexcpForeColor, lngBegin, Col_目标版本, lngEnd, Col_检查结果) = &H80000008
        '前置准备
        pgbGetUpInfo.value = 5
        For i = lngBegin To lngEnd
            mrsSysInfo.Filter = "系统编号=" & Val(.TextMatrix(i, Col_编号))
            mrsSysInfo.Update "配置文件", .Cell(flexcpData, i, Col_配置文件)
            strMaxVer = "": strCurMaxVer = ""
            If .Cell(flexcpData, i, Col_配置文件) <> "" Then
                Set mrsSysUpFiles = GetUpgradeFiles(mrsSysUpFiles, Val(.TextMatrix(i, Col_编号)), .TextMatrix(i, Col_当前版本), mrsSysInfo!配置文件, mrsSysInfo!中止信息 & "", mrsSysInfo!提前中止信息 & "", strMaxVer, strCurMaxVer)
            End If
            mrsSysInfo.Update Array("最终版本", "目标版本"), Array(strMaxVer, strCurMaxVer)
            pgbGetUpInfo.value = (i - lngBegin + 1) / (lngEnd - lngBegin + 1) * 75 + 5
        Next
    End With
    pgbGetUpInfo.value = 80
    Call RecToLog(mrsSysUpFiles, "系统编号,FullSPVer,SysType,FileType", "文件记录集")
    mrsSysInfo.Filter = "系统编号=0"
    strMaxTools = IIf(mrsSysInfo!目标版本 & "" = "", mrsSysInfo!系统版本号, mrsSysInfo!目标版本)
    mrsSysInfo.Filter = strFilterSys & IIf(strFilterSys <> "", " And ", "") & "可升级=1"
    Do While Not mrsSysInfo.EOF
        If mrsSysInfo!系统编号 <> 0 Then
            If mrsSysInfo!目标版本 & "" <> "" Then
                '管理工具目标版本支持不了应用系统升迁到目标版本
                If VerFull(mrsSysInfo!目标版本) <= VerFull("10.34.0") Then '10.34.0以后，管理工具与应用系统版本不对应，只比较主版本
                    If VerFull(GetPrimaryVer(GetContractVersion(mrsSysInfo!目标版本))) > VerFull(GetPrimaryVer(strMaxTools)) Then
                        mrsSysInfo.Update Array("可升级", "检查结果"), Array(0, "管理工具不能支持应用系统升迁到""" & mrsSysInfo!目标版本 & """!")
                    End If
                Else '10.34.0以后，管理工具与应用系统版本已经对应，可以比较SP版本
                    If VerFull(GetPrimaryVer(mrsSysInfo!目标版本)) > VerFull(GetPrimaryVer(strMaxTools)) Then
                        mrsSysInfo.Update Array("可升级", "检查结果"), Array(0, "管理工具不能支持应用系统升迁到""" & mrsSysInfo!目标版本 & """!")
                    End If
                End If
            Else
                mrsSysInfo.Update Array("可升级", "检查结果"), Array(0, "该系统没有可执行的升级脚本，不能升迁!")
            End If
        ElseIf mrsSysInfo!目标版本 & "" = "" Then
            mrsSysInfo.Update Array("可升级", "检查结果"), Array(0, "管理工具没有可执行的升级脚本，不能升迁!")
        End If
        mrsSysInfo.MoveNext
    Loop
    pgbGetUpInfo.value = 85
    '先判断应用系统能否常规升迁，应用不能常规升迁，则不能提前升迁
    Call RecUpdate(mrsSysInfo, strFilterSys & IIf(strFilterSys <> "", " And ", "") & "可升级=0", "可提前升级", 0)
    Call RecUpdate(mrsSysInfo, strFilterSys & IIf(strFilterSys <> "", " And ", "") & "可提前升级=0 And 提前检查结果=Null", "提前检查结果", "!检查结果")
    pgbGetUpInfo.value = 90
    '获取提前执行的目标版本
    If optUpMode(1).value Then
        mrsSysInfo.Filter = strFilterSys & IIf(strFilterSys <> "", " And ", "") & "可提前升级=1"
        Do While Not mrsSysInfo.EOF
            strFilter = "系统编号=" & mrsSysInfo!系统编号 & " And SysType<>" & ST_History & " And FileType=" & FT_Before
            mrsSysUpFiles.Filter = strFilter: mrsSysUpFiles.Sort = "FullSPVer Desc": strMaxVer = ""
            If Not mrsSysUpFiles.EOF Then
                strMaxVer = mrsSysUpFiles!SPVer
                mrsSysUpFiles.Filter = strFilter & " And 配置版本>'" & VerFull(mrsSysInfo!系统版本号) & "'": mrsSysUpFiles.Sort = "FullSPVer"
                If Not mrsSysUpFiles.EOF Then
                    mrsSysUpFiles.Filter = strFilter & " And FullSPVer<'" & mrsSysUpFiles!FullSPVer & "'": mrsSysUpFiles.Sort = "FullSPVer Desc"
                    If Not mrsSysUpFiles.EOF Then
                        strMaxVer = mrsSysUpFiles!SPVer
                    Else
                        strMaxVer = ""
                        mrsSysInfo.Update Array("可提前升级", "提前检查结果"), Array(0, "没有可执行的提前升级脚本，不能提前升迁！")
                    End If
                End If
            Else
                mrsSysInfo.Update Array("可提前升级", "提前检查结果"), Array(0, "没有提前升级脚本，不能提前升迁！")
            End If
            mrsSysInfo.Update "提前目标版本", strMaxVer
            '删除非提前脚本,不删除历史库主要是因为历史库可能版本较低，需要额外读取，此时需要完整的脚本来截取上次发生中止以后的脚本
            Call RecDelete(mrsSysUpFiles, "系统编号=" & mrsSysInfo!系统编号 & " And SysType<>" & ST_History & " And FileType<>" & FT_Before)
            '删除大于提前目标版本的提前升级脚本
            Call RecDelete(mrsSysUpFiles, strFilter & " And FullSPVer>'" & VerFull(strMaxVer) & "'")
            mrsSysInfo.MoveNext
        Loop
    End If
    '数据展现
    pgbGetUpInfo.value = 95
    With vsSysSel
        For i = lngBegin To lngEnd
            mrsSysInfo.Filter = "系统编号=" & Val(.TextMatrix(i, Col_编号))
            .RowData(i) = Val(IIf(optUpMode(1).value, mrsSysInfo!可提前升级, mrsSysInfo!可升级) & "")
            If .RowData(i) = 0 And Val(.TextMatrix(i, Col_Sel)) <> 0 Then
                .TextMatrix(i, Col_Sel) = 0
            ElseIf Val(.TextMatrix(i, Col_Sel)) <> 0 Then
                blnUp = True
            End If
            mrsSysInfo.Update "升级", IIf(Val(.TextMatrix(i, Col_Sel)) <> 0, 1, 0)
            .TextMatrix(i, Col_目标版本) = VerPAD(IIf(optUpMode(1).value, mrsSysInfo!提前目标版本, mrsSysInfo!目标版本) & "")
            .TextMatrix(i, Col_检查结果) = IIf(optUpMode(1).value, mrsSysInfo!提前检查结果, mrsSysInfo!检查结果) & ""
            If .RowData(i) = 0 Then
                .Cell(flexcpForeColor, i, Col_目标版本, i, Col_检查结果) = &H2222B2 '火砖红
            End If
            If Val(.TextMatrix(i, Col_Sel)) = 0 Then '不能升级，则删除脚本
                Call RecDelete(mrsSysUpFiles, "系统编号=" & Val(vsSysSel.TextMatrix(i, Col_编号)))
            End If
        Next
    End With
    pgbGetUpInfo.value = 100
    pgbGetUpInfo.Visible = False: pgbGetUpInfo.value = 0
    VilidateUpgrade = blnUp
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    pgbGetUpInfo.Visible = False
End Function

Private Sub SetSelBeach(Optional ByVal blnSel As Boolean)
'功能：设置批量选择
'参数：blnSel=True：批量选择，False:批量取消
    Dim intSel As Integer
    Dim i As Long
    
    intSel = IIf(blnSel, 1, 0)
    With vsSysSel
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                .TextMatrix(i, Col_Sel) = intSel
            Else
                .TextMatrix(i, Col_Sel) = 0
            End If
        Next
    End With
End Sub
