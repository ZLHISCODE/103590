VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.8#0"; "ZLIDKIND.OCX"
Begin VB.Form frmServiceChangeNum 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "预约换诊"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10950
   Icon            =   "frmServiceChangeNum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   300
      Left            =   705
      TabIndex        =   38
      Top             =   495
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "宋体"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox picInfo 
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      Height          =   1380
      Left            =   15
      ScaleHeight     =   1320
      ScaleWidth      =   10830
      TabIndex        =   0
      Top             =   1365
      Width           =   10890
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   -45
         TabIndex        =   39
         Top             =   1290
         Width           =   11145
      End
      Begin VB.TextBox txtNO 
         Enabled         =   0   'False
         Height          =   330
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   540
         Width           =   1530
      End
      Begin VB.TextBox txtDoc 
         Enabled         =   0   'False
         Height          =   330
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   945
         Width           =   1530
      End
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3225
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   945
         Width           =   1530
      End
      Begin VB.TextBox txtAdd 
         Height          =   330
         Left            =   7530
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   945
         Width           =   3240
      End
      Begin VB.TextBox txtFee 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   945
         Width           =   1110
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   -75
         TabIndex        =   25
         Top             =   480
         Width           =   11115
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   330
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   45
         Width           =   1530
      End
      Begin VB.TextBox txtGender 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   45
         Width           =   570
      End
      Begin VB.TextBox txtAge 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   45
         Width           =   810
      End
      Begin VB.TextBox txtPhone 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   45
         Width           =   1110
      End
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
         Height          =   330
         Left            =   7530
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   45
         Width           =   3240
      End
      Begin VB.TextBox txtAppTime 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3225
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   1530
      End
      Begin VB.TextBox txtRegNO 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   540
         Width           =   1110
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   330
         Left            =   7530
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   540
         Width           =   3240
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "项目"
         Height          =   180
         Left            =   2835
         TabIndex        =   35
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "换诊说明"
         Height          =   180
         Left            =   6765
         TabIndex        =   33
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "挂号费"
         Height          =   180
         Left            =   5025
         TabIndex        =   28
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         Height          =   180
         Left            =   375
         TabIndex        =   26
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Left            =   390
         TabIndex        =   24
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Left            =   2475
         TabIndex        =   23
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   180
         Left            =   3570
         TabIndex        =   22
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "联系电话"
         Height          =   180
         Left            =   4845
         TabIndex        =   21
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "现住址"
         Height          =   180
         Left            =   6945
         TabIndex        =   20
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "预约时间"
         Height          =   180
         Left            =   2475
         TabIndex        =   19
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "预约单号"
         Height          =   180
         Left            =   30
         TabIndex        =   18
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "号码"
         Height          =   180
         Left            =   5205
         TabIndex        =   17
         Top             =   615
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         Height          =   180
         Left            =   7125
         TabIndex        =   16
         Top             =   615
         Width           =   360
      End
   End
   Begin VB.PictureBox picReg 
      BorderStyle     =   0  'None
      Height          =   6180
      Left            =   120
      ScaleHeight     =   6180
      ScaleWidth      =   9015
      TabIndex        =   29
      Top             =   2865
      Width           =   9015
      Begin VB.PictureBox picSplit 
         BorderStyle     =   0  'None
         Height          =   100
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   105
         ScaleWidth      =   3855
         TabIndex        =   41
         Top             =   2565
         Width           =   3855
      End
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   2730
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   75
         Width           =   1125
      End
      Begin VB.PictureBox picTime 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4005
         ScaleHeight     =   330
         ScaleWidth      =   2745
         TabIndex        =   31
         Top             =   60
         Visible         =   0   'False
         Width           =   2745
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   315
            Left            =   750
            TabIndex        =   4
            Top             =   0
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   155320322
            CurrentDate     =   42340
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "预约时间"
            Height          =   180
            Left            =   0
            TabIndex        =   32
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.TextBox txtFilter 
         Height          =   315
         Left            =   465
         TabIndex        =   2
         Top             =   68
         Width           =   1320
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
         Height          =   2100
         Left            =   60
         TabIndex        =   6
         Top             =   465
         Width           =   3360
         _cx             =   5927
         _cy             =   3704
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   15658734
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   23
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmServiceChangeNum.frx":058A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   3120
         Left            =   60
         TabIndex        =   7
         Top             =   2670
         Width           =   5925
         _cx             =   10451
         _cy             =   5503
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmServiceChangeNum.frx":08BD
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
      Begin XtremeCalendarControl.DatePicker dtpMain 
         Height          =   2100
         Left            =   5190
         TabIndex        =   5
         Top             =   450
         Width           =   3060
         _Version        =   1048579
         _ExtentX        =   5397
         _ExtentY        =   3704
         _StockProps     =   64
         AutoSize        =   0   'False
         ShowTodayButton =   0   'False
         ShowNoneButton  =   0   'False
         Show3DBorder    =   0
         MaxSelectionCount=   1
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "预约时段"
         Height          =   180
         Left            =   1980
         TabIndex        =   40
         Top             =   135
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "号码"
         Height          =   180
         Left            =   60
         TabIndex        =   30
         Top             =   135
         Width           =   360
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   7545
      Top             =   3675
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmServiceChangeNum.frx":09C9
      Left            =   2445
      Top             =   2415
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmServiceChangeNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr时段s As String
Private mbln是否收费 As Boolean
Private mdbl收费金额 As Double
Private mblnUnload As Boolean, mblnNotClick As Boolean
Public mlng消息ID As Long
Private mstrName As String, mstrGender As String
Private mstrAge As String, mlng预约有效时间 As Long
Private mstrPhone As String
Private mstrNo As String
Private mstrAppTime As String
Private mstrInfo As String
Private mstrRegNo As String
Private mstrDoc As String, mblnInit As Boolean, mblnAppointmentChange As Boolean
Private mblnChangeByCode As Boolean, mlngRow As Long, msngTime As Single
Private mstrPriceGrade As String

Public Sub ShowMe(frmMain As Object)
    Me.Show 1, frmMain
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    
    Err = 0: On Error GoTo errHandle
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 '去除菜单栏前缀
    cbsThis.ActiveMenuBar.Visible = False
    
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ModifyStyle &H400000, 0 '去除菜单栏前缀
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, 3839, "换诊")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    DefMainCommandBars = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitPanel()
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHandle
    Set objPane = dkpMain.CreatePane(1, 145, 80, DockTopOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    objPane.Handle = picInfo.Hwnd
    objPane.MaxTrackSize.Height = 88
    objPane.MinTrackSize.Height = 88
    
    Set objPane = dkpMain.CreatePane(3, 145, 120, DockBottomOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Title = "出诊安排信息"
    objPane.Handle = picReg.Hwnd
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
        .PaintManager.HighlighActiveCaption = False
    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case 3839 '换诊
            If SaveData = True Then Unload Me
        Case Else
            Unload Me
    End Select
End Sub

Private Function SaveData() As Boolean
    '换诊操作
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim datApp As Date
    Dim strTemp As String
    
    If vsfPlan.RowData(vsfPlan.Row) = "" Or vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("号码")) = "" Then
        MsgBox "请选择一个安排进行换诊!", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsfPlan
        strTemp = .TextMatrix(.Row, .ColIndex("号码")) & "-" & .TextMatrix(.Row, .ColIndex("科室")) & _
                IIf(.TextMatrix(.Row, .ColIndex("医生")) = "", "", "(" & .TextMatrix(.Row, .ColIndex("医生")) & ")")
    End With
    
    If MsgBox("是否确定将号码:" & txtRegNO.Text & "-" & txtDept.Text & _
                IIf(txtDoc.Text = "", "", "(" & txtDoc.Text & ")") & " 换诊到" & vbCrLf & _
                "号码:" & strTemp & "?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then
        Exit Function
    End If
    
    If CheckLimit(Val(vsfPlan.RowData(vsfPlan.Row))) = False Then Exit Function
    
    If vsfList.Visible Then
        If vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col) = "" Then
            MsgBox "请选择一个有效的序号进行换诊!", vbInformation, gstrSysName
            Exit Function
        End If
        If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("分时段"))) = 1 Then
            If Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)) <> 0 Then
                strSQL = "Select 开始时间 From 临床出诊序号控制 Where 记录ID=[1] And 序号=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)))
                If Not rsTemp.EOF Then
                    datApp = CDate(Format(rsTemp!开始时间, "yyyy-mm-dd hh:mm:ss"))
                Else
                    datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
                End If
            Else
                datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
            End If
        Else
            datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
        End If
    Else
        datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
    End If
    
    
    If datApp < DateAdd("n", -1 * mlng预约有效时间, zlDatabase.Currentdate) Then
        MsgBox "预约时间小于了可预约时间(" & Format(DateAdd("n", -1 * mlng预约有效时间, zlDatabase.Currentdate), "hh:mm:ss") & "),无法预约!", vbInformation, gstrSysName
        Exit Function
    End If
    If Not (Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("分时段"))) = 1 And Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("序号控制"))) = 1) Then
        If Check有效时间段(Val(vsfPlan.RowData(vsfPlan.Row)), datApp) = False Then
            MsgBox "当前选择的出诊记录在" & Format(datApp, "yyyy-mm-dd hh:mm:ss") & "不出诊,请调整挂号时间!", vbInformation, gstrSysName
            If dtpTime.Enabled And dtpTime.Visible Then dtpTime.SetFocus
            Exit Function
        End If
    End If
    
    strSQL = "Select Zl_临床出诊限制_Check([1],[2],[3]) As 适用性检查 From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), txtGender.Text, txtAge.Text)
    If rsTemp.EOF Then
        MsgBox "当前选择的病人不适用该号别!", vbInformation, gstrSysName
        Exit Function
    Else
        If Val(Mid(Nvl(rsTemp!适用性检查), 1, 1)) <> 0 Then
            MsgBox "当前选择的病人不适用该号别!" & vbCrLf & "原因:" & Mid(Nvl(rsTemp!适用性检查), InStr(Nvl(rsTemp!适用性检查), "|") + 1), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSQL = "Select 1 From 临床出诊记录 Where Id=[1] And [2] Between 开始时间 And 终止时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, vsfPlan.RowData(vsfPlan.Row), datApp)
    If rsTemp.EOF Then
        MsgBox "当前预约时间不在该号码的有效时间内,请重新选择预约时间!", vbInformation, gstrSysName
        Exit Function
    End If
    strSQL = "Zl_患者服务中心_换诊(" & mlng消息ID & ",'"
    strSQL = strSQL & Trim(txtNO.Text) & "',"
    If vsfList.Visible Then
        strSQL = strSQL & Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)) & ","
    Else
        strSQL = strSQL & "Null,"
    End If
    strSQL = strSQL & "To_Date('" & datApp & "','yyyy-mm-dd hh24:mi:ss')" & ","
    strSQL = strSQL & vsfPlan.RowData(vsfPlan.Row) & ","
    strSQL = strSQL & "'" & txtAdd.Text & "',"
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    strSQL = strSQL & "'" & mstrPriceGrade & "')" '价格等级
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
End Function

Private Sub dtpTime_Change()
    Dim str日期 As String, i As Integer, lngRow As Long
    Dim str发生时间 As String
    If Not dtpMain.Visible Then Exit Sub
    If Not dtpMain.Enabled Then Exit Sub
    
    str日期 = Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-MM-dd")

    If str日期 = "" Then str日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    str发生时间 = str日期 & " " & Format(dtpTime.Value, "hh:mm:00")
    lngRow = 0
    If CDate(str发生时间) > CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("终止时间"))) Then
        '超出时间的安排，重新寻找定位
        For i = 1 To vsfPlan.Rows - 1
            With vsfPlan
                If .TextMatrix(.Row, .ColIndex("号码")) = .TextMatrix(i, .ColIndex("号码")) And _
                    CDate(vsfPlan.TextMatrix(i, vsfPlan.ColIndex("终止时间"))) >= CDate(str发生时间) Then
                    lngRow = i
                    Exit For
                End If
            End With
        Next i
    ElseIf CDate(str发生时间) < CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("开始时间"))) Then
        '超出时间的安排，重新寻找定位
        For i = 1 To vsfPlan.Rows - 1
            With vsfPlan
                If .TextMatrix(.Row, .ColIndex("号码")) = .TextMatrix(i, .ColIndex("号码")) And _
                    CDate(vsfPlan.TextMatrix(i, vsfPlan.ColIndex("开始时间"))) <= CDate(str发生时间) Then
                    lngRow = i
                    Exit For
                End If
            End With
        Next i
    End If
    If lngRow <> 0 Then
        mblnAppointmentChange = True
        vsfPlan.Select lngRow, 1
        mblnAppointmentChange = False
    End If
End Sub

Private Function Check有效时间段(lng记录ID As Long, datTime As Date) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    With vsfPlan
        '序号控制分时段号,不检查序号时间是否在出诊记录时间内
        If Val(.TextMatrix(.Row, .ColIndex("序号控制"))) = 1 And Val(.TextMatrix(.Row, .ColIndex("分时段"))) = 1 Then
            strSQL = "Select 1 From 临床出诊记录 Where ID=[1] And [2] Between Nvl(停诊开始时间,To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(停诊终止时间,To_Date('3000-01-01', 'yyyy-mm-dd')) "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, datTime)
            If rsTemp.EOF Then
                Check有效时间段 = True
            Else
                Check有效时间段 = False
            End If
            Exit Function
        End If
    End With
    
    strSQL = "Select 1 From 临床出诊记录 Where ID=[1] And [2] Between 开始时间 And 终止时间 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, datTime)
    
    If rsTemp.EOF Then
        Check有效时间段 = False
    Else
        strSQL = "Select 1 From 临床出诊记录 Where ID=[1] And [2] Between Nvl(停诊开始时间,To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(停诊终止时间,To_Date('3000-01-01', 'yyyy-mm-dd')) "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, datTime)
        If rsTemp.EOF Then
            Check有效时间段 = True
        Else
            Check有效时间段 = False
        End If
    End If
    
End Function

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 3839
        Control.Enabled = vsfPlan.RowData(vsfPlan.Row) <> ""
    End Select
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub dtpMain_SelectionChanged()
    Dim datNow As Date
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd") > Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
        MsgBox "不能换诊当前时间之前的安排!", vbInformation, gstrSysName
        dtpMain.SelectRange datNow, datNow
        dtpMain.Select datNow
        dtpMain.EnsureVisibleSelection
        dtpMain.RedrawControl
    End If
    Call LoadPlan
End Sub

Private Sub InitGrid()
    Dim i As Integer
    With vsfPlan
        .MergeRow(0) = True
        .Rows = 2
        For i = 0 To .Rows - 1
            .RowHeight(i) = 350
        Next i
    End With
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call dtpTime_Change
End Sub

Private Sub Form_Activate()
    mblnInit = True
    If mblnUnload Then
        mblnUnload = False
        Unload Me
        Exit Sub
    End If
    Call vsfPlan_EnterCell
    If txtFilter.Visible And txtFilter.Enabled Then txtFilter.SetFocus
    mblnInit = False
End Sub
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picInfo.Hwnd
    Case 3
        Item.Handle = picReg.Hwnd
    End Select
End Sub
Private Sub InitPara()
'    mlng预约有效时间 = Val(Split(zlDatabase.GetPara("预约限制时间", glngSys, 1111, "1|60") & "|", "|")(1))
    mlng预约有效时间 = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer, strSQL As String, rsTemp As ADODB.Recordset
    Dim datTime As Date
    mblnInit = True
    If mstrInfo = "" Then
        SetWithValue
    Else
        Call SetValue
    End If
    Set txtAdd.Container = Me
    txtAdd.Left = txtAdd.Left + 15
    txtAdd.Top = txtAdd.Top + 540
    txtAdd.Locked = False
    Call DefMainCommandBars
    Call InitPanel
    Call InitGrid
    Call InitPara
    dtpTime.Value = Now
    If Format(txtAppTime.Text, "yyyy-mm-dd") < Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
        datTime = CDate(zlDatabase.Currentdate)
    Else
        dtpTime = CDate(txtAppTime.Text)
    End If
    
    dtpMain.SelectRange dtpTime, dtpTime
    dtpMain.Select dtpTime
    dtpMain.EnsureVisibleSelection
    dtpMain.EnsureVisible dtpTime
    dtpMain.RedrawControl
    
    dtpMain.HighlightToday = False
    dtpMain.ShowNonMonthDays = False
    dtpMain.PaintManager.ControlBackColor = &H8000000F
    dtpMain.PaintManager.DayBackColor = &H8000000F
    dtpMain.PaintManager.DaysOfWeekBackColor = &H8000000F
    
    strSQL = "Select 1 From 门诊费用记录 Where NO=[1] And 记录性质=4 And 结帐ID Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txtNO.Text))
    If Not rsTemp.EOF Then
        mbln是否收费 = False
    Else
        mbln是否收费 = True
        strSQL = "Select Sum(Nvl(结帐金额, 0)) As 金额" & vbNewLine & _
                "From 门诊费用记录" & vbNewLine & _
                "Where NO = [1] And 记录性质 = 4 And" & vbNewLine & _
                "      收费细目id Not In" & vbNewLine & _
                "      (Select 收费细目id" & vbNewLine & _
                "       From 收费特定项目" & vbNewLine & _
                "       Where 特定项目 = '病历费'" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select 从项id From 收费从属项目 Where 主项id In (Select 收费细目id From 收费特定项目 Where 特定项目 = '病历费'))"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txtNO.Text))
        mdbl收费金额 = Val(Nvl(rsTemp!金额))
    End If
    mblnUnload = False
    Call LoadPlan
    IDKind.RaisEffect picInfo, -1
    mblnInit = False
End Sub

Public Sub InitValue(strName As String, strGender As String, strAge As String, strPhone As String, _
    strNO As String, strAppTime As String, strInfo As String, ByVal strPriceGrade As String)
    mstrName = strName
    mstrGender = strGender
    mstrAge = strAge
    mstrPhone = strPhone
    mstrNo = strNO
    mstrAppTime = strAppTime
    mstrInfo = strInfo
    mstrPriceGrade = strPriceGrade '价格等级
End Sub

Public Sub InitWithValue(strName As String, strGender As String, strAge As String, strPhone As String, _
    strNO As String, strAppTime As String, strRegNo As String, strDoc As String, ByVal strPriceGrade As String)
    mstrName = strName
    mstrGender = strGender
    mstrAge = strAge
    mstrPhone = strPhone
    mstrNo = strNO
    mstrAppTime = strAppTime
    mstrRegNo = strRegNo
    mstrDoc = strDoc
    mstrPriceGrade = strPriceGrade '价格等级
End Sub

Private Sub SetValue()
    Dim strArray() As String
    With Me
        .txtName = mstrName
        .txtGender = mstrGender
        .txtAge = mstrAge
        .txtPhone = mstrPhone
        .txtNO = mstrNo
        .txtAppTime = mstrAppTime
        strArray = Split(mstrInfo, "   ")
        .txtRegNO = Split(strArray(0), ":")(1) '& "[" & Split(strArray(1), ":")(1) & "]"
        .txtItem = Split(strArray(3), ":")(1)
        .txtDept = Split(strArray(2), ":")(1)
        .txtDoc = Split(strArray(4), ":")(1)
    End With
End Sub

Private Sub SetWithValue()
    Dim strArray() As String
    With Me
        .txtName = mstrName
        .txtGender = mstrGender
        .txtAge = mstrAge
        .txtPhone = mstrPhone
        .txtNO = mstrNo
        .txtAppTime = mstrAppTime
        .txtRegNO = mstrRegNo
        .txtDoc = mstrDoc
    End With
End Sub

Private Sub LoadPlan()
    Dim strSQL As String, rsPlan As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim datApp As Date, i As Integer, dblMoney As Double, blnAdd As Boolean
    Dim strTime() As String, lngLeft As Long
    Dim dbl项目金额 As Double
    
    datApp = dtpMain.Selection.Blocks(0).DateBegin
    strSQL = "Select a.id, b.号类, b.号码 As 号码, c.名称 As 科室, c.简码 As 科室简码, b.科室Id, a.上班时段 As 时段, " & _
            "        d.名称 As 项目, zlSpellcode(d.名称) As 项目简码, a.替诊医生姓名, a.替诊医生ID, a.医生Id, a.医生姓名 As 医生, e.简码 As 医生简码, " & _
            "        a.项目id, a.限约数 As 限约, a.已约数 As 已约, Nvl(a.是否分时段,0) As 分时段, Nvl(a.是否序号控制,0) As 序号控制, " & _
            "        a.出诊日期, a.缺省预约时间 , a.替诊开始时间 , a.替诊终止时间 , a.开始时间, a.终止时间 " & vbNewLine & _
            "From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 收费项目目录 D, 人员表 E" & vbNewLine & _
            "Where a.号源id = b.Id  And Nvl(C.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And a.项目id = d.Id And b.科室id = c.Id And (c.站点 Is Null Or c.站点 = '" & gstrNodeNo & "') And (a.出诊日期 = [1] Or a.出诊日期 = [2]) And (a.开始时间 < Nvl(a.停诊开始时间,a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间,a.开始时间) Or Exists (Select 1 From 临床出诊序号控制 C,临床出诊记录 D Where D.ID=A.ID And C.记录ID=D.ID And Nvl(C.是否停诊,0) = 0 And D.是否序号控制 =1 And D.是否分时段 = 1 And C.开始时间 <> C.终止时间)) " & _
            "      And Nvl(a.是否发布,0)=1 And a.医生Id = e.Id(+) And Nvl(a.预约控制,0) <> 1 " & _
            "      And Not Exists (Select 1 From 临床出诊记录 Where Id=a.Id And 终止时间 < [3]) And a.开始时间 >= [4] And Sysdate + zl_Fun_GetAppointmentDays + Decode(Nvl(B.预约天数," & gint预约天数 & "),0,15,Nvl(B.预约天数," & gint预约天数 & ")" & ") >= [1] " & _
            "      And [3] Not Between Nvl(a.停诊开始时间,a.终止时间) And Nvl(a.停诊终止时间,a.开始时间) "
    
    If Format(datApp, "yyyy-mm-dd") = Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
        Set rsPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, datApp, datApp - 1, zlDatabase.Currentdate, gdatRegistTime)
    Else
        Set rsPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, datApp, datApp - 1, datApp, gdatRegistTime)
    End If
    mstr时段s = ""
    mblnNotClick = True
    cboTime.Clear
    cboTime.AddItem "所有"
    mblnNotClick = False
    vsfPlan.Redraw = flexRDNone
    vsfPlan.Clear 1
    vsfPlan.Rows = 2
    Do While Not rsPlan.EOF
        blnAdd = True
        
        dbl项目金额 = Get项目金额(Val(Nvl(rsPlan!项目ID)), mstrPriceGrade)
        If mbln是否收费 Then
            '收费的挂号记录,必须金额相同的才能换诊
            If RoundEx(mdbl收费金额, 6) <> RoundEx(dbl项目金额, 6) Then
                blnAdd = False
            End If
        End If
        
        If blnAdd Then
            With vsfPlan
                .RowData(.Rows - 1) = Val(Nvl(rsPlan!ID))
                .TextMatrix(.Rows - 1, .ColIndex("号类")) = Nvl(rsPlan!号类)
                .TextMatrix(.Rows - 1, .ColIndex("号码")) = Nvl(rsPlan!号码)
                .TextMatrix(.Rows - 1, .ColIndex("科室")) = Nvl(rsPlan!科室)
                .TextMatrix(.Rows - 1, .ColIndex("时段")) = Nvl(rsPlan!时段)
                If InStr("," & mstr时段s & ",", "," & Nvl(rsPlan!时段) & ",") = 0 Then
                    mstr时段s = mstr时段s & "," & Nvl(rsPlan!时段)
                End If
                .TextMatrix(.Rows - 1, .ColIndex("项目")) = Nvl(rsPlan!项目)
                .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format(dbl项目金额, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("限约")) = Nvl(rsPlan!限约)
                .TextMatrix(.Rows - 1, .ColIndex("已约")) = Nvl(rsPlan!已约)
                .TextMatrix(.Rows - 1, .ColIndex("分时段")) = Nvl(rsPlan!分时段)
                .TextMatrix(.Rows - 1, .ColIndex("序号控制")) = Nvl(rsPlan!序号控制)
                .TextMatrix(.Rows - 1, .ColIndex("医生")) = Nvl(rsPlan!医生)
                .TextMatrix(.Rows - 1, .ColIndex("出诊日期")) = Format(Nvl(rsPlan!出诊日期), "yyyy-mm-dd")
                .TextMatrix(.Rows - 1, .ColIndex("医生简码")) = Nvl(rsPlan!医生简码)
                .TextMatrix(.Rows - 1, .ColIndex("科室简码")) = Nvl(rsPlan!科室简码)
                .TextMatrix(.Rows - 1, .ColIndex("项目简码")) = Nvl(rsPlan!项目简码)
                .TextMatrix(.Rows - 1, .ColIndex("预约时间")) = Format(Nvl(rsPlan!缺省预约时间), "yyyy-mm-dd hh:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("开始时间")) = Format(Nvl(rsPlan!开始时间), "yyyy-mm-dd hh:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("终止时间")) = Format(Nvl(rsPlan!终止时间), "yyyy-mm-dd hh:mm:ss")
                If Nvl(rsPlan!替诊医生姓名) <> "" Then
                    .Cell(flexcpData, .Rows - 1, .ColIndex("替诊医生")) = Nvl(rsPlan!替诊医生姓名) & "(" & Format(Nvl(rsPlan!替诊开始时间), "hh:mm") & "-" & Format(Nvl(rsPlan!替诊终止时间), "hh:mm") & ")"
                    .TextMatrix(.Rows - 1, .ColIndex("替诊医生")) = ""
                    .TextMatrix(.Rows - 1, .ColIndex("替诊医生姓名")) = Nvl(rsPlan!替诊医生姓名)
                    .TextMatrix(.Rows - 1, .ColIndex("替诊医生ID")) = Nvl(rsPlan!替诊医生id)
                    .TextMatrix(.Rows - 1, .ColIndex("替诊开始时间")) = Nvl(rsPlan!替诊开始时间)
                    .TextMatrix(.Rows - 1, .ColIndex("替诊终止时间")) = Nvl(rsPlan!替诊终止时间)
                End If
                .Rows = .Rows + 1
            End With
        End If
        rsPlan.MoveNext
    Loop
    mblnNotClick = True
    If mstr时段s <> "" Then
        mstr时段s = Mid(mstr时段s, 2)
        strTime = Split(mstr时段s, ",")
        For i = 0 To UBound(strTime)
            cboTime.AddItem strTime(i)
        Next i
    End If
    cboTime.ListIndex = 0
    mblnNotClick = False
    If rsPlan.RecordCount = 0 Then
        vsfPlan.Redraw = flexRDDirect
        vsfList.Visible = False
        picSplit.Visible = False
        vsfPlan.Height = 4915
        vsfPlan.Select 1, 1
        Exit Sub
    End If
    Call ShowRow
    If vsfPlan.Rows <> 2 Then vsfPlan.Rows = vsfPlan.Rows - 1
    For i = 0 To vsfPlan.Rows - 1
        vsfPlan.RowHeight(i) = 322
    Next i
'    vsfPlan.AutoSize 0, vsfPlan.Cols - 1
    zl_vsGrid_Para_Restore 1115, vsfPlan, Me.Name, "vsfPlan"
    vsfPlan.Redraw = flexRDDirect
    If vsfPlan.TextMatrix(1, vsfPlan.ColIndex("号码")) = "" Then
'        MsgBox "当前预约记录没有可以换诊的记录,无法换诊!", vbInformation, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save 1115, vsfPlan, Me.Name, "vsfPlan"
End Sub

Private Sub picReg_Resize()
    On Error Resume Next
    With dtpMain
        .Left = picReg.ScaleWidth - .Width - 15
    End With
    With vsfPlan
        .Width = dtpMain.Left - 60
    End With
    With vsfList
        .Width = picReg.ScaleWidth - 150
        .Height = picReg.ScaleHeight - 2600
    End With
    picSplit.Width = picReg.ScaleWidth
End Sub

Private Sub ShowRow()
    Dim i As Integer, blnHide As Boolean
    Dim blnEnable As Boolean, strTimeRange As String
    If cboTime.Text <> "所有" Then strTimeRange = cboTime.Text
    With vsfPlan
        For i = 1 To .Rows - 1
            blnHide = False
            If txtFilter <> "" Then
                blnHide = True
                If .TextMatrix(i, .ColIndex("号码")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("科室")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("号类")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("项目")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("医生")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("科室简码"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("医生简码"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("项目简码"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
            End If
            If cboTime.Text <> "所有" Then
                If InStr(strTimeRange & ",", .TextMatrix(i, .ColIndex("时段"))) = 0 Then blnHide = True
            End If
            .RowHidden(i) = blnHide
        Next i
    End With
    blnEnable = False
    With vsfPlan
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                .Select i, 1
                blnEnable = True
                Call vsfPlan_EnterCell
                Exit For
            End If
        Next i
        If blnEnable = False Then
            picTime.Visible = True
            vsfList.Visible = False
            picSplit.Visible = False
            vsfPlan.Height = 4915
        End If
    End With
End Sub

Private Sub txtFilter_Change()
    Call ShowRow
End Sub

Public Function CheckLimit(lng记录ID As Long) As Boolean
    On Error GoTo errHandle
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim rsUsed As ADODB.Recordset, lng合作单位已用数量 As Long
    Dim rsUnit As ADODB.Recordset, lng合作单位数量 As Long
    
    strSQL = "Select Nvl(限约数,限号数) As 限约数,已约数,Nvl(是否独占,0) As 是否独占,是否序号控制,是否分时段 From 临床出诊记录 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    strSQL = "Select 名称 As 合作单位, 控制方式, 序号, 数量 From 临床出诊挂号控制记录 Where 记录id = [1] And 类型 = 1"
    Set rsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    strSQL = "Select Count(1) As 数量 From 病人挂号记录 Where 出诊记录id = [1] And 合作单位 Is Not Null And 记录状态=1"
    Set rsUsed = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID)
    If Not rsUsed.EOF Then
        lng合作单位已用数量 = Val(Nvl(rsUsed!数量))
    End If
    If Not rsTemp.EOF Then
        If Val(Nvl(rsTemp!是否序号控制)) = 1 Then
            If rsUnit.EOF Then
                lng合作单位数量 = 0
            Else
                If Val(Nvl(rsUnit!控制方式)) = 2 Then
                    If Val(rsTemp!是否独占) = 0 Then
                        lng合作单位数量 = 0
                    Else
                        Do While Not rsUnit.EOF
                            lng合作单位数量 = lng合作单位数量 + Val(Nvl(rsUnit!数量))
                            rsUnit.MoveNext
                        Loop
                    End If
                ElseIf Val(Nvl(rsUnit!控制方式)) = 1 Then
                    If Val(rsTemp!是否独占) = 0 Then
                        lng合作单位数量 = 0
                    Else
                        Do While Not rsUnit.EOF
                            lng合作单位数量 = lng合作单位数量 + Int(Val(Nvl(rsUnit!数量)) * Val(Nvl(rsTemp!限约数)) / 100)
                            rsUnit.MoveNext
                        Loop
                    End If
                ElseIf Val(Nvl(rsUnit!控制方式)) = 3 Then
                    Do While Not rsUnit.EOF
                        lng合作单位数量 = lng合作单位数量 + Val(Nvl(rsUnit!数量))
                        rsUnit.MoveNext
                    Loop
                End If
            End If
            If Not IsNull(rsTemp!限约数) Then
                If Val(Nvl(rsTemp!已约数)) + lng合作单位数量 - lng合作单位已用数量 >= Val(Nvl(rsTemp!限约数)) Then
                    MsgBox "当前预约号码超过了限制数量" & Val(Nvl(rsTemp!限约数)) & "(其中包含合作单位限制数量" & lng合作单位数量 & "),不能继续预约!", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            If Val(Nvl(rsTemp!是否分时段)) = 1 Then
                If rsUnit.EOF Then
                    lng合作单位数量 = 0
                Else
                    If Val(Nvl(rsUnit!控制方式)) = 3 Then
                        rsUnit.Filter = "序号=" & Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col))
                        If rsUnit.EOF Then
                            If Not IsNull(rsTemp!限约数) Then
                                If Val(Nvl(rsTemp!已约数)) >= Val(Nvl(rsTemp!限约数)) Then
                                    MsgBox "当前预约号码超过了限制数量" & Val(Nvl(rsTemp!限约数)) & ",不能继续预约!", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                        Else
                            If Not IsNull(rsTemp!限约数) Then
                                If Val(Nvl(rsTemp!已约数)) >= Val(Nvl(rsTemp!限约数)) Then
                                    MsgBox "当前预约号码超过了限制数量" & Val(Nvl(rsTemp!限约数)) & ",不能继续预约!", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If Val(Nvl(rsUnit!控制方式)) = 2 Then
                            If Val(rsTemp!是否独占) = 0 Then
                                lng合作单位数量 = 0
                            Else
                                Do While Not rsUnit.EOF
                                    lng合作单位数量 = lng合作单位数量 + Val(Nvl(rsUnit!数量))
                                    rsUnit.MoveNext
                                Loop
                            End If
                        ElseIf Val(Nvl(rsUnit!控制方式)) = 1 Then
                            If Val(rsTemp!是否独占) = 0 Then
                                lng合作单位数量 = 0
                            Else
                                Do While Not rsUnit.EOF
                                    lng合作单位数量 = lng合作单位数量 + Int(Val(Nvl(rsUnit!数量)) * Val(Nvl(rsTemp!限约数)) / 100)
                                    rsUnit.MoveNext
                                Loop
                            End If
                        End If
                        If Not IsNull(rsTemp!限约数) Then
                            If Val(Nvl(rsTemp!已约数)) + lng合作单位数量 - lng合作单位已用数量 >= Val(Nvl(rsTemp!限约数)) Then
                                MsgBox "当前预约号码超过了限制数量" & Val(Nvl(rsTemp!限约数)) & "(其中包含合作单位限制数量" & lng合作单位数量 & "),不能继续预约!", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Else
                If rsUnit.EOF Then
                    lng合作单位数量 = 0
                Else
                    If Val(Nvl(rsUnit!控制方式)) = 2 Then
                        If Val(rsTemp!是否独占) = 0 Then
                            lng合作单位数量 = 0
                        Else
                            Do While Not rsUnit.EOF
                                lng合作单位数量 = lng合作单位数量 + Val(Nvl(rsUnit!数量))
                                rsUnit.MoveNext
                            Loop
                        End If
                    ElseIf Val(Nvl(rsUnit!控制方式)) = 1 Then
                        If Val(rsTemp!是否独占) = 0 Then
                            lng合作单位数量 = 0
                        Else
                            Do While Not rsUnit.EOF
                                lng合作单位数量 = lng合作单位数量 + Int(Val(Nvl(rsUnit!数量)) * Val(Nvl(rsTemp!限约数)) / 100)
                                rsUnit.MoveNext
                            Loop
                        End If
                    End If
                End If
                If Not IsNull(rsTemp!限约数) Then
                    If Val(Nvl(rsTemp!已约数)) + lng合作单位数量 - lng合作单位已用数量 >= Val(Nvl(rsTemp!限约数)) Then
                        MsgBox "当前预约号码超过了限制数量" & Val(Nvl(rsTemp!限约数)) & "(其中包含合作单位限制数量" & lng合作单位数量 & "),不能继续预约!", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    CheckLimit = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsfList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnNotClick = False Then
        With vsfList
            If .TextMatrix(NewRow, NewCol) = "" Then Cancel = True
            If .Cell(flexcpFontStrikethru, NewRow, NewCol) = True Then Cancel = True
            If .Cell(flexcpForeColor, NewRow, NewCol) <> vbBlack And .Cell(flexcpForeColor, NewRow, NewCol) <> 2 Then Cancel = True
        End With
    End If
End Sub

Private Sub vsfList_DblClick()
'    vsfList.Select vsfList.Row, vsfList.Col
    If SaveData = True Then Unload Me
End Sub

Private Sub vsfList_EnterCell()
    If vsfList.Row >= vsfList.Rows Then Exit Sub
    If vsfList.Col >= vsfList.Cols Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), ":") = 0 Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "-") = 0 Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "预约") > 0 Then
        dtpTime.Value = Split(Split(vsfList.TextMatrix(vsfList.Row, vsfList.Col), vbCrLf)(0), "-")(0)
    Else
        dtpTime.Value = Split(Split(vsfList.TextMatrix(vsfList.Row, vsfList.Col), vbCrLf)(1), "-")(0)
    End If
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "替") > 0 Then
        vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生"))
    Else
        vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = ""
    End If
    
End Sub

Private Sub vsfplan_DblClick()
    If SaveData = True Then Unload Me
End Sub

Private Sub vsfPlan_EnterCell()
    Dim i As Integer, j As Integer
    Dim sngTime As Single, datApp As Date
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim blnFind As Boolean
    If mblnInit = False Then
        If mblnChangeByCode Then Exit Sub
        sngTime = Timer
        If Format(sngTime, "0.000") - Format(msngTime, "0.000") < 0.1 Then
            mblnChangeByCode = True
            If mlngRow <> 0 Then vsfPlan.Select mlngRow, 0
            mblnChangeByCode = False
            Exit Sub
        End If
        msngTime = Timer
        mlngRow = vsfPlan.Row
    End If
    With vsfPlan
        If Val(.TextMatrix(.Row, .ColIndex("分时段"))) = 1 Then
            With vsfList
                For i = 0 To .Rows - 1
                    .RowHeight(i) = 500
                    .Cell(flexcpFontBold, i, 0) = True
                    .Cell(flexcpFontSize, i, 0) = 16
                Next i
            End With
            picTime.Visible = False
            picSplit.Visible = True
            vsfList.Visible = True
            vsfPlan.Height = picSplit.Top - vsfPlan.Top - 15
            Call LoadTimePlan
            blnFind = False
            With vsfList
                For i = 0 To .Rows - 1
                    If blnFind = False Then
                        For j = 1 To .Cols - 1
                            If (vsfList.Cell(flexcpForeColor, i, j) = vbBlack Or vsfList.Cell(flexcpForeColor, i, j) = 2) And vsfList.Cell(flexcpFontStrikethru, i, j) = False Then
                                .Select i, j
                                blnFind = True
                                Exit For
                            End If
                        Next j
                    End If
                Next i
            End With
        Else
            If mblnAppointmentChange = False Then
                If vsfPlan.TextMatrix(.Row, .ColIndex("预约时间")) <> "" And IsDate(vsfPlan.TextMatrix(.Row, .ColIndex("预约时间"))) = True Then
                    If Format(vsfPlan.TextMatrix(.Row, .ColIndex("出诊日期")), "yyyy-mm-dd") <> Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
                        dtpTime.Value = vsfPlan.TextMatrix(.Row, .ColIndex("终止时间"))
                    Else
                        dtpTime.Value = vsfPlan.TextMatrix(.Row, .ColIndex("预约时间"))
                    End If
                Else
                    dtpTime.Value = zlDatabase.Currentdate
                End If
            End If
            If Val(vsfPlan.TextMatrix(.Row, .ColIndex("序号控制"))) = 0 Then
                picTime.Visible = True
                vsfList.Visible = False
                picSplit.Visible = False
                vsfPlan.Height = 4915
            Else
                With vsfList
                    For i = 0 To .Rows - 1
                        .RowHeight(i) = 350
                        For j = 0 To .Cols - 1
                            .Cell(flexcpFontBold, i, j) = True
                            .Cell(flexcpFontSize, i, j) = 16
                        Next j
                    Next i
                End With
                picTime.Visible = True
                vsfList.Visible = True
                picSplit.Visible = True
                vsfPlan.Height = picSplit.Top - vsfPlan.Top - 15
                Call LoadSerialPlan
                blnFind = False
                With vsfList
                    For i = 0 To .Rows - 1
                        If blnFind = False Then
                            For j = 1 To .Cols - 1
                                If (vsfList.Cell(flexcpForeColor, i, j) = vbBlack Or vsfList.Cell(flexcpForeColor, i, j) = 2) And vsfList.Cell(flexcpFontStrikethru, i, j) = False Then
                                    .Select i, j
                                    blnFind = True
                                    Exit For
                                End If
                            Next j
                        End If
                    Next i
                End With
            End If
            If vsfList.Visible Then
                If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("分时段"))) = 1 Then
                    If Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)) <> 0 Then
                        strSQL = "Select 开始时间 From 临床出诊序号控制 Where 记录ID=[1] And 序号=[2]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)))
                        If Not rsTemp.EOF Then
                            datApp = CDate(Format(rsTemp!开始时间, "yyyy-mm-dd hh:mm:ss"))
                        Else
                            datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
                        End If
                    Else
                        datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
                    End If
                Else
                    datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
                End If
            Else
                datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
            End If
            If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间")) <> "" Then
                If datApp >= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间"))) And datApp <= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间"))) Then
                    vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = ""
                Else
                    vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生"))
                End If
            End If
        End If
    End With
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsfPlan.Height + Y < 500 Or vsfList.Height - Y < 500 Then Exit Sub
                
        picSplit.Top = picSplit.Top + Y
        vsfPlan.Height = vsfPlan.Height + Y
        vsfList.Top = vsfList.Top + Y
        vsfList.Height = vsfList.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub LoadSerialPlan()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim intCurrentTime As Integer, intCol As Integer
    Dim blnFind As Boolean, i As Integer, j As Integer
    
    vsfList.Redraw = flexRDNone
    vsfList.Clear
    vsfList.Rows = 2
    vsfList.Cols = 10
    vsfList.FixedRows = 0
    vsfList.FixedCols = 0
    intCol = 0
    
    strSQL = "Select 序号, 开始时间, 终止时间, 是否预约, 挂号状态, 名称, 类型 From 临床出诊序号控制 Where 记录id = [1] Order By 序号,开始时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
    Do While Not rsTemp.EOF
        With vsfList
            .TextMatrix(.Rows - 1, intCol) = Nvl(rsTemp!序号)
            .Cell(flexcpData, .Rows - 1, intCol) = Nvl(rsTemp!序号)
            Select Case Val(Nvl(rsTemp!挂号状态))
                Case 0
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlack
                Case 1 '已挂
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = vbRed
                    .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
                Case 2
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = vbGreen
                Case 3
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlue
                Case 4
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = vbRed
                Case 5
                    .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
            End Select
            intCol = intCol + 1
            If intCol > 9 Then
                intCol = 0
                .Rows = .Rows + 1
            End If
        End With
        rsTemp.MoveNext
    Loop
    With vsfList
        For i = 0 To .Rows - 1
            .RowHeight(i) = 400
        Next i
        For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                .Cell(flexcpFontBold, i, j) = True
                .Cell(flexcpFontSize, i, j) = 12
            Next j
        Next i
        For i = 0 To .Cols - 1
            .ColWidth(i) = 1000
            .ColAlignment(i) = flexAlignCenterCenter
        Next i
    End With
    blnFind = False
    With vsfList
        For i = 0 To .Rows - 1
            If blnFind = False Then
                For j = 0 To .Cols - 1
                    If (vsfList.Cell(flexcpForeColor, i, j) = vbBlack Or vsfList.Cell(flexcpForeColor, i, j) = 2) And vsfList.Cell(flexcpFontStrikethru, i, j) = False And vsfList.TextMatrix(i, j) <> "" Then
                        .Select i, j
                        Call vsfList_EnterCell
                        blnFind = True
                        Exit For
                    End If
                Next j
            End If
        Next i
        mblnNotClick = True
        If blnFind = False Then .Select 0, 0
        mblnNotClick = False
    End With
    vsfList.RowHidden(0) = True
    vsfList.Redraw = flexRDDirect
End Sub

Private Sub LoadTimePlan()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim intCurrentTime As Integer, intCol As Integer, datNow As Date
    Dim rsUnit As ADODB.Recordset, lng合作单位人数 As Long, lng已挂人数 As Long
    Dim i As Long, datTime As Date, blnFind As Boolean, j As Integer
    Dim rsTmp As ADODB.Recordset
    vsfList.Redraw = flexRDNone
    vsfList.Clear
    vsfList.Rows = 1
    vsfList.Cols = 2
    vsfList.FixedRows = 0
    vsfList.FixedCols = 1
    intCol = 0
    datTime = dtpMain.Selection.Blocks(0).DateBegin
    datNow = zlDatabase.Currentdate
    intCurrentTime = -1
    If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("序号控制"))) = 1 Then
        strSQL = "Select 序号, To_Char(开始时间,'hh24:mi:ss') As 开始时间, 开始时间 As 序号时间, To_Char(终止时间,'hh24:mi:ss') As 终止时间, 是否预约, 挂号状态, 名称, 类型 From 临床出诊序号控制 Where 记录id = [1] And 开始时间 <> 终止时间 Order By 序号,开始时间"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        strSQL = "Select 序号 From 临床出诊挂号控制记录 Where 记录id=[1] And 类型=1 And 控制方式=3"
        Set rsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        Do While Not rsTemp.EOF
            rsUnit.Filter = "序号=" & Val(rsTemp!序号)
            If rsUnit.EOF Then
                lng合作单位人数 = 0
            Else
                lng合作单位人数 = 1
            End If
            With vsfList
                If intCurrentTime = -1 Then
                    intCurrentTime = Val(Split(Nvl(rsTemp!开始时间), ":")(0))
                    .TextMatrix(.Rows - 1, 0) = Format(intCurrentTime, "00") & ":00"
                    intCol = intCol + 1
                Else
                    If intCurrentTime = Val(Split(Nvl(rsTemp!开始时间), ":")(0)) Then
                        intCol = intCol + 1
                    Else
                        .Rows = .Rows + 1
                        intCurrentTime = Val(Split(Nvl(rsTemp!开始时间), ":")(0))
                        .TextMatrix(.Rows - 1, 0) = Format(intCurrentTime, "00") & ":00"
                        intCol = 1
                    End If
                End If
                
                If intCol >= .Cols Then .Cols = .Cols + 1
                If vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) <> "" And _
                   Format(Nvl(rsTemp!序号时间), "yyyy-mm-dd hh:mm:ss") >= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间")), "yyyy-mm-dd hh:mm:ss") And _
                   Format(Nvl(rsTemp!序号时间), "yyyy-mm-dd hh:mm:ss") <= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊终止时间")), "yyyy-mm-dd hh:mm:ss") Then
                    .TextMatrix(.Rows - 1, intCol) = Nvl(rsTemp!序号) & "(替)" & vbCrLf & Format(Nvl(rsTemp!开始时间), "hh:mm") & "-" & Format(Nvl(rsTemp!终止时间), "hh:mm")
                Else
                    .TextMatrix(.Rows - 1, intCol) = Nvl(rsTemp!序号) & vbCrLf & Format(Nvl(rsTemp!开始时间), "hh:mm") & "-" & Format(Nvl(rsTemp!终止时间), "hh:mm")
                End If
                .Cell(flexcpData, .Rows - 1, intCol) = Nvl(rsTemp!序号)
                Select Case Val(Nvl(rsTemp!挂号状态))
                    Case 0
                        If lng合作单位人数 = 0 Then
                            .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlack
                        Else
                            .Cell(flexcpForeColor, .Rows - 1, intCol) = &HFF00FF
                        End If
                    Case 1 '已挂
                        .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
                    Case 2
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = vbGreen
                    Case 3
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlue
                    Case 4
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = vbRed
                    Case 5
                        .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
                    Case 6
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                End Select
                If CDate(Format(Nvl(rsTemp!序号时间), "yyyy-mm-dd hh:mm:ss")) < DateAdd("n", -1 * mlng预约有效时间, datNow) Then
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                End If
            End With
            rsTemp.MoveNext
        Loop
    Else
        strSQL = "Select 序号, To_Char(开始时间,'hh24:mi:ss') As 开始时间, 开始时间 As 序号时间, To_Char(终止时间,'hh24:mi:ss') As 终止时间, 数量, 是否预约, 挂号状态, 名称, 类型 From 临床出诊序号控制 Where 记录id = [1] And Nvl(是否预约,0) = 1 And 预约顺序号 Is Null Order By 序号,开始时间"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        strSQL = "Select Sum(Nvl(数量,0)) As 合作单位数量,序号 From 临床出诊挂号控制记录 Where 记录id=[1] And 类型=1 And 控制方式=3 Group By 序号"
        Set rsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        strSQL = "Select Count(1) As 已挂数量,序号 From 临床出诊序号控制 Where 记录ID=[1] And 预约顺序号 Is Null And Nvl(挂号状态,0) <> 0 Group By 序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        Do While Not rsTemp.EOF
            If Val(Nvl(rsTemp!数量)) <> 0 Then
                rsUnit.Filter = "序号=" & Val(rsTemp!序号)
                If rsUnit.EOF Then
                    lng合作单位人数 = 0
                Else
                    lng合作单位人数 = Val(Nvl(rsUnit!合作单位数量))
                End If
                rsTmp.Filter = "序号=" & Val(rsTemp!序号)
                If rsTmp.EOF Then
                    lng已挂人数 = 0
                Else
                    lng已挂人数 = Val(Nvl(rsTmp!已挂数量))
                End If
                With vsfList
                    If intCurrentTime = -1 Then
                        intCurrentTime = Val(Split(Nvl(rsTemp!开始时间), ":")(0))
                        .TextMatrix(.Rows - 1, 0) = Format(intCurrentTime, "00") & ":00"
                        intCol = intCol + 1
                    Else
                        If intCurrentTime = Val(Split(Nvl(rsTemp!开始时间), ":")(0)) Then
                            intCol = intCol + 1
                        Else
                            .Rows = .Rows + 1
                            intCurrentTime = Val(Split(Nvl(rsTemp!开始时间), ":")(0))
                            .TextMatrix(.Rows - 1, 0) = Format(intCurrentTime, "00") & ":00"
                            intCol = 1
                        End If
                    End If
                    
                    If intCol >= .Cols Then .Cols = .Cols + 1
                    If vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) <> "" And _
                       Format(Nvl(rsTemp!序号时间), "yyyy-mm-dd hh:mm:ss") >= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间")), "yyyy-mm-dd hh:mm:ss") And _
                       Format(Nvl(rsTemp!序号时间), "yyyy-mm-dd hh:mm:ss") <= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊终止时间")), "yyyy-mm-dd hh:mm:ss") Then
                        .TextMatrix(.Rows - 1, intCol) = Format(Nvl(rsTemp!开始时间), "hh:mm") & "-" & Format(Nvl(rsTemp!终止时间), "hh:mm") & vbCrLf & "预约" & Val(Nvl(rsTemp!数量)) - lng合作单位人数 - lng已挂人数 & "人(替)"
                    Else
                        .TextMatrix(.Rows - 1, intCol) = Format(Nvl(rsTemp!开始时间), "hh:mm") & "-" & Format(Nvl(rsTemp!终止时间), "hh:mm") & vbCrLf & "预约" & Val(Nvl(rsTemp!数量)) - lng合作单位人数 - lng已挂人数 & "人"
                    End If
                    .Cell(flexcpData, .Rows - 1, intCol) = Nvl(rsTemp!序号)
'                    Select Case Val(Nvl(rsTemp!挂号状态))
'                        Case 0
'                            .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlack
'                        Case 1 '已挂
'                            .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
'                        Case 2
'                            .Cell(flexcpForeColor, .Rows - 1, intCol) = vbGreen
'                        Case 3
'                            .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlue
'                        Case 4
'                            .Cell(flexcpForeColor, .Rows - 1, intCol) = vbRed
'                        Case 5
'                            .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
'                    End Select
                    If Val(Nvl(rsTemp!数量)) - lng合作单位人数 - lng已挂人数 = 0 Then
                        .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlack
                    End If
                    If CDate(Format(Nvl(rsTemp!序号时间), "yyyy-mm-dd hh:mm:ss")) < DateAdd("n", -1 * mlng预约有效时间, datNow) Then
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                    End If
                End With
            End If
            rsTemp.MoveNext
        Loop
    End If
    With vsfList
        For i = 0 To .Rows - 1
            .RowHeight(i) = 500
            .Cell(flexcpFontBold, i, 0) = True
            .Cell(flexcpFontSize, i, 0) = 20
        Next i
        For i = 0 To .Cols - 1
            .ColWidth(i) = 1500
            If i = 0 Then
                .ColAlignment(i) = flexAlignCenterTop
            Else
                .ColAlignment(i) = flexAlignCenterCenter
            End If
        Next i
    End With
    blnFind = False
    With vsfList
        For i = 0 To .Rows - 1
            If blnFind = False Then
                For j = 1 To .Cols - 1
                    If (vsfList.Cell(flexcpForeColor, i, j) = vbBlack Or vsfList.Cell(flexcpForeColor, i, j) = 2) And vsfList.Cell(flexcpFontStrikethru, i, j) = False And vsfList.TextMatrix(i, j) <> "" Then
                        .Select i, j
                        Call vsfList_EnterCell
                        blnFind = True
                        Exit For
                    End If
                Next j
            End If
        Next i
        mblnNotClick = True
        If blnFind = False Then .Select 0, 0
        mblnNotClick = False
    End With
    vsfList.Redraw = flexRDDirect
End Sub

Private Sub cboTime_Click()
    If mblnNotClick Then Exit Sub
    Call ShowRow
End Sub

Private Sub cboTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub vsfPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfPlan
        If OldRow < .Rows Then
            If OldRow Mod 2 = 1 Then
                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000005
            Else
                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &HEEEEEE
            End If
        End If
        .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = 16772055
    End With
End Sub

