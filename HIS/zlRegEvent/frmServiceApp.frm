VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "CODEJOCK.CALENDAR.V16.3.1.OCX"
Begin VB.Form frmServiceApp 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picReg 
      BorderStyle     =   0  'None
      Height          =   5280
      Left            =   2115
      ScaleHeight     =   5280
      ScaleWidth      =   8115
      TabIndex        =   25
      Top             =   525
      Width           =   8115
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   75
         Width           =   1125
      End
      Begin VB.PictureBox picSplit 
         Height          =   50
         Left            =   15
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   4065
         TabIndex        =   54
         Top             =   3105
         Width           =   4065
      End
      Begin VB.CommandButton cmdDirectApp 
         Height          =   315
         Left            =   5295
         Picture         =   "frmServiceApp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   68
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   4380
         TabIndex        =   43
         Top             =   75
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   163971074
         CurrentDate     =   42340
      End
      Begin VB.TextBox txtFilter 
         Height          =   315
         Left            =   465
         TabIndex        =   41
         ToolTipText     =   "可以通过输入号码,医生,科室,项目名称及其简码进行快速过滤查找"
         Top             =   68
         Width           =   1320
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2415
         Left            =   60
         TabIndex        =   46
         Top             =   4555
         Width           =   5925
         _cx             =   10451
         _cy             =   4260
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
         HighLight       =   0
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
         FormatString    =   $"frmServiceApp.frx":058A
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
      Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
         Height          =   2415
         Left            =   60
         TabIndex        =   45
         Top             =   450
         Width           =   3360
         _cx             =   5927
         _cy             =   4260
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
         Rows            =   6
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   322
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "时间段"
         Height          =   180
         Left            =   1875
         TabIndex        =   55
         Top             =   135
         Width           =   540
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "预约时间"
         Height          =   180
         Left            =   3630
         TabIndex        =   47
         Top             =   135
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "号码"
         Height          =   180
         Left            =   60
         TabIndex        =   40
         Top             =   135
         Width           =   360
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   1065
      ScaleHeight     =   2550
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   405
      Width           =   4845
      Begin VB.TextBox txtMarriage 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2130
         Width           =   1710
      End
      Begin VB.TextBox txtJob 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2130
         Width           =   1710
      End
      Begin VB.TextBox txtNation 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1725
         Width           =   1710
      End
      Begin VB.TextBox txtRace 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1725
         Width           =   1710
      End
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1335
         Width           =   4230
      End
      Begin VB.TextBox txtPhone 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   930
         Width           =   1710
      End
      Begin VB.TextBox txtFeeType 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   930
         Width           =   1710
      End
      Begin VB.TextBox txtBirth 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   525
         Width           =   1710
      End
      Begin VB.TextBox txtAge 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   525
         Width           =   690
      End
      Begin VB.TextBox txtGender 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   525
         Width           =   465
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   120
         Width           =   1710
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   1710
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "婚姻状况"
         Height          =   180
         Left            =   2340
         TabIndex        =   23
         Top             =   2205
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "职业"
         Height          =   180
         Left            =   165
         TabIndex        =   21
         Top             =   2205
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "国籍"
         Height          =   180
         Left            =   165
         TabIndex        =   19
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "民族"
         Height          =   180
         Left            =   2700
         TabIndex        =   17
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "现住址"
         Height          =   180
         Left            =   -15
         TabIndex        =   15
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "联系电话"
         Height          =   180
         Left            =   2340
         TabIndex        =   13
         Top             =   1005
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "费别"
         Height          =   180
         Left            =   165
         TabIndex        =   11
         Top             =   1005
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "出生日期"
         Height          =   180
         Left            =   2340
         TabIndex        =   9
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   180
         Left            =   1185
         TabIndex        =   7
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Left            =   165
         TabIndex        =   5
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "身份证号"
         Height          =   180
         Left            =   2340
         TabIndex        =   3
         Top             =   195
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Left            =   165
         TabIndex        =   1
         Top             =   195
         Width           =   360
      End
   End
   Begin VB.PictureBox picApp 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   885
      ScaleHeight     =   4245
      ScaleWidth      =   4860
      TabIndex        =   26
      Top             =   2040
      Width           =   4890
      Begin VB.TextBox txtMoney 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1260
         Width           =   1500
      End
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1260
         Width           =   1500
      End
      Begin VB.TextBox txtRegTime 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   1500
      End
      Begin VB.TextBox txtReger 
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   480
         Width           =   1500
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1650
         Width           =   1500
      End
      Begin VB.TextBox txtDoc 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1650
         Width           =   1500
      End
      Begin VB.TextBox txtTimeEnd 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   870
         Width           =   1500
      End
      Begin VB.TextBox txtTimeBegin 
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   870
         Width           =   1500
      End
      Begin VB.TextBox txtState 
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   90
         Width           =   3915
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "挂号金额"
         Height          =   180
         Left            =   2550
         TabIndex        =   53
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "预约项目"
         Height          =   180
         Left            =   135
         TabIndex        =   51
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "登记时间"
         Height          =   180
         Left            =   2550
         TabIndex        =   39
         Top             =   555
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "登记人"
         Height          =   180
         Left            =   330
         TabIndex        =   37
         Top             =   555
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "预约科室"
         Height          =   180
         Left            =   150
         TabIndex        =   35
         Top             =   1725
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "预约医生"
         Height          =   180
         Left            =   2550
         TabIndex        =   33
         Top             =   1725
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "时间范围                     至"
         Height          =   180
         Left            =   150
         TabIndex        =   30
         Top             =   945
         Width           =   2790
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "登记原因"
         Height          =   180
         Left            =   150
         TabIndex        =   28
         Top             =   165
         Width           =   720
      End
   End
   Begin VB.PictureBox picDate 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3315
      Left            =   600
      ScaleHeight     =   3315
      ScaleWidth      =   4800
      TabIndex        =   48
      Top             =   2985
      Width           =   4800
      Begin XtremeCalendarControl.DatePicker dtpMain 
         Height          =   2895
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   3870
         _Version        =   1048579
         _ExtentX        =   6826
         _ExtentY        =   5106
         _StockProps     =   64
         AutoSize        =   0   'False
         ShowTodayButton =   0   'False
         ShowNoneButton  =   0   'False
         Show3DBorder    =   0
         MaxSelectionCount=   1
         AskDayMetrics   =   -1  'True
      End
   End
   Begin VB.Image imgApp 
      Height          =   240
      Left            =   1860
      Picture         =   "frmServiceApp.frx":0696
      Top             =   645
      Width           =   240
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   1230
      Top             =   1515
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmServiceApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnNotClick As Boolean
Private mrsInfo As ADODB.Recordset
Private mlng消息ID As Long, mfrmMain As Object
Private mstr时段s As String, mlng预约有效时间
Private mblnUnload As Boolean, mdatCache As Date
Private mblnFirst As Boolean, mint同科限约数 As Integer
Private mint专家号预约限制 As Integer, mint病人预约科室数 As Integer
Private mblnInit As Boolean
Private mblnKeyPress As Boolean, mblnAppointmentChange As Boolean
Private mstrPriceGrade As String

Private Sub InitPanel()
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHandle
    Set objPane = dkpMain.CreatePane(1, 145, 80, DockLeftOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Title = "病人基本信息"
    objPane.Handle = picInfo.Hwnd
    objPane.MaxTrackSize.Width = 325
    objPane.MinTrackSize.Width = 325
    objPane.MaxTrackSize.Height = 170
    objPane.MinTrackSize.Height = 170
    
    
    Set objPane = dkpMain.CreatePane(2, 145, 90, DockBottomOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Title = "预约登记信息"
    objPane.Handle = picApp.Hwnd
    objPane.MaxTrackSize.Height = 138
    objPane.MinTrackSize.Height = 138
    
    Set objPane = dkpMain.CreatePane(3, 145, 120, DockBottomOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    objPane.Handle = picDate.Hwnd
    
    
    Set objPane = dkpMain.CreatePane(4, 320, 400, DockRightOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
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

Private Function zlGet当前星期几(Optional strDate As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当日是星期几
    '编制:刘兴洪
    '日期:2010-02-04 14:42:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, bln当前日期 As Boolean, strTemp As String
    If strDate = "" Then
        strSQL = "Select Decode(To_Char(Sysdate,'D'),'1','日','2','一','3','二','4','三','5','四','6','五','7','六',NULL) as 星期  From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        strSQL = "Select Decode(To_Char([1],'D'),'1','日','2','一','3','二','4','三','5','四','6','五','7','六','') As 星期 From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strDate))
    End If
    
    If rsTemp.EOF = True Then
        Exit Function
    End If
    strTemp = Nvl(rsTemp!星期)
    zlGet当前星期几 = strTemp
End Function

Private Function Check复诊(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long) As Boolean
'功能:判断病人是否再次到“相同临床性质的临床科室”挂号
'     包括挂过号的,或住过院的,复诊不好确定时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select a.临床科室id" & vbNewLine & _
    "       From (Select 执行部门id 临床科室id From 病人挂号记录 Where 病人id = [1] and 记录性质=1 and 记录状态=1 " & vbNewLine & _
    "             Union All" & vbNewLine & _
    "             Select 出院科室id 临床科室id From 病案主页 Where 病人id = [1]) a" & vbNewLine & _
    "       Where Exists (Select 1" & vbNewLine & _
    "                    From 临床部门 b" & vbNewLine & _
    "                    Where b.部门id = a.临床科室id And b.工作性质 = (Select 工作性质 From 临床部门 Where 部门id = [2] And Rownum=1))"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng执行部门ID)
    Check复诊 = Not rsTmp.EOF
End Function

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

Public Function SaveData() As Boolean
    On Error GoTo errHandle
    Dim i As Integer, k As Integer, int价格父号 As Integer, strDay As String, j As Integer
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim str登记时间 As String, lngSN As Long, str发生时间 As String, str付款方式 As String
    Dim lng挂号科室ID As Long, byt复诊 As Byte, strNO As String, cllPro As New Collection, rsCheck As ADODB.Recordset
    Dim str医生 As String, lng医生ID As Long, strSQL As String, blnNoDoc As Boolean, dat发生时间 As Date
    Dim bytMode As Byte, dat预约时间 As Date
    Dim strResult As String, bln专家号 As Boolean
    
    If vsfPlan.RowData(vsfPlan.Row) = "" Then
        MsgBox "请选择一个安排进行预约!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If vsfList.Visible Then
        If vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col) = "" Then
            MsgBox "请选择一个有效的序号进行预约!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If CheckLimit(Val(vsfPlan.RowData(vsfPlan.Row))) = False Then Exit Function
    
    If Not mrsInfo Is Nothing Then
        strSQL = "Select Zl_Fun_病人挂号记录_Check([1],[2],[3],[4],[5],[6]) As 检查结果 From Dual"
        bytMode = 1
        dat预约时间 = dtpMain.Selection.Blocks(0).DateBegin
        
        bln专家号 = vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("医生")) <> ""
        Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytMode, Val(Nvl(mrsInfo!病人ID)), vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("号别")), _
                                                Val(vsfPlan.RowData(vsfPlan.Row)), dat预约时间, IIf(bln专家号, 1, 0))
        If Not rsCheck.EOF Then
            strResult = Nvl(rsCheck!检查结果)
            If Val(Mid(strResult, 1, 1)) <> 0 Then
                MsgBox Mid(strResult, 3), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "有效性检查失败,无法继续！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    ReadRegistPrice Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("项目ID"))), False, False, txtFeeType.Text, rsItems, rsIncomes
    str登记时间 = "To_Date('" & zlDatabase.Currentdate & "','yyyy-mm-dd hh24:mi:ss')"
    strDay = zlGet当前星期几(dtpMain.Selection.Blocks(0).DateBegin)
    If vsfList.Visible Then
        lngSN = vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)
    End If
    If lngSN <> 0 Then
        If MsgBox("是否预约序号(" & lngSN & ")?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Function
    Else
        If MsgBox("是否预约该号?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Function
    End If
    
    If lngSN <> 0 Then
        If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("分时段"))) = 1 Then
            strSQL = "Select 开始时间 From 临床出诊序号控制 Where 记录ID=[1] And 序号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), lngSN)
            If Not rsTemp.EOF Then
                dat发生时间 = CDate(Format(rsTemp!开始时间, "yyyy-mm-dd hh:mm:ss"))
                str发生时间 = "To_Date('" & Format(rsTemp!开始时间, "yyyy-mm-dd hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
            Else
                dat发生时间 = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
                str发生时间 = "To_Date('" & Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00") & " ','YYYY-MM-DD HH24:MI:SS')"
            End If
        Else
            dat发生时间 = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
            str发生时间 = "To_Date('" & Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00") & " ','YYYY-MM-DD HH24:MI:SS')"
        End If
    Else
        dat发生时间 = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
        str发生时间 = "To_Date('" & Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00") & " ','YYYY-MM-DD HH24:MI:SS')"
    End If
    
    If dat发生时间 < DateAdd("n", -1 * mlng预约有效时间, zlDatabase.Currentdate) Then
        MsgBox "预约时间小于了可预约时间(" & Format(DateAdd("n", -1 * mlng预约有效时间, zlDatabase.Currentdate), "hh:mm:ss") & "),无法预约!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not (Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("分时段"))) = 1 And Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("序号控制"))) = 1) Then
        If Check有效时间段(Val(vsfPlan.RowData(vsfPlan.Row)), dat发生时间) = False Then
            MsgBox "当前选择的出诊记录在" & Format(dat发生时间, "yyyy-mm-dd hh:mm:ss") & "不出诊,请调整挂号时间!", vbInformation, gstrSysName
            If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
            Exit Function
        End If
    End If
    
    If lngSN = 0 Then
        strSQL = "Select Zl_Fun_Get临床出诊预约状态([1],[2]) As 预约检查 From Dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), dat发生时间)
    Else
        strSQL = "Select Zl_Fun_Get临床出诊预约状态([1],[2],[3]) As 预约检查 From Dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), dat发生时间, lngSN)
    End If
    If rsTemp.EOF Then
        MsgBox "当前选择的出诊记录无法预约!", vbInformation, gstrSysName
        Exit Function
    Else
        If Val(Mid(Nvl(rsTemp!预约检查), 1, 1)) <> 0 Then
            MsgBox "当前选择的出诊记录无法预约!" & vbCrLf & "原因:" & Mid(Nvl(rsTemp!预约检查), InStr(Nvl(rsTemp!预约检查), "|") + 1), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSQL = "Select Zl_临床出诊限制_Check([1],[2],[3]) As 适用性检查 From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), txtGender.Text, txtAge.Text)
'    If rsTemp.EOF Then
'        MsgBox "当前选择的病人不适用该号别!", vbInformation, gstrSysName
'        Exit Function
    If Not rsTemp.EOF Then
        If Val(Mid(Nvl(rsTemp!适用性检查), 1, 1)) <> 0 Then
            MsgBox "当前选择的病人不适用该号别!" & vbCrLf & "原因:" & Mid(Nvl(rsTemp!适用性检查), InStr(Nvl(rsTemp!适用性检查), "|") + 1), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    lng挂号科室ID = Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("科室ID")))
    byt复诊 = IIf(Check复诊(Val(txtName.Tag), lng挂号科室ID), 1, 0)
    strNO = zlDatabase.GetNextNo(12)
    
    With vsfPlan
        If .TextMatrix(.Row, .ColIndex("替诊医生姓名")) <> "" Then
            If dat发生时间 >= CDate(.TextMatrix(.Row, .ColIndex("替诊开始时间"))) And dat发生时间 <= CDate(.TextMatrix(.Row, .ColIndex("替诊终止时间"))) Then
                str医生 = .TextMatrix(.Row, .ColIndex("替诊医生姓名"))
                lng医生ID = .TextMatrix(.Row, .ColIndex("替诊医生ID"))
            Else
                str医生 = .TextMatrix(.Row, .ColIndex("医生"))
                lng医生ID = Val(.TextMatrix(.Row, .ColIndex("医生ID")))
            End If
        End If
    End With
    
    
    strSQL = "Select 编码 From 医疗付款方式 Where 名称 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, Nvl(mrsInfo!医疗付款方式))
    If rsTemp.RecordCount <> 0 Then
        str付款方式 = Nvl(rsTemp!编码)
    Else
        strSQL = "Select 编码 From 医疗付款方式 Where 缺省标志 = 1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
        If rsTemp.RecordCount <> 0 Then
            str付款方式 = Nvl(rsTemp!编码)
        End If
    End If
    
    k = 1: rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        int价格父号 = k
        rsIncomes.Filter = "项目ID=" & rsItems!项目ID
        For j = 1 To rsIncomes.RecordCount
            strSQL = _
            "zl_病人挂号记录_出诊_INSERT(" & ZVal(vsfPlan.RowData(vsfPlan.Row)) & "," & Val(Nvl(mrsInfo!病人ID)) & "," & IIf(IsNull(mrsInfo!门诊号), "NULL", mrsInfo!门诊号) & ",'" & txtName.Text & "','" & txtGender.Text & "'," & _
                     "'" & txtAge.Text & "','" & str付款方式 & "','" & txtFeeType.Text & "','" & strNO & "'," & _
                     "'" & "" & "'," & k & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & IIf(rsItems!性质 = 2, 1, "NULL") & "," & _
                     "'" & rsItems!类别 & "'," & rsItems!项目ID & "," & rsItems!数次 & "," & rsIncomes!单价 & "," & _
                     rsIncomes!收入项目ID & ",'" & rsIncomes!收据费目 & "','" & "" & "'," & _
                      rsIncomes!应收 & "," & rsIncomes!实收 & "," & _
                     lng挂号科室ID & "," & UserInfo.部门ID & "," & IIf(rsItems!执行科室ID = 0, lng挂号科室ID, rsItems!执行科室ID) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                     str发生时间 & "," & str登记时间 & "," & _
                     "'" & str医生 & "'," & ZVal(lng医生ID) & "," & IIf(rsItems!性质 = 3, 1, IIf(rsItems!性质 = 4, 2, 0)) & "," & Val(Nvl(mrsInfo!项目特性)) & "," & _
                     "'" & vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("号别")) & "','" & IIf(str医生 = UserInfo.姓名, "", "") & "'," & ZVal(0) & "," & "NULL" & "," & _
                     ZVal(IIf(k = 1, 0, 0)) & "," & ZVal(IIf(k = 1, 0, 0)) & "," & _
                     ZVal(IIf(k = 1, 0, 0)) & "," & ZVal(Nvl(rsItems!保险大类id, 0)) & "," & _
                     ZVal(Nvl(rsItems!保险项目否, 0)) & "," & ZVal(Nvl(rsIncomes!统筹金额, 0)) & "," & _
                     "'" & "" & "'," & 1 & "," & 0 & ",'" & rsItems!保险编码 & "'," & byt复诊 & "," & ZVal(lngSN) & ",Null," & _
                     1 & ",'" & "" & "'," & _
                     0 & ","
            '卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSQL = strSQL & "NULL" & ","
            '结算卡序号_In 病人预交记录.结算卡序号%Type := Null,
            strSQL = strSQL & "NULL" & ","
            '卡号_In       病人预交记录.卡号%Type := Null,
            strSQL = strSQL & "'" & "" & "',"
            '交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSQL = strSQL & " NULL,"
            '交易说明_In   病人预交记录.交易说明%Type := Null,
            strSQL = strSQL & " NULL,"
            '合作单位_In   病人预交记录.合作单位%Type := Null
            strSQL = strSQL & " NULL,"
            '  操作类型_In   Number:=0
            strSQL = strSQL & 0 & ","
            '  险类_IN       病人挂号记录.险类%type:=null,
            strSQL = strSQL & "NULL" & ","
            '  结算模式_IN   NUMBER :=0,
            strSQL = strSQL & 0 & ","
            '  记帐费用_IN Number:=0
            strSQL = strSQL & 0 & ","
            '  退号重用_IN Number:=1
            strSQL = strSQL & 0 & ","
            '  冲预交病人ids_In Varchar2 := Null
            strSQL = strSQL & "'" & "" & "',"
            '  修正病人费别_In Number := 0
            strSQL = strSQL & 0 & ")"
            
            Call zlAddArray(cllPro, strSQL)
            '问题:31187:将挂号汇总单独出来
            If Val(vsfPlan.RowData(vsfPlan.Row)) <> 0 And k = 1 Then
                If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("医生")) = "" Then blnNoDoc = True
                strSQL = "zl_病人挂号汇总_Update("
                '  医生姓名_In   挂号安排.医生姓名%Type,
                strSQL = strSQL & IIf(blnNoDoc, "Null,", "'" & str医生 & "',")
                '  医生id_In     挂号安排.医生id%Type,
                strSQL = strSQL & "" & IIf(blnNoDoc, "0,", ZVal(lng医生ID) & ",")
                '  收费细目id_In 门诊费用记录.收费细目id%Type,
                strSQL = strSQL & "" & Val(Nvl(rsItems!项目ID)) & ","
                '  执行部门id_In 门诊费用记录.执行部门id%Type,
                strSQL = strSQL & "" & IIf(Val(Nvl(rsItems!执行科室ID)) = 0, lng挂号科室ID, Val(Nvl(rsItems!执行科室ID))) & ","
                '  发生时间_In   门诊费用记录.发生时间%Type,
                strSQL = strSQL & "" & str发生时间 & ","
                '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收,3-收费预约
                strSQL = strSQL & 1 & ","
                '  号码_In       挂号安排.号码%Type := Null
                strSQL = strSQL & "'" & vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("号别")) & "',0,"
                strSQL = strSQL & "" & Val(vsfPlan.RowData(vsfPlan.Row)) & ")"
                Call zlAddArray(cllPro, strSQL)
            End If
            
            k = k + 1
            rsIncomes.MoveNext
            Next j
        rsItems.MoveNext
    Next i
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, False, False
    
    strSQL = "Select ID From 病人挂号记录 Where No=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    strSQL = "Zl_患者服务中心_更新("
    strSQL = strSQL & mlng消息ID & ","
    strSQL = strSQL & "Null,'"
    strSQL = strSQL & UserInfo.姓名 & "','"
    strSQL = strSQL & UserInfo.编号 & "',"
    strSQL = strSQL & Val(rsTemp!ID) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Call mfrmMain.RefreshData
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Check有效时间段(lng记录ID As Long, datTime As Date) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
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

Private Sub cboTime_Click()
    If mblnNotClick Then Exit Sub
    Call ShowRow
End Sub

Private Sub cboTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmdDirectApp_Click()
    Call SaveData
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Cancel = True
End Sub

Public Sub LoadData(frmMain As Object, ByVal lngID As Long)
    On Error GoTo errHandle
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim datBegin As Date, datEnd As Date
    Dim datNow As Date
    Set mfrmMain = frmMain
    Call ClearData
    mlng消息ID = lngID
    strSQL = "Select a.病人Id, b.门诊号, b.姓名, b.身份证号, b.医疗付款方式, d.项目特性, b.性别, b.年龄, b.出生日期, b.费别, b.家庭电话, b.家庭地址, b.国籍, b.民族, b.职业, b.婚姻状况, a.通知原因 As 登记原因, a.登记人, a.开始时间," & vbNewLine & _
            "       a.终止时间, c.名称 As 预约科室, a.医生姓名 As 预约医生, a.项目ID, d.名称 As 项目名称, a.登记时间 " & vbNewLine & _
            "From 病人服务信息记录 A, 病人信息 B, 部门表 C, 收费项目目录 D" & vbNewLine & _
            "Where a.Id = [1] And a.病人id = b.病人id And a.科室id = c.Id And a.项目Id = d.Id "
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If mrsInfo.EOF Then
        MsgBox "读取病人信息失败,无法处理该条消息!"
        Exit Sub
    End If
    
    '价格等级
    If gintPriceGradeStartType >= 2 Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!病人ID)), 0, Nvl(mrsInfo!医疗付款方式, ""), , , mstrPriceGrade)
    Else
        mstrPriceGrade = gstrPriceGrade
    End If
    
    txtName.Text = Nvl(mrsInfo!姓名)
    txtName.Tag = Nvl(mrsInfo!病人ID)
    txtID.Text = Nvl(mrsInfo!身份证号)
    txtGender.Text = Nvl(mrsInfo!性别)
    txtAge.Text = Nvl(mrsInfo!年龄)
    txtBirth.Text = Nvl(mrsInfo!出生日期)
    txtFeeType.Text = Nvl(mrsInfo!费别)
    txtPhone.Text = Nvl(mrsInfo!家庭电话)
    txtAddress.Text = Nvl(mrsInfo!家庭地址)
    txtNation.Text = Nvl(mrsInfo!国籍)
    txtRace.Text = Nvl(mrsInfo!民族)
    txtJob.Text = Nvl(mrsInfo!职业)
    txtMarriage.Text = Nvl(mrsInfo!婚姻状况)
    txtState.Text = Nvl(mrsInfo!登记原因)
    txtReger.Text = Nvl(mrsInfo!登记人)
    txtTimeBegin.Text = Format(Nvl(mrsInfo!开始时间), "yyyy-mm-dd")
    txtTimeEnd.Text = Format(Nvl(mrsInfo!终止时间), "yyyy-mm-dd")
    txtDept.Text = Nvl(mrsInfo!预约科室)
    txtDoc.Text = Nvl(mrsInfo!预约医生)
    txtItem.Text = Nvl(mrsInfo!项目名称)
    txtMoney.Text = Format(Get项目金额(Val(Nvl(mrsInfo!项目ID)), mstrPriceGrade), "0.00")
    txtRegTime.Text = Format(Nvl(mrsInfo!登记时间), "yyyy-mm-dd hh:mm:ss")
    mblnInit = True
    mblnNotClick = True
    txtFilter.Text = txtDoc.Text
    mblnNotClick = False
    datBegin = CDate(txtTimeBegin.Text & " 00:00:00")
    datEnd = CDate(txtTimeEnd.Text & " 23:59:59")
    datNow = zlDatabase.Currentdate
    strSQL = "Select 出诊日期 " & vbNewLine & _
            "From 临床出诊记录" & vbNewLine & _
            "Where 替诊医生姓名 Is Null And 医生姓名 = [1] And 出诊日期 Between [2] And [3] Order By 出诊日期"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtDoc.Text, datBegin, datEnd)
    mblnNotClick = True
    If rsTemp.EOF Then
        If datNow > datBegin Then
            dtpMain.SelectRange datNow, datNow
            dtpMain.Select datNow
            dtpMain.EnsureVisibleSelection
            dtpMain.RedrawControl
        Else
            dtpMain.SelectRange datBegin, datBegin
            dtpMain.Select datBegin
            dtpMain.EnsureVisibleSelection
            dtpMain.RedrawControl
        End If
    Else
        If datNow > CDate(rsTemp!出诊日期) Then
            dtpMain.SelectRange datNow, datNow
            dtpMain.Select datNow
            dtpMain.EnsureVisibleSelection
            dtpMain.RedrawControl
        Else
            dtpMain.SelectRange CDate(rsTemp!出诊日期), CDate(rsTemp!出诊日期)
            dtpMain.Select CDate(rsTemp!出诊日期)
            dtpMain.EnsureVisibleSelection
            dtpMain.RedrawControl
        End If
    End If
    Do While Not rsTemp.EOF
        '日历染色
        rsTemp.MoveNext
    Loop
    mblnNotClick = False
    
    Call LoadPlan
    Call ShowRow
    mblnInit = False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadPlan()
    Dim strSQL As String, rsPlan As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim datApp As Date, i As Integer, dblMoney As Double, blnAdd As Boolean
    Dim strTime() As String, lngLeft As Long, intChar As Integer
    Dim str项目ids As String
    On Error GoTo errH
    datApp = dtpMain.Selection.Blocks(0).DateBegin
    mdatCache = datApp
    strSQL = "Select a.id, b.号类, b.号码 As 号别, c.名称 As 科室, c.简码 As 科室简码, b.科室Id, a.上班时段 As 时段, " & _
            "        d.名称 As 项目, zlSpellcode(d.名称) As 项目简码, a.替诊医生ID, a.替诊医生姓名, a.医生Id, a.医生姓名 As 医生, e.简码 As 医生简码, " & _
            "        a.项目id, a.限约数 As 限约, a.已约数 As 已约, Nvl(a.是否分时段,0) As 分时段, Nvl(a.是否序号控制,0) As 序号控制, " & _
            "        a.出诊日期, a.缺省预约时间, a.替诊开始时间, a.替诊终止时间, a.开始时间, a.终止时间 " & vbNewLine & _
            "From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 收费项目目录 D, 人员表 E" & vbNewLine & _
            "Where a.号源id = b.Id  And Nvl(C.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And a.项目id = d.Id And b.科室id = c.Id And (c.站点 Is Null Or c.站点 = '" & gstrNodeNo & "') " & _
            "      And (a.出诊日期 = [1] Or a.出诊日期 = [2]) And (a.开始时间 < Nvl(a.停诊开始时间,a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间,a.开始时间) Or Exists (Select 1 From 临床出诊序号控制 C,临床出诊记录 D Where D.ID=A.ID And C.记录ID=D.ID And Nvl(C.是否停诊,0) = 0 And D.是否序号控制 =1 And D.是否分时段 = 1 And C.开始时间 <> C.终止时间)) " & _
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
    vsfPlan.Rows = 1
    vsfPlan.Clear 1
    vsfPlan.Rows = 2
    If rsPlan.RecordCount <> 0 Then
        str项目ids = ""
        Do While Not rsPlan.EOF
            If InStr("," & str项目ids & ",", "," & Val(Nvl(rsPlan!项目ID)) & ",") = 0 Then
                str项目ids = str项目ids & "," & Val(Nvl(rsPlan!项目ID))
            End If
            rsPlan.MoveNext
        Loop
        
        rsPlan.MoveFirst
    End If
    
    If str项目ids <> "" Then
        str项目ids = Mid(str项目ids, 2)
    End If
    Set rsTemp = Get项目信息(str项目ids, mstrPriceGrade)
    
    Do While Not rsPlan.EOF
        blnAdd = True
        
        If blnAdd Then
            With vsfPlan
                rsTemp.Filter = "项目ID=" & Val(Nvl(rsPlan!项目ID))
                .RowData(.Rows - 1) = Val(Nvl(rsPlan!ID))
                .TextMatrix(.Rows - 1, .ColIndex("号类")) = Nvl(rsPlan!号类)
                .TextMatrix(.Rows - 1, .ColIndex("号别")) = Nvl(rsPlan!号别)
                .TextMatrix(.Rows - 1, .ColIndex("科室")) = Nvl(rsPlan!科室)
                .TextMatrix(.Rows - 1, .ColIndex("时段")) = Nvl(rsPlan!时段)
                If InStr("," & mstr时段s & ",", "," & Nvl(rsPlan!时段) & ",") = 0 Then
                    mstr时段s = mstr时段s & "," & Nvl(rsPlan!时段)
                End If
                .TextMatrix(.Rows - 1, .ColIndex("医生")) = Nvl(rsPlan!医生)
                If Nvl(rsPlan!替诊医生姓名) <> "" Then
                    .Cell(flexcpData, .Rows - 1, .ColIndex("替诊医生")) = Nvl(rsPlan!替诊医生姓名) & "(" & Format(Nvl(rsPlan!替诊开始时间), "hh:mm") & "-" & Format(Nvl(rsPlan!替诊终止时间), "hh:mm") & ")"
                    .TextMatrix(.Rows - 1, .ColIndex("替诊医生")) = ""
                    .TextMatrix(.Rows - 1, .ColIndex("替诊医生姓名")) = Nvl(rsPlan!替诊医生姓名)
                    .TextMatrix(.Rows - 1, .ColIndex("替诊医生ID")) = Nvl(rsPlan!替诊医生id)
                    .TextMatrix(.Rows - 1, .ColIndex("替诊开始时间")) = Nvl(rsPlan!替诊开始时间)
                    .TextMatrix(.Rows - 1, .ColIndex("替诊终止时间")) = Nvl(rsPlan!替诊终止时间)
                End If
                .TextMatrix(.Rows - 1, .ColIndex("项目")) = Nvl(rsPlan!项目)
                If rsTemp.EOF Then
                    .TextMatrix(.Rows - 1, .ColIndex("金额")) = "0.00"
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format(Val(Nvl(rsTemp!金额)), "0.00")
                End If
                .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format(Get项目金额(Val(Nvl(rsPlan!项目ID)), mstrPriceGrade), "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("限约")) = Nvl(rsPlan!限约)
                .TextMatrix(.Rows - 1, .ColIndex("已约")) = Val(Nvl(rsPlan!已约))
                .TextMatrix(.Rows - 1, .ColIndex("分时段")) = Nvl(rsPlan!分时段)
                .TextMatrix(.Rows - 1, .ColIndex("序号控制")) = Nvl(rsPlan!序号控制)
                .TextMatrix(.Rows - 1, .ColIndex("项目ID")) = Nvl(rsPlan!项目ID)
                .TextMatrix(.Rows - 1, .ColIndex("科室ID")) = Nvl(rsPlan!科室ID)
                .TextMatrix(.Rows - 1, .ColIndex("医生ID")) = Nvl(rsPlan!医生ID)
                .TextMatrix(.Rows - 1, .ColIndex("医生简码")) = Nvl(rsPlan!医生简码)
                .TextMatrix(.Rows - 1, .ColIndex("科室简码")) = Nvl(rsPlan!科室简码)
                .TextMatrix(.Rows - 1, .ColIndex("项目简码")) = Nvl(rsPlan!项目简码)
                .TextMatrix(.Rows - 1, .ColIndex("出诊日期")) = Format(Nvl(rsPlan!出诊日期), "yyyy-mm-dd")
                .TextMatrix(.Rows - 1, .ColIndex("预约时间")) = Format(Nvl(rsPlan!缺省预约时间), "yyyy-mm-dd hh:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("开始时间")) = Format(Nvl(rsPlan!开始时间), "yyyy-mm-dd hh:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("终止时间")) = Format(Nvl(rsPlan!终止时间), "yyyy-mm-dd hh:mm:ss")
                
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
        mfrmMain.ShowPanelText "当前日期没有可以预约的号码,无法预约!"
        vsfPlan.Redraw = flexRDDirect
        vsfList.Visible = False
        vsfPlan.Height = picReg.ScaleHeight - 500
        vsfPlan.Select 1, 1
        mblnUnload = True
        Exit Sub
    Else
        mfrmMain.ShowPanelText ""
    End If
    Call ShowRow
    If vsfPlan.Rows <> 2 Then vsfPlan.Rows = vsfPlan.Rows - 1
    vsfPlan.Select 1, 1
    vsfPlan.AutoSize 0, vsfPlan.Cols - 1
    zl_vsGrid_Para_Restore 1115, vsfPlan, Me.Name, "vsfPlan"
    vsfPlan.Redraw = flexRDDirect
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ShowRow()
    Dim i As Integer, blnHide As Boolean
    Dim blnEnable As Boolean, strTimeRange As String
    On Error GoTo errH
    If vsfPlan.Rows = 2 And vsfPlan.TextMatrix(1, vsfPlan.ColIndex("号别")) = "" Then Exit Sub
    If cboTime.Text <> "所有" Then strTimeRange = cboTime.Text
    With vsfPlan
        For i = 1 To .Rows - 1
            blnHide = False
            If txtFilter <> "" Then
                blnHide = True
                If .TextMatrix(i, .ColIndex("号别")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("科室")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("号类")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("项目")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("医生")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("科室简码"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("医生简码"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("项目简码"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
            End If
            If strTimeRange <> .TextMatrix(i, .ColIndex("时段")) And strTimeRange <> "" Then blnHide = True
'            If InStr(strTimeRange & ",", .TextMatrix(i, .ColIndex("时段"))) > 0 Then blnHide = True
            .RowHidden(i) = blnHide
        Next i
    End With
    blnEnable = False
    With vsfPlan
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                blnEnable = True
                .Select i, 1
                Call vsfPlan_EnterCell
                Exit For
            End If
        Next i
    End With
    If blnEnable = False Then
        If mblnKeyPress Then
            vsfList.Visible = False
            vsfPlan.Height = picReg.ScaleHeight - 500
            vsfPlan.Select 1, 1
        Else
            txtFilter.Text = ""
        End If
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ClearData()
    On Error GoTo errHandle
    txtName.Text = ""
    txtID.Text = ""
    txtGender.Text = ""
    txtAge.Text = ""
    txtBirth.Text = ""
    txtFeeType.Text = ""
    txtPhone.Text = ""
    txtAddress.Text = ""
    txtNation.Text = ""
    txtRace.Text = ""
    txtJob.Text = ""
    txtMarriage.Text = ""
    txtState.Text = ""
    txtReger.Text = ""
    txtTimeBegin.Text = ""
    txtTimeEnd.Text = ""
    txtDept.Text = ""
    txtDoc.Text = ""
    txtItem.Text = ""
    txtMoney.Text = ""
    cboTime.Clear
    cboTime.AddItem "所有"
    cboTime.ListIndex = cboTime.NewIndex
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dtpDate_Change()
    Dim str日期 As String, i As Integer, lngRow As Long
    Dim str发生时间 As String
    If Not dtpMain.Visible Then Exit Sub
    If Not dtpMain.Enabled Then Exit Sub
    
    str日期 = Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-MM-dd")

    If str日期 = "" Then str日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    str发生时间 = str日期 & " " & Format(dtpDate.Value, "hh:mm:00")
    lngRow = 0
    If CDate(str发生时间) > CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("终止时间"))) Then
        '超出时间的安排，重新寻找定位
        For i = 1 To vsfPlan.Rows - 1
            With vsfPlan
                If .TextMatrix(.Row, .ColIndex("号别")) = .TextMatrix(i, .ColIndex("号别")) And _
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
                If .TextMatrix(.Row, .ColIndex("号别")) = .TextMatrix(i, .ColIndex("号别")) And _
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

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call dtpDate_Change
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Call InitPanel
    Call InitGrid
    Call InitPara
End Sub

Private Sub InitPara()
    mlng预约有效时间 = -1 * Val(Split(zlDatabase.GetPara("预约限制时间", glngSys, 1111, "1|60") & "|", "|")(1))
    mint专家号预约限制 = Val(zlDatabase.GetPara("专家号预约限制", glngSys, , 0))
    mint病人预约科室数 = Val(zlDatabase.GetPara("病人预约科室数", glngSys, 1111, 0))
    mint同科限约数 = Val(zlDatabase.GetPara("病人同科限约N个号", glngSys, 1111, 0))
End Sub

Private Sub InitGrid()
    Dim i As Integer
    With vsfPlan
        .Cols = 26
        .Rows = 2
        .TextMatrix(0, 0) = "号类"
        .TextMatrix(0, 1) = "号别"
        .TextMatrix(0, 2) = "科室"
        .TextMatrix(0, 3) = "时段"
        .TextMatrix(0, 4) = "项目"
        .TextMatrix(0, 5) = "医生"
        .TextMatrix(0, 6) = "替诊医生"
        .TextMatrix(0, 7) = "金额"
        .TextMatrix(0, 8) = "限约"
        .TextMatrix(0, 9) = "已约"
        .TextMatrix(0, 10) = "分时段"
        .TextMatrix(0, 11) = "序号控制"
        .TextMatrix(0, 12) = "项目ID"
        .TextMatrix(0, 13) = "科室ID"
        .TextMatrix(0, 14) = "医生ID"
        .TextMatrix(0, 15) = "科室简码"
        .TextMatrix(0, 16) = "医生简码"
        .TextMatrix(0, 17) = "项目简码"
        .TextMatrix(0, 18) = "出诊日期"
        .TextMatrix(0, 19) = "预约时间"
        .TextMatrix(0, 20) = "替诊医生姓名"
        .TextMatrix(0, 21) = "替诊医生ID"
        .TextMatrix(0, 22) = "替诊开始时间"
        .TextMatrix(0, 23) = "替诊终止时间"
        .TextMatrix(0, 24) = "开始时间"
        .TextMatrix(0, 25) = "终止时间"
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "分时段" Or .ColKey(i) = "序号控制" Or .ColKey(i) = "项目ID" Or _
                .ColKey(i) = "科室ID" Or .ColKey(i) = "医生ID" Or .ColKey(i) = "科室简码" _
                Or .ColKey(i) = "医生简码" Or .ColKey(i) = "项目简码" Or .ColKey(i) = "预约时间" _
                Or .ColKey(i) = "替诊医生姓名" Or .ColKey(i) = "替诊医生ID" Or .ColKey(i) = "替诊开始时间" _
                Or .ColKey(i) = "替诊终止时间" Or .ColKey(i) = "开始时间" Or .ColKey(i) = "终止时间" Then .ColHidden(i) = True
        Next i
    End With
    With vsfList
        For i = 0 To .Rows - 1
            .RowHeight(i) = 500
            .Cell(flexcpFontBold, i, 0) = True
            .Cell(flexcpFontSize, i, 0) = 16
        Next i
    End With
    vsfList.Visible = False
    vsfPlan.Height = picReg.ScaleHeight - 500
End Sub

Private Sub dtpMain_SelectionChanged()
    Dim datNow As Date
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd") > Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
        If Format(datNow, "yyyy-mm-dd") > Format(mdatCache, "yyyy-mm-dd") Then
            mdatCache = datNow
        End If
        dtpMain.SelectRange mdatCache, mdatCache
        dtpMain.Select mdatCache
        dtpMain.EnsureVisibleSelection
        dtpMain.RedrawControl
    End If
    If mblnNotClick Then Exit Sub
    Call LoadPlan
    Call ShowRow
End Sub


Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save 1115, vsfPlan, Me.Name, "vsfPlan"
End Sub

Private Sub picDate_Resize()
    With dtpMain
'        .Height = picDate.ScaleHeight
        .Width = picDate.ScaleWidth
    End With
End Sub

Private Sub picReg_Resize()
    On Error Resume Next
    With vsfPlan
        .Width = picReg.ScaleWidth - 150
        If vsfList.Visible Then
            .Height = vsfList.Top - .Top - 100
            picSplit.Top = vsfList.Top - 60
        Else
            .Height = picReg.ScaleHeight - 500
        End If
    End With
    picSplit.Width = picReg.ScaleWidth
    With vsfList
        .Width = picReg.ScaleWidth - 150
        .Height = picReg.ScaleHeight - vsfList.Top
    End With
End Sub

Private Sub txtFilter_Change()
    If mblnNotClick Then Exit Sub
    mblnKeyPress = True
    Call ShowRow
    mblnKeyPress = False
End Sub

Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub vsfList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfList.TextMatrix(Row, Col) = "" Then Cancel = True
    If Not ((vsfList.Cell(flexcpForeColor, Row, Col) = vbBlack Or vsfList.Cell(flexcpForeColor, Row, Col) = 2) And vsfList.Cell(flexcpFontStrikethru, Row, Col) = False) Then Cancel = True
    vsfList.ComboList = "..."
    vsfList.CellButtonPicture = imgApp
End Sub


Private Sub vsfList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call SaveData
End Sub

Private Sub vsfPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer, blnMark As Boolean
    With vsfPlan
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                If blnMark Then
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HEEEEEE
                    blnMark = False
                Else
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80000005
                    blnMark = True
                End If
                If i = .Row Then
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = 16772055
                End If
            End If
        Next i
'        If OldRow < .Rows Then
'            If OldRow Mod 2 = 1 Then
'                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000005
'            Else
'                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &HEEEEEE
'            End If
'        End If
'        .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = 16772055
    End With
End Sub

Private Sub vsfPlan_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer, blnMark As Boolean
    With vsfPlan
'        For i = 1 To .Rows - 1
'            If .RowHidden(i) = False Then
'                If blnMark Then
'                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HEEEEEE
'                    blnMark = False
'                Else
'                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80000005
'                    blnMark = True
'                End If
'                If i = .Row Then
'                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = 16772055
'                End If
'            End If
'        Next i
'        If OldRow < .Rows Then
'            If OldRow Mod 2 = 1 Then
'                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000005
'            Else
'                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &HEEEEEE
'            End If
'        End If
'        .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = 16772055
    End With
End Sub

Private Sub vsfPlan_EnterCell()
    Dim i As Integer, j As Integer, datApp As Date
    Dim strSQL As String, rsTemp As ADODB.Recordset
    With vsfPlan
        If Val(.TextMatrix(.Row, .ColIndex("分时段"))) = 1 Then
            dtpDate.Enabled = False
            With vsfList
                For i = 0 To .Rows - 1
                    .RowHeight(i) = 500
                    .Cell(flexcpFontBold, i, 0) = True
                    .Cell(flexcpFontSize, i, 0) = 16
                Next i
            End With
            vsfList.Visible = True
            picSplit.Visible = True
            vsfPlan.Height = vsfList.Top - .Top - 100
            picSplit.Top = vsfList.Top - 60
            vsfList.Height = picReg.ScaleHeight - vsfList.Top
            Call LoadTimePlan
        Else
            dtpDate.Enabled = True
            If mblnAppointmentChange = False Then
                If vsfPlan.TextMatrix(.Row, .ColIndex("预约时间")) <> "" And IsDate(vsfPlan.TextMatrix(.Row, .ColIndex("预约时间"))) = True Then
                    If Format(vsfPlan.TextMatrix(.Row, .ColIndex("出诊日期")), "yyyy-mm-dd") <> Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
                        dtpDate.Value = vsfPlan.TextMatrix(.Row, .ColIndex("终止时间"))
                    Else
                        dtpDate.Value = vsfPlan.TextMatrix(.Row, .ColIndex("预约时间"))
                    End If
                Else
                    dtpDate.Value = zlDatabase.Currentdate
                End If
            End If

            If Val(vsfPlan.TextMatrix(.Row, .ColIndex("序号控制"))) = 0 Then
                vsfList.Visible = False
                picSplit.Visible = False
                vsfPlan.Height = picReg.ScaleHeight - 500
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
                vsfList.Visible = True
                picSplit.Visible = True
                vsfPlan.Height = vsfList.Top - .Top - 100
                picSplit.Top = vsfList.Top - 60
                vsfList.Height = picReg.ScaleHeight - vsfList.Top
                Call LoadSerialPlan
            End If
            If vsfList.Visible Then
                If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("分时段"))) = 1 Then
                    If Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)) <> 0 Then
                        strSQL = "Select 开始时间 From 临床出诊序号控制 Where 记录ID=[1] And 序号=[2]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)))
                        If Not rsTemp.EOF Then
                            datApp = CDate(Format(rsTemp!开始时间, "yyyy-mm-dd hh:mm:ss"))
                        Else
                            datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
                        End If
                    Else
                        datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
                    End If
                Else
                    datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
                End If
            Else
                datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
            End If
            If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间")) <> "" And vsfPlan.Row <> 0 Then
                If datApp >= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊开始时间"))) And datApp <= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊终止时间"))) Then
                    vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = ""
                Else
                    vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生"))
                End If
            End If
        End If
    End With
    cmdDirectApp.Visible = vsfList.Visible = False
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

Private Sub vsfList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnNotClick = False Then
        With vsfList
            If .TextMatrix(NewRow, NewCol) = "" Then Cancel = True
            If Not ((.Cell(flexcpForeColor, NewRow, NewCol) = vbBlack Or .Cell(flexcpForeColor, NewRow, NewCol) = 2) And .Cell(flexcpFontStrikethru, NewRow, NewCol) = False) Then Cancel = True
        End With
    End If
End Sub

Private Sub vsfList_EnterCell()
    If vsfList.Row >= vsfList.Rows Then Exit Sub
    If vsfList.Col >= vsfList.Cols Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), ":") = 0 Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "-") = 0 Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "预约") > 0 Then
        dtpDate.Value = Split(Split(vsfList.TextMatrix(vsfList.Row, vsfList.Col), vbCrLf)(0), "-")(0)
    Else
        dtpDate.Value = Split(Split(vsfList.TextMatrix(vsfList.Row, vsfList.Col), vbCrLf)(1), "-")(0)
    End If
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "替") > 0 Then
        vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("替诊医生"))
    Else
        vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("替诊医生")) = ""
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
    
    strSQL = "Select 序号, 开始时间, 终止时间, 是否预约, 挂号状态, 名称, 类型 From 临床出诊序号控制 Where 记录id = [1] And Nvl(是否预约,0)=1 Order By 序号,开始时间"
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
            If Val(Nvl(rsTemp!是否预约)) = 0 Then
                .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
            End If
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
                .Cell(flexcpFontSize, i, j) = 10
            Next j
        Next i
        For i = 0 To .Cols - 1
            .ColWidth(i) = 700
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
    vsfList.Rows = vsfList.Rows - 1
    vsfList.RowHidden(0) = True
    vsfList.Redraw = flexRDDirect
End Sub

Private Sub LoadTimePlan()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim intCurrentTime As Integer, intCol As Integer
    Dim i As Integer, j As Integer
    Dim blnFind As Boolean, rsUsed As ADODB.Recordset
    Dim rsUnit As ADODB.Recordset, lng合作单位人数 As Long, lng已挂人数 As Long
    Dim datTime As Date
    Dim datNow As Date
    vsfList.Redraw = flexRDNone
    vsfList.Clear
    vsfList.Rows = 1
    vsfList.Cols = 2
    vsfList.FixedRows = 0
    vsfList.FixedCols = 1
    intCol = 0
    intCurrentTime = -1
    datTime = dtpMain.Selection.Blocks(0).DateBegin
    datNow = zlDatabase.Currentdate
    If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("序号控制"))) = 1 Then
        strSQL = "Select 序号, To_Char(开始时间,'hh24:mi:ss') As 开始时间, 开始时间 As 序号时间, To_Char(终止时间,'hh24:mi:ss') As 终止时间, 是否预约, Decode(是否停诊,1,6,挂号状态) As 挂号状态" & _
                " From 临床出诊序号控制 Where 记录id = [1] And Nvl(是否预约,0)=1 And 开始时间 <> 终止时间 Order By 序号,开始时间"
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
                If Val(Nvl(rsTemp!是否预约)) = 0 Then
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                End If
                If CDate(Format(Nvl(rsTemp!序号时间), "yyyy-mm-dd hh:mm:ss")) < DateAdd("n", -1 * mlng预约有效时间, datNow) Then
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                End If
                If Format(Nvl(rsTemp!序号时间), "yyyy-mm-dd") <> Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                End If
            End With
            rsTemp.MoveNext
        Loop
    Else
        strSQL = "Select 序号, To_Char(开始时间,'hh24:mi:ss') As 开始时间, 开始时间 As 序号时间, To_Char(终止时间,'hh24:mi:ss') As 终止时间, 数量, 是否预约 From 临床出诊序号控制 Where 记录id = [1] And 预约顺序号 Is Null And Nvl(是否预约,0)=1 Order By 序号,开始时间"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        strSQL = "Select Sum(Nvl(数量,0)) As 合作单位数量,序号  From 临床出诊挂号控制记录 Where 记录id=[1] And 类型=1 And 控制方式=3 Group By 序号"
        Set rsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        strSQL = "Select Count(1) As 已挂数量,序号 From 临床出诊序号控制 Where 记录ID=[1] And 预约顺序号 Is Null And Nvl(挂号状态,0) <> 0 Group By 序号"
        Set rsUsed = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        Do While Not rsTemp.EOF
            If Val(Nvl(rsTemp!数量)) <> 0 Then
                rsUnit.Filter = "序号=" & Val(rsTemp!序号)
                If rsUnit.EOF Then
                    lng合作单位人数 = 0
                Else
                    lng合作单位人数 = Val(Nvl(rsUnit!合作单位数量))
                End If
                rsUsed.Filter = "序号=" & Val(rsTemp!序号)
                If rsUsed.EOF Then
                    lng已挂人数 = 0
                Else
                    lng已挂人数 = Val(Nvl(rsUsed!已挂数量))
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
                    If Val(Nvl(rsTemp!数量)) - lng合作单位人数 - lng已挂人数 = 0 Then
                        .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                    End If
                    If Val(Nvl(rsTemp!是否预约)) = 0 Then
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                    End If
                    If CDate(Format(Nvl(rsTemp!序号时间), "yyyy-mm-dd hh:mm:ss")) < DateAdd("n", -1 * mlng预约有效时间, datNow) Then
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                    End If
                    If Format(Nvl(rsTemp!序号时间), "yyyy-mm-dd") <> Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
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

Private Sub vsfPlan_GotFocus()
    With vsfPlan
        .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = 16772055
    End With
End Sub

