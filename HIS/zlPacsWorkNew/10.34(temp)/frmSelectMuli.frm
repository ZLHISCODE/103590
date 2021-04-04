VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmSelectMuli 
   Appearance      =   0  'Flat
   Caption         =   "选择序列"
   ClientHeight    =   6795
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   14985
   Icon            =   "frmSelectMuli.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   14985
   StartUpPosition =   1  '所有者中心
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picList 
      Height          =   5175
      Left            =   240
      ScaleHeight     =   5115
      ScaleWidth      =   7635
      TabIndex        =   3
      Top             =   240
      Width           =   7695
      Begin VB.PictureBox picCommand 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   6135
         TabIndex        =   10
         Top             =   4200
         Width           =   6135
         Begin VB.CommandButton cmdDel 
            Caption         =   "删 除(&D)"
            Height          =   400
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "取 消(&C)"
            Height          =   400
            Left            =   3600
            TabIndex        =   12
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "确 定(&S)"
            Height          =   400
            Left            =   2400
            TabIndex        =   11
            Top             =   120
            Width           =   1095
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfTree 
         Height          =   2055
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   3855
         _cx             =   6800
         _cy             =   3625
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         SelectionMode   =   1
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
      Begin VB.Frame frmFilter 
         Caption         =   "过滤条件"
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   7455
         Begin VB.Frame frmTime 
            Height          =   615
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   7215
            Begin VB.OptionButton optDays 
               Caption         =   "2天"
               Height          =   180
               Index           =   1
               Left            =   840
               TabIndex        =   24
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optDays 
               Caption         =   "半月"
               Height          =   180
               Index           =   5
               Left            =   3240
               TabIndex        =   23
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optDays 
               Caption         =   "7天"
               Height          =   180
               Index           =   4
               Left            =   2640
               TabIndex        =   22
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optDays 
               Caption         =   "5天"
               Height          =   180
               Index           =   3
               Left            =   2040
               TabIndex        =   21
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optDays 
               Caption         =   "3天"
               Height          =   180
               Index           =   2
               Left            =   1440
               TabIndex        =   20
               Top             =   240
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton optDays 
               Caption         =   "1天"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   19
               Top             =   240
               Width           =   615
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
               Height          =   300
               Left            =   5760
               TabIndex        =   25
               Top             =   195
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   234881027
               CurrentDate     =   40833
            End
            Begin MSComCtl2.DTPicker dtpStart 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
               Height          =   300
               Left            =   4080
               TabIndex        =   26
               Top             =   195
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   234881027
               CurrentDate     =   40833
            End
            Begin VB.Label Label3 
               Caption         =   "到"
               Height          =   255
               Left            =   5520
               TabIndex        =   27
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Left            =   5880
            TabIndex        =   15
            Top             =   930
            Width           =   1455
         End
         Begin VB.TextBox txtStudyNo 
            Height          =   300
            Left            =   3360
            TabIndex        =   14
            Top             =   930
            Width           =   1455
         End
         Begin VB.ComboBox cboModality 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   930
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "姓  名："
            Height          =   255
            Left            =   5040
            TabIndex        =   16
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "检 查 号："
            Height          =   255
            Left            =   2520
            TabIndex        =   13
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "影像类别："
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picImage 
      Height          =   3255
      Left            =   6840
      ScaleHeight     =   3195
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   2640
      Width           =   7215
      Begin VB.CheckBox chkViewImage 
         Caption         =   "预览图像"
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   2760
         Width           =   1095
      End
      Begin zl9PacsControl.ucSplitPage ucPage 
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   6210
         _extentx        =   10504
         _extenty        =   582
         pagecount       =   0
         pagerecord      =   9
      End
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   2295
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "双击可以显示大图"
         Top             =   240
         Width           =   5895
         _Version        =   262147
         _ExtentX        =   10398
         _ExtentY        =   4048
         _StockProps     =   35
         BackColor       =   0
      End
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   6510
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13714
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image imgImage 
      Height          =   240
      Left            =   1920
      Picture         =   "frmSelectMuli.frx":0CCA
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSeries 
      Height          =   240
      Left            =   1320
      Picture         =   "frmSelectMuli.frx":10B4
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgStudy 
      Height          =   240
      Left            =   720
      Picture         =   "frmSelectMuli.frx":144C
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   360
      Top             =   4080
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSelectMuli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const M_STR_STUDY_UID As String = "1.2.840.023.500903000005970.67031919.0"
Private Const M_STR_SERIES_UID As String = "1.2.840.023.500903000005970.67031919.1"

Public mblnOK As Boolean

Private mintSelectIndex As Integer

Private mstrTitle As String
Private mlngReleationType As Long       '1--取消关联；2--关联图像
Private mstrModality As String          '影像类别

Private mMultiRows As Integer
Private mMultiCols As Integer

Private mlngModule  As Long         '当前站点模块
Private mlngCurDeptId As Long     '当前科室ID
Private mlngAdviceID As Long      '当前医嘱ID
Private mblnMoved As Boolean        '检查是否转存
Private mblnSaveReportImage As Boolean  '是否保存报告图

Private mlngCurPageIndex As Long    '保存当前页索引
Private mlngPageCount As Long       '每页显示的图像数量

Private mdcmUID As New DicomGlobal


Private mrsStudyData As ADODB.Recordset
Private mrsSeriesData As ADODB.Recordset
Private mrsImageData As ADODB.Recordset

Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long




Public Function ShowImageReleation(ByVal lngModule As Long, ByVal lngAdviceID As Long, ByVal strPrivs As String, _
    ByVal blnMoveId As Boolean, ByVal blnSaveReportImage As Boolean, ByVal lngCurDeptId As Long, _
    Optional lngReleationType As Long = 1, Optional strModality = "")
    
    Dim curDate As Date
    
    mblnOK = False
    
    curDate = zlDatabase.Currentdate
    
    mlngModule = lngModule
    mlngAdviceID = lngAdviceID
    mblnMoved = blnMoveId
    mblnSaveReportImage = blnSaveReportImage
    mlngCurDeptId = lngCurDeptId
    
    mlngReleationType = lngReleationType
    mstrModality = strModality
    
    dtpStart.value = curDate - 2
    dtpEnd.value = curDate
    
    Me.Caption = IIf(mlngReleationType = 1, "取消关联", "关联图像")
    cmdDel.Visible = IIf(mlngReleationType = 1, False, True)
    cmdDel.Enabled = CheckPopedom(strPrivs, "删除临时影像")
    
    '如果是病理模块，则不需要显示其他类别的图像
    If glngModul = G_LNG_PATHOLSYS_NUM Then
        cboModality.Clear
        cboModality.AddItem "DG-病理"
        
        cboModality.ListIndex = 0
        
        Label1.Visible = False
        cboModality.Visible = False
        
        frmTime.Left = 120
        frmTime.Width = frmFilter.Width - 240
    Else
        '填充影像类别
        Call FillModality
    End If


    Call InitReleationList

    
    '刷新列表
    If mlngReleationType = 2 Then
        Call QueryReleationData(dtpStart.value, dtpEnd.value)
        Call FilterReleationData
    Else
        Call QueryCancelReleationData
    End If
    
    Call LoadReleationDataToFace
    
    On Error GoTo 0
    
    Call InitFaceScheme

    Me.Show 1
End Function



Private Sub InitReleationList()
'初始化关联列表
    With vsfTree
        
        ' structure
        .Cols = 4
        .Rows = 0
        .FixedCols = 0
        .FixedRows = 0
        .Left = 50
        
        ' appearance
        .GridLines = flexGridNone
        .BackColorBkg = .BackColor
        .SheetBorder = .BackColor
        .ExtendLastCol = True
        .Redraw = flexRDBuffered ' << new setting
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarCompleteLeaf
        .Ellipsis = flexEllipsisEnd
        
        ' behavior
        .AllowSelection = False
        .HighLight = flexHighlightAlways
        .ScrollTrack = True
        .AutoSearch = flexSearchFromCursor
        
        .ColDataType(0) = flexDTBoolean
        .ColWidth(0) = 800
        
        .ColHidden(1) = True
        
        .ColWidth(2) = 1600
        
        
    End With
End Sub


Private Sub QueryCancelReleationData()
'查询需要取消关联的数据
    Dim strSql As String
    
    '查询临时记录
    strSql = "select 影像类别,to_char(检查号) as 检查号,姓名,英文名,性别,年龄,检查uid,位置一,位置一,位置一,检查设备,接收日期 from 影像检查记录 where 医嘱ID=[1]"
    If mblnMoved Then strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    
    Set mrsStudyData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    
    '查询影像检查序列（使用规则时，查询才会使用索引）
    strSql = "select /*+ Rule*/ a.序列UID,a.检查UID,a.序列号,a.序列描述,a.采集时间 from 影像检查序列 a, 影像检查记录 b where a.检查UID=b.检查UID and b.医嘱ID=[1] order by a.序列号"
    If mblnMoved Then
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
        strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
    End If
    
    Set mrsSeriesData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    
    '查询影像检查图像
    strSql = "select /*+ Rule*/ a.图像UID, a.序列UID, b.检查UID, a.图像号, a.图像描述, a.采集时间 from 影像检查图象 a, 影像检查序列 b, 影像检查记录 c where a.序列UID=b.序列UID and b.检查UID=c.检查UID  and c.医嘱ID=[1] order by a.序列UID, a.图像号"
    If mblnMoved Then
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
        strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
        strSql = Replace(strSql, "影像检查图象", "H影像检查图象")
    End If
    
    Set mrsImageData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
End Sub


'Private Sub QueryReleationData(ByVal dtStartDate As Date, ByVal dtEndDate As Date)
''查询关联数据
'    Dim strSql As String
'
'    '查询临时记录
'    strSql = "select /*+ Rule*/ 影像类别,to_char(检查号) as 检查号,姓名,英文名,性别,年龄,检查uid,位置一,位置一,位置一,检查设备,接收日期 " & _
'            " from 影像临时记录 where 接收日期 between [1] and [2]"
'    Set mrsStudyData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate)
'
'
'    '查询临时序列
'    strSql = "select /*+ Rule*/ a.序列UID,a.检查UID,a.序列号,a.序列描述,a.采集时间 from 影像临时序列 a, 影像临时记录 b where a.检查UID=b.检查uid and b.接收日期 between [1] and [2] order by 序列号"
'    Set mrsSeriesData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate)
'
'
'    '查询临时图像
'    strSql = "select /*+ Rule*/ a.图像UID, a.序列UID, b.检查UID, a.图像号, a.图像描述, a.采集时间 from 影像临时图象 a, 影像临时序列 b, 影像临时记录 c  where a.序列UID=b.序列UID and b.检查UID=c.检查UID and b.接收日期 between [1] and [2]   order by a.序列UID, a.图像号"
'    Set mrsImageData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
'
'End Sub


Private Sub QueryReleationData(ByVal dtStartDate As Date, ByVal dtEndDate As Date)
'查询关联数据
    Dim strSql As String

    '查询临时记录
    strSql = "select /*+ Rule*/ 影像类别,to_char(检查号) as 检查号,姓名,英文名,性别,年龄,检查uid,位置一,位置一,位置一,检查设备,接收日期 " & _
            " from 影像临时记录 where 接收日期 between [1] and [2]"
    Set mrsStudyData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(Format(dtStartDate, "yyyy-mm-dd 00:00")), CDate(Format(dtEndDate, "yyyy-mm-dd 23:59")))

End Sub



Private Sub FilterReleationData()
'过滤关联数据
    Dim strFilter As String
    

    strFilter = ""
    
    If cboModality.ListIndex >= 0 Then
        strFilter = "影像类别 = '" & Split(cboModality.Text, "-")(0) & "'"
    End If
    
    If Trim(Replace(txtStudyNo.Text, "'", "")) <> "" Then
        strFilter = strFilter & " and 检查号='" & Replace(txtStudyNo.Text, "'", "") & "'"
    End If
    
    If Trim(Replace(txtName.Text, "'", "")) <> "" Then
        strFilter = strFilter & " and 姓名 like '" & Replace(txtName.Text, "'", "") & "%'"
    End If
    
    
'    '增加时间过滤条件
'    If Not optDays(5).value And Replace(txtName.Text, "'", "") = "" And Replace(txtStudyNo.Text, "'", "") = "" Then
'        strFilter = strFilter & " And 接收日期 > '" & Format(dtpStart.value, "yyyy-MM-dd 00:00") & "'" & " And 接收日期 < '" & Format(dtpEnd.value, "yyyy-MM-dd 23:59") & "'"
'    End If

    
    mrsStudyData.Filter = strFilter
    
    stb.Panels(1).Text = "共搜索到 " & mrsStudyData.RecordCount & " 条检查结果。"
End Sub



Private Sub LoadReleationDataToFace()
'载入关联数据到界面
    Dim i As Long
    
    vsfTree.Rows = 0
    Call vsfTree.Clear
    Call DViewer.Images.Clear
    
    If mrsStudyData.RecordCount <= 0 Then Exit Sub
    
    With vsfTree
    
        .Redraw = flexRDNone
        .Rows = 0
        
        '读取检查节点
        While Not mrsStudyData.EOF
            .AddItem ""
            
            .RowData(.Rows - 1) = 0
            
            .Cell(flexcpChecked, .Rows - 1, 0) = IIf(mlngReleationType = 1, True, False) '如果是取消关联，则自动选择检查数据
            .Cell(flexcpPicture, .Rows - 1, 2) = imgStudy
            
            .Cell(flexcpText, .Rows - 1, 1) = Nvl(mrsStudyData!检查uid)
            .Cell(flexcpText, .Rows - 1, 2) = Nvl(mrsStudyData!姓名) & "(" & Nvl(mrsStudyData!检查号) & ")"
'            .Cell(flexcpFontBold, .Rows - 1, 2) = True
            
            .Cell(flexcpText, .Rows - 1, 3) = "姓名:" & Nvl(mrsStudyData!姓名) & "  检查号:" & Nvl(mrsStudyData!检查号) & "  性别:" & Nvl(mrsStudyData!性别) & "  年龄:" & Nvl(mrsStudyData!年龄) & "  检查日期:" & Nvl(mrsStudyData!接收日期)
            .Cell(flexcpFontSize, .Rows - 1, 3) = 9
            .Cell(flexcpForeColor, .Rows - 1, 3) = vbGrayText
            
            .IsSubtotal(.Rows - 1) = True
            .RowOutlineLevel(.Rows - 1) = 1
            
            If mlngReleationType <> 2 Then
                .RowData(.Rows - 1) = 1
                
                '读取序列节点
                mrsSeriesData.Filter = "检查UID='" & Nvl(mrsStudyData!检查uid) & "'"
                If mrsSeriesData.RecordCount > 0 Then
                    While Not mrsSeriesData.EOF
                        .AddItem ""
    
                        .RowData(.Rows - 1) = 1
    
                        .Cell(flexcpChecked, .Rows - 1, 0) = IIf(mlngReleationType = 1, True, False) '如果是取消关联，则自动选择序列数据
                        .Cell(flexcpPicture, .Rows - 1, 2) = imgSeries
    
                        .Cell(flexcpText, .Rows - 1, 1) = Nvl(mrsSeriesData!序列UID)
                        .Cell(flexcpText, .Rows - 1, 2) = "序列" & Nvl(mrsSeriesData!序列号)
    '                    .Cell(flexcpFontBold, .Rows - 1, 2) = True
    
                        .Cell(flexcpText, .Rows - 1, 3) = "序列号:" & Nvl(mrsSeriesData!序列号) & "  序列描述:" & Nvl(mrsSeriesData!序列描述) & "  生成日期:" & Nvl(mrsSeriesData!采集时间)
                        .Cell(flexcpFontSize, .Rows - 1, 3) = 9
                        .Cell(flexcpForeColor, .Rows - 1, 3) = vbGrayText
    
                        .IsSubtotal(.Rows - 1) = True
                        .RowOutlineLevel(.Rows - 1) = 2

    
                        '读取图像节点
                        mrsImageData.Filter = "序列UID='" & Nvl(mrsSeriesData!序列UID) & "'"
                        If mrsImageData.RecordCount > 0 Then
                            While Not mrsImageData.EOF
                                .AddItem ""
    
                                .RowData(.Rows - 1) = 1
    
                                .Cell(flexcpChecked, .Rows - 1, 0) = IIf(mlngReleationType = 1, True, False) '如果是取消关联，则自动选择图像数据
                                .Cell(flexcpPicture, .Rows - 1, 2) = imgImage
    
                                .Cell(flexcpText, .Rows - 1, 1) = Nvl(mrsImageData!图像UID)
                                .Cell(flexcpText, .Rows - 1, 2) = "图像" & Nvl(mrsImageData!图像号)
                                .Cell(flexcpText, .Rows - 1, 3) = "图像号:" & Nvl(mrsImageData!图像号) & "  采集时间:" & Nvl(mrsImageData!采集时间)
    
                                .Cell(flexcpFontSize, .Rows - 1, 3) = 9
                                .Cell(flexcpForeColor, .Rows - 1, 3) = &HC0C0FF
    
                                .IsSubtotal(.Rows - 1) = True
                                .RowOutlineLevel(.Rows - 1) = 3
    
                                Call mrsImageData.MoveNext
                            Wend
                        End If
    
                        mrsSeriesData.MoveNext
                    Wend
                    
                End If
            End If
                    
            Call mrsStudyData.MoveNext
            
        Wend

        .Outline 1
        
        If .Rows > 0 Then
            .Row = 0
            .RowSel = 0
        End If

        .Redraw = flexRDBuffered
        
        For i = 0 To vsfTree.Rows - 1
            '折叠节点
            .IsCollapsed(i) = flexOutlineCollapsed
        Next i
    End With
End Sub



Private Sub InitPageControl(ByVal lngSearchType As Long, ByVal strSearchId As String)
'初始化分页控件
    Dim strFilter As String
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    If mlngReleationType = 2 Then
        Select Case lngSearchType
            Case 1
                strSql = "select count(1)  as 返回值 from 影像临时图象 a, 影像临时序列 b where a.序列UID=b.序列UID and b.检查UID=[1]"
            Case 2
                strSql = "select count(1)  as 返回值 from 影像临时图象  where  序列UID=[1]"
            Case 3
                strSql = "select count(1)  as 返回值 from 影像临时图象  where  图像UID=[1]"
        End Select
        
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strSearchId)
        If rsData.RecordCount > 0 Then
            lngRecordCount = Nvl(rsData!返回值)
        Else
            lngRecordCount = 0
        End If
    Else
        Select Case lngSearchType
            Case 1
                strFilter = "检查UID='" & strSearchId & "'"
            Case 2
                strFilter = "序列UID='" & strSearchId & "'"
            Case 3
                strFilter = "图像UID='" & strSearchId & "'"
        End Select
    
        mrsImageData.Filter = strFilter
        lngRecordCount = mrsImageData.RecordCount
    End If
    
    ucPage.RecordCount = lngRecordCount
End Sub



Private Function GetImageViewData(ByVal lngSearchType As Long, _
    ByVal strSearchId As String, ByVal lngCurPage As Long, ByVal lngPageRecord As Long) As ADODB.Recordset
'获取预览图像数据
'intSearchType:0-按检查uid搜索,1-按序列UID搜索,2-按图像UID搜索

    Dim strSql As String
    Dim lngStartRecord As Long
    Dim lngEndRecord As Long
    
    If mlngReleationType = 2 Then
        '关联图像
        strSql = "Select rownum as 顺序号,  A.图像号,d.FTP用户名 As User1, d.FTP密码 As Pwd1, d.Ip地址 As Host1," & _
                " '/' || d.Ftp目录 || '/' As Root1, " & _
                " Decode(C.接收日期, Null, '', To_Char(C.接收日期, 'YYYYMMDD') || '/') || C.检查uid || '/' || A.图像uid As URL, " & _
                " d.设备号 As 设备号1,A.图像UID,C.检查UID,B.序列UID,d.FTP用户名 As User2, d.FTP密码 As Pwd2," & _
                " d.Ip地址 As Host2, '/' || d.Ftp目录 || '/' As Root2, " & _
                " d.设备号 As 设备号2,A.动态图,A.编码名称, A.采集时间, A.录制长度 " & _
                " From 影像临时图象 A, 影像临时序列 B, 影像临时记录 C ,影像设备目录 D " & _
                " Where A.序列UID = B.序列UID And B.检查UID = C.检查UID And  C.位置一 = D.设备号 "
    Else
        '取消关联
        
        strSql = "Select rownum as 顺序号, A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1," & _
            "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1," & _
            "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
            "||C.检查UID||'/'||A.图像UID As URL,d.设备号 as 设备号1, " & _
            "E.FTP用户名 As User2,E.FTP密码 As Pwd2," & _
            "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2," & _
            "e.设备号 as 设备号2, A.图像UID,C.检查UID,B.序列UID,A.动态图,A.编码名称,A.采集时间, A.录制长度 " & _
            "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
            "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) "
            
        If mblnMoved Then
            strSql = Replace(strSql, "影像检查图象", "H影像检查图象")
            strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
            strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
        End If
    End If

    Select Case lngSearchType
        Case 1
            strSql = strSql & " and C.检查UID=[1]"
        Case 2
            strSql = strSql & " and B.序列UID=[1]"
        Case 3
            strSql = strSql & " and A.图像UID=[1]"
    End Select
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    strSql = "select /*+RULE*/ * from (" & strSql & " order by b.序列UID, a.图像号) where 顺序号>=" & lngStartRecord & " and 顺序号<=" & lngEndRecord
    
    Set GetImageViewData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strSearchId)
End Function


Private Sub LoadViewImageToFace(rsCurImageData As ADODB.Recordset)
'加载预览图像到界面
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    
    Dim iCols As Integer, iRows As Integer
    
    
    
    DViewer.Images.Clear
    
    If rsCurImageData.RecordCount > 0 Then
        '计算图像显示数量
        ResizeRegion rsCurImageData.RecordCount, DViewer.Width, DViewer.Height, iRows, iCols
        
        mMultiCols = iCols
        mMultiRows = iRows

        DViewer.MultiColumns = iCols
        DViewer.MultiRows = iRows
        
        '创建本地目录
        strCachePath = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
        MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsCurImageData("URL")))
        
        Do While Not rsCurImageData.EOF
            '循环加载图像到DicomViewer中
            strTmpFile = strCachePath & Nvl(rsCurImageData("URL"))
            
            If Nvl(rsCurImageData("动态图"), IMGTAG) = VIDEOTAG Then
                strTmpFile = IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\Avi.bmp", App.Path & "..\附加文件\Avi.bmp")
            ElseIf Nvl(rsCurImageData("动态图"), IMGTAG) = AUDIOTAG Then
                strTmpFile = IIf(Len(App.Path) > 3, App.Path & "\..\附加文件\wav.bmp", App.Path & "..\附加文件\wav.bmp")
            End If
            
            If Dir(strTmpFile) = vbNullString Then
                '本地缓存图像不存在，则读取FTP图像
                
                '建立FTP连接
                If Nvl(rsCurImageData("设备号1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(Nvl(rsCurImageData("Host1")), Nvl(rsCurImageData("User1")), Nvl(rsCurImageData("Pwd1"))) = 0 Then
                        If Nvl(rsCurImageData("设备号2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))) = 0 Then
                                MsgBoxD Me, "FTP不能正常连接，请检查网络设置。"
                                Exit Sub
                            End If
                        Else
                            MsgBoxD Me, "FTP不能正常连接，请检查网络设置。"
                            Exit Sub
                        End If
                    End If
                End If
                
                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL"))) <> 0 Then
                    '从设备号1提取图像失败，则从设备号2提取图像
                    If Nvl(rsCurImageData("设备号2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")))
                    End If
                End If
            End If
  
            If Dir(strTmpFile) <> vbNullString Then
               If Nvl(rsCurImageData("动态图"), IMGTAG) <> VIDEOTAG And Nvl(rsCurImageData("动态图"), IMGTAG) <> AUDIOTAG Then
                    Set curImage = DViewer.Images.ReadFile(strTmpFile)
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                Else
                    Set curImage = New DicomImage
                    
                    On Error GoTo continue
                        Call curImage.FileImport(strTmpFile, "DIB/BMP")
continue:
                    
                    Call AddVideoLabelToDicomImage(curImage, _
                        "采集时间：" & Nvl(rsCurImageData("采集时间")), _
                        "录制长度：" & Nvl(rsCurImageData("录制长度"), "0") & " 秒", _
                        "编码名称：" & Nvl(rsCurImageData("编码名称")))
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                    
                    Call DViewer.Images.Add(curImage)
                End If
                
                
                '取消自动剪影,因为DicomObjects控件本身对处理剪影有BUG，存在（0028，6100）时，会自动对图像进行剪影，
                '导致晋煤的DSA图像不能正常显示
                '虽然设置图像的mask=0 ,可以取消剪影，但是每次图像被添加到新的Dicomimages之后，自动又将mask设置成1了，
                '这样在程序中无法很好的控制，因此直接去掉（0028，6100）这个属性。
                If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
                    curImage.Attributes.Remove &H28, &H6100
                End If
            End If
            
            rsCurImageData.MoveNext
        Loop
        
        If DViewer.Images.Count > 0 Then
            DViewer.CurrentIndex = 1
            DViewer.Images(1).BorderColour = vbRed
        End If
        
        Inet1.FuncFtpDisConnect
        Inet2.FuncFtpDisConnect
    Else
        DViewer.MultiColumns = 1
        DViewer.MultiRows = 1
    End If
End Sub


Private Sub cboModality_Click()
    If Not cboModality.Visible Then Exit Sub
    
    If mlngReleationType = 2 Then '关联图像
        If cboModality.ListIndex < 0 Then Exit Sub
        
        Call FilterReleationData
        Call LoadReleationDataToFace
    End If
End Sub

Private Sub chkViewImage_Click()
On Error GoTo ErrHandle
    If Not vsfTree.Visible Then Exit Sub
    
    If chkViewImage.value <> 0 Then
        Call vsfTree_SelChange
    Else
        Call DViewer.Images.Clear
    End If
    
    '保存参数
    Call SetDeptPara(mlngCurDeptId, "预览关联图像", chkViewImage.value)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub


Private Function GetReleationImageIds() As ADODB.Recordset
'查询关联或者要取消关联的图像ID
    Dim i As Long, j As Long
    Dim strSql As String
    Dim strValues(0 To 80) As String
    Dim strValue As String
    Dim strUninTable As String
    Dim strFilter As String
    

    j = 0
    strUninTable = ""
    strFilter = ""
    strValue = ""
    
    
    '构造查询语句
    For i = 0 To vsfTree.Rows - 1
        If vsfTree.RowOutlineLevel(i) = 3 And vsfTree.TextMatrix(i, 0) = -1 Then '为3表示图像节点
            If j > 79 Then
                strFilter = strFilter & " Or 图像UID ='" & vsfTree.TextMatrix(i, 1) & "'"
            Else
                If zlCommFun.ActualLen(strValue) > 3600 Then
                     strValues(j) = Mid(strValue, 2)
                     strUninTable = strUninTable & " Union ALL  Select  Column_Value as 图像UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
                     
                     strValue = ""
                     j = j + 1
                End If
                
                strValue = strValue & "," & vsfTree.TextMatrix(i, 1)
            End If
        End If
    Next i
    
    If strValue <> "" Then
        strValues(j) = Mid(strValue, 2)
        strUninTable = strUninTable & " Union ALL  Select  Column_Value as 图像UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
    End If
    
    '如果没有需要查找的图像UID，则返回空数据集
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
        Set GetReleationImageIds = Nothing
        Exit Function
    End If
    
'    If strFilter <> "" Then strFilter = " and ( " & Mid(strFilter, 4) & ")"
    If strFilter <> "" Then strFilter = strUninTable & " Union All Select 图像UID from [影像图象] where  ( " & Mid(strFilter, 4) & ")"
    
    '根据移动的方向不同，源图有可能在“影像临时记录”或者“影像检查记录”中
    '关联时从临时记录搬移到正常记录，取消关联时从正常记录搬移到临时记录
    strSql = "Select /*+RULE*/ D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd, Decode(C.位置一,Null,C.位置二,C.位置一) as 设备号," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL,A.图像UID, c.姓名,c.性别,c.年龄,c.检查号 " & _
        "From 影像检查图象 A, 影像检查序列 B, 影像检查记录 C,影像设备目录 D,(" & Replace(strUninTable, "[影像图象]", "影像检查图象") & ") E " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And A.序列UID=B.序列UID and B.检查UID=C.检查UID and A.图像UID = E.图像UID " & _
        "Union All " & _
        "Select /*+RULE*/ D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd, Decode(C.位置一,Null,C.位置二,C.位置一) as 设备号," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL,A.图像UID, c.姓名,c.性别,c.年龄,c.检查号 " & _
        "From 影像临时图象 A,影像临时序列 B, 影像临时记录 C,影像设备目录 D,(" & Replace(strUninTable, "[影像图象]", "影像临时图象") & ") E " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And A.序列UID=B.序列UID and B.检查UID=C.检查UID and A.图像UID= E.图像UID"
        
    If mblnMoved Then
        strSql = Replace(strSql, "影像检查图象", "H影像检查图象")
        strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
    
    Set GetReleationImageIds = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strValues(0), strValues(1), strValues(2), strValues(3), _
        strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10), _
        strValues(11), strValues(12), strValues(13), strValues(14), strValues(15), strValues(16), strValues(17), _
        strValues(18), strValues(19), strValues(20), strValues(21), strValues(22), strValues(23), strValues(24), strValues(25), strValues(26), _
        strValues(27), strValues(28), strValues(29), strValues(30), strValues(31), strValues(32), strValues(33), strValues(34), strValues(35), strValues(36), _
        strValues(37), strValues(38), strValues(39), strValues(40), strValues(41), strValues(42), strValues(43), strValues(44), strValues(45), strValues(46), _
        strValues(47), strValues(48), strValues(49), strValues(50), strValues(51), strValues(52), strValues(53), strValues(54), strValues(55), strValues(56), _
        strValues(57), strValues(58), strValues(59), strValues(60), strValues(61), strValues(62), strValues(63), strValues(64), strValues(65), strValues(66), _
        strValues(67), strValues(68), strValues(69), strValues(70), strValues(71), strValues(72), strValues(73), strValues(74), strValues(75), strValues(76), _
        strValues(77), strValues(78), strValues(79), strValues(80))
End Function



Private Sub GetStorageDevice(ByVal lngAdviceID As Long, ByVal strNewStudyUID As String, _
    ByRef strDeviceNO As String, ByRef strFTPIP As String, _
    ByRef strFtpUrl As String, ByRef strVirtualPath As String, _
    ByRef strFTPUser As String, ByRef strFTPPwd As String)
'获取新的存储设备信息，如果设备存储信息部存在，则需要进行增加
'如果是取消关联，则使用strNewStudyUID将不能从数据库中查找到对应的数据
'strDeviceNum:设备号
'strFtpIp: ftp地址
'strFtpUrl: ftp目录
'strVirtualPath: ftp虚拟存储路径
'strFtpUser: ftp用户名
'strFtpPwd: ftp密码



    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim objDestFtp As New clsFtp
    Dim curDate As Date
    
    strFTPIP = ""
    strFtpUrl = ""
    strFTPUser = ""
    strFTPPwd = ""
    strDeviceNO = ""
    
    strSql = "Select D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd,C.位置一,C.位置二,C.位置三,C.接收日期," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL " & _
        "From 影像检查记录 C,影像设备目录 D " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And C.检查UID= [1] Union All " & _
        "Select D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd,C.位置一,C.位置二,C.位置三,C.接收日期," & _
        "D.IP地址 As Host," & _
        "'/'||D.Ftp目录||'/' As Root,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL " & _
        "From 影像临时记录 C,影像设备目录 D " & _
        "Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
        "And C.检查UID= [1]"
        
    If mblnMoved Then
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNewStudyUID)
    
    blnIsGetNewDevice = False
    
    If rsData.RecordCount <= 0 Then
        blnIsGetNewDevice = True
    Else
        '如果执行到这里，说明是执行图像关联,需要判断当前检查的存储设备是否有效，如果无效需生成新的存储设备
        If Trim(rsData!接收日期) = "" Then
            blnIsGetNewDevice = True
        Else
            strDeviceNO = Nvl(rsData!位置一)
            strFTPIP = Nvl(rsData!host)
            strFtpUrl = Nvl(rsData!Root)
            strFTPUser = Nvl(rsData!FtpUser)
            strFTPPwd = Nvl(rsData!FtpPwd)
            strVirtualPath = strFtpUrl & Nvl(rsData!Url)
        End If
    End If
    
    
    If blnIsGetNewDevice Then
        '生成新的检查UID和存储设备,如果执行到这里，说明是取消关联
        
        If mlngModule = 1290 Then
            '查询医技工作站中，检查所对应的存储设备
            strSql = "select d.参数值 " & _
                        " from 医技执行房间 a, 病人医嘱发送 b, 影像DICOM服务对 c, 影像DICOM服务参数 d " & _
                        " Where a.科室ID = b.执行部门id And a.执行间 = b.执行间 And a.检查设备 = c.设备号 " & _
                        " and c.服务功能='图像接收' and c.服务ID=d.服务ID and d.参数名称='存储设备' and b.医嘱id=[1]"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
            
            If rsTemp.RecordCount <= 0 Then
                MsgBoxD Me, "未找到图像存储设备,请确认当前检查所用设备是否在影像设备目录的服务配置中配置了图像存储。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strDeviceNO = Nvl(rsTemp!参数值)
        Else
            '查询非医技工作站中的图像存储设备
            strDeviceNO = GetDeptPara(mlngCurDeptId, "存储设备号")
            
            If Val(strDeviceNO) <= 0 Then
                MsgBoxD Me, "未找到图像存储设备,请确认在影像流程管理中是否对该科室配置了图像采集存储设备。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        
        strSql = "Select 设备号,设备名,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL,FTP用户名,FTP密码,IP地址 " & _
                    " From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Tag, strDeviceNO)
        
        '如果存储设备停用，则直接退出
        If rsTemp.RecordCount <= 0 Then
            MsgBoxD Me, "未找到存储设备,请确认设备号为 [" & strDeviceNO & "] 的设备是否启用。", vbInformation, gstrSysName
            Exit Sub
        End If
        
'        Call funGetStorageDevice(Me, strDeviceNO, strFtpUrl, strFTPIP, strFTPUser, strFTPPwd)
        strFtpUrl = Nvl(rsTemp("URL"))
        strFTPIP = Nvl(rsTemp("IP地址"))
        strFTPUser = Nvl(rsTemp("FTP用户名"))
        strFTPPwd = Nvl(rsTemp("FTP密码"))
        
        strFtpUrl = IIf(strFtpUrl = "/", "//", strFtpUrl)
        
        objDestFtp.FuncFtpConnect strFTPIP, strFTPUser, strFTPPwd
        On Error GoTo ErrHandle
            curDate = zlDatabase.Currentdate
            
            strVirtualPath = strFtpUrl & Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            '创建FTP目录
            objDestFtp.FuncFtpMkDir strFtpUrl, Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            
        Call objDestFtp.FuncFtpDisConnect
ErrHandle:
        Call objDestFtp.FuncFtpDisConnect
    End If
End Sub


Private Function DelTempImages(rsImageDatas As ADODB.Recordset) As Boolean
'删除ftp服务器中的文件
    Dim objSrcFtp As New clsFtp
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strImageUID As String
    Dim strVirtualPath As String
'    Dim lngResult As Long
    
    DelTempImages = False
    If rsImageDatas.RecordCount <= 0 Then Exit Function
    
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
    
        strImageUID = Nvl(rsImageDatas!图像UID)
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
            
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
        
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If
    

        '删除图像文件，当删除失败后，则退出执行
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        '删除可能存在的报告图像
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID & ".jpg")
        
'        If lngResult <> 0 Then
'            Call err.Raise(-1, "MoveImageToStudy", "Ftp服务器图像删除失败。 [图像UID:" & strImageUID & "]", err.HelpFile, err.HelpContext)
'            Exit Function
'        End If
    
        '图像删除成功后，同步删除数据库中的数据
        Call zlDatabase.ExecuteProcedure("ZL_影像检查_删除临时图像(3,'" & strImageUID & "')", Me.Caption)
        
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend


    objSrcFtp.FuncFtpDisConnect
    
    DelTempImages = True
    
    Exit Function
ErrHandle:
    objSrcFtp.FuncFtpDisConnect
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
    
End Function


Public Function MoveImageToStudy(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String, _
    ByVal strFTPIP As String, ByVal strFtpUrl As String, ByVal strFtpVirtualPath As String, _
    ByVal strFTPUser As String, ByVal strFTPPwd As String, ByRef objMoveList As Collection) As Boolean
'------------------------------------------------
'功能：将选定的检查图像移动到ftp上指定的检查中
'返回：True--成功；False－失败
'------------------------------------------------
    Dim objSrcFtp As New clsFtp
    Dim objDestFtp As New clsFtp
    Dim strVirtualPath As String
    Dim strDestPath As String
    Dim strTmpFile As String
    Dim aFiles() As String
    Dim i As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim lngResult As Long       '记录FTP操作的结果
    Dim strImageUID As String
    Dim dcmImages As New DicomImages
    Dim dcmImg As New DicomImage
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strFileList As String
    Dim blnIsMove As Boolean
    
On Error GoTo ErrHandle
    
    blnIsMove = False
    MoveImageToStudy = False
    If rsImageDatas.RecordCount <= 0 Then Exit Function

    '连接目标Ftp
    Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    strVirtualPath = ""
    strFileList = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
    
        strImageUID = Nvl(rsImageDatas!图像UID)
        
        If strVirtualPath <> Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url) Then
            strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
            strFileList = ""
        End If
    
        
        '当移动的文件不是相同的ftp地址时，则使用下载后再上传的方式转移文件
        If Nvl(rsImageDatas!host) <> strFTPIP Or Nvl(rsImageDatas!Root) <> strFtpUrl Then
        
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            
                strCurFtpIp = Nvl(rsImageDatas!host)
                strCurFtpUser = Nvl(rsImageDatas!FtpUser)
                strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
                
                Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
            End If
        
            strTmpFile = App.Path & "\TmpImage\" & strImageUID
            
            If strFileList = "" Then
                strFileList = objSrcFtp.FuncDirFiles(strVirtualPath)
            End If
            
            '如果源ftp设备中不存在该图像，则不进行移动
            If InStr(strFileList, strImageUID) > 0 Then
                lngResult = objSrcFtp.FuncDownloadFile(strVirtualPath, strTmpFile, strImageUID)
                If lngResult <> 0 Then
                    objSrcFtp.FuncFtpDisConnect
                    objDestFtp.FuncFtpDisConnect
        
                    Call err.Raise(-1, "MoveImageToStudy", "下载关联图像失败。 [图像UID:" & strImageUID & " 文件虚拟目录:" & strVirtualPath & " 本地路径:" & strTmpFile & "]", err.HelpFile, err.HelpContext)
                    Exit Function
                End If
        
                lngResult = objDestFtp.FuncUploadFile(strFtpVirtualPath, strTmpFile, strImageUID)
                If lngResult <> 0 Then
                    objSrcFtp.FuncFtpDisConnect
                    objDestFtp.FuncFtpDisConnect
        
                    Call err.Raise(-1, "MoveImageToStudy", "上传关联图像失败。 [图像UID:" & strImageUID & " 上传虚拟目录:" & strFtpVirtualPath & " 本地路径:" & strTmpFile & "]", err.HelpFile, err.HelpContext)
                    Exit Function
                End If
                
                blnIsMove = True
            End If
        Else
            If strFileList = "" Then
                strFileList = objDestFtp.FuncDirFiles(strVirtualPath)
            End If
            
            '如果源ftp设备中不存在该图像，则不进行移动
            If InStr(strFileList, strImageUID) > 0 Then
                lngResult = objDestFtp.FuncReNameFile(strVirtualPath & "/" & strImageUID, strFtpVirtualPath & "/" & strImageUID)
                If lngResult <> 0 Then
                    '如果文件移动失败，则端开连接重试一次
                    Call objDestFtp.FuncFtpDisConnect
                    Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
                    
                    lngResult = objDestFtp.FuncReNameFile(strVirtualPath & "/" & strImageUID, strFtpVirtualPath & "/" & strImageUID)
                    
                    If lngResult <> 0 Then
                        objSrcFtp.FuncFtpDisConnect
                        objDestFtp.FuncFtpDisConnect
                        
                        Call err.Raise(-1, "MoveImageToStudy", "在Ftp中移动文件时失败。 [图像UID:" & strImageUID & " 原虚拟目录:" & strVirtualPath & " 新虚拟目录:" & strFtpVirtualPath & "]", err.HelpFile, err.HelpContext)
                        Exit Function
                    End If
                End If
                
                blnIsMove = True
                
                '记录已经被移动过的文件，以便在处理数据失败的时候，还可对移动的图像进行恢复
                If Not objMoveList Is Nothing Then
                    Call objMoveList.Add(strVirtualPath & "/" & strImageUID & ">" & strFtpVirtualPath & "/" & strImageUID)
                End If
            End If
        End If
        

        If mblnSaveReportImage Then
            '上传ftp中的报告图
            
            If Nvl(rsImageDatas!host) <> strFTPIP Or Nvl(rsImageDatas!Root) <> strFtpUrl Then
                Call MoveReportImage(strTmpFile, strImageUID, objSrcFtp, objDestFtp, strVirtualPath, strFtpVirtualPath, objMoveList, 0)
            Else
                Call MoveReportImage(strTmpFile, strImageUID, objSrcFtp, objDestFtp, strVirtualPath, strFtpVirtualPath, objMoveList, 1)
            End If
        End If
        
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend


    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect
    
    '如果一个文件都没有被移动，则直接退出
    If Not blnIsMove Then Exit Function
    
    MoveImageToStudy = True
    
    Exit Function
ErrHandle:
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect

    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Private Sub MoveReportImage(ByVal strDicomFile As String, ByVal strImgUid As String, _
    objSrcFtp As clsFtp, objDestFtp As clsFtp, ByVal strSourceVirtualPath As String, ByVal strDestVirtualPath As String, _
    objMoveList As Collection, Optional ByVal lngWay As Long = 0)
On Error GoTo ErrHandle
'移动报告图
    Dim dcmImages As New DicomImages
    Dim dcmImg As New DicomImage
    Dim lngResult As Long
    
    If lngWay = 0 Then
        Call objSrcFtp.FuncDelFile(strSourceVirtualPath, strImgUid & ".jpg")
        
        '如果本地中存在从源ftp中下载的dicom图像，则将图像转换成jpg，并保存到目的ftp设备中
        If FileExists(strDicomFile) Then
            Call dcmImages.Clear
            Set dcmImg = dcmImages.ReadFile(strDicomFile)
    
            Call dcmImg.FileExport(strDicomFile & ".jpg", "JPG")
            Call objDestFtp.FuncUploadFile(strDestVirtualPath, strDicomFile & ".jpg", strImgUid & ".jpg")
            
            If FileExists(strDicomFile & ".jpg") Then Call Kill(strDicomFile & ".jpg")
        End If
    Else
        '如果源ftp设备中不存在该图像，则不进行移动
        If objDestFtp.FuncFtpFileExists(strSourceVirtualPath, strImgUid & ".jpg") Then
            lngResult = objDestFtp.FuncReNameFile(strSourceVirtualPath & "/" & strImgUid & ".jpg", strDestVirtualPath & "/" & strImgUid & ".jpg")
            
            If lngResult <> 0 Then
                '如果文件移动失败，则端开连接重试一次
                Call objDestFtp.FuncFtpDisConnect
'                Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
                Call objDestFtp.ResotreFtpConnect
                
                Call objDestFtp.FuncReNameFile(strSourceVirtualPath & "/" & strImgUid & ".jpg", strDestVirtualPath & "/" & strImgUid & ".jpg")
                
                '记录已经被移动过的文件，以便在处理数据失败的时候，还可对移动的图像进行恢复
                If Not objMoveList Is Nothing Then
                    Call objMoveList.Add(strSourceVirtualPath & "/" & strImgUid & ".jpg" & ">" & strDestVirtualPath & "/" & strImgUid & ".jpg")
                End If
            End If
        End If
    End If
Exit Sub
ErrHandle:
    Call OutputDebug("MoveReportImage", err)
End Sub


Private Sub ClearFtpImage(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String)
On Error GoTo ErrHandle
'转移图像成功后，在删除临时图像和原有FTP的图像和目录，清场操作出现错误可以不处理
    Dim objSrcFtp As New clsFtp
    Dim strTmpFile As String
    Dim strVirtualPath As String
    Dim strImageUID As String
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strNewDirectory
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    strNewDirectory = App.Path & "\TmpImage\" & Format(zlDatabase.Currentdate, "YYYYMMDD")
    
    If Not DirExists(strNewDirectory) Then MkDir strNewDirectory
    If Not DirExists(strNewDirectory & "\" & strNewStudyUID) Then MkDir strNewDirectory & "\" & strNewStudyUID
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strTmpFile = App.Path & "\TmpImage\" & Nvl(rsImageDatas!图像UID)
        
        strImageUID = Nvl(rsImageDatas!图像UID)
        
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
                
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If
        
'       为避免重新下载图像，如果本地存在图像文件，则不用进行删除
        
        If FileExists(strTmpFile) Then Call Kill(strTmpFile)
        If FileExists(strTmpFile & ".jpg") Then Call Kill(strTmpFile & ".jpg")

        '移动文件到新的位置
        Call MoveFile(App.Path & "\TmpImage\" & Nvl(rsImageDatas!Url) & "\" & Nvl(rsImageDatas!图像UID), _
            strNewDirectory & "\" & strNewStudyUID & "\" & Nvl(rsImageDatas!图像UID))
                

        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        '删除空的ftp目录
        Call objSrcFtp.FuncFtpDelDir(Replace(strVirtualPath, strImageUID, ""), strImageUID)
                
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend
    
    objSrcFtp.FuncFtpDisConnect
Exit Sub
ErrHandle:
    Call OutputDebug("ClearFtpImage", err)
End Sub


'撤销图像的移动
Private Sub CancelImageMove(ByVal strFTPIP As String, ByVal strFTPUser As String, ByVal strFTPPwd As String, objMoveList As Collection)
    Dim i As Long
    Dim objFtp As New clsFtp
    Dim strDestFile As String
    Dim strMoveFile As String
    
    If objMoveList Is Nothing Then Exit Sub
    If objMoveList.Count <= 0 Then Exit Sub
    
On Error GoTo ErrHandle

    Call objFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    For i = 1 To objMoveList.Count
        strDestFile = objMoveList.Item(i)
        
        strMoveFile = Mid(strDestFile, InStr(strDestFile, ">") + 1, 255)
        strDestFile = Mid(strDestFile, 1, InStr(strDestFile, ">") - 1)
        
        Call objFtp.FuncReNameFile(strMoveFile, strDestFile)
    Next i
        
ErrHandle:
    objFtp.FuncFtpDisConnect
End Sub


'取得关联提示信息
Private Function GetReleationHintInfo(lngAdviceID As Long, rsReleationImage As ADODB.Recordset) As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strResult As String
    Dim strStudyInf As String
    
    
    GetReleationHintInfo = ""
    
    If rsReleationImage.RecordCount <= 0 Then Exit Function
    
    Call rsReleationImage.MoveFirst
    While Not rsReleationImage.EOF
        strStudyInf = "[" & Nvl(rsReleationImage!姓名) & "(" & Nvl(rsReleationImage!检查号) & ") " & Nvl(rsReleationImage!性别) & " " & Nvl(rsReleationImage!年龄) & "]"
        
        If InStr(strResult, strStudyInf) <= 0 Then
            If strResult <> "" Then strResult = strResult & "+"
        
            strResult = strResult & strStudyInf
        End If
        Call rsReleationImage.MoveNext
    Wend
    
    strSql = "select 检查号,姓名,性别,年龄 from 影像检查记录 where 医嘱ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    
    GetReleationHintInfo = "是否确认将  " & strResult & "  的图像与  [" & Nvl(rsTemp!姓名) & "(" & Nvl(rsTemp!检查号) & ") " & Nvl(rsTemp!性别) & " " & Nvl(rsTemp!年龄) & "]  的检查进行关联操作？"
End Function


Private Function StartReleation(ByVal lngAdviceID As Long, rsImageDatas As ADODB.Recordset) As Boolean
'开始关联
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim strReportImageIds As String
    Dim strOldReportImages As String
    Dim lngReportImageLen As Long
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    
    blnBeginTrans = False
    StartReleation = False
    
    curDate = zlDatabase.Currentdate
    
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始查找关联检查数据... ]"
    
    strSql = "select 检查UID,接收日期 from 影像检查记录 where 医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsTmp.RecordCount <= 0 Then
        Call MsgBoxD(Me, "找不到待关联的检查信息。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始获取新的图像存储信息... ]"
    
    
    
    If Trim(Nvl(rsTmp!检查uid)) = "" Or Trim(Nvl(rsTmp!接收日期)) = "" Then
        
        '尚未采集图像，需要生成新的检查UID
        strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, Me.Caption)
            Exit Function
        End If
        
        '更新存储设备信息
        strSql = "Zl_影像检查_更新设备(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Else
        strNewStudyUID = Nvl(rsTmp!检查uid)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, Me.Caption)
            Exit Function
        End If
    End If
    
        
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始移动图像到新的图像存储位置... ]"
    
    '移动图像文件
    If Not MoveImageToStudy(rsImageDatas, strNewStudyUID, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd, objMoveList) Then
        Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
        
        Call MsgBoxD(Me, "图像移动失败，请检查FTP传输是否正常。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    
    
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始更新图像关联数据... ]"
    
    
    '获取报告图像信息
    strSql = "Select 检查UID,报告图象 From 影像检查记录 Where 医嘱ID=[1]"
    Set rsReportImage = zlDatabase.OpenSQLRecord(strSql, "关联影像", lngAdviceID)
    
    strOldReportImages = ""
    lngReportImageLen = 0
    
    If rsReportImage.RecordCount > 0 Then
        strOldReportImages = Nvl(rsReportImage!报告图象)
        lngReportImageLen = Len(strOldReportImages)
    End If
        
    '创建新的序列UID
    strNewSeriesUid = CreateSeriesUid(mdcmUID.NewUID)
    
    strReportImageIds = ""
    rsImageDatas.MoveFirst
                
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    While Not rsImageDatas.EOF
        '更新图像关联数据
        strSql = "Zl_影像检查_图像关联(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewSeriesUid & "','" & Nvl(rsImageDatas!图像UID) & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        '保存报告数据
        If mblnSaveReportImage Then
            If InStr(1, strOldReportImages & ";" & strReportImageIds, Nvl(rsImageDatas!图像UID)) <= 0 And Len(strReportImageIds) < 4000 - lngReportImageLen - 60 Then
                If strReportImageIds <> "" Then strReportImageIds = strReportImageIds & ";"
                strReportImageIds = strReportImageIds & Nvl(rsImageDatas!图像UID) & ".jpg"
            End If
        End If
    
        rsImageDatas.MoveNext
    Wend
    
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始更新报告图信息... ]"
    
    '如果需要保持报告图，则需要先查询目前已经保持的报告图像UID
    If mblnSaveReportImage Then
        
        If rsReportImage.RecordCount > 0 Then
            strReportImageIds = IIf(strOldReportImages <> "", strOldReportImages & ";", "") & strReportImageIds
            strReportImageIds = Replace(strReportImageIds, ";;", ";")
        End If
        
        strSql = "Zl_影像检查_更新报告图(" & mlngAdviceID & ",'" & strReportImageIds & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    End If
    
    '提交事务
    Call gcnOracle.CommitTrans
    
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始删除无效的FTP图像文件... ]"
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    StartReleation = True
    
    Exit Function
ErrHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    
    Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
    Call OutputDebug("StartReleation", err)
    
    Call RaiseErr(err)  '继续抛出错误
End Function

Private Function CancelReleation(ByVal lngAdviceID As Long, rsImageDatas As ADODB.Recordset) As Boolean
'撤销关联
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim strReportImageIds As String
    Dim strOldReportImages As Long
    Dim lngReportImageLen As Long
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    
    blnBeginTrans = False
    CancelReleation = False
    
    curDate = zlDatabase.Currentdate
    
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始获取新的图像存储信息... ]"
    
    '撤销图像关联
    strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
    
    Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
    If Trim(strNewFtpIp) = "" Then
        Call MsgBoxD(Me, "不能取得有效的存储设备，请检查存储设备配置。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始移动图像到新的图像存储位置... ]"
    
    '移动图像文件
    If Not MoveImageToStudy(rsImageDatas, strNewStudyUID, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd, objMoveList) Then
        Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
        
        Call MsgBoxD(Me, "图像移动失败，请检查FTP传输是否正常。", vbInformation, Me.Caption)
        Exit Function
    End If
    
    
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始更新图像关联数据... ]"
    
    strSql = "Select 检查UID,报告图象 From 影像检查记录 Where 医嘱ID=[1]"
    Set rsReportImage = zlDatabase.OpenSQLRecord(strSql, "关联影像", mlngAdviceID)
    
    If rsReportImage.RecordCount > 0 Then
        strReportImageIds = Nvl(rsReportImage!报告图象)
        strReportImageIds = Replace(strReportImageIds, " ", "") '采集图像时，可能会在报告图数据后添加空格
    End If
        
        
    '更新数据
    rsImageDatas.MoveFirst
    
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    While Not rsImageDatas.EOF
        strSql = "Zl_影像检查_撤销关联(" & mlngAdviceID & ",'" & Nvl(rsImageDatas!图像UID) & "','" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
                                        
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        '修改报告图数据
        strReportImageIds = Replace(strReportImageIds, Nvl(rsImageDatas!图像UID) & ".jpg;", "")
        strReportImageIds = Replace(strReportImageIds, Nvl(rsImageDatas!图像UID) & ".jpg", "")
        
        rsImageDatas.MoveNext
    Wend
    
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始更新报告图信息... ]"
    
    '更新报告图像
    strSql = "Zl_影像检查_更新报告图(" & mlngAdviceID & ",'" & strReportImageIds & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call gcnOracle.CommitTrans
    
    stb.Panels(2).Text = "正在执行，请等待!  [ 开始删除无效的FTP图像文件... ]"
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    CancelReleation = True
Exit Function
ErrHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    
    Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
    Call OutputDebug("CancelReleation", err)
    
    Call RaiseErr(err)
End Function


Public Function ReleationImage() As Boolean
'-----------------------------------------------------------------------------
'功能:关联图像，移动FTP图像到新的位置，修改数据库记录，从临时表转到正式表中
'返回：无
'-----------------------------------------------------------------------------
    Dim rsImageDatas As ADODB.Recordset
    Dim strHint As String
    Dim blnResult As Boolean
    
    On Error GoTo ErrHandle
        ReleationImage = False

        
        '在数据库中查询图像数据
        Set rsImageDatas = GetReleationImageIds()
    
        If rsImageDatas Is Nothing Then
            Call MsgBoxD(Me, "请选择需要进行关联的检查图像。", vbInformation, Me.Caption)
            Exit Function
        End If
        
        '当前检查UID在数据库中不存在，则退出本程序
        If rsImageDatas.RecordCount <= 0 Then
            Call MsgBoxD(Me, "请选择需要进行关联的检查图像。", vbInformation, Me.Caption)
            Exit Function
        End If
        
        
        If mlngReleationType = 2 Then
            '关联图像提示
            strHint = GetReleationHintInfo(mlngAdviceID, rsImageDatas)
            
            If strHint = "" Then
                Call MsgBoxD(Me, "不能查询到需要关联的数据信息，结束关联。", vbOKOnly, Me.Caption)
                Exit Function
            End If
            
            If MsgBoxD(Me, strHint, vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
            
        Else
            '取消关联提示
            If MsgBoxD(Me, "是否确认对所选图像进行取消关联操作？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        End If

        If mlngReleationType = 2 Then '等于2表示关联图像
            blnResult = StartReleation(mlngAdviceID, rsImageDatas)
        Else
            blnResult = CancelReleation(mlngAdviceID, rsImageDatas)
        End If
        

        stb.Panels(2).Text = "当前操作已执行完毕。"
        
        ReleationImage = blnResult
        
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Function


Private Sub cmdDel_Click()
'删除临时图像，只有在关联图像的窗口中才能进行删除，撤销关联时，不能执行该操作。
On Error GoTo ErrHandle
    Dim rsImageDatas As ADODB.Recordset
    Dim i As Long
    
    '在数据库中查询图像数据
    Set rsImageDatas = GetReleationImageIds()

    If rsImageDatas Is Nothing Then
        Call MsgBoxD(Me, "请选择需要删除的检查图像。", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    '当前检查UID在数据库中不存在，则退出本程序
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgBoxD(Me, "请选择需要删除的检查图像。", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    
    If MsgBoxD(Me, "是否确认删除所选图像？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    
    If DelTempImages(rsImageDatas) Then
        For i = vsfTree.Rows - 1 To 0 Step -1
            
            If vsfTree.TextMatrix(i, 0) = -1 Then
                Call vsfTree.RemoveItem(i)
            Else
                If vsfTree.GetNode(i).Children <= 0 And vsfTree.RowOutlineLevel(i) < 3 And vsfTree.RowData(i) = 1 Then
                    Call vsfTree.RemoveItem(i)
                End If
            End If
        Next i
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHandle
    mblnOK = ReleationImage
    
    If mblnOK Then
        Call MsgBoxD(Me, "当前操作已执行完毕。", vbInformation, Me.Caption)
        
        stb.Panels(2).Text = ""
        Unload Me
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
On Error GoTo ErrHandle
    Cancel = IIf((dkpMain.Panes(1).Hidden Or dkpMain.Panes(2).Hidden) And Action = 8 Or ((Action = 4 Or Action = 6 Or Action = 5) And Not Pane.Hidden), True, False)
ErrHandle:
End Sub


Private Sub dtpEnd_Change()
On Error GoTo ErrHandle
    Call QueryReleationData(dtpStart.value, dtpEnd.value)
        
    Call FilterReleationData
    Call LoadReleationDataToFace
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpStart_Change()
On Error GoTo ErrHandle
    Call QueryReleationData(dtpStart.value, dtpEnd.value)
        
    Call FilterReleationData
    Call LoadReleationDataToFace
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub DViewer_DblClick()
    If DViewer.Images.Count = 0 Then Exit Sub
    If mintSelectIndex <= 0 Then Exit Sub

    If DViewer.MultiColumns = 1 And DViewer.MultiRows = 1 Then
        DViewer.MultiColumns = mMultiCols
        DViewer.MultiRows = mMultiRows
        DViewer.CurrentIndex = 1
    Else
        mMultiCols = DViewer.MultiColumns
        mMultiRows = DViewer.MultiRows
        DViewer.MultiColumns = 1
        DViewer.MultiRows = 1
        DViewer.CurrentIndex = mintSelectIndex
    End If
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    If DViewer.Images.Count = 0 Then Exit Sub
    
    If Button = 1 And Shift = 0 Then
        mintSelectIndex = DViewer.ImageIndex(X, Y)
        
        If mintSelectIndex <= 0 Then Exit Sub
        
        For i = 1 To DViewer.Images.Count
            DViewer.Images(i).BorderColour = vbWhite
        Next i
        DViewer.Images(mintSelectIndex).BorderColour = vbBlue
    End If
End Sub





Private Sub Form_Activate()
    vsfTree.SetFocus
End Sub


Private Sub FillModality()
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select 编码,名称 from 影像检查类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "影像检查类别")
    
    cboModality.Clear
    Do Until rsTemp.EOF
        cboModality.AddItem rsTemp!编码 & "-" & rsTemp!名称
        If rsTemp!编码 = mstrModality Then cboModality.ListIndex = cboModality.ListCount - 1
        rsTemp.MoveNext
    Loop
    
    If cboModality.ListIndex = -1 Then
        If cboModality.ListCount >= 1 Then
            cboModality.ListIndex = 1
        End If
    End If
End Sub




Private Sub InitFaceScheme()
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    With Me.dkpMain
        .CloseAll
'        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 150, 200, DockLeftOf, Nothing)
    Pane1.Title = "病人列表"
    Pane1.Handle = PicList.hWnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable 'Or PaneNoHideable Or PaneNoCaption
    Pane1.MinTrackSize.Width = 520
    
    
    Set Pane2 = dkpMain.CreatePane(2, 150, 200, DockRightOf, Pane1)
    Pane2.Title = "图像预览"
    Pane2.Handle = picImage.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable 'Or PaneNoHideable Or PaneNoCaption
    Pane2.MinTrackSize.Width = 450

End Sub

Private Sub Form_Load()
    '恢复窗口状态
    Call RestoreWinState(Me, App.ProductName)
    
    chkViewImage.value = GetDeptPara(mlngCurDeptId, "预览关联图像", 0)
    ucPage.PageRecord = GetDeptPara(mlngCurDeptId, "预览关联图像数量", 9)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '保存窗口状态
    Call SaveWinState(Me, App.ProductName)
    
    Call SetDeptPara(mlngCurDeptId, "预览关联图像数量", Val(ucPage.PageRecord))
End Sub

Private Sub optDays_Click(Index As Integer)
On Error GoTo ErrHandle
    Dim i As Integer
    Dim dtNow As Date
    
    If mlngReleationType = 2 Then '关联图像
    
        dtpStart.Enabled = True
        dtpEnd.Enabled = True
        
        dtNow = zlDatabase.Currentdate
                        
        '增加时间过滤条件
        For i = 0 To 5
            If optDays(i).value = True Then
                Select Case i
                    Case 0
                        dtpStart.value = dtNow
                        dtpEnd.value = dtNow
                    Case 1
                        dtpStart.value = DateAdd("d", -1, dtNow)
                        dtpEnd.value = dtNow
                    Case 2
                        dtpStart.value = DateAdd("d", -2, dtNow)
                        dtpEnd.value = dtNow
                    Case 3
                        dtpStart.value = DateAdd("d", -4, dtNow)
                        dtpEnd.value = dtNow
                    Case 4
                        dtpStart.value = DateAdd("d", -6, dtNow)
                        dtpEnd.value = dtNow
                    Case 5
                        dtpStart.value = DateAdd("d", -14, dtNow)
                        dtpEnd.value = dtNow
'                        dtpStart.Enabled = False
'                        dtpEnd.Enabled = False
                End Select
            End If
        Next i
    
        Call QueryReleationData(dtpStart.value, dtpEnd.value)
        
        Call FilterReleationData
        Call LoadReleationDataToFace
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picCommand_Resize()
On Error GoTo ErrHandle
    cmdCancel.Left = picCommand.ScaleWidth - cmdCancel.Width - 120
    cmdCancel.Top = 60
    
    CmdOK.Left = cmdCancel.Left - CmdOK.Width - 120
    CmdOK.Top = 60
    
    cmdDel.Left = 120
    cmdDel.Top = 60
ErrHandle:
End Sub

Private Sub picImage_Resize()
On Error GoTo ErrHandle
    Dim iCols As Integer, iRows As Integer
    
    DViewer.Left = 0
    DViewer.Top = 0
    DViewer.Width = picImage.ScaleWidth
    DViewer.Height = picImage.ScaleHeight - ucPage.Height - 60 - stb.Height
    
    ucPage.Left = 0
    ucPage.Top = picImage.ScaleHeight - ucPage.Height - stb.Height
    
    chkViewImage.Left = ucPage.Left + ucPage.Width + 120
    chkViewImage.Top = ucPage.Top + 30

    ResizeRegion DViewer.Images.Count, DViewer.Width, DViewer.Height, iRows, iCols
    DViewer.MultiRows = iRows
    DViewer.MultiColumns = iCols
ErrHandle:
End Sub

Private Sub picList_Resize()
On Error GoTo ErrHandle
    frmFilter.Top = PicList.Height - frmFilter.Height - picCommand.Height - 240 - stb.Height
    frmFilter.Width = PicList.ScaleWidth - 180
    
    If mlngReleationType = 1 Then    '取消关联
        frmFilter.Visible = False
        vsfTree.Height = PicList.ScaleHeight - picCommand.Height - 240 - stb.Height
    ElseIf mlngReleationType = 2 Then    '关联图像
        frmFilter.Visible = True
        vsfTree.Height = PicList.ScaleHeight - frmFilter.Height - picCommand.Height - 240 - stb.Height
    End If
    
    vsfTree.Left = 0
    vsfTree.Top = 0
    vsfTree.Width = PicList.ScaleWidth
    
    picCommand.Left = 0
    picCommand.Top = PicList.Height - picCommand.Height - 120 - stb.Height
    picCommand.Width = PicList.ScaleWidth
ErrHandle:
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle
    If KeyAscii <> 13 Then Exit Sub
    
    Call FilterReleationData
    Call LoadReleationDataToFace
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub txtStudyNo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle
    If KeyAscii <> 13 Then Exit Sub
    
    Call FilterReleationData
    Call LoadReleationDataToFace
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucPage_OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)
On Error GoTo ErrHandle
    Dim rsData As ADODB.Recordset
    
    Dim strSearchId As String
    Dim lngSearchType As Long
    
    If Not vsfTree.Visible Then Exit Sub
    
    If chkViewImage.value = 0 Then Exit Sub
    
    strSearchId = vsfTree.TextMatrix(vsfTree.Row, 1)
    lngSearchType = vsfTree.RowOutlineLevel(vsfTree.Row)
    
    Set rsData = GetImageViewData(lngSearchType, strSearchId, lngPageIndex, lngPageCount)
    Call LoadViewImageToFace(rsData)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vsfTree_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Dim lngCurLevel As Long
    Dim i As Long
    
    If Col <> 0 Then Exit Sub
    
    lngCurLevel = vsfTree.RowOutlineLevel(Row)
    

    For i = Row + 1 To vsfTree.Rows - 1
        If vsfTree.RowOutlineLevel(i) <= lngCurLevel Then Exit For
        
        vsfTree.Cell(flexcpChecked, i, 0) = vsfTree.Cell(flexcpChecked, Row, Col)
    Next i
    
    
    i = Row - 1
    While i >= 0
        If vsfTree.RowOutlineLevel(i) < lngCurLevel Then
            If vsfTree.Cell(flexcpChecked, Row, 0) = 2 Then
                vsfTree.Cell(flexcpChecked, i, 0) = False
                lngCurLevel = vsfTree.RowOutlineLevel(i)
            End If
        End If
        
        i = i - 1
    Wend

    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub vsfTree_DblClick()
On Error GoTo ErrHandle
    If vsfTree.Rows <= 0 Then Exit Sub
        
    
    If vsfTree.IsCollapsed(vsfTree.Row) = flexOutlineCollapsed Then
        vsfTree.IsCollapsed(vsfTree.Row) = flexOutlineExpanded
    Else
        vsfTree.IsCollapsed(vsfTree.Row) = flexOutlineCollapsed
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadSeriesNode(ByVal lngStudyRow As Long)
'加载序列节点
    Dim strStudyUID As String
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim blnChecked As Boolean
    
    If vsfTree.GetNode(lngStudyRow).Children > 0 Then Exit Sub
    
    vsfTree.RowData(lngStudyRow) = 1
    
    strStudyUID = vsfTree.TextMatrix(lngStudyRow, 1)
    blnChecked = vsfTree.Cell(flexcpChecked, lngStudyRow, 0) = 1
    
    strSql = "select  序列UID, 检查UID, 序列号, 序列描述, 采集时间 from 影像临时序列 where 检查UID=[1] order by 序列号 desc"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strStudyUID)
    If rsData.RecordCount <= 0 Then Exit Sub
    
    With vsfTree
        '开始加载序列
        While Not rsData.EOF
            .AddItem "", lngStudyRow + 1
    
            .RowData(lngStudyRow + 1) = 1
    
            .Cell(flexcpChecked, lngStudyRow + 1, 0) = blnChecked
            .Cell(flexcpPicture, lngStudyRow + 1, 2) = imgSeries
    
            .Cell(flexcpText, lngStudyRow + 1, 1) = Nvl(rsData!序列UID)
            .Cell(flexcpText, lngStudyRow + 1, 2) = "序列" & Nvl(rsData!序列号)
    
            .Cell(flexcpText, lngStudyRow + 1, 3) = "序列号:" & Nvl(rsData!序列号) & "  序列描述:" & Nvl(rsData!序列描述) & "  生成日期:" & Nvl(rsData!采集时间)
            .Cell(flexcpFontSize, lngStudyRow + 1, 3) = 9
            .Cell(flexcpForeColor, lngStudyRow + 1, 3) = vbGrayText
    
            .IsSubtotal(lngStudyRow + 1) = True
            .RowOutlineLevel(lngStudyRow + 1) = 2
            
            Call LoadImageNode(lngStudyRow + 1)
            
            .IsCollapsed(lngStudyRow + 1) = flexOutlineCollapsed
            
            rsData.MoveNext
        Wend
    End With
End Sub

Private Sub LoadImageNode(ByVal lngSeriesRow As Long)
'加载图像节点
    Dim strSeriesUID As String
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim blnChecked As Boolean
    
    If vsfTree.GetNode(lngSeriesRow).Children > 0 Then Exit Sub
    
    vsfTree.RowData(lngSeriesRow) = 1
    
    strSeriesUID = vsfTree.TextMatrix(lngSeriesRow, 1)
    blnChecked = vsfTree.Cell(flexcpChecked, lngSeriesRow, 0) = 1
    
    strSql = "select  图像UID, 序列UID, 图像号, 图像描述, 采集时间  from 影像临时图象 where 序列UID=[1] order by 图像号 desc"
        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strSeriesUID)
    If rsData.RecordCount <= 0 Then Exit Sub
    
    
    With vsfTree
        '开始加载图像节点
        While Not rsData.EOF
            .AddItem "", lngSeriesRow + 1

            .RowData(lngSeriesRow + 1) = 1

            .Cell(flexcpChecked, lngSeriesRow + 1, 0) = blnChecked
            .Cell(flexcpPicture, lngSeriesRow + 1, 2) = imgImage

            .Cell(flexcpText, lngSeriesRow + 1, 1) = Nvl(rsData!图像UID)
            .Cell(flexcpText, lngSeriesRow + 1, 2) = "图像" & Nvl(rsData!图像号)
            .Cell(flexcpText, lngSeriesRow + 1, 3) = "图像号:" & Nvl(rsData!图像号) & "  采集时间:" & Nvl(rsData!采集时间)

            .Cell(flexcpFontSize, lngSeriesRow + 1, 3) = 9
            .Cell(flexcpForeColor, lngSeriesRow + 1, 3) = &HC0C0FF

            .IsSubtotal(lngSeriesRow + 1) = True
            .RowOutlineLevel(lngSeriesRow + 1) = 3

            Call rsData.MoveNext
        Wend
    End With
    
End Sub


Private Sub vsfTree_SelChange()
On Error GoTo ErrHandle
    Dim rsData As ADODB.Recordset
    Dim strSearchId As String
    Dim lngSearchType As Long
    
    ucPage.RecordCount = 0
    
    If vsfTree.Row < 0 Then Exit Sub
    If vsfTree.RowSel < 0 Then Exit Sub
    
    
    If mlngReleationType = 2 Then
        '如果节点为检查，则需要判断是否有子层节点，如果没有则加载
        If vsfTree.RowOutlineLevel(vsfTree.Row) = 1 Then
            Call LoadSeriesNode(vsfTree.Row)
        End If
        
'        '如果节点为序列，则需要判断是否有子层节点，如果没有则加载
'        If vsfTree.RowOutlineLevel(vsfTree.Row) = 2 Then
'            Call LoadImageNode(vsfTree.Row)
'        End If
    End If
    
    '没有启用图像预览时，这不加载图像
    If chkViewImage.value = 0 Then Exit Sub
    
    strSearchId = vsfTree.TextMatrix(vsfTree.Row, 1)
    lngSearchType = vsfTree.RowOutlineLevel(vsfTree.Row)
    
    
    Call InitPageControl(lngSearchType, strSearchId)
    
    Set rsData = GetImageViewData(lngSearchType, strSearchId, 1, ucPage.PageRecord)
    Call LoadViewImageToFace(rsData)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vsfTree_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
