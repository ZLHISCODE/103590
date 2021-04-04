VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugList 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "病人用药清单"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18540
   Icon            =   "frmDrugList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10185
   ScaleWidth      =   18540
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picAdviceFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   2040
      ScaleHeight     =   2625
      ScaleWidth      =   2985
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton cmdAdviceQuit 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   300
         Left            =   1560
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdviceOK 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "确定(&O)"
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkZY 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "中药"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   1680
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkXY 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "西药"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1680
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chk住院 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "住院"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   1230
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkMZ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "门诊"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   1230
         Value           =   1  'Checked
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   300
         Left            =   975
         TabIndex        =   6
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   201981955
         CurrentDate     =   40976
      End
      Begin MSComCtl2.DTPicker dtpStopTime 
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   201981955
         CurrentDate     =   40976
      End
      Begin VB.Label lblFilter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "用药自动提取过滤"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   960
         TabIndex        =   22
         Top             =   60
         Width           =   1560
      End
      Begin VB.Image imgIco 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   600
         Picture         =   "frmDrugList.frx":6852
         Stretch         =   -1  'True
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblAdvice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药品分类："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1777
         Width           =   900
      End
      Begin VB.Label lblAdvice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医嘱来源："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1327
         Width           =   900
      End
      Begin VB.Label lblAdvice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   900
         Width           =   900
      End
      Begin VB.Label lblAdvice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   420
         Width           =   900
      End
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C000&
      Height          =   465
      Left            =   720
      ScaleHeight     =   465
      ScaleWidth      =   10455
      TabIndex        =   14
      Top             =   1200
      Width           =   10455
      Begin VB.PictureBox picTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6480
         ScaleHeight     =   375
         ScaleWidth      =   3975
         TabIndex        =   23
         Top             =   38
         Width           =   3975
         Begin VB.OptionButton OptTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "所有"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   2
            Top             =   0
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton OptTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "近一月"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   3
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton OptTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "近三月"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   2160
            TabIndex        =   4
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton OptTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "近半年"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   3120
            TabIndex        =   5
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblTime 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "时间"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   24
            Top             =   90
            Width           =   360
         End
      End
      Begin VB.Label lblPati 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病人信息："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1050
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   9825
      Width           =   18540
      _ExtentX        =   32703
      _ExtentY        =   635
      SimpleText      =   $"frmDrugList.frx":D0A4
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugList.frx":D0EB
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   27623
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
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   5925
      _cx             =   10451
      _cy             =   6271
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
      MouseIcon       =   "frmDrugList.frx":D97F
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   400
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   10000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugList.frx":E259
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox pictmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1920
         ScaleHeight     =   240
         ScaleWidth      =   480
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   360
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDrugList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnMod As Boolean   'true 窗体模态显示，false非模态显示
Private mlng病人ID As Long   '病人id
Private mlng主页ID As Long   '主页id
Private mlngEditTag As Long  '界面状态设置，0-查阅状态，1-编辑状态
Private mblnReturn As Boolean
Private mstrLike As String
Private mint简码 As Integer
Private mstrTip As String
Private mlngLastColor As Long   '上次选择项颜色

Private Type PointAPI
        X As Long
        Y As Long
End Type

Private Const GRD_UNEDITCELL_COLOR = &H8000000B  '未编辑的单元格颜色：灰蓝色
Private Const Red_COLOR = &HC0C0FF  '淡红色


Private Enum COL用药清单
    '隐藏列
    COL_ID = 1
    COL_病人ID = 2
    COL_主页ID = 3
    col_组号 = 4
    COL_用药来源 = 5
    COL_诊疗项目ID = 6
    COL_收费细目ID = 7
    COL_频率间隔 = 8
    COL_间隔单位 = 9
    COL_用法id = 10
    col_煎法id = 11
    COL_终止时间 = 12
    '可见列
    COL_开始时间 = 13
    col_药品类别 = 14
    col_用药内容 = 15
    COL_用法 = 16
    COL_单次用量 = 17
    COL_单量单位 = 18
    COL_总给予量 = 19
    COL_总量单位 = 20
    COL_天数 = 21
    COL_执行频次 = 22
    COL_备注 = 23
    
    '隐藏列
    COL_登记人 = 24
    COL_登记时间 = 25
    col_配方数据 = 26 '格式:[配方数据]中药名称<Data>诊疗项目ID<Data>收费细目ID<Data>单量<Data>脚注<Data>单位
    col_是否修改 = 27
    COL_频率次数 = 28
End Enum

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Private Function LoadDrug()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim intTime As Integer
    Dim i As Long
    
    On Error GoTo errH
    With vsAdvice
        For i = 0 To OptTime.Count - 1
            If OptTime(i).value = True Then
                intTime = decode(OptTime(i).Caption, "近一月", 1, "近三月", 3, "近半年", 6)
                Exit For
            End If
        Next
        strSQL = "Select a.Id, a.病人id, a.主页id, a.组号, a.用药来源, a.药品类别, a.用药内容, a.诊疗项目id, a.收费细目id, a.天数, a.开始时间, a.终止时间, a.登记时间, a. 登记人,a. 总给予量, a.单次用量, a. 执行频次, a.频率次数, a.频率间隔, a.间隔单位, a.用法id, a.煎法id, a.备注, b.计算单位,C.名称 as 用法,D.住院单位" & _
                " From 病人用药清单 A,诊疗项目目录 B, 诊疗项目目录 C, 药品规格 D Where a.诊疗项目id = b.Id(+) And a.收费细目id=D.药品ID(+) And A.用法ID=C.ID(+) And a.病人id = [1]" & IIF(intTime = 0, "", " And a.开始时间 Between add_months(sysdate,-[2]) And Sysdate") & " Order By a.开始时间,a.组号,a.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, intTime)
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             .Rows = .FixedRows + rsTmp.RecordCount
             For i = 1 To rsTmp.RecordCount
                '隐藏列
                .TextMatrix(i, COL_ID) = Val(rsTmp!ID & "")
                .TextMatrix(i, COL_病人ID) = Val(rsTmp!病人ID & "")
                .TextMatrix(i, COL_主页ID) = Val(rsTmp!主页ID & "")
                .TextMatrix(i, col_组号) = Val(rsTmp!组号 & "")
                .TextMatrix(i, COL_用药来源) = Val(rsTmp!用药来源 & "")
                .TextMatrix(i, COL_诊疗项目ID) = Val(rsTmp!诊疗项目ID & "")
                .TextMatrix(i, COL_收费细目ID) = Val(rsTmp!收费细目ID & "")
                .TextMatrix(i, COL_频率间隔) = Val(rsTmp!频率间隔 & "")
                .TextMatrix(i, COL_间隔单位) = rsTmp!间隔单位 & ""
                .TextMatrix(i, COL_用法id) = Val(rsTmp!用法ID & "")
                .TextMatrix(i, col_煎法id) = Val(rsTmp!煎法ID & "")
                .TextMatrix(i, COL_终止时间) = Format(rsTmp!终止时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_开始时间) = Format(rsTmp!开始时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, col_药品类别) = decode(rsTmp!药品类别 & "", "5", "西成药", "6", "中成药", "中草药")
                .TextMatrix(i, col_用药内容) = rsTmp!用药内容 & ""
                .TextMatrix(i, COL_用法) = rsTmp!用法 & ""
                .TextMatrix(i, COL_单次用量) = IIF(.TextMatrix(i, col_药品类别) = "中草药", "", FormatEx(NVL(rsTmp!单次用量), 5))
                .TextMatrix(i, COL_总给予量) = FormatEx(NVL(rsTmp!总给予量), 5)
                .TextMatrix(i, COL_单量单位) = IIF(.TextMatrix(i, col_药品类别) = "中草药", "", rsTmp!计算单位 & "")
                .TextMatrix(i, COL_总量单位) = IIF(.TextMatrix(i, col_药品类别) = "中草药", "付", rsTmp!住院单位 & "")
                .TextMatrix(i, COL_天数) = FormatEx(NVL(rsTmp!天数), 5)
                .TextMatrix(i, COL_执行频次) = rsTmp!执行频次 & ""
                .TextMatrix(i, COL_备注) = rsTmp!备注 & ""
                .TextMatrix(i, COL_登记人) = rsTmp!登记人 & ""
                .TextMatrix(i, COL_登记时间) = Format(rsTmp!登记时间 & "", "yyyy-mm-dd hh:mm")
                
                '缓存数据
                .Cell(flexcpData, i, COL_执行频次) = .TextMatrix(i, COL_执行频次)
                .Cell(flexcpData, i, COL_用法) = .TextMatrix(i, COL_用法)
                .Cell(flexcpData, i, col_用药内容) = .TextMatrix(i, col_用药内容)
                .Cell(flexcpData, i, col_药品类别) = decode(.TextMatrix(i, col_药品类别), "西成药", "5", "中成药", "6", "中草药", "8")
                rsTmp.MoveNext
             Next
             .Redraw = flexRDDirect
        Else
            .Rows = .FixedRows + 1
        End If
        .WordWrap = True
        '自动调整行高
        .AutoSize col_用药内容
        .Cell(flexcpBackColor, .FixedRows, col_药品类别, .Rows - 1, col_药品类别) = GRD_UNEDITCELL_COLOR      '灰蓝色
        .Cell(flexcpBackColor, .FixedRows, COL_单量单位, .Rows - 1, COL_单量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
        .Cell(flexcpBackColor, .FixedRows, COL_总量单位, .Rows - 1, COL_总量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, 0) = GRD_UNEDITCELL_COLOR
        SetTag一并给药
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function SaveDrug()
    Dim lngID As Long
    Dim i As Long, j As Long
    Dim arrSQL As Variant
    Dim dtNow As Date
    Dim blnTran As Boolean
    Dim arrTime As Variant, arrTmp As Variant
    Dim lng组号 As Long
    
    arrSQL = Array()
    On Error GoTo errH
    dtNow = zlDatabase.Currentdate
    With vsAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col_用药内容) <> "" Then
                lng组号 = 0
                If Val(.TextMatrix(i, COL_ID)) <> 0 Then
                    If .TextMatrix(i, col_是否修改) = "1" Then
                        If .RowHidden(i) = True Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_病人用药清单_Delete(" & Val(.TextMatrix(i, COL_ID)) & ")"
                        Else
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            
                            arrSQL(UBound(arrSQL)) = "Zl_病人用药清单_Update(" & Val(.TextMatrix(i, COL_ID)) & "," & ZVal(.TextMatrix(i, col_组号)) & "," & Val(.TextMatrix(i, COL_用药来源)) & ",'" & _
                                Val(.Cell(flexcpData, i, col_药品类别)) & "','" & .TextMatrix(i, col_用药内容) & "'," & ZVal(.TextMatrix(i, COL_诊疗项目ID)) & "," & ZVal(.TextMatrix(i, COL_收费细目ID)) & "," & _
                                ZVal(.TextMatrix(i, COL_天数)) & ",To_Date('" & Format(.TextMatrix(i, COL_开始时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & IIF(.TextMatrix(i, COL_终止时间) = "", "Null", "To_Date('" & Format(.TextMatrix(i, COL_终止时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')") & "," & _
                                "To_Date('" & Format(dtNow, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & .TextMatrix(i, COL_登记人) & "'," & ZVal(.TextMatrix(i, COL_总给予量)) & "," & ZVal(.TextMatrix(i, COL_单次用量)) & ",'" & _
                                .TextMatrix(i, COL_执行频次) & "'," & ZVal(.TextMatrix(i, COL_频率次数)) & "," & ZVal(.TextMatrix(i, COL_频率间隔)) & ",'" & .TextMatrix(i, COL_间隔单位) & "'," & ZVal(.TextMatrix(i, COL_用法id)) & "," & _
                                ZVal(.TextMatrix(i, col_煎法id)) & ",'" & .TextMatrix(i, COL_备注) & "')"
                        
                            If .TextMatrix(i, col_药品类别) = "中草药" Or .Cell(flexcpData, i, col_配方数据) <> "" Then
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_病人用药配方_Delete(" & Val(.TextMatrix(i, COL_ID)) & ")"
                            End If
                            If .TextMatrix(i, col_药品类别) = "中草药" Then
                                arrTime = Split(.TextMatrix(i, col_配方数据), "[配方数据]")
                                For j = 1 To UBound(arrTime)
                                    arrTmp = Split(arrTime(j), "<Data>")
                                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                    arrSQL(UBound(arrSQL)) = "Zl_病人用药配方_Insert(" & Val(.TextMatrix(i, COL_ID)) & "," & j & "," & Val(arrTmp(1)) & "," & Val(arrTmp(2)) & "," & ZVal(arrTmp(3)) & ",'" & arrTmp(4) & "')"
                                Next
                            End If
                        End If
                    End If
                Else
                    lngID = zlDatabase.GetNextID("病人用药清单")
                    
                    '转换一并给药ID
                    If .TextMatrix(i, col_药品类别) <> "中草药" And Val(.TextMatrix(i, col_组号)) <> 0 Then
                        If Val(.TextMatrix(i, col_组号)) < 0 Then
                            If Val(.TextMatrix(i, col_组号)) = Val(.TextMatrix(Abs(Val(.TextMatrix(i, col_组号))), col_组号)) Then
                                If i = Abs(Val(.TextMatrix(i, col_组号))) Then
                                   lng组号 = lngID
                                Else
                                   lng组号 = Val(.Cell(flexcpData, Abs(Val(.TextMatrix(i, col_组号))), col_组号))
                                End If
                            End If
                            .Cell(flexcpData, i, col_组号) = lng组号
                        Else
                            lng组号 = Val(.TextMatrix(i, col_组号))
                        End If
                    End If
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人用药清单_Insert(" & lngID & "," & mlng病人ID & "," & mlng主页ID & "," & ZVal(lng组号) & "," & Val(.TextMatrix(i, COL_用药来源)) & ",'" & _
                        Val(.Cell(flexcpData, i, col_药品类别)) & "','" & .TextMatrix(i, col_用药内容) & "'," & ZVal(.TextMatrix(i, COL_诊疗项目ID)) & "," & ZVal(.TextMatrix(i, COL_收费细目ID)) & "," & _
                        ZVal(.TextMatrix(i, COL_天数)) & ",To_Date('" & Format(.TextMatrix(i, COL_开始时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & IIF(.TextMatrix(i, COL_终止时间) = "", "Null", "To_Date('" & Format(.TextMatrix(i, COL_终止时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')") & "," & _
                        "To_Date('" & Format(dtNow, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & .TextMatrix(i, COL_登记人) & "'," & ZVal(.TextMatrix(i, COL_总给予量)) & "," & ZVal(.TextMatrix(i, COL_单次用量)) & ",'" & _
                        .TextMatrix(i, COL_执行频次) & "'," & ZVal(.TextMatrix(i, COL_频率次数)) & "," & ZVal(.TextMatrix(i, COL_频率间隔)) & ",'" & .TextMatrix(i, COL_间隔单位) & "'," & ZVal(.TextMatrix(i, COL_用法id)) & "," & _
                        ZVal(.TextMatrix(i, col_煎法id)) & ",'" & .TextMatrix(i, COL_备注) & "')"
                        
                    If .TextMatrix(i, col_药品类别) = "中草药" Then
                        arrTime = Split(.TextMatrix(i, col_配方数据), "[配方数据]")
                        For j = 1 To UBound(arrTime)
                            arrTmp = Split(arrTime(j), "<Data>")
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_病人用药配方_Insert(" & lngID & "," & j & "," & Val(arrTmp(1)) & "," & Val(arrTmp(2)) & "," & ZVal(arrTmp(3)) & ",'" & arrTmp(4) & "')"
                        Next
                    End If
                    .Cell(flexcpData, i, COL_收费细目ID) = lngID  '缓存id
                End If
                
            End If
        Next
    End With
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    Screen.MousePointer = 0
    
    SaveDrug = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function





Private Sub UpdateDrug()
    '功能：更新当前数据
    Dim i As Long, lngRows As Long

    With vsAdvice
        lngRows = .Rows - 1
        For i = lngRows To 1 Step -1
            If .TextMatrix(i, col_是否修改) = "1" Then .TextMatrix(i, col_是否修改) = ""
            If Val(.Cell(flexcpData, i, col_组号)) > 0 And Val(.TextMatrix(i, COL_ID)) = 0 And Val(.TextMatrix(i, col_组号)) < 0 Then .TextMatrix(i, col_组号) = Val(.Cell(flexcpData, i, col_组号))
            .Cell(flexcpData, i, col_组号) = ""
            If .Cell(flexcpData, i, COL_收费细目ID) <> "" Then .TextMatrix(i, COL_ID) = Val(.Cell(flexcpData, i, COL_收费细目ID)): .Cell(flexcpData, i, COL_收费细目ID) = ""
            .Cell(flexcpBackColor, i, 0, i, 0) = GRD_UNEDITCELL_COLOR
            If .RowHidden(i) = True Then .RemoveItem (i)
        Next
        vsAdvice.Tag = ""
    End With
End Sub


Public Function ShowMe(frmParent As Object, ByVal blnMod As Boolean, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：病人用药清单
'参数：frmParent 父窗体
'      blnMod 是否是模态方式显示
'      lng病人ID,
'      lng主页ID,
'返回：
    mblnMod = blnMod
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    Me.Show IIF(blnMod, 1, 0), frmParent
End Function

Private Function checkDrug()
    Dim i As Long, j As Long
    
    mstrTip = ""
    With vsAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col_用药内容) <> "" And .RowHidden(i) = False Then
                '恢复颜色
                If .Cell(flexcpBackColor, i, COL_用法, i, COL_用法) = Red_COLOR Then .Cell(flexcpBackColor, i, COL_用法, i, COL_用法) = 0
                If .Cell(flexcpBackColor, i, col_用药内容, i, col_用药内容) = Red_COLOR Then .Cell(flexcpBackColor, i, col_用药内容, i, col_用药内容) = 0
                If .Cell(flexcpBackColor, i, COL_开始时间, i, COL_开始时间) = Red_COLOR Then .Cell(flexcpBackColor, i, COL_开始时间, i, COL_开始时间) = 0
                
                If .TextMatrix(i, COL_开始时间) = "" Then
                    .Cell(flexcpBackColor, i, COL_开始时间, i, COL_开始时间) = Red_COLOR
                    MsgBox "用药清单的开始时间为必填项,请录入。", vbInformation, gstrSysName
                    mstrTip = i & "|" & COL_开始时间 & "|" & "用药清单开始时间为必填项,请录入。"
                    .Row = i: .Col = COL_开始时间: Call vsAdvice.ShowCell(.Row, .Col)
                    Exit Function
                End If
                
                If .TextMatrix(i, COL_用法) = "" Then
                    .Cell(flexcpBackColor, i, COL_用法, i, COL_用法) = Red_COLOR
                    MsgBox "用药清单的用法为必填项,请录入。", vbInformation, gstrSysName
                    mstrTip = i & "|" & COL_用法 & "|" & "用药清单的用法为必填项,请录入。"
                    .Row = i: .Col = COL_用法: Call vsAdvice.ShowCell(.Row, .Col)
                    Exit Function
                End If
                
                If i <> .Rows - 1 And .TextMatrix(i, col_药品类别) <> "中草药" Then '检查是否存在相同用药清单
                    For j = .Rows - 1 To i + 1 Step -1
                        If .Cell(flexcpBackColor, i, col_用药内容, i, col_用药内容) = Red_COLOR Then .Cell(flexcpBackColor, i, col_用药内容, i, col_用药内容) = 0
                        If .TextMatrix(j, col_用药内容) <> "" And .RowHidden(j) = False Then
                            If .TextMatrix(j, COL_开始时间) & "|" & .TextMatrix(j, col_用药内容) & "|" & .TextMatrix(j, col_药品类别) & "|" & .TextMatrix(j, COL_诊疗项目ID) & "|" & .TextMatrix(j, COL_收费细目ID) = .TextMatrix(i, COL_开始时间) & "|" & .TextMatrix(i, col_用药内容) & "|" & .TextMatrix(i, col_药品类别) & "|" & .TextMatrix(i, COL_诊疗项目ID) & "|" & .TextMatrix(i, COL_收费细目ID) Then
                                .Cell(flexcpBackColor, j, col_用药内容, j, col_用药内容) = Red_COLOR
                                MsgBox "发现两条重复的用药清单,请检查。", vbInformation, gstrSysName
                                mstrTip = j & "|" & col_用药内容 & "|" & "发现两条重复的用药清单,请检查。"
                                .Row = j: .Col = col_用药内容: Call vsAdvice.ShowCell(.Row, .Col)
                                Exit Function
                            End If
                        End If
                    Next
                End If
            End If
        Next
        checkDrug = True
    End With
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim Pt As PointAPI
    Dim i As Long, lngTmp As Long
    Dim lngUpRow As Long
    With vsAdvice
        Select Case Control.ID
        Case conMenu_Edit_Save  '保存记录
            If checkDrug = True Then
                If SaveDrug Then
                    Call UpdateDrug
                End If
            End If
        Case conMenu_Edit_DrugAuto '自动提取
            GetCursorPos Pt
            picAdviceFilter.Left = Pt.X + (picAdviceFilter.Width / 2): picAdviceFilter.Top = Pt.Y + 300
            picAdviceFilter.Visible = Not picAdviceFilter.Visible
            picAdviceFilter.Enabled = picAdviceFilter.Visible
            If picAdviceFilter.Visible = True Then
                cmdAdviceOK.SetFocus
            Else
                .SetFocus
            End If
        Case conMenu_Edit_NewItem '新增用药记录
            If .TextMatrix(.Rows - 1, col_用药内容) = "" Then
                .Row = .Rows - 1: .Col = COL_开始时间
                .ShowCell .Row, COL_开始时间
            Else
                .Rows = vsAdvice.Rows + 1
                .Cell(flexcpBackColor, .FixedRows, col_药品类别, .Rows - 1, col_药品类别) = GRD_UNEDITCELL_COLOR      '灰蓝色
                .Cell(flexcpBackColor, .FixedRows, COL_单量单位, .Rows - 1, COL_单量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
                .Cell(flexcpBackColor, .FixedRows, COL_总量单位, .Rows - 1, COL_总量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
                .Cell(flexcpBackColor, .Rows - 1, 0, vsAdvice.Rows - 1, 0) = Red_COLOR
                .Row = vsAdvice.Rows - 1: .Col = COL_开始时间
                .ShowCell .Row, COL_开始时间
            End If
        Case conMenu_Edit_Modify '修改用药记录
            mlngLastColor = 0
            mlngEditTag = 1
            vsAdvice.Editable = flexEDKbdMouse
            lblTime.Visible = mlngEditTag = 0
             OptTime(0).Visible = lblTime.Visible: OptTime(1).Visible = lblTime.Visible: OptTime(2).Visible = lblTime.Visible: OptTime(3).Visible = lblTime.Visible
            If vsAdvice.Col = col_用药内容 Then Call Get用药配方(vsAdvice.Row)
            staThis.Panels(2).Text = "当前模式为：" & IIF(mlngEditTag = 0, "查阅用药清单", "编辑用药清单")
        Case conMenu_Edit_ItemUndo '退出编辑
            If Val(vsAdvice.Tag) = 1 Then
                If MsgBox("当前还有未保存的用药记录,确定要取消编辑吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            mlngEditTag = 0
            .Editable = flexEDNone
            Call LoadDrug
            lblTime.Visible = mlngEditTag = 0
            OptTime(0).Visible = lblTime.Visible: OptTime(1).Visible = lblTime.Visible: OptTime(2).Visible = lblTime.Visible: OptTime(3).Visible = lblTime.Visible
            picAdviceFilter.Visible = False
            staThis.Panels(2).Text = "当前模式为：" & IIF(mlngEditTag = 0, "查阅用药清单", "编辑用药清单")
        Case conMenu_Edit_Delete '删除用药记录
            Call DeteleRow
        Case conMenu_Edit_DrugGrp '用药记录一并给药
            If Control.Checked = True Then
                If .TextMatrix(.Row, 0) = "┗" And .TextMatrix(GetUpRow(.Row), 0) <> "┏" And Val(.TextMatrix(.Row, col_组号)) <> Val(.TextMatrix(.Row, COL_ID)) And Val(.TextMatrix(.Row, col_组号)) <> -.Row Then
                    .TextMatrix(.Row, 0) = ""
                    lngTmp = Val(.TextMatrix(.Row, col_组号))
                    .TextMatrix(.Row, col_组号) = ""
                    Call SetTag一并给药(lngTmp)
                Else
                    If MsgBox("要将该组一并给药的药品全部取消为单独给药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                    
                    lngTmp = Val(.TextMatrix(.Row, col_组号))
                    For i = .FixedRows To .Rows - 1
                        If lngTmp = Val(.TextMatrix(i, col_组号)) Then
                            .TextMatrix(i, 0) = ""
                            .TextMatrix(i, col_组号) = ""
                            .TextMatrix(i, col_是否修改) = "1"
                            .Cell(flexcpBackColor, i, 0, i, 0) = Red_COLOR
                        End If
                    Next
                End If
            Else
                If Not Check一并给药(vsAdvice.Row) Then Exit Sub
                lngUpRow = GetUpRow(.Row)
                If Val(.TextMatrix(lngUpRow, col_组号)) = 0 Then
                    If Val(.TextMatrix(lngUpRow, COL_ID)) <> 0 Then
                        .TextMatrix(lngUpRow, col_组号) = Val(.TextMatrix(lngUpRow, COL_ID))
                    Else
                        .TextMatrix(lngUpRow, col_组号) = -lngUpRow
                    End If
                    .TextMatrix(lngUpRow, col_是否修改) = "1"
                    .Cell(flexcpBackColor, lngUpRow, 0, lngUpRow, 0) = Red_COLOR
                End If
                
                '一并给药数据同步
                .TextMatrix(.Row, col_组号) = .TextMatrix(lngUpRow, col_组号)
                .TextMatrix(.Row, COL_开始时间) = .TextMatrix(lngUpRow, COL_开始时间)
                .TextMatrix(.Row, COL_用法) = .TextMatrix(lngUpRow, COL_用法)
                .TextMatrix(.Row, COL_用法id) = .TextMatrix(lngUpRow, COL_用法id)
                .TextMatrix(.Row, COL_执行频次) = .TextMatrix(lngUpRow, COL_执行频次)
                .TextMatrix(.Row, COL_频率间隔) = .TextMatrix(lngUpRow, COL_频率间隔)
                .TextMatrix(.Row, COL_间隔单位) = .TextMatrix(lngUpRow, COL_间隔单位)
                .TextMatrix(.Row, COL_频率次数) = .TextMatrix(lngUpRow, COL_频率次数)
                .TextMatrix(.Row, COL_天数) = .TextMatrix(lngUpRow, COL_天数)
                
                .Cell(flexcpData, .Row, COL_执行频次) = .TextMatrix(lngUpRow, COL_执行频次)
                .Cell(flexcpData, .Row, COL_用法) = .TextMatrix(lngUpRow, COL_用法)
                Call SetTag一并给药(.TextMatrix(.Row, col_组号))
            End If
            .Tag = "1"
            .TextMatrix(.Row, col_是否修改) = "1"
            .Cell(flexcpBackColor, .Row, 0, .Row, 0) = Red_COLOR
        Case conMenu_File_Exit '退出
            Unload Me
        End Select
    End With
End Sub

Private Function Change一并给药ID(ByVal lngRow As Long) As Long
    Dim i As Long
    Dim lngTmp As Long
    With vsAdvice
        If Val(.TextMatrix(lngRow, col_组号)) <> 0 And (Val(.TextMatrix(.Row, col_组号)) = Val(.TextMatrix(.Row, COL_ID)) Or Val(.TextMatrix(lngRow, col_组号)) = -lngRow) Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, col_组号)) = Val(.TextMatrix(lngRow, col_组号)) And i <> lngRow Then
                    If lngTmp = 0 Then
                        lngTmp = IIF(Val(.TextMatrix(i, COL_ID)) <> 0, Val(.TextMatrix(i, COL_ID)), -i)
                    End If
                    .TextMatrix(i, col_组号) = lngTmp
                End If
            Next
        End If
        Change一并给药ID = lngTmp
    End With
End Function


Private Sub DeteleRow()
    Dim i As Long
    Dim blnTmp As Boolean
    Dim lngTmp As Long
    Dim lng组号 As Long

    With vsAdvice
        If .Row < 1 Then Exit Sub
        If Val(.TextMatrix(.Row, COL_ID)) = 0 Then
            If .TextMatrix(.Row, COL_诊疗项目ID) <> "" Then
                If MsgBox("确实要删除该行用药记录吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    '处理一并给药
                    If (.TextMatrix(.Row, 0) = "┏" And .TextMatrix(GetDownRow(.Row), 0) = "┗") Or (.TextMatrix(.Row, 0) = "┗" And .TextMatrix(GetUpRow(.Row), 0) = "┏") Then
                        lng组号 = Val(.TextMatrix(.Row, col_组号))
                        For i = .FixedRows To .Rows - 1
                            If lng组号 = Val(.TextMatrix(i, col_组号)) Then
                                .TextMatrix(i, 0) = ""
                                .TextMatrix(i, col_组号) = ""
                                .TextMatrix(i, col_是否修改) = "1"
                                .Cell(flexcpBackColor, i, 0, i, 0) = Red_COLOR
                            End If
                        Next
                    Else
                        lngTmp = Change一并给药ID(.Row)
                        If lngTmp = 0 Then lngTmp = Val(.TextMatrix(.Row, col_组号))
                    End If
                    mlngLastColor = 0
                    .RemoveItem .Row
                    If lngTmp <> 0 Then SetTag一并给药 (lngTmp)
                Else
                    Exit Sub
                End If
            Else
                mlngLastColor = 0
                .RemoveItem .Row
            End If
            
        Else
            If .RowHidden(.Row) = False Then
                If MsgBox("确实要删除该行用药记录吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    '处理一并给药
                    If (.TextMatrix(.Row, 0) = "┏" And .TextMatrix(GetDownRow(.Row), 0) = "┗") Or (.TextMatrix(.Row, 0) = "┗" And .TextMatrix(GetUpRow(.Row), 0) = "┏") Then
                        lng组号 = Val(.TextMatrix(.Row, col_组号))
                        For i = .FixedRows To .Rows - 1
                            If lng组号 = Val(.TextMatrix(i, col_组号)) Then
                                .TextMatrix(i, 0) = ""
                                .TextMatrix(i, col_组号) = ""
                                .TextMatrix(i, col_是否修改) = "1"
                                .Cell(flexcpBackColor, i, 0, i, 0) = Red_COLOR
                            End If
                        Next
                    Else
                        lngTmp = Change一并给药ID(.Row)
                        If lngTmp = 0 Then lngTmp = Val(.TextMatrix(.Row, col_组号))
                    End If
                    mlngLastColor = 0
                    .RowHidden(.Row) = True
                    .TextMatrix(.Row, col_组号) = ""
                    If lngTmp <> 0 Then SetTag一并给药 (lngTmp)
                    .TextMatrix(.Row, col_是否修改) = "1"
                Else
                    Exit Sub
                End If
            End If
        End If
        
        
        '寻找上一个焦点
        If .RowHidden(.Row) = True Then
            For i = .Row To 1 Step -1
                If .RowHidden(i) = False Then
                    .Row = i: .Col = col_用药内容: .ShowCell i, col_用药内容: Exit For
                End If
            Next
        End If
        
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                blnTmp = True: Exit For
            End If
        Next
        If .Rows = 1 Or (Not blnTmp) Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, .FixedRows, col_药品类别, .Rows - 1, col_药品类别) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Cell(flexcpBackColor, .FixedRows, COL_单量单位, .Rows - 1, COL_单量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Cell(flexcpBackColor, .FixedRows, COL_总量单位, .Rows - 1, COL_总量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Row = .Rows - 1: .Col = COL_开始时间
            .ShowCell .Rows - 1, COL_开始时间
        End If
        .Tag = "1"
    End With
End Sub

Private Function Get用药配方(lngRow As Long) As String
    Dim strSQL As String, strTmp
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_ID)) <> 0 And .TextMatrix(lngRow, col_药品类别) = "中草药" And .TextMatrix(lngRow, col_配方数据) = "" Then
            strSQL = "Select a.配方id, a.序号, a.诊疗项目id, a.收费细目id, a.单量, a.脚注, b.名称, b.计算单位 From 病人用药配方 A, 诊疗项目目录 B Where a.诊疗项目id = b.Id And a.配方id =[1]  order by A.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)))
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    strTmp = strTmp & "[配方数据]" & rsTmp!名称 & "<Data>" & Val(rsTmp!诊疗项目ID) & "<Data>" & Val(rsTmp!收费细目ID) & "<Data>" & FormatEx(NVL(rsTmp!单量), 5) & "<Data>" & rsTmp!脚注 & "<Data>" & rsTmp!计算单位
                    rsTmp.MoveNext
                Next
            End If
            Get用药配方 = strTmp
            .TextMatrix(lngRow, col_配方数据) = strTmp
            .Cell(flexcpData, lngRow, col_配方数据) = strTmp
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdAdviceOK_Click()
    '提取病人历史医嘱记录
    Dim strSQL As String
    Dim strType As String
    Dim strTime As String
    Dim str来源 As String
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, i As Long
    
    On Error GoTo errH
    strType = "(a.诊疗类别 In (" & IIF(chkXY.value = 1, "'5', '6'", "") & IIF(chkZY.value = 1, IIF(chkXY.value = 1, ",", "") & "'7'", "") & ")" & IIF(chkZY.value = 1, "Or a.诊疗类别 = 'E' And (c.操作类型 = '3'))", ")")
    strTime = " And A.开始执行时间 between [3] and [4]"
    If chkMZ.value = 1 And chkZY.value = 1 Then
        str来源 = ""
    Else
        str来源 = IIF(chkMZ.value = 1, " And a.主页id is null", " And A.挂号单 is null")
    End If
    strSQL = "Select a.Id, a.相关id As 组号, a.诊疗类别 As 药品类别, a.医嘱内容 As 用药内容, a.医生嘱托 As 医生嘱托, a.诊疗项目id, a.收费细目id, a.天数, a.开始执行时间 As 开始时间," & vbNewLine & _
            "       a.执行终止时间 As 终止时间, decode(a.病人来源,1,a.总给予量/e.门诊包装,2,a.总给予量/e.住院包装,a.总给予量) as 总给予量, a.单次用量, a.执行频次, a.频率次数, a.频率间隔, a.间隔单位, b.诊疗项目id As 用药id, c.计算单位, b.医嘱内容 As 用法, d.名称 As 中药用法,E.住院单位" & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱记录 B, 诊疗项目目录 C, 诊疗项目目录 D,药品规格 E" & vbNewLine & _
            "Where a.相关id = b.Id And a.诊疗项目id = c.Id And A.收费细目id=E.药品id(+) And b.诊疗项目id = d.Id " & vbNewLine & _
            " And a.病人id = [1] And (nvl(a.主页id,0) <> [2])" & vbNewLine & _
            " And " & strType & strTime & str来源 & vbNewLine & _
            " and nvl(a.医嘱状态,0)<>4" & vbNewLine & _
            "Order By a.病人id,a.主页id,a.挂号单,a.序号,a.开始执行时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, CDate(dtpStartTime.value), CDate(dtpStopTime.value))
    With vsAdvice
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             If .TextMatrix(.Rows - 1, col_用药内容) = "" And Val(.TextMatrix(.Rows - 1, COL_ID)) = 0 Then .Rows = .Rows - 1
             For i = 1 To rsTmp.RecordCount
                If (rsTmp!药品类别 & "" = "7" Or rsTmp!药品类别 & "" = "E") And Val(.Cell(flexcpData, .Rows - 1, col_组号)) = Val(rsTmp!组号 & "") Then
                    If rsTmp!药品类别 & "" = "7" Then
                        .TextMatrix(.Rows - 1, col_配方数据) = .TextMatrix(.Rows - 1, col_配方数据) & "[配方数据]" & rsTmp!用药内容 & "<Data>" & Val(rsTmp!诊疗项目ID & "") & "<Data>" & Val(rsTmp!收费细目ID & "") & "<Data>" & FormatEx(NVL(rsTmp!单次用量), 5) & "<Data>" & rsTmp!医生嘱托 & "<Data>" & rsTmp!计算单位
                    Else
                        .TextMatrix(.Rows - 1, col_煎法id) = Val(rsTmp!诊疗项目ID & "")
                    End If
                Else
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                        
                    '隐藏列
                    .TextMatrix(lngRow, COL_病人ID) = mlng病人ID
                    .TextMatrix(lngRow, COL_主页ID) = mlng主页ID
                    .TextMatrix(lngRow, COL_用药来源) = 1
                    .TextMatrix(lngRow, COL_诊疗项目ID) = Val(rsTmp!诊疗项目ID & "")
                    .TextMatrix(lngRow, COL_收费细目ID) = Val(rsTmp!收费细目ID & "")
                    .TextMatrix(lngRow, COL_频率间隔) = Val(rsTmp!频率间隔 & "")
                    .TextMatrix(lngRow, COL_间隔单位) = rsTmp!间隔单位 & ""
                    .TextMatrix(lngRow, COL_用法id) = Val(rsTmp!用药id & "")
                    .TextMatrix(lngRow, COL_终止时间) = Format(rsTmp!终止时间 & "", "yyyy-mm-dd hh:mm")
                    .TextMatrix(lngRow, COL_开始时间) = Format(rsTmp!开始时间 & "", "yyyy-mm-dd hh:mm")
                    .TextMatrix(lngRow, col_药品类别) = decode(rsTmp!药品类别 & "", "5", "西成药", "6", "中成药", "中草药")
                    .TextMatrix(lngRow, col_用药内容) = IIF(.TextMatrix(lngRow, col_药品类别) = "中草药", rsTmp!用法 & "", rsTmp!用药内容 & "")
                    .TextMatrix(lngRow, COL_用法) = IIF(.TextMatrix(lngRow, col_药品类别) = "中草药", rsTmp!中药用法 & "", rsTmp!用法 & "")
                    .TextMatrix(lngRow, COL_单次用量) = IIF(.TextMatrix(lngRow, col_药品类别) = "中草药", "", FormatEx(NVL(rsTmp!单次用量), 5))
                    .TextMatrix(lngRow, COL_总给予量) = FormatEx(NVL(rsTmp!总给予量), 5)
                    .TextMatrix(lngRow, COL_单量单位) = IIF(.TextMatrix(lngRow, col_药品类别) = "中草药", "", rsTmp!计算单位 & "")
                    .TextMatrix(lngRow, COL_总量单位) = IIF(.TextMatrix(lngRow, col_药品类别) = "中草药", "付", rsTmp!住院单位 & "")
                    .Cell(flexcpData, lngRow, col_组号) = Val(rsTmp!组号 & "")
                    
                    If .TextMatrix(lngRow, col_药品类别) = "中草药" Then
                        .TextMatrix(lngRow, col_组号) = ""
                    Else
                        If .Cell(flexcpData, lngRow, col_组号) = .Cell(flexcpData, lngRow - 1, col_组号) And .Cell(flexcpData, lngRow, col_组号) <> "" Then
                            If .TextMatrix(lngRow - 1, col_组号) = "" Then .TextMatrix(lngRow - 1, col_组号) = -(lngRow - 1)
                            .TextMatrix(lngRow, col_组号) = .TextMatrix(lngRow - 1, col_组号)
                        End If
                    End If

                    If rsTmp!天数 & "" = "" Then
                        If rsTmp!终止时间 & "" <> "" And rsTmp!开始时间 & "" <> "" Then
                            .TextMatrix(lngRow, COL_天数) = FormatEx(NVL(DateDiff("d", CDate(rsTmp!开始时间 & ""), CDate(rsTmp!终止时间 & ""))), 5)
                        End If
                    Else
                        .TextMatrix(lngRow, COL_天数) = FormatEx(NVL(rsTmp!天数), 5)
                    End If
                    .TextMatrix(lngRow, COL_执行频次) = rsTmp!执行频次 & ""
                    
                    If rsTmp!药品类别 & "" = "7" Then
                        .TextMatrix(lngRow, col_配方数据) = .TextMatrix(lngRow, col_配方数据) & "[配方数据]" & rsTmp!用药内容 & "<Data>" & Val(rsTmp!诊疗项目ID & "") & "<Data>" & Val(rsTmp!收费细目ID & "") & "<Data>" & FormatEx(NVL(rsTmp!单次用量), 5) & "<Data>" & rsTmp!医生嘱托 & "<Data>" & rsTmp!计算单位
                    ElseIf rsTmp!药品类别 & "" = "E" Then
                        .TextMatrix(lngRow, col_煎法id) = Val(rsTmp!诊疗项目ID & "")
                    End If
                    
                    '缓存数据
                    .Cell(flexcpData, lngRow, COL_执行频次) = .TextMatrix(lngRow, COL_执行频次)
                    .Cell(flexcpData, lngRow, COL_用法) = .TextMatrix(lngRow, COL_用法)
                    .Cell(flexcpData, lngRow, col_用药内容) = .TextMatrix(lngRow, col_用药内容)
                    .Cell(flexcpData, lngRow, col_药品类别) = decode(.TextMatrix(lngRow, col_药品类别), "西成药", "5", "中成药", "6", "中草药", "8")
                    .Cell(flexcpBackColor, lngRow, 0, lngRow, 0) = Red_COLOR
                End If
                rsTmp.MoveNext
             Next
             .Redraw = flexRDDirect
             .Tag = "1"
        End If
        Call SetTag一并给药
        .Cell(flexcpBackColor, .FixedRows, col_药品类别, .Rows - 1, col_药品类别) = GRD_UNEDITCELL_COLOR      '灰蓝色
        .Cell(flexcpBackColor, .FixedRows, COL_单量单位, .Rows - 1, COL_单量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
        .Cell(flexcpBackColor, .FixedRows, COL_总量单位, .Rows - 1, COL_单量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
    End With
    picAdviceFilter.Visible = False
    picAdviceFilter.Enabled = picAdviceFilter.Visible
    vsAdvice.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function get中药配方(lng项目id As Long) As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, i As Long
    
    On Error GoTo errH
    '输入了配方项目
    strSQL = "Select A.ID,A.名称,b.收费细目id as 药品id,A.计算单位,B.单次用量,B.医生嘱托,C.规格" & _
        " From 诊疗项目目录 A,诊疗项目组合 B,收费项目目录 C" & _
        " Where A.ID=B.诊疗项目ID And B.诊疗组合ID=[1] And c.Id(+) = b.收费细目id" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) And A.服务对象 IN(1,2,3) Order By B.序号"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id)
    If rsTmp.EOF Then
        MsgBox "该中药配方当前无有效的配方组成，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Function
    Else
        For i = 1 To rsTmp.RecordCount
             strTmp = strTmp & "[配方数据]" & rsTmp!名称 & "<Data>" & Val(rsTmp!ID & "") & "<Data>" & Val(rsTmp!药品ID & "") & "<Data>" & FormatEx(NVL(rsTmp!单次用量), 5) & "<Data>" & rsTmp!医生嘱托 & "<Data>" & rsTmp!计算单位
            rsTmp.MoveNext
        Next
        get中药配方 = strTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cmdAdviceQuit_Click()
    If picAdviceFilter.Visible = False Then Exit Sub
    picAdviceFilter.Visible = False
    picAdviceFilter.Enabled = picAdviceFilter.Visible
    vsAdvice.SetFocus
End Sub

Private Sub Form_Activate()
    vsAdvice.SetFocus
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picFilter.Top = 520
    picFilter.Left = 10
    picFilter.Width = Me.ScaleWidth
    
    picTime.Left = picFilter.Width - picTime.Width - 50
    lblTime.Left = picTime.Left - lblTime.Width - 50
    
    vsAdvice.Left = 0
    vsAdvice.Top = picFilter.Top + picFilter.Height + 10
    vsAdvice.Width = picFilter.Width
    
    vsAdvice.Height = Me.ScaleHeight - staThis.Height - vsAdvice.Top
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Modify
        Control.Enabled = mlngEditTag = 0
    Case conMenu_Edit_ItemUndo
        Control.Enabled = mlngEditTag = 1
    Case conMenu_Edit_Save
        Control.Enabled = (mlngEditTag = 1 And Val(vsAdvice.Tag) = 1)
        Control.Visible = mlngEditTag = 1
    Case conMenu_Edit_DrugAuto
        Control.Enabled = mlngEditTag = 1
    Case conMenu_Edit_NewItem
        Control.Enabled = mlngEditTag = 1
    Case conMenu_Edit_Delete
        Control.Enabled = mlngEditTag = 1
    Case conMenu_Edit_DrugGrp
        Control.Enabled = mlngEditTag = 1
        If Control.Enabled Then
            Control.Checked = Val(vsAdvice.TextMatrix(vsAdvice.Row, col_组号)) <> 0
        End If
    End Select
    If Control.ID <> conMenu_Edit_Save Then Control.Visible = Control.Enabled
End Sub

Private Sub Form_Load()

    Call InitCommandBar
    Call InitAdviceTable
    
    OptTime(0).value = True

    dtpStartTime.value = DateAdd("m", -3, Now())
    dtpStopTime.value = Now()
    
    '输入匹配
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    '简码匹配方式：0-拼音,1-五笔
    mint简码 = Val(zlDatabase.GetPara("简码方式"))
    mlngEditTag = 0
    mlngLastColor = 0
    
    staThis.Panels(2).Text = "当前模式为：" & IIF(mlngEditTag = 0, "查阅用药清单", "编辑用药清单")
    Call LoadPatiInfo
    vsAdvice.Editable = flexEDNone
    Call LoadDrug
    vsAdvice.Row = vsAdvice.Rows - 1: vsAdvice.Col = COL_开始时间
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    '处理win10教育版 图标显示异常
    imgIco.Top = 45: imgIco.Left = 600: imgIco.Height = 240: imgIco.Width = 240
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mlngEditTag = 1 And Val(vsAdvice.Tag) = 1 Then
        If MsgBox("当前还有未保存的用药记录,确定要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub LoadPatiInfo()
'功能：加载病人信息
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = _
        " Select B.住院号,b.姓名,b.性别,b.年龄,B.出院病床," & _
        " B.住院医师,B.出院科室ID,C.名称 as 科室,B.险类,B.病人性质 " & _
        " From  病案主页 B,部门表 C" & _
        " Where B.出院科室ID=C.ID And b.病人ID=[1] And B.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    lblPati.Caption = "姓名：" & rsTmp!姓名 & "　住院号：" & NVL(rsTmp!住院号) & _
        "　床号：" & NVL(rsTmp!出院病床) & "　科室：" & NVL(rsTmp!科室) & "　年龄：" & NVL(rsTmp!年龄)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    '工具栏----------------------------------------------
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
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlcommfun.GetPubIcons

    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)

    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "编辑清单")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ItemUndo, "取消编辑"): objControl.IconId = 5019
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DrugAuto, "自动提取")
            objControl.IconId = 3587
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DrugGrp, "一并给药")
            objControl.IconId = 3064
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " 退出(&X)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With
End Sub

Private Sub InitAdviceTable()
'功能：初始化医嘱清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
              
    strHead = "ID;病人ID;主页ID;组号;用药来源;诊疗项目ID;收费细目ID;频率间隔;间隔单位;用法ID;煎法ID;终止时间;" & _
                "开始时间,2000,1;药品类别,850,4;用药内容,7000,1;用法,2000,1;单量,850,4;单位,600,4;总量,850,4;单位,600,4;天数,450,4;执行频次,1000,4;备注,1000,1;登记人;登记时间;配方数据;是否修改;频率次数"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionFree
        .FocusRect = flexFocusSolid
        .HighLight = flexHighlightWithFocus
        .BackColorSel = &H404040


        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    '为了支持zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '记录原始列宽用于列选择器
        Next
        .Editable = flexEDKbdMouse
        .WordWrap = True
        .AutoSize col_用药内容
    End With
End Sub



Public Function AdviceCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsAdvice
        If .ColHidden(lngCol) Then Exit Function
        '必须先输入用药内容
        If (lngCol = col_药品类别 Or lngCol = COL_单量单位 Or lngCol = COL_登记时间 Or lngCol = COL_登记人 Or lngCol = COL_总量单位) Then Exit Function
        If .TextMatrix(lngRow, col_用药内容) = "" Then
            If lngCol > col_用药内容 Then Exit Function
        End If
        If lngCol = COL_单次用量 And .TextMatrix(lngRow, col_药品类别) = "中草药" Then Exit Function
    End With
    AdviceCellEditable = True
End Function

Private Sub EnterNextCellAdvice()
    Dim i As Long, j As Long

    With vsAdvice
        '从下一单元开始循环搜索
        If .Row < .FixedRows Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, .FixedRows, col_药品类别, .Rows - 1, col_药品类别) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Cell(flexcpBackColor, .FixedRows, COL_单量单位, .Rows - 1, COL_单量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Cell(flexcpBackColor, .FixedRows, COL_总量单位, .Rows - 1, COL_总量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 0) = Red_COLOR
            .ShowCell .Rows - 1, COL_开始时间
        End If
        For i = .Row To .Rows - 1
            For j = IIF(i = .Row, .Col + 1, COL_开始时间) To COL_备注
                If AdviceCellEditable(i, j) Then Exit For
            Next
            If j <= COL_备注 Then Exit For
        Next
        If i <= .Rows - 1 Then
            Call .Select(i, j)
            .ShowCell .Row, .Col
        ElseIf i = .Rows And j > COL_备注 And .TextMatrix(.Rows - 1, col_用药内容) <> "" Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, .FixedRows, col_药品类别, .Rows - 1, col_药品类别) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Cell(flexcpBackColor, .FixedRows, COL_单量单位, .Rows - 1, COL_单量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Cell(flexcpBackColor, .FixedRows, COL_总量单位, .Rows - 1, COL_总量单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 0) = Red_COLOR
            .ShowCell .Rows - 1, COL_开始时间
        End If
    End With
End Sub


Private Sub OptTime_Click(Index As Integer)
    LoadDrug
End Sub


Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If NewRow = -1 Or NewCol = -1 Then Exit Sub
    If vsAdvice.Editable = flexEDNone Then Exit Sub
    
    If OldRow > 0 And OldCol > 0 Then
        If OldRow < vsAdvice.Rows Then
            vsAdvice.Cell(flexcpBackColor, OldRow, OldCol, OldRow, OldCol) = mlngLastColor
        End If
    End If
    If NewRow > 0 Then
        mlngLastColor = vsAdvice.Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol)
        vsAdvice.Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol) = &HC0FFC0
    End If
    
    If (Not AdviceCellEditable(NewRow, NewCol)) Then
        vsAdvice.ComboList = ""
        vsAdvice.FocusRect = flexFocusLight
    Else
        vsAdvice.FocusRect = flexFocusSolid
        Select Case NewCol
            Case col_用药内容, COL_用法, COL_执行频次
                If NewCol = col_用药内容 Then Call Get用药配方(NewRow)
                vsAdvice.ComboList = "..."
            Case Else
                vsAdvice.ComboList = ""
        End Select
    End If
End Sub


Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsAdvice
        .EditSelStart = 0
        .EditSelLength = zlcommfun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not AdviceCellEditable(Row, Col) Then
        Cancel = True
    End If
End Sub


Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    If vsAdvice.Editable = flexEDNone Then Exit Sub
    With vsAdvice
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            mblnReturn = True
            Call EnterNextCellAdvice
        Else
            If .ComboList = "..." Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsAdvice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub



Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsAdvice
        If Not KeyAscii = vbKeyReturn Then
            If Col = COL_开始时间 Or Col = COL_登记时间 Then
                If InStr("0123456789-" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            ElseIf Col = COL_单次用量 Or Col = COL_总给予量 Or Col = COL_天数 Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            ElseIf Col = col_用药内容 And vsAdvice.TextMatrix(Row, col_药品类别) = "中草药" Then
                KeyAscii = 0
            End If
            mblnReturn = False
        Else
            mblnReturn = True
        End If
    End With
End Sub



Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsAdvice.Editable = flexEDNone Then Exit Sub
    With vsAdvice
        If KeyCode = vbKeyF4 Then
            If .ComboList = "..." Then
                Call zlcommfun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Col = col_用药内容 Then
                DeteleRow
            ElseIf .Col = COL_用法 Then
                .TextMatrix(.Row, COL_用法) = ""
                .TextMatrix(.Row, COL_用法id) = ""
                .Cell(flexcpData, .Row, COL_用法) = ""
                .TextMatrix(.Row, col_是否修改) = "1"
            ElseIf .Col = COL_执行频次 Then
                   .Cell(flexcpData, .Row, COL_执行频次) = ""
                    .TextMatrix(.Row, COL_频率间隔) = ""
                    .TextMatrix(.Row, COL_频率次数) = ""
                    .TextMatrix(.Row, COL_间隔单位) = ""
                    .TextMatrix(.Row, col_是否修改) = "1"
            End If
            .Tag = "1"
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsAdvice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strSeek As String, int类型 As Integer
    Dim lng项目id As Long
    Dim bytOK As Byte
    Dim strData As String, lng煎法ID As Long
    Dim blnAuto As Boolean, lngUpRow As Long
    
    On Error GoTo errH
   With vsAdvice
        Select Case Col
            Case col_用药内容
                If .TextMatrix(Row, col_药品类别) = "中草药" Then
                    strData = .TextMatrix(Row, col_配方数据)
                    lng煎法ID = .TextMatrix(Row, col_煎法id)
                    If frmDrugListEditEx.ShowEdit(Me, strData, lng煎法ID) Then
                        If strData = "" Then
                            .TextMatrix(Row, col_用药内容) = ""
                            .Cell(flexcpData, Row, col_用药内容) = ""
                            .TextMatrix(Row, COL_诊疗项目ID) = ""
                            .TextMatrix(Row, COL_收费细目ID) = ""
                            .TextMatrix(Row, col_药品类别) = ""
                            .Cell(flexcpData, Row, col_药品类别) = ""
                            .TextMatrix(Row, COL_单量单位) = ""
                            .TextMatrix(Row, COL_总量单位) = ""
                            .TextMatrix(Row, COL_用法) = ""
                            .Cell(flexcpData, Row, COL_用法) = ""
                            .TextMatrix(Row, COL_用法id) = ""
                            .TextMatrix(Row, COL_执行频次) = ""
                            .TextMatrix(Row, COL_频率次数) = ""
                            .Cell(flexcpData, Row, COL_执行频次) = ""
                            .TextMatrix(Row, COL_频率间隔) = ""
                            .TextMatrix(Row, COL_间隔单位) = ""
                            .TextMatrix(Row, col_配方数据) = ""
                            .TextMatrix(Row, col_煎法id) = ""
                            .TextMatrix(Row, COL_单次用量) = ""
                            Exit Sub
                        Else
                            .TextMatrix(Row, col_用药内容) = Set中药配方(Row, strData, lng煎法ID)
                            .Cell(flexcpData, Row, col_用药内容) = .TextMatrix(Row, col_用药内容)
                        End If
                        .Tag = "1"
                        .TextMatrix(Row, col_是否修改) = "1"
                        .Cell(flexcpBackColor, Row, 0, .Row, 0) = Red_COLOR
                    Else
                        Exit Sub
                    End If
                Else
                    Set rsTmp = frmDrugSelect.ShowSelect(Me, bytOK)
                    If bytOK = 1 And (Not rsTmp Is Nothing) Then
                        If rsTmp!类别 & "" = "配方" Or rsTmp!类别 & "" = "中草药" Then
                            If Val(.TextMatrix(Row, col_组号)) <> 0 Then
                                MsgBox "该组一并给药的药品必须都为西成药或中成药。", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If rsTmp!类别 & "" = "配方" Then
                                strData = get中药配方(Val(rsTmp!诊疗项目ID & ""))
                            Else
                                strData = "[配方数据]" & rsTmp!名称 & "<Data>" & Val(rsTmp!诊疗项目ID & "") & "<Data>" & Val(rsTmp!收费细目ID & "") & "<Data>0<Data><Data>" & rsTmp!计算单位
                            End If
                            
                            If frmDrugListEditEx.ShowEdit(Me, strData, lng煎法ID) Then
                                If strData = "" Then
                                    Exit Sub
                                Else
                                    .TextMatrix(Row, col_用药内容) = Set中药配方(Row, strData, lng煎法ID)
                                    .Cell(flexcpData, Row, col_用药内容) = .TextMatrix(Row, col_用药内容)
                                    .TextMatrix(Row, COL_诊疗项目ID) = ""
                                    .TextMatrix(Row, COL_收费细目ID) = ""
                                End If
                            Else
                                Exit Sub
                            End If
                        Else
                        
                            '新加入行自动一并给药
                            If Val(.TextMatrix(Row, col_组号)) = 0 And Val(.TextMatrix(GetUpRow(Row), col_组号)) <> 0 And .Cell(flexcpData, Row, col_用药内容) = "" And Val(.TextMatrix(Row, COL_ID)) = 0 Then
                                .TextMatrix(Row, col_组号) = Val(.TextMatrix(GetUpRow(Row), col_组号))
                                Call SetTag一并给药(Val(.TextMatrix(Row, col_组号)))
                                blnAuto = True
                            End If
                        
                            .TextMatrix(Row, col_用药内容) = rsTmp!名称 & IIF(rsTmp!规格 & "" = "", "", "(" & rsTmp!规格 & ")")
                            .Cell(flexcpData, Row, col_用药内容) = .TextMatrix(Row, col_用药内容)
                            .TextMatrix(Row, COL_诊疗项目ID) = Val(rsTmp!诊疗项目ID & "")
                            .TextMatrix(Row, COL_收费细目ID) = Val(rsTmp!收费细目ID & "")
                            .TextMatrix(Row, col_煎法id) = ""
                            .TextMatrix(Row, col_配方数据) = ""
                        End If

                        .TextMatrix(Row, col_药品类别) = IIF(rsTmp!类别 & "" = "配方", "中草药", rsTmp!类别 & "")
                        .Cell(flexcpData, Row, col_药品类别) = decode(.TextMatrix(Row, col_药品类别), "西成药", "5", "中成药", "6", "中草药", "8")
                        .TextMatrix(Row, COL_单量单位) = IIF(.TextMatrix(Row, col_药品类别) = "中草药", "", rsTmp!计算单位 & "")
                        .TextMatrix(Row, COL_总量单位) = IIF(.TextMatrix(Row, col_药品类别) = "中草药", "付", rsTmp!总量单位 & "")
                        .TextMatrix(Row, COL_单次用量) = IIF(.TextMatrix(Row, col_药品类别) = "中草药", "", .TextMatrix(Row, COL_单次用量))
                        If Val(.TextMatrix(Row, col_组号)) = 0 Then
                            .TextMatrix(Row, COL_用法) = ""
                            .Cell(flexcpData, Row, COL_用法) = ""
                            .TextMatrix(Row, COL_用法id) = ""
                            .TextMatrix(Row, COL_执行频次) = ""
                            .TextMatrix(.Row, COL_频率次数) = ""
                            .Cell(flexcpData, Row, COL_执行频次) = ""
                            .TextMatrix(Row, COL_频率间隔) = ""
                            .TextMatrix(Row, COL_间隔单位) = ""
                            .TextMatrix(Row, COL_天数) = ""
                        ElseIf blnAuto Then
                            '自动一并给药是同步数据
                            lngUpRow = GetUpRow(Row)
                            .TextMatrix(Row, COL_开始时间) = .TextMatrix(lngUpRow, COL_开始时间)
                            .TextMatrix(Row, COL_用法) = .TextMatrix(lngUpRow, COL_用法)
                            .Cell(flexcpData, Row, COL_用法) = .Cell(flexcpData, lngUpRow, COL_用法)
                            .TextMatrix(Row, COL_用法id) = .TextMatrix(lngUpRow, COL_用法id)
                            .TextMatrix(Row, COL_执行频次) = .TextMatrix(lngUpRow, COL_执行频次)
                            .TextMatrix(.Row, COL_频率次数) = .TextMatrix(lngUpRow, COL_频率次数)
                            .Cell(flexcpData, Row, COL_执行频次) = .Cell(flexcpData, lngUpRow, COL_执行频次)
                            .TextMatrix(Row, COL_频率间隔) = .TextMatrix(lngUpRow, COL_频率间隔)
                            .TextMatrix(Row, COL_间隔单位) = .TextMatrix(lngUpRow, COL_间隔单位)
                            .TextMatrix(Row, COL_天数) = .TextMatrix(lngUpRow, COL_天数)
                        End If
                        
                        .Tag = "1"
                        .TextMatrix(Row, col_是否修改) = "1"
                        .Cell(flexcpBackColor, Row, 0, .Row, 0) = Red_COLOR
                    End If
                End If
            Case COL_用法
                int类型 = IIF(.TextMatrix(Row, col_药品类别) = "中草药", 4, 2)
                
                lng项目id = Val(.TextMatrix(.Row, COL_诊疗项目ID))
                If int类型 = 2 Then '只取有效范围的给药途径(无设置或仅一个时可任选)
                    If Val(.TextMatrix(.Row, COL_收费细目ID)) = 0 Then
                        strSQL = " And (A.ID IN(Select 用法ID From 诊疗用法用量 Where 项目ID=[2] And 性质>0)" & _
                            " Or (Select Count(A.用法ID) From 诊疗用法用量 A,诊疗项目目录 B" & _
                            " Where A.用法ID=B.ID And B.服务对象 IN(1,2,3) And A.项目ID=[2] And A.性质>0)<=1)"
                    Else
                        lng项目id = Val(.TextMatrix(.Row, COL_收费细目ID))
                        strSQL = " And (A.ID IN(Select 用法ID From 药品用法用量 Where 药品ID=[2] And 性质=1)" & _
                            " Or (Select Count(A.用法ID) From 药品用法用量 A,诊疗项目目录 B" & _
                            " Where A.用法ID=B.ID And B.服务对象 IN(1,2,3) And A.药品ID=[2] And A.性质=1)<=1)"
                    End If
                End If
                strSQL = "Select Distinct A.ID,A.编码,A.名称,C.名称 as 分类,A.执行分类 as 执行分类ID" & _
                    " From 诊疗项目目录 A,诊疗分类目录 C" & _
                    " Where A.分类ID=C.ID(+) And A.类别='E' And A.操作类型=[1] And A.服务对象 IN(1,2,3)" & strSQL & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " Order by A.编码"
                 vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "给药途径", False, strSeek, "", False, False, True, _
                    vPoint.X, vPoint.Y, IIF(.RowHeight(Row) < .RowHeightMin, .RowHeightMin, .RowHeight(Row)), blnCancel, False, True, CStr(int类型), lng项目id)
                    
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "没有可用的给药途径，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                    End If
                    Exit Sub
                Else
                    .TextMatrix(Row, COL_用法) = rsTmp!名称 & ""
                    .TextMatrix(Row, COL_用法id) = Val(rsTmp!ID & "")
                    .Cell(flexcpData, Row, COL_用法) = .TextMatrix(Row, COL_用法)
                    .TextMatrix(Row, col_是否修改) = "1"
                    .Cell(flexcpBackColor, Row, 0, .Row, 0) = Red_COLOR
                    If Val(.TextMatrix(Row, col_组号)) <> 0 Then Set同步一并给药 (Row)
                    .Tag = "1"
                End If
            Case COL_执行频次
                strSQL = _
                    " Select Rownum as ID,A.编码,A.名称,A.简码," & _
                    " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位,A.适用范围 as 范围ID" & _
                    " From 诊疗频率项目 A" & _
                    " Where A.适用范围<>[1]" & _
                    " Order by A.适用范围,A.编码"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "执行频次", False, strSeek, "", False, False, True, _
                    vPoint.X, vPoint.Y, IIF(.RowHeight(Row) < .RowHeightMin, .RowHeightMin, .RowHeight(Row)), blnCancel, False, True, IIF(.TextMatrix(.Row, col_药品类别) = "中草药", "1", "2"))
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "没有可用的诊疗频率项目，请先到医嘱频率管理中设置。", vbInformation, gstrSysName
                    End If
                    Exit Sub
                Else
                    .TextMatrix(Row, COL_执行频次) = rsTmp!名称 & ""
                    .Cell(flexcpData, Row, COL_执行频次) = .TextMatrix(Row, COL_执行频次)
                    .TextMatrix(Row, COL_频率间隔) = rsTmp!频率间隔 & ""
                    .TextMatrix(Row, COL_间隔单位) = rsTmp!间隔单位 & ""
                    .TextMatrix(.Row, COL_频率次数) = rsTmp!频率次数 & ""
                    .Tag = "1"
                    .TextMatrix(Row, col_是否修改) = "1"
                    If Val(.TextMatrix(Row, col_组号)) <> 0 Then Set同步一并给药 (Row)
                    .Cell(flexcpBackColor, Row, 0, .Row, 0) = Red_COLOR
                End If
        End Select
   End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strInput As String
    
    With vsAdvice
        Select Case Col
            Case COL_开始时间
                strInput = Format(zlStr.FullDate(.TextMatrix(Row, Col)), "yyyy-mm-dd hh:mm")
                If Not IsDate(strInput) Then
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Else
                    .TextMatrix(Row, Col) = strInput
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    If Val(.TextMatrix(Row, col_组号)) <> 0 Then Set同步一并给药 (Row)
                End If
            Case COL_天数
                If Val(.TextMatrix(Row, col_组号)) <> 0 Then Set同步一并给药 (Row)
        End Select
    End With
End Sub

Private Sub vsAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'vsAdvice_ValidateEdit事件
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strSeek As String, int类型 As Integer
    Dim lng项目id As Long, strLike As String
    Dim strInput As String
    Dim lngMax As Long
    Dim strData As String
    Dim lng煎法ID As Long
    Dim lngUpRow As Long, blnAuto As Boolean

    On Error GoTo errH
   With vsAdvice
        strLike = mstrLike
        If Len(.EditText) < 2 Then strLike = "" '优化
        Select Case Col
            Case col_用药内容
                If .EditText = "" Then
                    .EditText = .Cell(flexcpData, Row, col_用药内容)
                    If mblnReturn Then Call EnterNextCellAdvice
                ElseIf .EditText = .Cell(flexcpData, Row, col_用药内容) Then
                    If mblnReturn Then Call EnterNextCellAdvice
                Else
                    strInput = " And (A.编码 Like [1] And E.码类=[3]" & _
                        " Or E.名称 Like [2] And E.码类=[3] Or E.简码 Like [2] And E.码类 IN([3],3))"
                
                    If IsNumeric(.EditText) Then
                        '1X.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                        If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And (A.编码 Like [1] And E.码类=[3] Or E.简码 Like [2] And E.码类=3)"
                    ElseIf zlcommfun.IsCharAlpha(.EditText) Then
                        'X1.输入全是字母时只匹配简码
                        If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And E.简码 Like [2] And E.码类=[3]"
                    ElseIf zlcommfun.IsCharChinese(.EditText) Then
                        '包含汉字,则只匹配名称a
                        strInput = " And E.名称 Like [2] And E.码类=[3]"
                    End If
                    
                    strSQL = "Select distinct a.Id, b.Id As 收费细目id,decode(a.类别,'5','西成药','6','中成药','7','中草药','8','配方') as 类别, a.名称, b.规格, a.计算单位, d.药品剂型,C.住院单位 as 总量单位" & _
                    " From 诊疗项目目录 A, 收费项目目录 B, 药品规格 C, 药品特性 D,诊疗项目别名 E " & _
                    " Where c.药品id= b.Id(+) And a.Id =c.药名id(+) And c.药名id = d.药名id(+) And A.ID=E.诊疗项目ID(+) And a.类别 in ('5','6','7','8') and (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & strInput
                    
                    vPoint = zlControl.GetCoordPos(vsAdvice.hwnd, vsAdvice.CellLeft, vsAdvice.CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药品目录", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, IIF(.RowHeight(Row) < .RowHeightMin, .RowHeightMin, .RowHeight(Row)), blnCancel, False, True, UCase(.EditText) & "%", strLike & UCase(.EditText) & "%", mint简码 + 1)
                    
                    If rsTmp Is Nothing Then
                        If Not blnCancel Then
                            MsgBox "未找到可用的药品，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            .EditText = .Cell(flexcpData, Row, col_用药内容)
                        End If
                        Exit Sub
                    Else
                        If rsTmp!类别 & "" = "配方" Or rsTmp!类别 & "" = "中草药" Then
                            If Val(.TextMatrix(Row, col_组号)) <> 0 Then
                                MsgBox "该组一并给药的药品必须都为西成药或中成药。", vbInformation, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            If rsTmp!类别 & "" = "配方" Then
                                strData = get中药配方(Val(rsTmp!ID & ""))
                            Else
                                strData = "[配方数据]" & rsTmp!名称 & "<Data>" & Val(rsTmp!ID & "") & "<Data>" & Val(rsTmp!收费细目ID & "") & "<Data>0<Data><Data>" & rsTmp!计算单位
                            End If
                            If frmDrugListEditEx.ShowEdit(Me, strData, lng煎法ID) Then
                                If strData = "" Then
                                    .EditText = .Cell(flexcpData, Row, col_用药内容)
                                    Exit Sub
                                Else
                                    .EditText = Set中药配方(Row, strData, lng煎法ID)
                                    .TextMatrix(Row, col_用药内容) = .EditText
                                    .Cell(flexcpData, Row, col_用药内容) = .EditText
                                    .TextMatrix(Row, COL_诊疗项目ID) = ""
                                    .TextMatrix(Row, COL_收费细目ID) = ""
                                End If
                            Else
                                .EditText = .Cell(flexcpData, Row, col_用药内容)
                                Exit Sub
                            End If
                        Else
                            '新加入行自动一并给药
                            If Val(.TextMatrix(Row, col_组号)) = 0 And Val(.TextMatrix(GetUpRow(Row), col_组号)) <> 0 And .Cell(flexcpData, Row, col_用药内容) = "" Then
                                .TextMatrix(Row, col_组号) = Val(.TextMatrix(GetUpRow(Row), col_组号))
                                Call SetTag一并给药(Val(.TextMatrix(Row, col_组号)))
                                blnAuto = True
                            End If
                            .EditText = rsTmp!名称 & IIF(rsTmp!规格 & "" = "", "", "(" & rsTmp!规格 & ")")
                            .TextMatrix(Row, col_用药内容) = rsTmp!名称 & IIF(rsTmp!规格 & "" = "", "", "(" & rsTmp!规格 & ")")
                            .Cell(flexcpData, Row, col_用药内容) = .TextMatrix(Row, col_用药内容)
                            .TextMatrix(Row, COL_诊疗项目ID) = Val(rsTmp!ID & "")
                            .TextMatrix(Row, COL_收费细目ID) = Val(rsTmp!收费细目ID & "")
                            .TextMatrix(Row, col_煎法id) = ""
                            .TextMatrix(Row, col_配方数据) = ""
                        End If
                        
                        .TextMatrix(Row, col_药品类别) = IIF(rsTmp!类别 & "" = "配方", "中草药", rsTmp!类别 & "")
                        .Cell(flexcpData, Row, col_药品类别) = decode(.TextMatrix(Row, col_药品类别), "西成药", "5", "中成药", "6", "中草药", "8")
                        .TextMatrix(Row, COL_单量单位) = IIF(.TextMatrix(Row, col_药品类别) = "中草药", "", rsTmp!计算单位 & "")
                        .TextMatrix(Row, COL_总量单位) = IIF(.TextMatrix(Row, col_药品类别) = "中草药", "付", rsTmp!总量单位 & "")
                        .TextMatrix(Row, COL_单次用量) = IIF(.TextMatrix(Row, col_药品类别) = "中草药", "", .TextMatrix(Row, COL_单次用量))
                        
                        If Val(.TextMatrix(Row, col_组号)) = 0 Then
                            .TextMatrix(Row, COL_用法) = ""
                            .Cell(flexcpData, Row, COL_用法) = ""
                            .TextMatrix(Row, COL_用法id) = ""
                            .TextMatrix(Row, COL_执行频次) = ""
                            .TextMatrix(.Row, COL_频率次数) = ""
                            .Cell(flexcpData, Row, COL_执行频次) = ""
                            .TextMatrix(Row, COL_频率间隔) = ""
                            .TextMatrix(Row, COL_间隔单位) = ""
                            .TextMatrix(Row, COL_天数) = ""
                        ElseIf blnAuto Then
                            '自动一并给药是同步数据
                            lngUpRow = GetUpRow(Row)
                            .TextMatrix(Row, COL_开始时间) = .TextMatrix(lngUpRow, COL_开始时间)
                            .TextMatrix(Row, COL_用法) = .TextMatrix(lngUpRow, COL_用法)
                            .Cell(flexcpData, Row, COL_用法) = .Cell(flexcpData, lngUpRow, COL_用法)
                            .TextMatrix(Row, COL_用法id) = .TextMatrix(lngUpRow, COL_用法id)
                            .TextMatrix(Row, COL_执行频次) = .TextMatrix(lngUpRow, COL_执行频次)
                            .TextMatrix(.Row, COL_频率次数) = .TextMatrix(lngUpRow, COL_频率次数)
                            .Cell(flexcpData, Row, COL_执行频次) = .Cell(flexcpData, lngUpRow, COL_执行频次)
                            .TextMatrix(Row, COL_频率间隔) = .TextMatrix(lngUpRow, COL_频率间隔)
                            .TextMatrix(Row, COL_间隔单位) = .TextMatrix(lngUpRow, COL_间隔单位)
                            .TextMatrix(Row, COL_天数) = .TextMatrix(lngUpRow, COL_天数)
                        End If
                    End If
                End If
                lngMax = 1000
            Case COL_用法
                If .EditText = "" Then
                    .EditText = .Cell(flexcpData, Row, COL_用法)
                    If mblnReturn Then Call EnterNextCellAdvice
                ElseIf .EditText = .Cell(flexcpData, Row, COL_用法) Then
                    If mblnReturn Then Call EnterNextCellAdvice
                Else
                    int类型 = IIF(vsAdvice.TextMatrix(Row, col_药品类别) = "中草药", 4, 2)
                    lng项目id = Val(.TextMatrix(.Row, COL_诊疗项目ID))
                    If int类型 = 2 Then '只取有效范围的给药途径(无设置或仅一个时可任选)
                        If Val(.TextMatrix(.Row, COL_收费细目ID)) = 0 Then
                            strSQL = " And (A.ID IN(Select 用法ID From 诊疗用法用量 Where 项目ID=[4] And 性质>0)" & _
                                " Or (Select Count(A.用法ID) From 诊疗用法用量 A,诊疗项目目录 B" & _
                                " Where A.用法ID=B.ID And B.服务对象 IN(1,2,3) And A.项目ID=[4] And A.性质>0)<=1)"
                        Else
                            lng项目id = Val(.TextMatrix(.Row, COL_收费细目ID))
                            strSQL = " And (A.ID IN(Select 用法ID From 药品用法用量 Where 药品ID=[4] And 性质=1)" & _
                                " Or (Select Count(A.用法ID) From 药品用法用量 A,诊疗项目目录 B" & _
                                " Where A.用法ID=B.ID And B.服务对象 IN(1,3) And A.药品ID=[4] And A.性质=1)<=1)"
                        End If
                    End If
         
                    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.执行分类 as 执行分类ID" & _
                        " From 诊疗项目目录 A,诊疗项目别名 B" & _
                        " Where A.ID=B.诊疗项目ID" & _
                        " And A.类别='E' And A.操作类型=[3] And A.服务对象 IN(1,2,3)" & strSQL & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2])" & _
                        decode(mint简码, 0, " And B.码类 IN([5],3)", 1, " And B.码类 IN([5],3)", "") & _
                        " Order by A.编码"
                     vPoint = zlControl.GetCoordPos(vsAdvice.hwnd, vsAdvice.CellLeft, vsAdvice.CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "给药途径", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, IIF(.RowHeight(Row) < .RowHeightMin, .RowHeightMin, .RowHeight(Row)), blnCancel, False, True, UCase(.EditText) & "%", strLike & UCase(.EditText) & "%", CStr(int类型), lng项目id, mint简码 + 1)
                    If rsTmp Is Nothing Then
                        If Not blnCancel Then
                            MsgBox "未找到可用的给药途径，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            .EditText = .Cell(flexcpData, Row, COL_用法)
                        End If
                        Exit Sub
                    Else
                        .EditText = rsTmp!名称 & ""
                        .TextMatrix(Row, COL_用法) = rsTmp!名称 & ""
                        .Cell(flexcpData, Row, COL_用法) = .TextMatrix(Row, COL_用法)
                        .TextMatrix(Row, COL_用法id) = Val(rsTmp!ID & "")
                        If Val(.TextMatrix(Row, col_组号)) <> 0 Then Set同步一并给药 (Row)
                    End If
                End If
            Case COL_执行频次
                strSQL = _
                    " Select Rownum as ID,A.编码,A.名称,A.简码," & _
                    " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位,A.适用范围 as 范围ID" & _
                    " From 诊疗频率项目 A" & _
                    " Where A.适用范围<>[1] And (A.编码 Like [2] Or Upper(A.名称) Like [3]" & _
                    " Or Upper(A.简码) Like [3] Or Upper(A.英文名称) Like [3])" & _
                    " Order by A.适用范围,A.编码"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "执行频次", False, strSeek, "", False, False, True, _
                    vPoint.X, vPoint.Y, IIF(.RowHeight(Row) < .RowHeightMin, .RowHeightMin, .RowHeight(Row)), blnCancel, False, True, IIF(.TextMatrix(.Row, col_药品类别) = "中草药", "1", "2"), UCase(.EditText) & "%", strLike & UCase(.EditText) & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "未找到可用的诊疗频率项目，请先到医嘱频率管理中设置。", vbInformation, gstrSysName
                        Cancel = True
                    Else
                        .EditText = .Cell(flexcpData, Row, COL_执行频次)
                    End If
                    Exit Sub
                Else
                    .EditText = rsTmp!名称 & ""
                    .TextMatrix(Row, COL_执行频次) = rsTmp!名称 & ""
                    .Cell(flexcpData, Row, COL_执行频次) = .TextMatrix(Row, COL_执行频次)
                    .TextMatrix(Row, COL_频率间隔) = rsTmp!频率间隔 & ""
                    .TextMatrix(Row, COL_间隔单位) = rsTmp!间隔单位 & ""
                    .TextMatrix(Row, COL_频率次数) = rsTmp!频率次数 & ""
                    If Val(.TextMatrix(Row, col_组号)) <> 0 Then Set同步一并给药 (Row)
                End If
                lngMax = 20
            Case COL_单次用量
                lngMax = 10
            Case COL_总给予量
                lngMax = 10
            Case COL_天数
                lngMax = 10
            Case COL_备注
                lngMax = 1000
            Case COL_开始时间
        End Select
        
        If LenB(StrConv(.EditText, vbFromUnicode)) > lngMax And lngMax <> 0 Then
            MsgBox "不能超过" & lngMax & "个字符的长度。", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
        
        mblnReturn = False
        .TextMatrix(Row, col_是否修改) = "1"
        .Cell(flexcpBackColor, Row, 0, .Row, 0) = Red_COLOR
        .Tag = "1"
   End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Set中药配方(ByVal lngRow As Long, ByVal strData As String, ByVal lng煎法ID As Long) As String
    Dim strTmp As String
    Dim arrTime As Variant, arrTmp As Variant
    Dim i As Long
    With vsAdvice
        .TextMatrix(lngRow, col_煎法id) = lng煎法ID
        .TextMatrix(lngRow, col_配方数据) = strData
        arrTime = Split(strData, "[配方数据]")
        For i = 1 To UBound(arrTime)
            If i = 1 Then strTmp = "中药配方:"
            arrTmp = Split(arrTime(i), "<Data>")
            strTmp = strTmp & arrTmp(0) & " " & FormatEx(NVL(arrTmp(3)), 5) & arrTmp(5) & " " & arrTmp(4) & ","
        Next
        If strTmp <> "" Then strTmp = Left(strTmp, Len(strTmp) - 1)
        Set中药配方 = strTmp
    End With
End Function




Private Sub vsAdvice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    
    With vsAdvice
        If mstrTip <> "" Then
            If .MouseRow = Val(Split(mstrTip, "|")(0)) And .MouseCol = Val(Split(mstrTip, "|")(1)) Then
                strInfo = Split(mstrTip, "|")(2)
            End If
        End If
    End With
    Call zlcommfun.ShowTipInfo(vsAdvice.hwnd, strInfo, True, True)
End Sub

Private Sub SetTag一并给药(Optional ByVal lng组号 As Long)
'功能：在一并给药的医嘱前加标志
    Dim i As Long
    Dim lngUpRow As Long

    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If lng组号 = 0 Then .TextMatrix(i, 0) = ""
            If lng组号 <> 0 And Val(.TextMatrix(i, col_组号)) = lng组号 Then .TextMatrix(i, 0) = ""
            If Val(.TextMatrix(i, col_组号)) <> 0 And ((lng组号 = Val(.TextMatrix(i, col_组号)) And lng组号 <> 0) Or lng组号 = 0) And .RowHidden(i) = False Then
                lngUpRow = GetUpRow(i)
                If lngUpRow = 0 Then
                    .TextMatrix(i, 0) = "┏"
                Else
                    If Val(.TextMatrix(i, col_组号)) = Val(.TextMatrix(lngUpRow, col_组号)) And i <> lngUpRow Then
                        If .TextMatrix(lngUpRow, 0) = "┗" Then
                            .TextMatrix(lngUpRow, 0) = "┃"
                        End If
                        .TextMatrix(i, 0) = "┗"
                    Else
                        .TextMatrix(i, 0) = "┏"
                    End If
                End If
            End If
        Next
    End With
End Sub




Private Function Check一并给药(ByVal lngRow As Long) As Boolean
    Dim lngUpRow As Long
    With vsAdvice
        lngUpRow = GetUpRow(lngRow)
        If lngUpRow = 0 Then
             MsgBox "前面没有可以一并给药的用药行。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If .TextMatrix(lngRow, col_药品类别) = "中草药" Or .TextMatrix(lngUpRow, col_药品类别) = "中草药" Then
            MsgBox "中药配方不能设置为一并给药。", vbInformation, gstrSysName
            Exit Function
        End If
        Check一并给药 = True
    End With
End Function

Private Function GetUpRow(ByVal lngRow As Long) As Long
'功能：取上一行有效行
    Dim i As Long

    With vsAdvice
        lngRow = lngRow - 1
        For i = lngRow To 1 Step -1
            If .RowHidden(i) = False Then
                GetUpRow = i: Exit For
            End If
        Next
    End With
End Function

Private Function GetDownRow(ByVal lngRow As Long) As Long
'功能：取下一行有效行
    Dim i As Long

    With vsAdvice
        lngRow = lngRow + 1
        For i = lngRow To .Rows - 1
            If .RowHidden(i) = False Then
                GetDownRow = i: Exit For
            End If
        Next
    End With
End Function


Private Function Set同步一并给药(ByVal lngRow As Long) As Long
'功能：同步一组一并给药行数据
    Dim i As Long
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngRow, col_组号)) = Val(.TextMatrix(i, col_组号)) And Val(.TextMatrix(i, col_组号)) <> 0 And .RowHidden(i) = False Then
                .TextMatrix(i, col_组号) = .TextMatrix(lngRow, col_组号)
                .TextMatrix(i, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
                .TextMatrix(i, COL_用法) = .TextMatrix(lngRow, COL_用法)
                .TextMatrix(i, COL_用法id) = .TextMatrix(lngRow, COL_用法id)
                .TextMatrix(i, COL_执行频次) = .TextMatrix(lngRow, COL_执行频次)
                .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                .TextMatrix(i, COL_天数) = .TextMatrix(lngRow, COL_天数)
                
                .Cell(flexcpData, i, COL_执行频次) = .TextMatrix(lngRow, COL_执行频次)
                .Cell(flexcpData, i, COL_用法) = .TextMatrix(lngRow, COL_用法)
                .TextMatrix(i, col_是否修改) = "1"
                .Cell(flexcpBackColor, i, 0, i, 0) = Red_COLOR
            End If
        Next
    End With
End Function

