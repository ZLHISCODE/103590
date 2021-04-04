VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendFileOpen 
   BackColor       =   &H00FFFFFF&
   Caption         =   "护理记录查阅"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   Icon            =   "frmTendFileOpen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   8205
      ScaleHeight     =   2610
      ScaleWidth      =   2730
      TabIndex        =   8
      Top             =   1245
      Visible         =   0   'False
      Width           =   2760
      Begin VSFlex8Ctl.VSFlexGrid vfgThisPrint 
         Height          =   1695
         Left            =   105
         TabIndex        =   11
         Top             =   645
         Width           =   2340
         _cx             =   4128
         _cy             =   2990
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   6
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   250
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
         AutoResize      =   -1  'True
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
      Begin VB.TextBox txtLength 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1005
         Left            =   465
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1485
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Label lblSubHeadPrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:##"
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblTitlePrint 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "一般护理记录单"
         Height          =   180
         Left            =   405
         TabIndex        =   9
         Top             =   75
         Width           =   1275
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   555
      Left            =   120
      TabIndex        =   5
      Top             =   2670
      Visible         =   0   'False
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   979
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHead 
      Height          =   555
      Left            =   120
      TabIndex        =   4
      Top             =   2130
      Visible         =   0   'False
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   979
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendFileOpen.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17224
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
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   4170
      Left            =   135
      TabIndex        =   1
      Top             =   1320
      Width           =   7500
      _cx             =   13229
      _cy             =   7355
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   6
      FixedRows       =   3
      FixedCols       =   0
      RowHeightMin    =   250
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
      AutoResize      =   -1  'True
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
   Begin RichTextLib.RichTextBox rtbHead 
      Height          =   1200
      Left            =   2730
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   2117
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmTendFileOpen.frx":0E1C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbFoot 
      Height          =   1200
      Left            =   2730
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6540
      Visible         =   0   'False
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   2117
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmTendFileOpen.frx":0EB9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThisRowHeight 
      Height          =   1695
      Left            =   8430
      TabIndex        =   13
      Top             =   4335
      Width           =   2340
      _cx             =   4128
      _cy             =   2990
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   6
      FixedRows       =   3
      FixedCols       =   0
      RowHeightMin    =   250
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
      AutoResize      =   -1  'True
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
   Begin VB.Label lblSubhead 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名:##"
      Height          =   180
      Left            =   315
      TabIndex        =   3
      Top             =   1050
      Width           =   630
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "一般护理记录单"
      Height          =   180
      Left            =   3105
      TabIndex        =   2
      Top             =   510
      Width           =   1275
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   135
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmTendFileOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
'窗体变量
'######################################################################################################################
Private mblnHead As Boolean, mblnFoot As Boolean
Private mlngPatiId As Long, mlngPageId As Long, mlngDeptId As Long, mintBaby As Integer
Private mstrPeriod As String
Private mbyt护理级别 As Byte
Private mintTabTiers As Integer     '表头层次
Private mintTagFormHour As Integer  '开始时间条件
Private mintTagToHour As Integer    '截止时间条件
Private mobjTitleFont As New StdFont, lngTitleFontSize As Long '标题字体颜色
Private mobjSubFont  As New StdFont, lngSubFontSize As Long  '表格字体大小
Private mobjTagFont As New StdFont, lngTagFontSize As Long   '条件样式字体
Private mobjTagFontPrint As New StdFont '打印时条件字体
Private mlngTagColor As Long        '条件样式颜色
Private mdblRowHeightMin As Double     '表格最小高度
Private mstrPaperSet As String      '格式
Private mstrPageHead As String      '页眉
Private mstrPageFoot As String      '页脚
Private mblnChildForm As Boolean
Private mstrSubhead As String       '表上标签
Private mstrTabHead As String       '表头单元
Private mstrColWidth As String      '列宽序列串
Private mstrSQL As String           '组织产生的内容查询语句
Private mlngFileID As Long
Private mbln日期时间合并 As Boolean
Private mbln时间列隐藏 As Boolean
Private mintStartCOLCount As Long, mintEndColCount As Long '开始固定行数和结束固定行数，开始固定行为表格固定列+日期、时间，结束固定行为护士、签名人、签名时间、签名日期
'临时变量
Private cbrControl As CommandBarControl
Private cbrMenuBar As CommandBarPopup
Private cbrToolBar As CommandBar
Private mrsSumCol As New ADODB.Recordset
Private rsTemp As New ADODB.Recordset
Private lngCount As Long
Private mblnStartUp As Boolean
Private strTemp As String
Private lngCurColor As Long, strCurFont As String, objFont As StdFont

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const gconLineHigh = 30

Public WithEvents zlEvent_Print As zlPrintMethod
Attribute zlEvent_Print.VB_VarHelpID = -1
Public Event zlAfterPrint(ByVal lngFileID As Long)

Private mbytFontSize As Byte        '字体大小0-9号字体,1-12号字体
'######################################################################################################################

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-18 15:16
    '问题:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '编制:刘鹏飞
    '日期:2012-06-18 15:16
    '问题:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim lngCount As Long
    Dim aryItem() As String
    Dim blnTag As Boolean
    Dim lngReDraw As Long
    
    mobjSubFont.Size = lngSubFontSize
    Set CtlFont = mobjSubFont
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = BlowUp(CtlFont.Size)
    Set Me.Font = CtlFont
    '标题字体
    mobjTitleFont.Size = lngTitleFontSize
    Set CtlFont = mobjTitleFont
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = BlowUp(CtlFont.Size)
    lblTitle.AutoSize = True
    Set lblTitle.Font = CtlFont
    lblTitle.AutoSize = False
    '文本字体
    Set lblSubhead.Font = Me.Font
    '表格字体
    With vfgThis
         Set .Font = Me.Font
        lngReDraw = .Redraw
        .Redraw = flexRDNone
        .RowHeightMin = BlowUp(mdblRowHeightMin)
        '列宽设置
        aryItem = Split(mstrColWidth, ",")
        For lngCount = 2 To .Cols - 1
            .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - 2), "`")(0)))
        Next lngCount
        
         '条件格式
         mobjTagFont.Size = lngTagFontSize
        Set CtlFont = mobjTagFont
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = BlowUp(CtlFont.Size)
        For lngCount = .FixedRows To .Rows - 1
            blnTag = False
            If IsDate(.TextMatrix(lngCount, 0)) Then
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                End If
            End If
            If blnTag Then
                Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = CtlFont
                .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
            End If
        Next
        '89729:刘鹏飞,设置字体大小后需重新设置属性，自动调整表格高度
        .AutoSizeMode = flexAutoSizeRowHeight
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        .AutoSize 0, .Cols - 1
        For lngCount = 0 To .Rows - 1
            If .ROWHEIGHT(lngCount) < .RowHeightMin Then .ROWHEIGHT(lngCount) = .RowHeightMin
        Next
        .Redraw = lngReDraw
    End With
    
    If mblnChildForm = False Then
        Set CtlFont = cbsThis.Options.Font
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = mbytFontSize
        Set cbsThis.Options.Font = CtlFont
        
        cbsThis.RecalcLayout
    Else
        Call cbsThis_Resize
    End If
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '放大：字体，单元格宽度
    BlowUp = dblChange + (dblChange * IIf(mbytFontSize = 12, 1, 0) / 3)
End Function

Private Function SumTend(ByVal lngFile As Long, ByVal lngPatientKey As Long, ByVal lngPageKey As Long) As Boolean
    '*****************************************************************************************************************
    '功能： 汇总所有范围内的护理数据
    '参数：
    '返回：
    '*****************************************************************************************************************
    Dim rsGroup As New ADODB.Recordset
    Dim strSQL As String
    Dim rsTimePeriod As New ADODB.Recordset
    
    On Error GoTo errHand

    mrsSumCol.Filter = ""
    If mrsSumCol.RecordCount = 0 Then Exit Function
    
    '查找病人的医嘱记录,有无特殊护理的入出汇总的长期医嘱
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select a.开始执行时间,Nvl(a.执行终止时间,Sysdate+365) As 执行终止时间,b.标本部位 From 病人医嘱记录 a,诊疗项目目录 b Where a.病人id=[1] And a.主页id=[2] And b.ID=a.诊疗项目id And a.医嘱状态 Not In (1,2,4) And b.操作类型='12' And b.类别='Z' Order By a.开始执行时间"
    Set rsTimePeriod = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientKey, lngPageKey)
    If rsTimePeriod.BOF Then Exit Function
    
    '统计的时间段,如果没有时间段,退出不进行汇总
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select d.对象序号, d.对象属性, d.内容行次, d.内容文本 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '汇总时段' And d.内容文本 Is Not Null" & _
        " Order By d.对象序号, d.内容行次"
    Set rsGroup = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngFile)
    If rsGroup.BOF Then Exit Function
        
    With vfgThis
        
        Do While Not rsTimePeriod.EOF
                        
            Call SumRangeTend(rsGroup, Format(rsTimePeriod("开始执行时间").Value, "yyyy-MM-dd HH:mm:ss"), _
                                        Format(rsTimePeriod("执行终止时间").Value, "yyyy-MM-dd HH:mm:ss"), _
                                        Format(.TextMatrix(2, 1), "yyyy-MM-dd HH:mm:ss"), _
                                        Format(.TextMatrix(.Rows - 1, 1), "yyyy-MM-dd HH:mm:ss"), _
                                        zlCommFun.NVL(rsTimePeriod("标本部位").Value))
                        
            rsTimePeriod.MoveNext
        Loop

    End With
        
    SumTend = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function SumRangeTend(ByVal rsGroup As ADODB.Recordset, ByVal strStartTime As String, ByVal strEndTime As String, ByVal strMinTime As String, ByVal strMaxTime As String, Optional ByVal strColumn As String) As Boolean
    '*****************************************************************************************************************
    '功能： 汇总指定范围内的护理数据
    '参数：
    '返回：
    '*****************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim aryTmp As Variant
    Dim lngLoop As Long
    Dim strStart As String
    Dim strEnd As String
    Dim strSum As String
    Dim strTime As String
    Dim strSvrTime As String
    Dim rsResult As New ADODB.Recordset
    Dim lngLen As Long
    Dim lngRow As Long
    Dim intStartCol As Integer
    Dim intEndCol As Integer
    Dim strSumCol As String
    Dim lngDyas As Long
    Dim str终止时间 As String
    Dim str开始时间 As String
    Dim lngStartRow As Long, lngEndRow As Long
    Dim blnAllow As Boolean
    
    On Error GoTo errHand
    
    
    If strMaxTime < Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") Then strMaxTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    
    '初始化处理
    '------------------------------------------------------------------------------------------------------------------
    Set rsResult = New ADODB.Recordset
    With rsResult
        .Fields.Append "列号", adBigInt
        .Fields.Append "时间", adVarChar, 30
        .Fields.Append "结果", adVarChar, 100
        .Fields.Append "标题", adVarChar, 100
        .Open
        
    End With
    
    mrsSumCol.Filter = ""
    mrsSumCol.MoveFirst
    
    '从时间范围确认开始行,结束行
    '------------------------------------------------------------------------------------------------------------------
    With vfgThis
        For lngLoop = 3 To .Rows - 1    '表头固定是3行,有可能隐蔽了部分固定行,请搜索:mintTabTiers
            strTime = Format(.TextMatrix(lngLoop, 1), "yyyy-MM-dd HH:mm:ss")
            If lngStartRow = 0 Then
                If strTime >= strStartTime Then
                    lngStartRow = lngLoop
                End If
            End If
            
            If lngEndRow = 0 Then
                If strTime > strEndTime Then
                    lngEndRow = lngLoop - 1
                    Exit For
                End If
            End If
        Next
        If lngLoop = .Rows Then
            If lngEndRow = 0 Then lngEndRow = .Rows - 1
        End If
    End With
    If Not (lngStartRow > 0 And lngStartRow <= lngEndRow) Then Exit Function
    
    '求统计点
    '------------------------------------------------------------------------------------------------------------------
    
    lngDyas = DateDiff("d", CDate(strStartTime), CDate(strEndTime))
    rsGroup.MoveFirst
    Do While Not rsGroup.EOF
        strTmp = zlCommFun.NVL(rsGroup!内容文本)
        If strTmp <> "" Then
            aryTmp = Split(strTmp, ",")
            If UBound(aryTmp) >= 2 Then
                
                If InStr(aryTmp(1), ":") = 0 Then aryTmp(1) = aryTmp(1) & ":00"
                If InStr(aryTmp(2), ":") = 0 Then aryTmp(2) = aryTmp(2) & ":00"
                
                For lngLoop = 0 To lngDyas
                    
                    str开始时间 = Format(DateAdd("d", lngLoop, CDate(strStartTime)), "yyyy-MM-dd") & " " & Format(aryTmp(1), "HH:mm:ss")
                    
                    If str开始时间 < strStartTime Then
                        strSvrTime = strStartTime
                    Else
                        strSvrTime = str开始时间
                    End If
                    
                    If Format(aryTmp(1), "HH:mm:ss") < Format(aryTmp(2), "HH:mm:59") Then
                        '同一天
                        blnAllow = True
                        str终止时间 = Format(DateAdd("d", lngLoop, CDate(strStartTime)), "yyyy-MM-dd") & " " & Format(aryTmp(2), "HH:mm:59")
                    Else
                        '不是同一天
                        blnAllow = False
                        str终止时间 = Format(strSvrTime, "yyyy-MM-dd") & " " & Format(aryTmp(2), "HH:mm:59")
                        strSvrTime = Format(DateAdd("d", -1, CDate(strSvrTime)), "yyyy-MM-dd") & " " & Format(aryTmp(1), "HH:mm:ss")
                    End If
                                                            
                    If str终止时间 > strMaxTime Then Exit For
                                                            
                    If str终止时间 >= strStartTime And str终止时间 <= strEndTime And str终止时间 >= strMinTime Then
                        
                        mrsSumCol.Filter = ""
                        If strColumn <> "" Then mrsSumCol.Filter = "列名='" & strColumn & "'"
                        If mrsSumCol.RecordCount > 0 Then
                            mrsSumCol.MoveFirst
                            Do While Not mrsSumCol.EOF
                                
                                rsResult.AddNew
                                rsResult("列号").Value = mrsSumCol("列号").Value
                                rsResult("时间").Value = str终止时间
                                rsResult("结果").Value = 0
                                
                                lngLen = DateDiff("n", CDate(strSvrTime), CDate(str终止时间))
                                strSum = Format(lngLen \ 60, "00") & "小时" & Format(lngLen Mod 60, "00") & "分"
                                
                                If blnAllow Then
                                    rsResult("标题").Value = Format(str终止时间, "MM-dd") & " " & aryTmp(0) & "(" & strSum & ")"
                                Else
                                    rsResult("标题").Value = Format(strSvrTime, "MM-dd") & " " & aryTmp(0) & "(" & strSum & ")"
                                End If
                                
                                mrsSumCol.MoveNext
                            Loop
                        End If
                    End If
                Next
            End If
        End If
        rsGroup.MoveNext
    Loop
    
    '分组进行汇总，并填写到对应的rsResult记录集中
    '------------------------------------------------------------------------------------------------------------------
    strSum = ""
    rsGroup.MoveFirst
    With vfgThis
        .AddItem "": vfgThisPrint.AddItem ""
        .TextMatrix(.Rows - 1, 1) = Format(DateAdd("d", 1, CDate(.TextMatrix(.Rows - 2, 1))), "yyyy-MM-dd") & " 23:59:59"
        vfgThisPrint.TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 1, 1)
        lngEndRow = lngEndRow + 1
        
        '根据明细生成汇总行数据
        '--------------------------------------------------------------------------------------------------------------
        Do While Not rsGroup.EOF
            
            strTmp = zlCommFun.NVL(rsGroup!内容文本)
            If strTmp <> "" Then
                aryTmp = Split(strTmp, ",")
                If UBound(aryTmp) >= 2 Then
                    
                    mrsSumCol.Filter = ""
                    If strColumn <> "" Then mrsSumCol.Filter = "列名='" & strColumn & "'"
                    If mrsSumCol.RecordCount > 0 Then
                        mrsSumCol.MoveFirst
                        Do While Not mrsSumCol.EOF
    
                            If InStr(aryTmp(1), ":") = 0 Then aryTmp(1) = aryTmp(1) & ":00"
                            If InStr(aryTmp(2), ":") = 0 Then aryTmp(2) = aryTmp(2) & ":00"
                                        
                            For lngLoop = lngStartRow To lngEndRow
                                                                                        
                                strTime = Format(.TextMatrix(lngLoop, 1), "yyyy-MM-dd HH:mm:ss")
                                
                                If strTime > strEnd And strEnd <> "" Then
                                    '填写
                                    rsResult.Filter = ""
                                    rsResult.Filter = "时间='" & strEnd & "' And 列号=" & mrsSumCol("列号").Value & " And 标题 Like '*" & Split(rsGroup!内容文本, ",")(0) & "*'"
                                    If rsResult.RecordCount = 0 Then
                                        rsResult.AddNew
                                        rsResult("列号").Value = mrsSumCol("列号").Value
                                        rsResult("时间").Value = strEnd
                                    End If
                                    rsResult("结果").Value = Val(strSum)
                                        
                                    strSum = ""
                                    strStart = ""
                                    strEnd = ""
                                    strSvrTime = ""
                                End If
                                
                                '确定具体的时间段
                                If Format(aryTmp(1), "HH:mm:ss") < Format(aryTmp(2), "HH:mm:59") Then
                                    '时间范围在同一天的情况
                                    
                                    If (strStart = "" And strEnd = "") Or Not (strTime >= strStart And strTime <= strEnd) Then
                                        
                                        '判断上次和当前是否是同一天,如果不是,说明中间有无数据的统计
                                        If strSvrTime <> "" And strTime <> "" Then
                                            If CDate(strSvrTime) <> CDate(strTime) Then
                                                strSvrTime = IIf(strSvrTime = "", strStartTime, strTime)
                                            End If
                                        End If
                                        
                                        strSvrTime = IIf(strSvrTime = "", strStartTime, strTime)
                                        strStart = Format(strTime, "yyyy-MM-dd") & " " & Format(aryTmp(1), "HH:mm:ss")
                                        strEnd = Format(strTime, "yyyy-MM-dd") & " " & Format(aryTmp(2), "HH:mm:59")
                                        If strTime > strEnd Then
                                            strStart = ""
                                            strEnd = ""
                                        End If
    
                                    End If
                                    
                                    If strTime >= strStart And strTime <= strEnd And strEnd <> "" Then
                                        strSum = Val(strSum) + Val(.TextMatrix(lngLoop, mrsSumCol("列号").Value))
                                    End If
                                
                                Else
                                    '时间范围不在同一天的情况,即超过一天
                                    
                                    If (strStart = "" And strEnd = "") Or Not (strTime >= strStart Or strTime <= strEnd) Then
                                        
'                                        strEnd = Format(CDate(strTime), "yyyy-MM-dd") & " " & Format(aryTmp(2), "HH:mm:59")
'                                        strStart = Format(DateAdd("d", -1, CDate(strTime)), "yyyy-MM-dd") & " " & Format(aryTmp(1), "HH:mm:ss")
                                        strSvrTime = IIf(strSvrTime = "", strStartTime, strTime)
                                        strStart = Format(strTime, "yyyy-MM-dd") & " " & Format(aryTmp(1), "HH:mm:ss")
                                        strEnd = Format(DateAdd("d", 1, CDate(strTime)), "yyyy-MM-dd") & " " & Format(aryTmp(2), "HH:mm:59")
    
                                    End If
                                                                    
                                    If (strTime >= strStart And strTime <= strEnd) And strEnd <> "" Then
                                        strSum = Val(strSum) + Val(.TextMatrix(lngLoop, mrsSumCol("列号").Value))
                                    End If
                                
                                End If
                                    
                            Next
                        
                            If strEnd <> "" Then
                                rsResult.Filter = ""
                                rsResult.Filter = "时间='" & strEnd & "' And 列号=" & mrsSumCol("列号").Value & " And 标题 Like '*" & Split(rsGroup!内容文本, ",")(0) & "*'"
                                If rsResult.RecordCount = 0 Then
                                    rsResult.AddNew
                                    rsResult("列号").Value = mrsSumCol("列号").Value
                                    rsResult("时间").Value = strEnd
                                End If
                                rsResult("结果").Value = Val(strSum)
                                    
                                strSum = ""
                                strStart = ""
                                strEnd = ""
                                strSvrTime = ""
                            End If
                            
                            mrsSumCol.MoveNext
                        Loop
                    End If
                End If
            End If
            
            rsGroup.MoveNext
        Loop
        
        rsResult.Filter = "结果>0"
        Do While Not rsResult.EOF
            Debug.Print rsResult!时间 & ","; rsResult!结果 & "," & rsResult!标题
            rsResult.MoveNext
        Loop
        
        '根据已生成的汇总数据进行插入显示
        '--------------------------------------------------------------------------------------------------------------
        Call ShowSumTend(rsResult, lngStartRow, lngEndRow, strColumn)
        
    End With
        
    SumRangeTend = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If

End Function

Private Function ShowSumTend(ByVal rsResult As ADODB.Recordset, ByVal lngStartRow As Long, ByVal lngEndRow As Long, Optional ByVal strColumn As String) As Boolean
    '*****************************************************************************************************************
    '功能： 显示汇总数据
    '参数：
    '返回：
    '*****************************************************************************************************************
    Dim lngRow As Long
    Dim lngLoop As Long
    Dim intEndCol As Integer
    Dim strTmp As String
    Dim aryTmp As Variant
    Dim intStartCol As Integer
    
    On Error GoTo errHand
    If rsResult.RecordCount = 0 Then Exit Function
    
    rsResult.MoveFirst
    rsResult.Sort = "时间"
    
    mrsSumCol.Filter = ""
    If strColumn <> "" Then mrsSumCol.Filter = "列名='" & strColumn & "'"
    If mrsSumCol.RecordCount = 0 Then Exit Function
    
    mrsSumCol.Sort = "列号"
    If mbln日期时间合并 = False Then
        intStartCol = 2
    Else
        intStartCol = 1
    End If
    intEndCol = mrsSumCol("列号").Value
    If intEndCol <= 2 Then Exit Function

    With vfgThis
        
        lngLoop = lngStartRow
        
        Do While Not rsResult.EOF

            If Format(.TextMatrix(lngLoop, 1), "yyyy-MM-dd HH:mm:ss") > Format(rsResult("时间").Value, "yyyy-MM-dd HH:mm:ss") And .Cell(flexcpData, lngLoop, 1, lngLoop, 1) = "" Then
                If .Cell(flexcpData, lngLoop, 1, lngLoop, 1) <> rsResult("标题").Value Then
                    
                    If .Cell(flexcpData, lngLoop - 1, 1, lngLoop - 1, 1) <> rsResult("标题").Value Then
                        .AddItem "", lngLoop
                        
                        .MergeRow(lngLoop) = True
                        .Cell(flexcpText, lngLoop, intStartCol, lngLoop, intEndCol) = rsResult("标题").Value
                        .Cell(flexcpAlignment, lngLoop, intStartCol, lngLoop, intEndCol) = flexAlignCenterCenter
                        .Cell(flexcpData, lngLoop, 1, lngLoop, 1) = rsResult("标题").Value
                        .Cell(flexcpForeColor, lngLoop, 0, lngLoop, .Cols - 1) = 255
'                        .Cell(flexcpForeColor, lngLoop, intStartCol, lngLoop, intEndCol) = 255
                        
                        vfgThisPrint.AddItem "", lngLoop
                        
                        vfgThisPrint.MergeRow(lngLoop) = True
                        vfgThisPrint.Cell(flexcpText, lngLoop, intStartCol, lngLoop, intEndCol) = rsResult("标题").Value
                        vfgThisPrint.Cell(flexcpAlignment, lngLoop, intStartCol, lngLoop, intEndCol) = flexAlignCenterCenter
                        vfgThisPrint.Cell(flexcpData, lngLoop, 1, lngLoop, 1) = rsResult("标题").Value
                        vfgThisPrint.Cell(flexcpForeColor, lngLoop, 0, lngLoop, .Cols - 1) = 255
                        
                        lngEndRow = lngEndRow + 1
                        
                        lngRow = lngLoop
                    Else
                        lngRow = lngLoop - 1
                    End If
                    
                    If rsResult("列号").Value Mod 2 = 1 Then
                        .TextMatrix(lngRow, rsResult("列号").Value) = rsResult("结果").Value
                    Else
                        .TextMatrix(lngRow, rsResult("列号").Value) = " " & rsResult("结果").Value
                    End If
                    
                    .Cell(flexcpAlignment, lngRow, rsResult("列号").Value, lngRow, rsResult("列号").Value) = .Cell(flexcpAlignment, 2, rsResult("列号").Value, 2, rsResult("列号").Value)
                    vfgThisPrint.TextMatrix(lngRow, rsResult("列号").Value) = .TextMatrix(lngRow, rsResult("列号").Value)
                    vfgThisPrint.Cell(flexcpAlignment, lngRow, rsResult("列号").Value, lngRow, rsResult("列号").Value) = .Cell(flexcpAlignment, lngRow, rsResult("列号").Value, lngRow, rsResult("列号").Value)
                End If
                
                rsResult.MoveNext
            Else
                lngLoop = lngLoop + 1
                If lngLoop > lngEndRow Then Exit Do
            End If

        Loop

        
        .Rows = .Rows - 1
        
        '处理最后一天的情况(最后一天如果是在日间时间范围内停止的,就没有全日总结)
        '--------------------------------------------------------------------------------------------------------------
        strTmp = ""
        For lngLoop = .Rows - 1 To 1 Step -1
            If .Cell(flexcpData, lngLoop - 1, 1, lngLoop - 1, 1) = "" Then
                Exit For
            ElseIf .Cell(flexcpData, lngLoop, 1, lngLoop, 1) <> "" Then
                strTmp = strTmp & "," & lngLoop
            End If
        Next

        If strTmp <> "" Then
            aryTmp = Split(strTmp, ",")
            For lngLoop = 0 To UBound(aryTmp)
                If Val(aryTmp(lngLoop)) > 0 Then
                    .RemoveItem Val(aryTmp(lngLoop))
                    vfgThisPrint.RemoveItem Val(aryTmp(lngLoop))
                End If
            Next
        End If
    End With
    
    ShowSumTend = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    
End Function

Public Function zlPrintTend(Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDeviceName As String) As Boolean
    
    '1-预览,2-打印
    
    Select Case bytMode
    Case 1
        Call zlRptPrint(2, strPrintDeviceName)
    Case 2
        Call zlRptPrint(1, strPrintDeviceName)
    Case 3
        Call zlRptPrint(3, strPrintDeviceName)
    End Select
   
    '
End Function

Public Sub ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, lngDeptId As Long, Optional ByVal intBaby As Integer = 0, Optional ByVal strPeriod As String, Optional ByVal blnChildForm As Boolean = False, Optional ByVal byt护理级别 As Byte = 3, Optional ByVal blnDataMoved As Boolean, Optional ByVal bytSize As Byte = 0)
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngFileID           护理文件格式句柄
    '       lngPatiID           病人id
    '       lngPageID           主页id
    '       lngDeptID           要显示护理记录的科室
    '       intBaby             婴儿标志
    '       bytSize             '字体大小 0-9号字体 ，1-12号字体
    '返回： 无
    '******************************************************************************************************************
'    Dim bln护理级别 As Boolean
    
    Err = 0
    Dim stdObjFont As StdFont
    Dim rsItem As New ADODB.Recordset
    Dim rsCol As New ADODB.Recordset
    
    On Error GoTo errHand
    mlngFileID = lngFileID
    
    mblnChildForm = blnChildForm
    mblnMoved_HL = blnDataMoved
    If mblnChildForm = False Then
        Call InitForm
    Else
        
        If mblnStartUp Then
            'Me.WindowState = 2
            Call FormSetCaption(Me, False, False)
            
            stbThis.Visible = Not mblnChildForm
            cbsThis.ActiveMenuBar.Visible = False
            cbsThis.RecalcLayout
            mblnStartUp = False
        End If
        
    End If
    
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    picPrint.Visible = False
    lblSubHeadPrint.Caption = "": lblSubHeadPrint.Tag = ""
    
    Set mrsSumCol = New ADODB.Recordset
    With mrsSumCol
        .Fields.Append "列号", adBigInt
        .Fields.Append "列名", adVarChar, 50
        .Fields.Append "列标题", adVarChar, 100
        .Open
    End With
    
    '65164:刘鹏飞,2013-08-27
    Set rsCol = New ADODB.Recordset
    With rsCol
        .Fields.Append "序号", adBigInt
        .Open
    End With
    '提取汇总护理项目
    gstrSQL = "Select 项目名称 From 护理记录项目 where 项目类型=0 And 项目表示=4"
    Call zlDatabase.OpenRecordset(rsItem, gstrSQL, "提取汇总护理项目")
    
    '定义样式获取
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select l.名称 From 病历文件列表 l Where l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    If mblnChildForm = False Then Me.Caption = "护理记录查阅 - " & rsTemp!名称
    mlngPatiId = lngPatiID
    mlngPageId = lngPageId
    mlngDeptId = lngDeptId
    mintBaby = intBaby
    mstrPeriod = strPeriod
    mbyt护理级别 = byt护理级别
    mbln日期时间合并 = False
    mbln时间列隐藏 = False
    mintStartCOLCount = 0
    mintEndColCount = 0
'    bln护理级别 = (Val(zlDatabase.GetPara("按护理级别分组", glngSys, 1255, "0")) = 1)
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "表头层数": mintTabTiers = Val("" & !内容文本)
            Case "总列数":  Me.vfgThis.Cols = Val("" & !内容文本): Me.vfgThisPrint.Cols = Me.vfgThis.Cols
            Case "最小行高": Me.vfgThis.RowHeightMin = Val("" & !内容文本): mdblRowHeightMin = Me.vfgThis.RowHeightMin: Me.vfgThisPrint.RowHeightMin = Me.vfgThis.RowHeightMin
            Case "文本字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set Me.vfgThis.Font = objFont
                Set Me.lblSubhead.Font = Me.vfgThis.Font
                Set Me.Font = Me.lblSubhead.Font
                Set mobjSubFont = objFont
                lngSubFontSize = objFont.Size
                
                Set stdObjFont = New StdFont
                With stdObjFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set Me.lblSubHeadPrint.Font = stdObjFont
                Set Me.vfgThisPrint.Font = stdObjFont
                Set Me.picPrint.Font = stdObjFont
            Case "文本颜色": Me.vfgThis.ForeColor = Val("" & !内容文本): Me.vfgThisPrint.ForeColor = Val("" & !内容文本)
            Case "表格颜色"
                Me.vfgThis.GridColor = Val("" & !内容文本): Me.vfgThis.GridColorFixed = Me.vfgThis.GridColor
                Me.vfgThisPrint.GridColor = Val("" & !内容文本): Me.vfgThisPrint.GridColorFixed = Me.vfgThis.GridColor
            Case "标题文本": Me.lblTitle.Caption = "" & !内容文本: Me.lblTitlePrint.Caption = Me.lblTitle.Caption
            Case "标题字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                Set stdObjFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set Me.lblTitle.Font = objFont
                Me.lblTitle.AutoSize = False
                Set mobjTitleFont = objFont
                lngTitleFontSize = objFont.Size
                
                With stdObjFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set Me.lblTitlePrint.Font = stdObjFont
                Me.lblTitlePrint.AutoSize = False
                
            Case "开始时间": mintTagFormHour = Val("" & !内容文本)
            Case "终止时间": mintTagToHour = Val("" & !内容文本)
            Case "条件字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
                lngTagFontSize = objFont.Size
                
                Set stdObjFont = New StdFont
                With stdObjFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjTagFontPrint = stdObjFont
            Case "条件颜色": mlngTagColor = Val("" & !内容文本)
            Case "日期时间合并": mbln日期时间合并 = (Val("" & !内容文本) = 1)
            '65502:刘鹏飞,2013-11-12
            Case "时间列隐藏": mbln时间列隐藏 = (Val("" & !内容文本) = 1)
            End Select
            .MoveNext
        Loop
    End With
    
    If mbln时间列隐藏 = True Then mbln日期时间合并 = False
    
    
    gstrSQL = "Select 种类||'-'||编号 AS KEY,格式, 页眉, 页脚,报表 From 病历页面格式 Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!格式: mstrPageHead = "" & rsTemp!页眉: mstrPageFoot = "" & rsTemp!页脚
        mblnHead = ReadPageHead(rtbHead, rsTemp!Key)
        mblnFoot = ReadPageFoot(rtbFoot, rsTemp!Key)
        
        mbyt护理级别 = rsTemp!报表
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !内容行次 - 1 & "," & !对象序号 & "," & !内容文本
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '查询语句组织
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql内 As String, strSql中 As String, strSql外 As String, strSql列 As String, strSQL条件 As String
    Dim bln日期 As Boolean, bln时间 As Boolean, bln护士 As Boolean
    Dim bln签名人 As Boolean, bln签名时间 As Boolean, bln签名日期 As Boolean
    Dim lngColumn As Long, lngNum As Long
    'lngNum 钦州因为项目名中有冒号,导致最外层不能读数据,最外层替换为虚拟组
    lngNum = 1

    
    gstrSQL = "Select d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    '65164:刘鹏飞,2013-08-27
    With rsTemp
        Do While Not .EOF
            rsCol.AddNew
            rsCol("序号") = Val(NVL(!对象序号, 0))
            rsCol.Update
        .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    
    With rsTemp
        lngColumn = 0: mstrColWidth = ""
        strSql内 = "": strSql中 = "": strSql外 = "": strSql列 = "": strSQL条件 = ""
        bln日期 = False: bln时间 = False: bln护士 = False
        bln签名人 = False: bln签名时间 = False: bln签名日期 = False
        Do While Not .EOF
            
            If lngColumn <> !对象序号 Then
                mstrColWidth = mstrColWidth & "," & !对象属性
                If NVL(!要素名称) <> "" Then
                    If strSql外 <> "" Then
                        strSql列 = strSql列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        strSql列 = strSql列 & ",'' As C" & Format(lngColumn, "00")
                    End If
                Else
                    If strSql外 <> "" Then
                        strSql列 = strSql列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        strSql列 = strSql列 & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql外 = ""
                lngColumn = !对象序号
            End If
            
            '53172:刘鹏飞,2013-04-25,修改提取护士由c.记录人改为l.保存人
            Select Case NVL(!要素名称)
            Case "日期"
                bln日期 = True
                strSql中 = strSql中 & ",日期"
                strSql内 = strSql内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名人"
                bln签名人 = True
                strSql中 = strSql中 & ",签名人"
                strSql内 = strSql内 & ",a.记录人 As 签名人"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名时间"
                bln签名时间 = True
                strSql中 = strSql中 & ",签名时间"
                strSql内 = strSql内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称,12,5)) As 签名时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名日期"
                bln签名日期 = True
                strSql中 = strSql中 & ",签名日期"
                strSql内 = strSql内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称, 1,11)) As 签名日期"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "时间"
                bln时间 = True
                strSql中 = strSql中 & ",时间"
                strSql内 = strSql内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
                strSql外 = strSql外 & "||" & !要素名称
            Case "护士"
                bln护士 = True
                strSql中 = strSql中 & ",护士"
                'strSql内 = strSql内 & ",c.记录人 As 护士"
                strSql内 = strSql内 & ",l.保存人 as 护士"
                strSql外 = strSql外 & "||" & !要素名称
            Case Else
                If NVL(!要素名称) <> "" Then
                    strSql中 = strSql中 & ",Max(""" & !要素名称 & """) As """ & "B" & Format(lngNum, "00") & """"
                    
                    strSQL条件 = strSQL条件 & " Or """ & !要素名称 & """ Is Not Null"
                    
                    strSql外 = strSql外 & "||""B" & Format(lngNum, "00") & """"
                    
                    If Trim("" & !内容文本) = "" And Trim("" & !要素单位) = "" Then
                        strSql内 = strSql内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,c.记录内容), '') As """ & !要素名称 & """"
'                        strSql外 = strSql外 & "||""" & !要素名称 & """"
                    Else
                        strSql内 = strSql内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,Null,'" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "')), '') As """ & !要素名称 & """"
'                        strSql外 = strSql外 & "||Decode(""" & !要素名称 & """, Null, Null,'" & !内容文本 & "'||""" & !要素名称 & """||'" & !要素单位 & "')"
                    End If
                    lngNum = lngNum + 1
                End If
            End Select
            
            '65164:刘鹏飞,2013-08-27,28版本前因为没有汇总项目标识，在样式构造中需要执行汇总列来判断是否汇总项目(只针对每列帮顶一个项目)
            '28版本护理记录项目增加了汇总标识，目前只要是汇总列都进行汇总。
            rsItem.Filter = "项目名称='" & "" & !要素名称 & "'"
            rsCol.Filter = "序号=" & lngColumn
            'If zlCommFun.NVL(!要素表示, 0) = 1 Then
            If rsItem.RecordCount > 0 And rsCol.RecordCount = 1 Then
                mrsSumCol.Filter = ""
                mrsSumCol.Filter = "列号=" & lngColumn + 1
                If mrsSumCol.RecordCount = 0 Then
                    mrsSumCol.AddNew
                    mrsSumCol("列号") = lngColumn + 1
                    mrsSumCol("列名") = "" & !要素名称
                    mrsSumCol("列标题") = "" & !要素名称
                    mrsSumCol.Update
                End If
            End If
            
            .MoveNext
        Loop
        If Mid(strSql外, 3) <> "" Then
            strSql列 = strSql列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
        Else
            strSql列 = strSql列 & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If strSQL条件 <> "" Then strSQL条件 = "(" & Mid(strSQL条件, 5) & ")"
        
        '如果没有出现日期，时间，护士，则内层需要补充，以保证中层分组的正常：
        If bln日期 = False Then strSql内 = strSql内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
        If bln时间 = False Then strSql内 = strSql内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
        If bln护士 = False Then strSql内 = strSql内 & ",l.保存人 as 护士"
        'If bln护士 = False Then strSql内 = strSql内 & ",c.记录人 As 护士"
        
        If bln签名人 = False Then strSql内 = strSql内 & ",a.记录人 As 签名人"
        If bln签名日期 = False Then strSql内 = strSql内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称,1,11)) As 签名日期"
        If bln签名时间 = False Then strSql内 = strSql内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称,12,5)) As 签名时间"
        
        If Mid(strSql中, 2) = "" Then
            ShowSimpleMsg "对不起，您没有定义当前护理单的显示列信息，请在病历文件管理中定义！"
            Exit Sub
        End If
        If bln日期 = True Then mintStartCOLCount = mintStartCOLCount + 1
        If bln时间 = True Then mintStartCOLCount = mintStartCOLCount + 1
        If bln护士 = True Then mintEndColCount = mintEndColCount + 1
        
        If bln签名人 = True Then mintEndColCount = mintEndColCount + 1
        If bln签名日期 = True Then mintEndColCount = mintEndColCount + 1
        If bln签名时间 = True Then mintEndColCount = mintEndColCount + 1
        
        mintStartCOLCount = mintStartCOLCount + 2
        mstrSQL = "Select 备用,发生时间," & Mid(strSql列, 12) & vbCrLf & _
                " From (Select 记录组号,时间 as 备用,发生时间," & Mid(strSql中, 2) & vbCrLf & _
                "        From (Select c.记录组号,发生时间," & Mid(strSql内, 2) & vbCrLf & _
                "               From 病人护理记录 l, 病人护理内容 c,病人护理内容 a " & vbCrLf & _
                "               Where l.Id = c.记录id And l.病人id = [1] And l.主页id = [2] And a.记录id(+)=l.ID And a.记录类型(+)=5 And a.终止版本(+) IS NULL And Nvl(l.婴儿,0)=[4] And c.终止版本 Is Null And c.记录类型<>5 And l.科室id + 0 = [3] And l.发生时间 Between [5] And [6] And l.护理级别<=[7])" & vbCrLf & _
                IIf(strSQL条件 <> "", "Where " & strSQL条件, "") & _
                "       Group By 日期, 时间, 发生时间,记录组号,护士,签名人,签名日期,签名时间" & _
                                "       Order By 日期, 时间, 发生时间,记录组号,护士,签名人,签名日期,签名时间)"
                
        mstrColWidth = Mid(mstrColWidth, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    Call zlRefresh
    '窗体显示
    If blnChildForm = False Then
        Call SetFontSize(bytSize)
        If frmParent Is Nothing Then
            Me.Show vbModal
        Else
            Me.Show vbModal, frmParent
        End If
        
        Unload Me
    Else
        '102173:电子病案审查打印护理记录单(病人存在两份或以上),打印第二份死机处理。
        mblnStartUp = False
'        Call cbsThis_Resize
    End If
    
    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitForm() As Boolean
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '数据元素形态设置
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
End Function

Private Sub zlRefresh(Optional ByVal blnReSize As Boolean = False)
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, strCell As String
    Dim blnTag As Boolean
    Dim strTmpSQL As String
    Dim aryPeriod() As String
    Dim strTmp As String
    '列宽设置
    Dim blnAlign As Boolean
    Dim dblWidth As Double  '日期与时间列的宽度
    Dim strCol As String
    
    Err = 0: On Error GoTo errHand
    
    '表上标签获取
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    lblSubHeadPrint.Caption = ""
    lblSubHeadPrint.Tag = ""
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as 信息 From Dual"
    aryItem = Split(mstrSubhead, "|")
    
    aryPeriod = Split(mstrPeriod, "～")
    aryPeriod(0) = Format(aryPeriod(0) & ":00", "yyyy-MM-dd HH:mm:ss")
    aryPeriod(1) = Format(aryPeriod(1) & ":59", "yyyy-MM-dd HH:mm:ss")
        
    For lngCount = 0 To UBound(aryItem)
        strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
        strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
        
        strTmp = strPrefix
        Select Case strItemName
        Case "当前病区"
        
            strTmpSQL = "Select b.名称" & vbNewLine & _
                        "From (Select 病区id, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a,部门表 b " & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [5] Between a.开始时间 And a.终止时间) And a.病区id Is Not Null And b.ID=a.病区id" & vbNewLine & _
                        "Order By a.开始时间"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            
        Case "当前床号"
        
            strTmpSQL = "Select a.床号" & vbNewLine & _
                        "From (Select 床号, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [5] Between a.开始时间 And a.终止时间) And a.床号 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"

            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "当前科室"
        
            strTmpSQL = "Select 名称 From 部门表 a Where a.ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngDeptId)
            
        Case "住院医师"
            strTmpSQL = "Select a.经治医师" & vbNewLine & _
                        "From (Select 经治医师, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [5] Between a.开始时间 And a.终止时间) And a.经治医师 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "责任护士"
        
            strTmpSQL = "Select a.责任护士" & vbNewLine & _
                        "From (Select 责任护士, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [5] Between a.开始时间 And a.终止时间) And a.责任护士 Is Not Null" & vbNewLine & _
                        "Order By a.开始时间"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "护理等级"
'            不知何义,注释先
'            strTmpSQL = "Select b.名称" & vbNewLine & _
'                        "From (Select 护理等级ID, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
'                        "            From 病人变动记录" & vbNewLine & _
'                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a,诊疗项目目录 b" & vbNewLine & _
'                        "Where ([4] Between a.开始时间 And a.终止时间 Or [5] Between a.开始时间 And a.终止时间) And a.护理等级ID Is Not Null And b.ID=a.护理等级ID" & vbNewLine & _
'                        "Order By a.开始时间"
                        
            strTmpSQL = "Select b.名称" & vbNewLine & _
                        "From (Select 护理等级ID, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                        "            From 病人变动记录" & vbNewLine & _
                        "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a,护理等级 b" & vbNewLine & _
                        "Where ([4] Between a.开始时间 And a.终止时间 Or [5] Between a.开始时间 And a.终止时间) And a.护理等级ID Is Not Null And b.序号=a.护理等级ID" & vbNewLine & _
                        "Order By a.开始时间"
            Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case Else
            strTmp = ""
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strPrefix, strItemName, mlngPatiId, mlngPageId, mintBaby)
        End Select
        
        If rsTemp.BOF = False Then
            If strTmp <> "" Then
                lblSubhead.Tag = lblSubhead.Tag & " " & strTmp & rsTemp.Fields(0).Value
            Else
                lblSubhead.Tag = lblSubhead.Tag & " " & rsTemp.Fields(0).Value
            End If
        End If
    Next
    lblSubhead.Tag = Trim(lblSubhead.Tag)
    lblSubHeadPrint.Tag = lblSubhead.Tag
    '表上标签分散处理
    Call zlLableBruit
    
    '装入数据
    gstrSQL = mstrSQL
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId, mintBaby, CDate(aryPeriod(0)), CDate(aryPeriod(1)), mbyt护理级别)
    
    '问题号：51746，刘鹏飞，2012-06-18 15:16，设置字体大小
    '打印表格
    With Me.vfgThisPrint
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp

        '表头填写
        '65164:刘鹏飞,2013-08-27,修改合并方式
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        strCol = ""
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + 1) = strCell
            
            If mbln日期时间合并 And InStr(1, ",日期,时间,", "," & strCell & ",") > 0 And strCell <> "" Then
                .ColHidden(lngCol + 1) = True
                strCol = strCol & "," & lngCol + 1
            End If
            If strCell = "时间" And mbln时间列隐藏 = True Then .ColHidden(lngCol + 1) = True
        Next
        
        '列宽设置
        blnAlign = False
        dblWidth = 0 '日期与时间列的宽度
        
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        aryItem = Split(mstrColWidth, ",")
        For lngCount = 2 To .Cols - 1
            .ColWidth(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(0))
            If mbln日期时间合并 And InStr(1, strCol & ",", "," & lngCount & ",") > 0 Then
                dblWidth = dblWidth + .ColWidth(lngCount)
            End If
            If InStr(1, aryItem(lngCount - 2), "`") <> 0 Then
                blnAlign = True
                .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(1))
            End If
        Next
        '将发生时间列显示出来,列宽为日期与时间列的总宽度
        If mbln日期时间合并 Then
            .ColHidden(1) = False
            .ColWidth(1) = dblWidth
            .TextMatrix(0, 1) = "发生时间"
            If mintTabTiers >= 2 Then .TextMatrix(1, 1) = "发生时间"
            If mintTabTiers >= 3 Then .TextMatrix(2, 1) = "发生时间"
        End If
        
        '条件格式
        For lngCount = .FixedRows To .Rows - 1
            blnTag = False
            If mintTagFormHour < mintTagToHour Then
                blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
            Else
                blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
            End If
            If blnTag Then
                Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFontPrint
                .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
            End If
        Next
        
        '自适应高度和其他格式调整
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeRowHeight
        '再按列合并
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        .AutoSize 0, .Cols - 1
        
        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .ROWHEIGHT(lngCount) < .RowHeightMin Then .ROWHEIGHT(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = 0
'            .ROWHEIGHT(2) = 0
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
        
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = .RowHeightMin
'            .ROWHEIGHT(2) = 0
            
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
            
            
        Case 3
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = .RowHeightMin
'            .ROWHEIGHT(2) = .RowHeightMin

            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        .Redraw = flexRDDirect
    End With
    
    With Me.vfgThis
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp

        '表头填写
        '65164:刘鹏飞,2013-08-27,修改合并方式
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        strCol = ""
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + 1) = strCell
            
            If mbln日期时间合并 And InStr(1, "日期,时间", strCell) > 0 And strCell <> "" Then
                .ColHidden(lngCol + 1) = True
                strCol = strCol & "," & lngCol + 1
            End If
            If strCell = "时间" And mbln时间列隐藏 = True Then .ColHidden(lngCol + 1) = True
        Next
        
        '列宽设置
        blnAlign = False
        dblWidth = 0 '日期与时间列的宽度
        
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        aryItem = Split(mstrColWidth, ",")
        For lngCount = 2 To .Cols - 1
            .ColWidth(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(0))
            If mbln日期时间合并 And InStr(1, strCol & ",", "," & lngCount & ",") > 0 Then
                dblWidth = dblWidth + .ColWidth(lngCount)
            End If
            If InStr(1, aryItem(lngCount - 2), "`") <> 0 Then
                blnAlign = True
                vfgThis.ColAlignment(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(1))
            End If
        Next
        '将发生时间列显示出来,列宽为日期与时间列的总宽度
        If mbln日期时间合并 Then
            .ColHidden(1) = False
            .ColWidth(1) = dblWidth
            .TextMatrix(0, 1) = "发生时间"
            If mintTabTiers >= 2 Then .TextMatrix(1, 1) = "发生时间"
            If mintTabTiers >= 3 Then .TextMatrix(2, 1) = "发生时间"
        End If
        
        '条件格式
        For lngCount = .FixedRows To .Rows - 1
            blnTag = False
            If mintTagFormHour < mintTagToHour Then
                blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
            Else
                blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
            End If
            If blnTag Then
                Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
            End If
        Next
        
        '自适应高度和其他格式调整
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeRowHeight
        '再按列合并
        For lngCount = 0 To vfgThis.Cols - 1
            vfgThis.MergeCol(lngCount) = True
        Next
        vfgThis.AutoSize 0, vfgThis.Cols - 1
        
        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .ROWHEIGHT(lngCount) < .RowHeightMin Then .ROWHEIGHT(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = 0
'            .ROWHEIGHT(2) = 0
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
        
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = .RowHeightMin
'            .ROWHEIGHT(2) = 0
            
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
            
            
        Case 3
'            .ROWHEIGHT(0) = .RowHeightMin
'            .ROWHEIGHT(1) = .RowHeightMin
'            .ROWHEIGHT(2) = .RowHeightMin

            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        Call SumTend(mlngFileID, mlngPatiId, mlngPageId)
        
        .Redraw = flexRDDirect
    End With
    
    '在模态床体下刷新需要重新设置字体大小
    If blnReSize = True Then Call ReSetFontSize
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub aa()

End Sub
Private Sub zlLableBruit()
    
    '根据宽度分散标签
    Dim aryRow() As String
    Dim lngSpaces As Long
'    aryRow = Split(Me.lblSubhead.Tag, vbCrLf)
'
'    For lngCount = 0 To UBound(aryRow)
'        If UBound(Split(aryRow(lngCount), Space(1))) > 0 Then
'            lngSpaces = 1
'            Do
'                If Me.TextWidth(Join(Split(aryRow(lngCount), Space(1)), Space(lngSpaces + 1))) > Me.vfgThis.Width Then
'                    If lngSpaces > 1 Then lngSpaces = lngSpaces - 1
'                    aryRow(lngCount) = Join(Split(aryRow(lngCount), Space(1)), Space(lngSpaces))
'                    Exit Do
'                End If
'                lngSpaces = lngSpaces + 1
'            Loop
'        End If
'    Next
'    Me.lblSubhead.Caption = Join(aryRow, vbCrLf)
    Me.lblSubhead.Caption = Me.lblSubhead.Tag
    Me.lblSubHeadPrint.Caption = Me.lblSubHeadPrint.Tag
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    Me.vfgThis.Move lngScaleLeft + 210, Me.lblSubhead.Top + Me.lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    Me.vfgThis.Height = lngScaleBottom - Me.vfgThis.Top - 210
End Sub

Private Sub VsfToMsh(Optional ByVal blnShow As Boolean = False)
    Dim dblHeight As Double
    Dim lngRow As Long, lngRows As Long
    Dim lngCol As Long, lngCols As Long
    On Error Resume Next
    
    '1、先转换表头
    lngRows = vfgThisPrint.FixedRows - 1
    lngCols = vfgThisPrint.Cols - 1
    '设置表头基础格式
    mshHead.Rows = lngRows + 2
    mshHead.FixedRows = lngRows + 1
    mshHead.Cols = lngCols + 1
    mshHead.FixedCols = vfgThisPrint.FixedCols
    mshHead.MergeCells = flexMergeFree
    mshHead.ROWHEIGHT(mshHead.Rows - 1) = 0
    Set mshHead.Font = vfgThisPrint.Font
    For lngRow = 0 To lngRows
        mshHead.Row = lngRow
        vfgThisPrint.Row = lngRow
        If vfgThisPrint.RowHidden(lngRow) Then
            mshHead.ROWHEIGHT(lngRow) = 0
        Else
            mshHead.ROWHEIGHT(lngRow) = vfgThisPrint.ROWHEIGHT(lngRow)
        End If
        dblHeight = dblHeight + vfgThisPrint.ROWHEIGHT(lngRow)
        For lngCol = 0 To lngCols
            mshHead.Col = lngCol
            vfgThisPrint.Col = lngCol
            mshHead.CellAlignment = vfgThisPrint.CellAlignment
            mshHead.TextMatrix(lngRow, lngCol) = vfgThisPrint.TextMatrix(lngRow, lngCol)
            If lngRow = lngRows Then '设置列宽
                If vfgThisPrint.ColHidden(lngCol) Then
                    mshHead.ColWidth(lngCol) = 0
                Else
                    mshHead.ColWidth(lngCol) = vfgThisPrint.ColWidth(lngCol)
                End If
            End If
        Next
        mshHead.MergeRow(lngRow) = True
    Next
    '依次列合并
    For lngCol = 0 To lngCols
        mshHead.MergeCol(lngCol) = True
    Next
    
    '2、再转换表体
    lngRows = vfgThisPrint.Rows - vfgThisPrint.FixedRows
    mshDetail.Rows = lngRows
    mshDetail.Cols = lngCols + 1
    mshDetail.FixedRows = 0
    mshHead.FixedCols = vfgThisPrint.FixedCols
    mshDetail.WordWrap = False
    '65164:刘鹏飞,2013-08-27,修改打印不能合并问题
    mshDetail.MergeCells = flexMergeFree
    Set mshDetail.Font = vfgThisPrint.Font
    Set picPrint.Font = vfgThisPrint.Font
    Set vfgThisRowHeight.Font = vfgThisPrint.Font
    
    For lngRow = 0 To lngRows - 1
        mshDetail.Row = lngRow
        If vfgThisPrint.RowHidden(lngRow + vfgThisPrint.FixedRows) Then
            mshDetail.ROWHEIGHT(lngRow) = 0
        Else
            mshDetail.ROWHEIGHT(lngRow) = vfgThisPrint.ROWHEIGHT(lngRow + vfgThisPrint.FixedRows)
        End If
        For lngCol = 0 To lngCols
            mshDetail.Col = lngCol
            vfgThisPrint.Row = lngRow + vfgThisPrint.FixedRows
            vfgThisPrint.Col = lngCol
            mshDetail.CellForeColor = vfgThisPrint.CellForeColor
            
            mshDetail.CellAlignment = vfgThisPrint.ColAlignment(lngCol)
            
            mshDetail.TextMatrix(lngRow, lngCol) = vfgThisPrint.TextMatrix(lngRow + vfgThisPrint.FixedRows, lngCol)
            If lngRow = lngRows - 1 Then '设置列宽
                If vfgThisPrint.ColHidden(lngCol) Then
                    mshDetail.ColWidth(lngCol) = 0
                Else
                    mshDetail.ColWidth(lngCol) = vfgThisPrint.ColWidth(lngCol)
                End If
            End If
        Next
        '65164:刘鹏飞,2013-08-27,修改打印不能合并问题
        mshDetail.MergeRow(lngRow) = vfgThisPrint.MergeRow(lngRow + vfgThisPrint.FixedRows)
    Next
    
    '设置大小，位置
    mshHead.Move vfgThisPrint.Left, vfgThisPrint.Top, vfgThisPrint.Width, dblHeight
    mshDetail.Move vfgThisPrint.Left, vfgThisPrint.Top + dblHeight, vfgThisPrint.Width, vfgThisPrint.Height - dblHeight
    
    mshHead.Visible = blnShow
    mshDetail.Visible = blnShow
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte, Optional ByVal strPrintDeviceName As String)
    Dim objPrint As New zlPrint2Grd, objAppRow As zlTabAppRow
    Dim lngWidth As Long, lngEmptyLR As Long, lngScaleWidth As Long
    On Error GoTo errHand
    
    If zlEvent_Print Is Nothing Then
        Set zlEvent_Print = VBA.GetObject("", "zl9PrintMode.zlPrintMethod")
    End If
    
    objPrint.EmptyUp = GetSetting("ZLSOFT", "公共模块\zl9PrintMode\Default", "PageUp", 20)
    objPrint.EmptyDown = GetSetting("ZLSOFT", "公共模块\zl9PrintMode\Default", "PageDown", 20)
    objPrint.EmptyLeft = GetSetting("ZLSOFT", "公共模块\zl9PrintMode\Default", "PageLeft", 20)
    objPrint.EmptyRight = GetSetting("ZLSOFT", "公共模块\zl9PrintMode\Default", "PageRight", 20)
        
    '设置打印格式
    If mblnHead Then mstrPageHead = Me.rtbHead.Text
    If mblnFoot Then mstrPageFoot = Me.rtbFoot.Text
    SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "PageHead", mstrPageHead
    SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "PageFoot", mstrPageFoot
    If UBound(Split(mstrPaperSet, ";")) >= 0 Then SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "PaperSize", Val(Split(mstrPaperSet, ";")(0))
    If UBound(Split(mstrPaperSet, ";")) >= 1 Then SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "Orientation", Val(Split(mstrPaperSet, ";")(1))
    If UBound(Split(mstrPaperSet, ";")) >= 2 Then SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "Height", Val(Split(mstrPaperSet, ";")(2))
    If UBound(Split(mstrPaperSet, ";")) >= 3 Then SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "Width", Val(Split(mstrPaperSet, ";")(3))
    lngEmptyLR = 0
    If UBound(Split(mstrPaperSet, ";")) >= 4 Then
        lngEmptyLR = lngEmptyLR + Val(Split(mstrPaperSet, ";")(4))
        objPrint.EmptyLeft = Round(Me.ScaleY(Val(Split(mstrPaperSet, ";")(4)), vbTwips, vbMillimeters), 2)
        SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "PageLeft", objPrint.EmptyLeft
    End If
    If UBound(Split(mstrPaperSet, ";")) >= 5 Then
        lngEmptyLR = lngEmptyLR + Val(Split(mstrPaperSet, ";")(5))
        objPrint.EmptyRight = Round(Me.ScaleY(Val(Split(mstrPaperSet, ";")(5)), vbTwips, vbMillimeters), 2)
        SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "PageRight", objPrint.EmptyRight
    End If
    If UBound(Split(mstrPaperSet, ";")) >= 6 Then
        objPrint.EmptyUp = Round(Me.ScaleX(Val(Split(mstrPaperSet, ";")(6)), vbTwips, vbMillimeters), 2)
        SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "PageUp", objPrint.EmptyUp
    End If
    If UBound(Split(mstrPaperSet, ";")) >= 7 Then
        objPrint.EmptyDown = Round(Me.ScaleX(Val(Split(mstrPaperSet, ";")(7)), vbTwips, vbMillimeters), 2)
        SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "PageDown", objPrint.EmptyDown
    End If
    
    '84140:LPF,标上标签内容输出需根据打印设置计算
    On Error Resume Next
    Printer.PaperSize = Val(Split(mstrPaperSet, ";")(0))
    Printer.Orientation = Val(Split(mstrPaperSet, ";")(1))
    
    If Printer.PaperSize = 256 Then
        Call SetCustonPager(Me.hWnd, Val(Split(mstrPaperSet, ";")(3)), Val(Split(mstrPaperSet, ";")(2)))
    End If
    lngScaleWidth = Printer.Width - lngEmptyLR
    On Error GoTo errHand

    Call VsfToMsh(False)
    
    Set objPrint.BodyHead = Me.mshHead
    Set objPrint.BodyGrid = Me.mshDetail
    objPrint.Title.Text = lblTitlePrint.Caption
    Set objPrint.Title.Font = lblTitlePrint.Font
    Set objPrint.AppFont = lblSubHeadPrint.Font
    
    Dim strLable As String, strAppRow As String, lngSpaces As Long
    Dim lngStart As Long, lngPos As Long, lngMAX As Long, lngNumber As Long, blnNumber As Boolean, lngAsc As Long
    lngSpaces = lblSubHeadPrint.Height / 210
    strLable = lblSubHeadPrint.Caption
    lngMAX = Len(strLable)
    lngNumber = 0
    lngStart = 1
    For lngPos = 1 To lngMAX
        '如果数学超长,则把数字移到下一行显示
        lngAsc = Asc(Mid(strLable, lngPos, 1))

        '检查是否超宽(长度超过行宽,或者遇到回车换行符)
        If picPrint.TextWidth(Mid(strLable, lngStart, lngPos - lngStart + 1) & "测") > lngScaleWidth Or lngPos = lngMAX Or lngAsc = 10 Then

            strAppRow = Mid(strLable, lngStart, lngPos - lngStart + 1)
            lngStart = lngPos + 1
            
            '输出表上项
            Set objAppRow = New zlTabAppRow
            Call objAppRow.Add(strAppRow)
            Call objPrint.UnderAppRows.Add(objAppRow)
        End If
    Next
    
    lngWidth = Val(Split(mstrPaperSet, ";")(3))
    If mstrPageHead <> "" Then objPrint.Header = mstrPageHead
    If mstrPageFoot <> "" Then
        mstrPageFoot = Replace(mstrPageFoot, "{打印时间}", Now)
        mstrPageFoot = Replace(mstrPageFoot, "{打印人}", gstrUserName)
        objPrint.Footer = mstrPageFoot ' LeftB(mstrPageFoot & Space(lngWidth), lngWidth - objPrint.EmptyLeft - objPrint.EmptyRight)
    End If
    
    
    If bytMode = 1 Then
        If strPrintDeviceName = "" Then
            bytMode = zlEvent_Print.zlPrintAsk(objPrint)
        Else
            SaveSetting "ZLSOFT", "公共模块\zl9PrintMode\Default", "DeviceName", strPrintDeviceName
        End If
        
        objPrint.Footer = mstrPageFoot
        Call ReSetTableRows(objPrint)
        If bytMode <> 0 Then zlEvent_Print.zlPrintOrView2Grd objPrint, bytMode
    Else
        Call ReSetTableRows(objPrint)
        zlEvent_Print.zlPrintOrView2Grd objPrint, bytMode
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strItemKey As String
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview:  Call zlRptPrint(0)
    Case conMenu_File_Print:    Call zlRptPrint(1)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_Exit:     Unload Me
    
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh: Call zlRefresh(True)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Err = 0: On Error Resume Next
    Me.lblTitle.Move lngScaleLeft, lngScaleTop + 120, lngScaleRight - lngScaleLeft
    With Me.lblSubhead
        .Left = lngScaleLeft + 210: .Width = lngScaleRight - lngScaleLeft - 210 * 2
        .Top = Me.lblTitle.Top + Me.lblTitle.Height + 120
    End With
    Me.vfgThis.Move lngScaleLeft + 210, Me.lblSubhead.Top + Me.lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    Me.vfgThis.Height = lngScaleBottom - Me.vfgThis.Top - 210
    
    '表上标签分散处理
    Call zlLableBruit
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.vfgThis.Rows > Me.vfgThis.FixedRows)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub Form_Load()
    mblnStartUp = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If mblnChildForm = False Then Call SaveWinState(Me, App.ProductName)
    Set mobjTagFont = Nothing
    Set cbrControl = Nothing
    Set cbrMenuBar = Nothing
    Set cbrToolBar = Nothing
    Set mrsSumCol = Nothing
    Set rsTemp = Nothing
    Set objFont = Nothing
    Set zlEvent_Print = Nothing
        
    Set mobjTitleFont = Nothing
    Set mobjSubFont = Nothing
    Set mobjTagFontPrint = Nothing
    mblnStartUp = False
End Sub

Private Sub vfgThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vfgThis.AutoSize 0, vfgThis.Cols - 1
End Sub

Private Sub zlEvent_Print_zlAfterPrint()
    RaiseEvent zlAfterPrint(mlngFileID)
End Sub

Private Function ReadPageHead(objHead As RichTextBox, ByVal StrKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(12, StrKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Head_S.RTF")
        objHead.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageHead = True
    Else
        objHead.Text = ""
    End If
End Function

Private Function ReadPageFoot(objFoot As RichTextBox, ByVal StrKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(13, StrKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Foot_S.RTF")
        objFoot.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageFoot = True
    Else
        objFoot.Text = ""
    End If
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Private Function UnzipTendPage(ByVal strZipFile As String, ByVal strTarFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    Dim mclsUnzip As New cUnzip
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp ' & "\TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FolderExists(strZipFileTmp) Then
        
        strZipFileName = gobjFSO.GetFile(strZipFileTmp & "\" & strTarFile)
        Call gobjFSO.CopyFile(strZipFileName, "C:\" & strTarFile)
        
        On Error Resume Next
        gobjFSO.DeleteFolder strZipPathTmp, True
        gobjFSO.DeleteFile strZipFile, True
        
        UnzipTendPage = "C:\" & strTarFile
    Else
        UnzipTendPage = ""
    End If
    
    Exit Function
    
errHand:
    Call SaveErrLog
End Function

'-------------------------------------------------------------------
'66724: 刘鹏飞,2014-1-15
'功能：根据控件的长度计算文本占用的行数
Private Function GetData(ByVal strInput As String, Optional ByVal strSplit As String = "'") As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long, lngLen As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.hWnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        lngLen = SendMessage(txtLength.hWnd, EM_GETLINE, lngRow - 1, strLine(0))
        Call ClearArray(strLine, lngLen)
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", strSplit) & strData
    Next
    GetData = Split(GetData, strSplit)
End Function

Private Sub ClearArray(strLine() As Byte, Optional ByVal lngPos As Long = 0)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = lngPos To intMax
        strLine(intDo) = 0
        If lngPos > 0 Then Exit Sub     '不为零,表示仅设置字符串结束符
    Next
    strLine(1) = 1
End Sub

Private Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function


Private Function ReSetTableRows(objsend As zlPrint2Grd) As Boolean
'----------------------------------------------------------------------------------------
'功能：根据纸张大小、表体数据内容重新整理表格的行数
'objSend:打印对象
'调用时机：预览打印之前调用(zlRptPrint)
'说明：该函数里面的算法参照打印部件的处理（有效区计算和输出表格处理），这样才能保持所算及所得
'----------------------------------------------------------------------------------------
    Dim sgnTitle As Single, sgnUpAppRow As Single, sgnDownAppRow As Single, sgnFixRow As Single '标题高度，表上项目高度,表下项目高度，表格固定列高度
    Dim sgnHeight As Single '表体部分有效输出高度
    Dim lngRow As Long, lngCol As Long, lngStartRow As Long
    Dim sgnTmpHeight As Single, sgnRHeight As Single, sgnTextHeight As Single
    Dim arrData, intDatas As Integer, intData As Integer, i As Integer
    Dim arrColText() As String, sgnRowHeight As Single, StrText As String, sgnRowHeightNew As Single
    Dim lngNum As Long '计数变量，用于记录新增的行数
    Dim sgnRowHeightCurrent As Single  '当前表格实际高度  问题号102102
    On Error GoTo errHand
    
    '一:计算实际输出表体的有效高度
    If Not zlGetPrinterSet Then Exit Function
    Set picPrint.Font = objsend.Title.Font
    sgnTitle = picPrint.TextHeight(objsend.Title.Text) + 2 * gconLineHigh
    Set picPrint.Font = objsend.AppFont
    sgnUpAppRow = (picPrint.TextHeight("jg") + gconLineHigh) * objsend.UnderAppRows.Count + gconLineHigh
    sgnDownAppRow = (picPrint.TextHeight("jg") + gconLineHigh) * objsend.BelowAppRows.Count + gconLineHigh
    
    For lngRow = 0 To Me.mshHead.FixedRows - 1
        sgnFixRow = sgnFixRow + Me.mshHead.ROWHEIGHT(lngRow)
    Next lngRow
    sgnHeight = Printer.ScaleHeight - (objsend.EmptyUp + objsend.EmptyDown) * conRatemmToTwip - sgnTitle - sgnUpAppRow - sgnDownAppRow - sgnFixRow - 2 * gconLineHigh
    
    If sgnHeight < vfgThisPrint.RowHeightMin Then ReSetTableRows = True: Exit Function
    
    '二：循环表格内容检查是否超出范围，如果超出将在本行下追加新的一行存放超出内容
    sgnTmpHeight = 0
    lngStartRow = 0
    lngNum = 0
    Set picPrint.Font = mshDetail.Font
PreForStart:
    For lngRow = lngStartRow To Me.mshDetail.Rows - 1
PreBegin:
        '汇总行不处理
        If mshDetail.MergeRow(lngRow) = False Then
            ReDim arrColText(0 To mshDetail.Cols - 1)
            If sgnTmpHeight + Me.mshDetail.ROWHEIGHT(lngRow) > sgnHeight Then
                sgnRHeight = sgnHeight - sgnTmpHeight '可输出的高度
                If sgnRHeight >= vfgThisPrint.RowHeightMin Then
                    '从时间列之后开始
                    sgnRowHeight = 0
                    For lngCol = mintStartCOLCount To mshDetail.Cols - mintEndColCount - 1
                            '获取文本内容占用的行数
                            arrColText(lngCol) = ""
                            With txtLength
                                .Width = mshDetail.ColWidth(lngCol)
                                .Text = Replace(Replace(Replace(mshDetail.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                .FontName = mshDetail.CellFontName
                                .FontSize = mshDetail.CellFontSize
                                .FontBold = mshDetail.CellFontBold
                                .FontItalic = mshDetail.CellFontItalic
                            End With
                            arrData = GetData(txtLength.Text)
                            intDatas = UBound(arrData)
                            '计算并记录超出的文本内容
                            sgnRowHeightCurrent = 0
                            sgnRowHeightCurrent = zlGetCurrentVSFHight(mshDetail.ColWidth(lngCol), intDatas, arrData)
                            If intDatas > 0 And sgnRowHeightCurrent > sgnRHeight Then
                                sgnTextHeight = 0
                                For intData = 0 To intDatas
                                    sgnRowHeightCurrent = zlGetCurrentVSFHight(mshDetail.ColWidth(lngCol), intData, arrData)
                                    If sgnRowHeightCurrent > sgnRHeight Then
                                        If intData = 0 Then GoTo PreEnd
                                        StrText = ""
                                        For i = 0 To intData - 1
                                            StrText = StrText & arrData(i)
                                        Next i
                                        mshDetail.TextMatrix(lngRow, lngCol) = StrText
                                        sgnRowHeightCurrent = zlGetCurrentVSFHight(mshDetail.ColWidth(lngCol), intData - 1, arrData)
                                        If sgnRowHeight < sgnRowHeightCurrent Then
                                            sgnRowHeight = sgnRowHeightCurrent
                                        End If
                                        For i = intData To intDatas
                                            arrColText(lngCol) = arrColText(lngCol) & arrData(i)
                                        Next i
                                        Exit For
                                    End If
                                Next intData
                            End If
                    Next lngCol
                    'sgnRowHeight > 0说明内容超出了可输出的范围
                    If sgnRowHeight > 0 Then
                        If sgnRowHeight < vfgThisPrint.RowHeightMin Then sgnRowHeight = vfgThisPrint.RowHeightMin
                        mshDetail.ROWHEIGHT(lngRow) = sgnRowHeight
                    End If
                    sgnRowHeight = 0
                    '计算超出内容的最大高度
                    For lngCol = mintStartCOLCount To mshDetail.Cols - mintEndColCount - 1
                        If arrColText(lngCol) <> "" Then
                            With txtLength
                                .Width = mshDetail.ColWidth(lngCol)
                                .Text = Replace(Replace(Replace(arrColText(lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                .FontName = mshDetail.CellFontName
                                .FontSize = mshDetail.CellFontSize
                                .FontBold = mshDetail.CellFontBold
                                .FontItalic = mshDetail.CellFontItalic
                            End With
                            
                            arrData = GetData(txtLength.Text)
                            intDatas = UBound(arrData)
                            sgnRowHeightNew = zlGetCurrentVSFHight(mshDetail.ColWidth(lngCol), intDatas, arrData)
                            If sgnRowHeight < sgnRowHeightNew Then
                                sgnRowHeight = sgnRowHeightNew
                            End If
                            
                        End If
                    Next lngCol
                    '完成表格行的添加和赋值
                    If sgnRowHeight > 0 Then
                        '数据跨页后，本条数据在两页都要显示日期、时间、护士、签名人、签名时间、签名日期
                        For i = 0 To mintStartCOLCount - 1
                            arrColText(i) = mshDetail.TextMatrix(lngRow, i)
                        Next i
                        '护士、签名人、签名日期、签名时间的赋值
                        For i = 0 To mintEndColCount - 1
                            arrColText(mshDetail.Cols - i - 1) = mshDetail.TextMatrix(lngRow, mshDetail.Cols - i - 1)
                        Next i
                        
                        If sgnRowHeight < vfgThisPrint.RowHeightMin Then sgnRowHeight = vfgThisPrint.RowHeightMin
                        mshDetail.AddItem "", lngRow + 1: lngNum = lngNum + 1
                        mshDetail.ROWHEIGHT(lngRow + 1) = sgnRowHeight
                        On Error Resume Next
                        For lngCol = 0 To mshDetail.Cols - 1
                            mshDetail.Row = lngRow + 1: mshDetail.Col = lngCol
                            vfgThisPrint.Row = lngRow + vfgThisPrint.FixedRows - (lngNum - 1)
                            vfgThisPrint.Col = lngCol
                            mshDetail.CellForeColor = vfgThisPrint.CellForeColor
                            mshDetail.CellAlignment = vfgThisPrint.ColAlignment(lngCol)
                            mshDetail.TextMatrix(lngRow + 1, lngCol) = arrColText(lngCol)
                        Next lngCol
                        mshDetail.MergeRow(lngRow + 1) = mshDetail.MergeRow(lngRow)
                        If Err <> 0 Then Err.Clear
                        On Error GoTo errHand
                    End If
                    sgnTmpHeight = 0
                    lngStartRow = lngRow + 1
                    GoTo PreForStart
                Else
PreEnd:
                    sgnTmpHeight = 0
                    GoTo PreBegin
                End If
            Else
                sgnTmpHeight = sgnTmpHeight + Me.mshDetail.ROWHEIGHT(lngRow)
            End If
        End If
    Next lngRow
    
    ReSetTableRows = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function zlGetCurrentVSFHight(ByVal ColWith As Long, ByVal Count As Long, ByVal arrData As Variant) As Single
'----------------------------------------------------------------------------------------
'功能：根据数据内容获取当前的行高
'colwith 列宽  count 当前数据行数 arrData 数据内容
'----------------------------------------------------------------------------------------
'102102
    Dim StrText As String
    Dim intData  As Long, intDatas As Long
    Dim i As Long
    vfgThisRowHeight.ColWidth(0) = ColWith
    vfgThisRowHeight.WordWrap = True
    vfgThisRowHeight.AutoSizeMode = flexAutoSizeRowHeight
    If Count < 0 Then Exit Function
    StrText = ""
    For i = 0 To Count
        StrText = StrText & arrData(i)
    Next i
    vfgThisRowHeight.TextMatrix(0, 0) = StrText
    vfgThisRowHeight.AutoSize 0, 0
    zlGetCurrentVSFHight = vfgThisRowHeight.ROWHEIGHT(0)
End Function

Private Function zlGetPrinterSet() As Boolean
    '------------------------------------------------
    '功能：读取本系统注册表的打印缺省设置
    '------------------------------------------------
    Dim iCount As Long
    Dim strDeviceName As String
    Dim intPaperSize As Integer
    Dim intPaperBin As Integer
    Dim intOrientation As Long
    
    If Printers.Count = 0 Then
        zlGetPrinterSet = False
        Exit Function
    End If
    
    strDeviceName = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName)
    If Printer.DeviceName <> strDeviceName Then
        For iCount = 0 To Printers.Count - 1
            If Printers(iCount).DeviceName = strDeviceName Then
                Set Printer = Printers(iCount)
                Exit For
            End If
        Next
    End If
    
    Err = 0
    On Error Resume Next
    Printer.PaperBin = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperBin", Printer.PaperBin)
    Printer.Orientation = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Orientation", Printer.Orientation)
    
    intPaperSize = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperSize", Printer.PaperSize)
    If intPaperSize = 256 Then
        Dim lngWidth As Long
        Dim lngHeight As Long
        
        lngWidth = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Width", Printer.Width)
        lngHeight = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "Height", Printer.Height)
        
        Call SetCustonPager(Me.hWnd, lngWidth, lngHeight)
    Else
        Printer.PaperSize = intPaperSize
    End If

    zlGetPrinterSet = True
End Function

'-----------------------------------------------------------------------------------------------------
