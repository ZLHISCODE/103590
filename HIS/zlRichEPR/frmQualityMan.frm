VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmQualityMan 
   Caption         =   "病历质量审查"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   9615
   Icon            =   "frmQualityMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   9615
   StartUpPosition =   3  '窗口缺省
   Begin XtremeSuiteControls.TabControl tbcThis 
      Height          =   1095
      Left            =   3330
      TabIndex        =   23
      Top             =   810
      Width           =   1005
      _Version        =   589884
      _ExtentX        =   1773
      _ExtentY        =   1931
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   2565
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityMan.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityMan.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityMan.frx":0EBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2085
      Left            =   225
      ScaleHeight     =   2085
      ScaleWidth      =   2325
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2325
      Begin VB.Label lblZZXD 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   22
         Top             =   1755
         Width           =   2580
      End
      Begin VB.Label lblSXWC 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   21
         Top             =   1485
         Width           =   2580
      End
      Begin VB.Label lblSXCS 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   20
         Top             =   1215
         Width           =   2580
      End
      Begin VB.Label lblZZSX 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   19
         Top             =   945
         Width           =   2580
      End
      Begin VB.Label lblRYRS 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   18
         Top             =   675
         Width           =   2580
      End
      Begin VB.Label lblBM 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   17
         Top             =   405
         Width           =   2580
      End
      Begin VB.Label lblMC 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   945
         TabIndex        =   16
         Top             =   135
         Width           =   2580
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "正在修订:"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   1755
         Width           =   870
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "书写完成:"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   1485
         Width           =   870
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "书写超时:"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   1215
         Width           =   870
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "正在书写:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   945
         Width           =   870
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "入院人数:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "科室编码:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   405
         Width           =   870
      End
      Begin VB.Label lblName1 
         BackStyle       =   0  'Transparent
         Caption         =   "科室名称:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   135
         Width           =   870
      End
   End
   Begin VB.PictureBox picDate 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   180
      ScaleHeight     =   1140
      ScaleWidth      =   2325
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   2325
      Begin VB.CommandButton cmdSearch 
         Caption         =   "重新统计(&R)"
         Height          =   350
         Left            =   450
         TabIndex        =   2
         Top             =   720
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   300
         Left            =   450
         TabIndex        =   1
         Top             =   390
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   69009411
         CurrentDate     =   38683
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   300
         Left            =   450
         TabIndex        =   0
         Top             =   45
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   69009411
         CurrentDate     =   38683
      End
      Begin VB.Label lblDateFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "从"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   105
         Width           =   180
      End
      Begin VB.Label lblDateTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   450
         Width           =   180
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5790
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQualityMan.frx":1258
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11880
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
   Begin XtremeSuiteControls.TaskPanel tplThis 
      Height          =   4425
      Left            =   45
      TabIndex        =   7
      Top             =   90
      Width           =   2805
      _Version        =   589884
      _ExtentX        =   4948
      _ExtentY        =   7805
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg住院病历 
      Height          =   1110
      Left            =   6030
      TabIndex        =   24
      Top             =   135
      Width           =   1410
      _cx             =   2487
      _cy             =   1958
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
      Rows            =   3
      Cols            =   6
      FixedRows       =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vfg门诊病历 
      Height          =   1110
      Left            =   4500
      TabIndex        =   25
      Top             =   135
      Width           =   1410
      _cx             =   2487
      _cy             =   1958
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
      Rows            =   3
      Cols            =   6
      FixedRows       =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vfg护理病历 
      Height          =   1110
      Left            =   7560
      TabIndex        =   26
      Top             =   135
      Width           =   1410
      _cx             =   2487
      _cy             =   1958
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
      Rows            =   3
      Cols            =   6
      FixedRows       =   1
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
   Begin VB.Image imgBG 
      Height          =   2295
      Left            =   7290
      Picture         =   "frmQualityMan.frx":1AEA
      Top             =   3465
      Visible         =   0   'False
      Width           =   2265
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   3240
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmQualityMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ID = 0: 编码: 名称: 正在书写: 书写超时: 已完成: 正在修订: 入院: 转入: 正常出院: 转出: 转院: 死亡
End Enum

Private Enum Enum病历种类
    门诊病历 = 1
    住院病历 = 2
    护理病历 = 4
End Enum
Private mvar病历种类 As Enum病历种类

Private Const ID_ViewFile = 802           '查看文件
Private Const ID_ViewPati = 803           '查看病人

Private cbp文件 As CommandBarPopup      '文件菜单
Private cbp视图 As CommandBarPopup      '视图菜单
Private cbp帮助 As CommandBarPopup      '帮助菜单
Private mfrmQualityViewFile As New frmQualityViewFile
Private mfrmQualityViewPati As New frmQualityViewPati

Private Bar常用 As CommandBar           '常用工具栏
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '       strSubhead，打印的副标题
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Select Case mvar病历种类
    Case 门诊病历
        Set objPrint.Body = Me.vfg门诊病历
        objPrint.Title.Text = "门诊病历质量审查"
    Case 住院病历
        Set objPrint.Body = Me.vfg住院病历
        objPrint.Title.Text = "住院病历质量审查"
    Case 护理病历
        Set objPrint.Body = Me.vfg护理病历
        objPrint.Title.Text = "护理病历质量审查"
    End Select
    objPrint.Title.Font.Name = "黑体"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("统计时间:" & Format(Me.dtpDateFrom.Value, "yyyy-MM-dd") & " ～ " & Format(Me.dtpDateTo.Value, "yyyy-MM-dd"))
    Call objPrint.UnderAppRows.Add(objAppRow)
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub FillGrid(ByVal strFrom As String, ByVal strTo As String)
    '填充数据
    Dim Rs As ADODB.Recordset, i As Long, lngCount(1 To 10) As Long
    Select Case mvar病历种类
    Case 门诊病历
        gstrSQL = "Select d.ID, d.编码, d.名称, l.正在书写, l.书写超时, l.已完成 " & _
            " From 部门表 d," & _
            "          (Select 科室id, Sum(Decode(完成时间, Null, 1, 0)) As 正在书写, " & _
            "                             Sum(Decode(完成时间, Null, Decode(Sign(Sysdate - 创建时间 - 1), 1, 1, 0), 0)) As 书写超时, " & _
            "                             Sum(Decode(完成时间, Null, 0, 1)) As 已完成 " & _
            "              From 电子病历记录 " & _
            "              Where 病历种类 = 1 And 创建时间 Between [1] And [2] " & _
            "              Group By 科室id) l " & _
            " Where D.ID = L.科室id " & _
            " Order By d.编码"
        Set Rs = OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(strFrom, "YYYY-MM-DD")), CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"))
        Call InitGrid(mvar病历种类)
        Me.vfg门诊病历.Rows = 2 + Rs.RecordCount
        stbThis.Panels(2).Text = "共计：" & Rs.RecordCount & "条记录。"
        i = 1
        Do While Not Rs.EOF
            With Me.vfg门诊病历
                .TextMatrix(i, mCol.ID) = NVL(Rs("ID"))
                .TextMatrix(i, mCol.编码) = NVL(Rs("编码"))
                .TextMatrix(i, mCol.名称) = NVL(Rs("名称"))
                .TextMatrix(i, mCol.正在书写) = IIf(NVL(Rs("正在书写")) = 0, "", NVL(Rs("正在书写"))): lngCount(1) = lngCount(1) + Val(.TextMatrix(i, mCol.正在书写))
                .TextMatrix(i, mCol.书写超时) = IIf(NVL(Rs("书写超时")) = 0, "", NVL(Rs("书写超时"))): lngCount(2) = lngCount(2) + Val(.TextMatrix(i, mCol.书写超时))
                .TextMatrix(i, mCol.已完成) = IIf(NVL(Rs("已完成")) = 0, "", NVL(Rs("已完成"))): lngCount(3) = lngCount(3) + Val(.TextMatrix(i, mCol.已完成))
            End With
            Rs.MoveNext
            i = i + 1
        Loop
        With Me.vfg门诊病历
            .TextMatrix(i, mCol.编码) = "合计"
            .TextMatrix(i, mCol.名称) = ""
            .TextMatrix(i, mCol.正在书写) = lngCount(1)
            .TextMatrix(i, mCol.书写超时) = lngCount(2)
            .TextMatrix(i, mCol.已完成) = lngCount(3)
        End With
        Rs.Close
        Set Rs = Nothing
        If Me.vfg门诊病历.Rows > 1 Then Me.vfg门诊病历.Row = 1
    Case 住院病历
        gstrSQL = "Select d.ID, d.编码, d.名称, l.正在书写, l.书写超时, l.已完成, l.正在修订, p.入院, t.转入, e.正常出院, e.死亡, e.转院, i.转出 " & _
            " From 部门表 d, (Select 出院科室id, Count(*) As 入院 From 病案主页 Where 入院日期 Between [1] And [2] Group By 出院科室id) p, " & _
            "          (Select 科室id, Sum(Decode(完成时间, Null, 1, 0)) As 正在书写, " & _
            "                             Sum(Decode(完成时间, Null, Decode(Sign(Sysdate - 创建时间 - 1), 1, 1, 0), 0)) As 书写超时, " & _
            "                             Sum(Decode(完成时间, Null, 0, 1)) As 已完成, " & _
            "                             Sum(Decode(完成时间, Null, 0, Decode(NVL(签名级别, 0), 0, 1, 0))) As 正在修订 " & _
            "              From 电子病历记录 " & _
            "              Where  病历种类 = 2 And 创建时间 Between [1] And [2] " & _
            "              Group By 科室id) l, " & _
            "          (Select b.Id, Count((a.病人id)) As 转入 " & _
            "              From 病人变动记录 a, 部门表 b " & _
            "              Where a.开始时间 Between [1] And [2] And a.开始原因 = 3 And Nvl(附加床位, 0) = 0 And b.Id = a.科室id " & _
            "              Group By b.Id) t, " & _
            "          (Select a.Id, Sum(Decode(b.出院方式, '正常', 1, 0)) As 正常出院, Sum(Decode(b.出院方式, '死亡', 1, 0)) As 死亡, " & _
            "                             Sum(Decode(b.出院方式, '转院', 1, 0)) As 转院 " & _
            "              From 部门表 a, 病案主页 b " & _
            "              Where b.出院日期 Between [1] And [2] And b.出院科室id = a.Id " & _
            "              Group By a.Id) e, " & _
            "          (Select b.Id, Count((a.病人id)) As 转出 " & _
            "              From 病人变动记录 a, 部门表 b " & _
            "              Where a.终止时间 Between [1] And [2] And a.终止原因 = 3 And Nvl(附加床位, 0) = 0 And b.Id = a.科室id " & _
            "              Group By b.Id) i " & _
            " Where d.Id = p.出院科室id And p.出院科室id = l.科室id(+) And t.Id = l.科室id And e.Id = l.科室id And i.Id = l.科室id " & _
            " Order By d.编码"
        Set Rs = OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(strFrom, "YYYY-MM-DD")), CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"))
        Call InitGrid(mvar病历种类)
        Me.vfg住院病历.Rows = 3 + Rs.RecordCount
        stbThis.Panels(2).Text = "共计：" & Rs.RecordCount & "条记录。"
        i = 2
        Do While Not Rs.EOF
            With Me.vfg住院病历
                .TextMatrix(i, mCol.ID) = NVL(Rs("ID"))
                .TextMatrix(i, mCol.编码) = NVL(Rs("编码"))
                .TextMatrix(i, mCol.名称) = NVL(Rs("名称"))
                .TextMatrix(i, mCol.正在书写) = IIf(NVL(Rs("正在书写")) = 0, "", NVL(Rs("正在书写"))): lngCount(1) = lngCount(1) + Val(.TextMatrix(i, mCol.正在书写))
                .TextMatrix(i, mCol.书写超时) = IIf(NVL(Rs("书写超时")) = 0, "", NVL(Rs("书写超时"))): lngCount(2) = lngCount(2) + Val(.TextMatrix(i, mCol.书写超时))
                .TextMatrix(i, mCol.已完成) = IIf(NVL(Rs("已完成")) = 0, "", NVL(Rs("已完成"))): lngCount(3) = lngCount(3) + Val(.TextMatrix(i, mCol.已完成))
                .TextMatrix(i, mCol.正在修订) = IIf(NVL(Rs("正在修订")) = 0, "", NVL(Rs("正在修订"))): lngCount(4) = lngCount(4) + Val(.TextMatrix(i, mCol.正在修订))
                .TextMatrix(i, mCol.入院) = IIf(NVL(Rs("入院")) = 0, "", NVL(Rs("入院"))): lngCount(5) = lngCount(5) + Val(.TextMatrix(i, mCol.入院))
                .TextMatrix(i, mCol.转入) = IIf(NVL(Rs("转入")) = 0, "", NVL(Rs("转入"))): lngCount(6) = lngCount(6) + Val(.TextMatrix(i, mCol.转入))
                .TextMatrix(i, mCol.正常出院) = IIf(NVL(Rs("正常出院")) = 0, "", NVL(Rs("正常出院"))): lngCount(7) = lngCount(7) + Val(.TextMatrix(i, mCol.正常出院))
                .TextMatrix(i, mCol.转出) = IIf(NVL(Rs("转出")) = 0, "", NVL(Rs("转出"))): lngCount(8) = lngCount(8) + Val(.TextMatrix(i, mCol.转出))
                .TextMatrix(i, mCol.转院) = IIf(NVL(Rs("转院")) = 0, "", NVL(Rs("转院"))): lngCount(9) = lngCount(9) + Val(.TextMatrix(i, mCol.转院))
                .TextMatrix(i, mCol.死亡) = IIf(NVL(Rs("死亡")) = 0, "", NVL(Rs("死亡"))): lngCount(10) = lngCount(10) + Val(.TextMatrix(i, mCol.死亡))
            End With
            Rs.MoveNext
            i = i + 1
        Loop
        With Me.vfg住院病历
            .TextMatrix(i, mCol.编码) = "合计"
            .TextMatrix(i, mCol.名称) = ""
            .TextMatrix(i, mCol.正在书写) = lngCount(1)
            .TextMatrix(i, mCol.书写超时) = lngCount(2)
            .TextMatrix(i, mCol.已完成) = lngCount(3)
            .TextMatrix(i, mCol.正在修订) = lngCount(4)
            .TextMatrix(i, mCol.入院) = lngCount(5)
            .TextMatrix(i, mCol.转入) = lngCount(6)
            .TextMatrix(i, mCol.正常出院) = lngCount(7)
            .TextMatrix(i, mCol.转出) = lngCount(8)
            .TextMatrix(i, mCol.转院) = lngCount(9)
            .TextMatrix(i, mCol.死亡) = lngCount(10)
            .Cell(flexcpFontBold, i, mCol.编码) = True
            .Cell(flexcpFontBold, i, mCol.名称) = True
            .Cell(flexcpFontBold, i, mCol.正在书写) = True
            .Cell(flexcpFontBold, i, mCol.书写超时) = True
            .Cell(flexcpFontBold, i, mCol.已完成) = True
            .Cell(flexcpFontBold, i, mCol.正在修订) = True
            .Cell(flexcpFontBold, i, mCol.入院) = True
            .Cell(flexcpFontBold, i, mCol.转入) = True
            .Cell(flexcpFontBold, i, mCol.正常出院) = True
            .Cell(flexcpFontBold, i, mCol.转出) = True
            .Cell(flexcpFontBold, i, mCol.转院) = True
            .Cell(flexcpFontBold, i, mCol.死亡) = True
        End With
        Rs.Close
        Set Rs = Nothing
        If Me.vfg住院病历.Rows > 2 Then Me.vfg住院病历.Row = 2: Call vfg住院病历_RowColChange
    Case 护理病历
        gstrSQL = "Select d.ID, d.编码, d.名称, l.正在书写, l.书写超时, l.已完成, l.正在修订, p.入院, t.转入, e.正常出院, e.死亡, e.转院, i.转出 " & _
            " From 部门表 d, (Select 出院科室id, Count(*) As 入院 From 病案主页 Where 入院日期 Between [1] And [2] Group By 出院科室id) p, " & _
            "          (Select 科室id, Sum(Decode(完成时间, Null, 1, 0)) As 正在书写, " & _
            "                             Sum(Decode(完成时间, Null, Decode(Sign(Sysdate - 创建时间 - 1), 1, 1, 0), 0)) As 书写超时, " & _
            "                             Sum(Decode(完成时间, Null, 0, 1)) As 已完成, " & _
            "                             Sum(Decode(完成时间, Null, 0, Decode(NVL(签名级别, 0), 0, 1, 0))) As 正在修订 " & _
            "              From 电子病历记录 " & _
            "              Where  病历种类 = 4 And 创建时间 Between [1] And [2] " & _
            "              Group By 科室id) l, " & _
            "          (Select b.Id, Count((a.病人id)) As 转入 " & _
            "              From 病人变动记录 a, 部门表 b " & _
            "              Where a.开始时间 Between [1] And [2] And a.开始原因 = 3 And Nvl(附加床位, 0) = 0 And b.Id = a.科室id " & _
            "              Group By b.Id) t, " & _
            "          (Select a.Id, Sum(Decode(b.出院方式, '正常', 1, 0)) As 正常出院, Sum(Decode(b.出院方式, '死亡', 1, 0)) As 死亡, " & _
            "                             Sum(Decode(b.出院方式, '转院', 1, 0)) As 转院 " & _
            "              From 部门表 a, 病案主页 b " & _
            "              Where b.出院日期 Between [1] And [2] And b.出院科室id = a.Id " & _
            "              Group By a.Id) e, " & _
            "          (Select b.Id, Count((a.病人id)) As 转出 " & _
            "              From 病人变动记录 a, 部门表 b " & _
            "              Where a.终止时间 Between [1] And [2] And a.终止原因 = 3 And Nvl(附加床位, 0) = 0 And b.Id = a.科室id " & _
            "              Group By b.Id) i " & _
            " Where d.Id = p.出院科室id And p.出院科室id = l.科室id(+) And t.Id = l.科室id And e.Id = l.科室id And i.Id = l.科室id " & _
            " Order By d.编码"
        Set Rs = OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(strFrom, "YYYY-MM-DD")), CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"))
        Call InitGrid(mvar病历种类)
        Me.vfg护理病历.Rows = 3 + Rs.RecordCount
        stbThis.Panels(2).Text = "共计：" & Rs.RecordCount & "条记录。"
        i = 2
        Do While Not Rs.EOF
            With Me.vfg护理病历
                .TextMatrix(i, mCol.ID) = NVL(Rs("ID"))
                .TextMatrix(i, mCol.编码) = NVL(Rs("编码"))
                .TextMatrix(i, mCol.名称) = NVL(Rs("名称"))
                .TextMatrix(i, mCol.正在书写) = IIf(NVL(Rs("正在书写")) = 0, "", NVL(Rs("正在书写"))): lngCount(1) = lngCount(1) + Val(.TextMatrix(i, mCol.正在书写))
                .TextMatrix(i, mCol.书写超时) = IIf(NVL(Rs("书写超时")) = 0, "", NVL(Rs("书写超时"))): lngCount(2) = lngCount(2) + Val(.TextMatrix(i, mCol.书写超时))
                .TextMatrix(i, mCol.已完成) = IIf(NVL(Rs("已完成")) = 0, "", NVL(Rs("已完成"))): lngCount(3) = lngCount(3) + Val(.TextMatrix(i, mCol.已完成))
                .TextMatrix(i, mCol.正在修订) = IIf(NVL(Rs("正在修订")) = 0, "", NVL(Rs("正在修订"))): lngCount(4) = lngCount(4) + Val(.TextMatrix(i, mCol.正在修订))
                .TextMatrix(i, mCol.入院) = IIf(NVL(Rs("入院")) = 0, "", NVL(Rs("入院"))): lngCount(5) = lngCount(5) + Val(.TextMatrix(i, mCol.入院))
                .TextMatrix(i, mCol.转入) = IIf(NVL(Rs("转入")) = 0, "", NVL(Rs("转入"))): lngCount(6) = lngCount(6) + Val(.TextMatrix(i, mCol.转入))
                .TextMatrix(i, mCol.正常出院) = IIf(NVL(Rs("正常出院")) = 0, "", NVL(Rs("正常出院"))): lngCount(7) = lngCount(7) + Val(.TextMatrix(i, mCol.正常出院))
                .TextMatrix(i, mCol.转出) = IIf(NVL(Rs("转出")) = 0, "", NVL(Rs("转出"))): lngCount(8) = lngCount(8) + Val(.TextMatrix(i, mCol.转出))
                .TextMatrix(i, mCol.转院) = IIf(NVL(Rs("转院")) = 0, "", NVL(Rs("转院"))): lngCount(9) = lngCount(9) + Val(.TextMatrix(i, mCol.转院))
                .TextMatrix(i, mCol.死亡) = IIf(NVL(Rs("死亡")) = 0, "", NVL(Rs("死亡"))): lngCount(10) = lngCount(10) + Val(.TextMatrix(i, mCol.死亡))
            End With
            Rs.MoveNext
            i = i + 1
        Loop
        With Me.vfg护理病历
            .TextMatrix(i, mCol.编码) = "合计"
            .TextMatrix(i, mCol.名称) = ""
            .TextMatrix(i, mCol.正在书写) = lngCount(1)
            .TextMatrix(i, mCol.书写超时) = lngCount(2)
            .TextMatrix(i, mCol.已完成) = lngCount(3)
            .TextMatrix(i, mCol.正在修订) = lngCount(4)
            .TextMatrix(i, mCol.入院) = lngCount(5)
            .TextMatrix(i, mCol.转入) = lngCount(6)
            .TextMatrix(i, mCol.正常出院) = lngCount(7)
            .TextMatrix(i, mCol.转出) = lngCount(8)
            .TextMatrix(i, mCol.转院) = lngCount(9)
            .TextMatrix(i, mCol.死亡) = lngCount(10)
            .Cell(flexcpFontBold, i, mCol.编码) = True
            .Cell(flexcpFontBold, i, mCol.名称) = True
            .Cell(flexcpFontBold, i, mCol.正在书写) = True
            .Cell(flexcpFontBold, i, mCol.书写超时) = True
            .Cell(flexcpFontBold, i, mCol.已完成) = True
            .Cell(flexcpFontBold, i, mCol.正在修订) = True
            .Cell(flexcpFontBold, i, mCol.入院) = True
            .Cell(flexcpFontBold, i, mCol.转入) = True
            .Cell(flexcpFontBold, i, mCol.正常出院) = True
            .Cell(flexcpFontBold, i, mCol.转出) = True
            .Cell(flexcpFontBold, i, mCol.转院) = True
            .Cell(flexcpFontBold, i, mCol.死亡) = True
        End With
        Rs.Close
        Set Rs = Nothing
        If Me.vfg护理病历.Rows > 2 Then Me.vfg护理病历.Row = 2
    End Select
End Sub

Private Sub InitGrid(ByVal 病历种类 As Enum病历种类)
    Dim i As Long
    Select Case 病历种类
    Case 门诊病历
        With Me.vfg门诊病历
            .Clear
            .Rows = 1
            .FixedRows = 1
            .Cols = 6
            .RowHeightMin = 300
            .WallPaper = imgBG.Picture
            .WallPaperAlignment = flexPicAlignRightBottom
            
    '        .BackColorAlternate = RGB(240, 240, 255)
            .BackColorSel = RGB(125, 125, 255)
            .ForeColorSel = vbWhite
            .Sort = flexSortCustom
            
            .TextMatrix(0, mCol.编码) = "编码"
            .TextMatrix(0, mCol.名称) = "名称"
            .TextMatrix(0, mCol.正在书写) = "正在书写病历数"
            .TextMatrix(0, mCol.书写超时) = "其中超时书写数"
            .TextMatrix(0, mCol.已完成) = "已完成病历数"
            
            For i = 0 To 5
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
            Next
            
            .ColWidth(mCol.ID) = 0
            .ColWidth(mCol.编码) = 600
            .ColWidth(mCol.名称) = 2500
            .ColWidth(mCol.正在书写) = 1600
            .ColWidth(mCol.书写超时) = 1600
            .ColWidth(mCol.已完成) = 1600
        End With
    Case 护理病历
        With Me.vfg护理病历
            .Clear
            .Rows = 3
            .FixedRows = 2
            .Cols = 13
            .RowHeightMin = 300
            .WallPaper = imgBG.Picture
            .WallPaperAlignment = flexPicAlignRightBottom
            
    '        .BackColorAlternate = RGB(240, 240, 255)
            .BackColorSel = RGB(125, 125, 255)
            .ForeColorSel = vbWhite
            .Sort = flexSortCustom
            
            .TextMatrix(0, mCol.编码) = "住院科室"
            .TextMatrix(0, mCol.名称) = "住院科室"
            .TextMatrix(0, mCol.正在书写) = "正在书写病历（份）"
            .TextMatrix(0, mCol.书写超时) = "正在书写病历（份）"
            .TextMatrix(0, mCol.已完成) = "已完成病历（份）"
            .TextMatrix(0, mCol.正在修订) = "已完成病历（份）"
            .TextMatrix(0, mCol.入院) = "增加人次"
            .TextMatrix(0, mCol.转入) = "增加人次"
            .TextMatrix(0, mCol.正常出院) = "减少人次"
            .TextMatrix(0, mCol.转出) = "减少人次"
            .TextMatrix(0, mCol.转院) = "减少人次"
            .TextMatrix(0, mCol.死亡) = "减少人次"
            .TextMatrix(1, mCol.编码) = "编码"
            .TextMatrix(1, mCol.名称) = "名称"
            .TextMatrix(1, mCol.正在书写) = "总数"
            .TextMatrix(1, mCol.书写超时) = "超时24小时"
            .TextMatrix(1, mCol.已完成) = "总数"
            .TextMatrix(1, mCol.正在修订) = "正在修订"
            .TextMatrix(1, mCol.入院) = "入院"
            .TextMatrix(1, mCol.转入) = "他科转入"
            .TextMatrix(1, mCol.正常出院) = "正常出院"
            .TextMatrix(1, mCol.转出) = "转出"
            .TextMatrix(1, mCol.转院) = "转院"
            .TextMatrix(1, mCol.死亡) = "死亡"
            
            .MergeRow(0) = True
            .MergeCells = flexMergeRestrictRows
            
            For i = 0 To 12
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
                .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
            Next
            
            .ColWidth(mCol.ID) = 0
            .ColWidth(mCol.编码) = 600
            .ColWidth(mCol.名称) = 1600
            .ColWidth(mCol.正在书写) = 600
            .ColWidth(mCol.书写超时) = 1200
            .ColWidth(mCol.已完成) = 600
            .ColWidth(mCol.正在修订) = 1200
            .ColWidth(mCol.入院) = 900
            .ColWidth(mCol.转入) = 900
            .ColWidth(mCol.正常出院) = 900
            .ColWidth(mCol.转出) = 600
            .ColWidth(mCol.转院) = 600
            .ColWidth(mCol.死亡) = 600
        End With
    Case 住院病历
        With Me.vfg住院病历
            .Clear
            .Rows = 3
            .FixedRows = 2
            .Cols = 13
            .RowHeightMin = 300
            .WallPaper = imgBG.Picture
            .WallPaperAlignment = flexPicAlignRightBottom
            
    '        .BackColorAlternate = RGB(240, 240, 255)
            .BackColorSel = RGB(125, 125, 255)
            .ForeColorSel = vbWhite
            .Sort = flexSortCustom
            
            .TextMatrix(0, mCol.编码) = "住院科室"
            .TextMatrix(0, mCol.名称) = "住院科室"
            .TextMatrix(0, mCol.正在书写) = "正在书写病历（份）"
            .TextMatrix(0, mCol.书写超时) = "正在书写病历（份）"
            .TextMatrix(0, mCol.已完成) = "已完成病历（份）"
            .TextMatrix(0, mCol.正在修订) = "已完成病历（份）"
            .TextMatrix(0, mCol.入院) = "增加人次"
            .TextMatrix(0, mCol.转入) = "增加人次"
            .TextMatrix(0, mCol.正常出院) = "减少人次"
            .TextMatrix(0, mCol.转出) = "减少人次"
            .TextMatrix(0, mCol.转院) = "减少人次"
            .TextMatrix(0, mCol.死亡) = "减少人次"
            .TextMatrix(1, mCol.编码) = "编码"
            .TextMatrix(1, mCol.名称) = "名称"
            .TextMatrix(1, mCol.正在书写) = "总数"
            .TextMatrix(1, mCol.书写超时) = "超时24小时"
            .TextMatrix(1, mCol.已完成) = "总数"
            .TextMatrix(1, mCol.正在修订) = "正在修订"
            .TextMatrix(1, mCol.入院) = "入院"
            .TextMatrix(1, mCol.转入) = "他科转入"
            .TextMatrix(1, mCol.正常出院) = "正常出院"
            .TextMatrix(1, mCol.转出) = "转出"
            .TextMatrix(1, mCol.转院) = "转院"
            .TextMatrix(1, mCol.死亡) = "死亡"
            
            .MergeRow(0) = True
            .MergeCells = flexMergeRestrictRows
            
            For i = 0 To 12
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
                .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
            Next
            
            .ColWidth(mCol.ID) = 0
            .ColWidth(mCol.编码) = 600
            .ColWidth(mCol.名称) = 1600
            .ColWidth(mCol.正在书写) = 600
            .ColWidth(mCol.书写超时) = 1200
            .ColWidth(mCol.已完成) = 600
            .ColWidth(mCol.正在修订) = 1200
            .ColWidth(mCol.入院) = 900
            .ColWidth(mCol.转入) = 900
            .ColWidth(mCol.正常出院) = 900
            .ColWidth(mCol.转出) = 600
            .ColWidth(mCol.转院) = 600
            .ColWidth(mCol.死亡) = 600
        End With
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        Call cmdSearch_Click
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case ID_ViewFile
        If mvar病历种类 = 住院病历 Then
            If Me.vfg住院病历.Row = 0 Or Me.vfg住院病历.Row = Me.vfg住院病历.Rows - 1 Then Exit Sub
            mfrmQualityViewFile.ShowMe mvar病历种类, Me, Me.vfg住院病历.TextMatrix(Me.vfg住院病历.Row, mCol.ID), Me.vfg住院病历.TextMatrix(Me.vfg住院病历.Row, mCol.名称), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        ElseIf mvar病历种类 = 门诊病历 Then
            If Me.vfg门诊病历.Row = 0 Or Me.vfg门诊病历.Row = Me.vfg门诊病历.Rows - 1 Then Exit Sub
            mfrmQualityViewFile.ShowMe mvar病历种类, Me, Me.vfg门诊病历.TextMatrix(Me.vfg门诊病历.Row, mCol.ID), Me.vfg门诊病历.TextMatrix(Me.vfg门诊病历.Row, mCol.名称), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        Else
            If Me.vfg护理病历.Row = 0 Or Me.vfg护理病历.Row = Me.vfg护理病历.Rows - 1 Then Exit Sub
            mfrmQualityViewFile.ShowMe mvar病历种类, Me, Me.vfg护理病历.TextMatrix(Me.vfg护理病历.Row, mCol.ID), Me.vfg护理病历.TextMatrix(Me.vfg护理病历.Row, mCol.名称), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        End If
    Case ID_ViewPati
        If mvar病历种类 = 住院病历 Then
            If Me.vfg住院病历.Row = 0 Or Me.vfg住院病历.Row = Me.vfg住院病历.Rows - 1 Then Exit Sub
            mfrmQualityViewPati.ShowMe mvar病历种类, Me, Me.vfg住院病历.TextMatrix(Me.vfg住院病历.Row, mCol.ID), Me.vfg住院病历.TextMatrix(Me.vfg住院病历.Row, mCol.名称), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        ElseIf mvar病历种类 = 门诊病历 Then
            If Me.vfg门诊病历.Row = 0 Or Me.vfg门诊病历.Row = Me.vfg门诊病历.Rows - 1 Then Exit Sub
            mfrmQualityViewPati.ShowMe mvar病历种类, Me, Me.vfg门诊病历.TextMatrix(Me.vfg门诊病历.Row, mCol.ID), Me.vfg门诊病历.TextMatrix(Me.vfg门诊病历.Row, mCol.名称), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        Else
            If Me.vfg护理病历.Row = 0 Or Me.vfg护理病历.Row = Me.vfg护理病历.Rows - 1 Then Exit Sub
            mfrmQualityViewPati.ShowMe mvar病历种类, Me, Me.vfg护理病历.TextMatrix(Me.vfg护理病历.Row, mCol.ID), Me.vfg护理病历.TextMatrix(Me.vfg护理病历.Row, mCol.名称), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        End If
    End Select
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    Me.cbsThis.GetClientRect Left, Top, Right, Bottom
    With Me.tplThis
        .Left = Left: .Width = 3050
        .Top = Top: .Height = Bottom - Top - stbThis.Height
    End With
    Me.tbcThis.Move Me.tplThis.Width, Top, Right - Left - Me.tplThis.Width, Me.tplThis.Height
'    vfg门诊病历.Move 0, 0, tbcThis.Width, tbcThis.Height
'    vfg住院病历.Move 0, 0, tbcThis.Width, tbcThis.Height
'    vfg护理病历.Move 0, 0, tbcThis.Width, tbcThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.vfg住院病历.Rows <> 0)
    Case conMenu_View_Jump '跳转
        If Me.tbcThis.Selected.Index + 1 <= Me.tbcThis.ItemCount - 1 Then
            Me.tbcThis.Item(Me.tbcThis.Selected.Index + 1).Selected = True
        Else
            Me.tbcThis.Item(0).Selected = True
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case ID_ViewFile:
        Select Case mvar病历种类
        Case 门诊病历
            Control.Enabled = Me.vfg门诊病历.Row > 0 And Me.vfg门诊病历.Row < Me.vfg门诊病历.Rows - 1
        Case 住院病历
            Control.Enabled = Me.vfg住院病历.Row > 1 And Me.vfg住院病历.Row < Me.vfg住院病历.Rows - 1
        Case 护理病历
            Control.Enabled = Me.vfg护理病历.Row > 1 And Me.vfg护理病历.Row < Me.vfg护理病历.Rows - 1
        End Select
    Case ID_ViewPati:
        Select Case mvar病历种类
        Case 门诊病历
            Control.Enabled = False 'Me.vfg门诊病历.Row > 0 And Me.vfg门诊病历.Row < Me.vfg门诊病历.Rows - 1
        Case 住院病历
            Control.Enabled = Me.vfg住院病历.Row > 1 And Me.vfg住院病历.Row < Me.vfg住院病历.Rows - 1
        Case 护理病历
            Control.Enabled = Me.vfg护理病历.Row > 1 And Me.vfg护理病历.Row < Me.vfg护理病历.Rows - 1
        End Select
    End Select
End Sub

Private Sub cmdSearch_Click()
    FillGrid Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
End Sub

Private Sub dtpDateFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpDateTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    Dim cbpPopup As CommandBarPopup                     '临时对象
    Dim cbpPopupSub As CommandBarPopup                  '临时对象
    Dim objControl As CommandBarControl                 '工具栏控件
    Dim objCustControl As CommandBarControlCustom       '自定义控件
    Dim Combo As CommandBarComboBox                     '工具栏下拉框控件
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    
    Set cbp文件 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "文件(&F)")
    With cbp文件.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "导出到&Excel")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With
    
    Set cbp视图 = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbp视图.ID = conMenu_ViewPopup
    With cbp视图.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set objControl = .Add(xtpControlButton, ID_ViewFile, "文件列表视图"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_ViewPati, "初入院病人视图")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
    End With
    
    Set cbp帮助 = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbp帮助.ID = conMenu_HelpPopup
    With cbp帮助.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    Set Bar常用 = cbsThis.Add("常用工具栏", xtpBarTop)
    With Bar常用.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_ViewFile, "文件列表视图"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_ViewPati, "初入院病人视图")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.BeginGroup = True
    End With
    For Each objControl In Bar常用.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '热键绑定
    cbsThis.KeyBindings.Add FCONTROL, Asc("Q"), conMenu_File_Exit
    cbsThis.KeyBindings.Add FCONTROL, Asc("P"), conMenu_File_Print
    cbsThis.KeyBindings.Add 0, vbKeyF5, conMenu_View_Refresh
    cbsThis.KeyBindings.Add 0, vbKeyF6, conMenu_View_Jump  '跳转
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet '打印设置
        .AddHiddenCommand conMenu_File_Excel '输出到Excel
        .AddHiddenCommand conMenu_View_Jump '跳转
    End With

    Set Group = tplThis.Groups.Add(0, "统计条件")
    Group.Special = True
    Set Item = Group.Items.Add(0, "书写日期范围:", xtpTaskItemTypeText)
    Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
    Set Item.Control = Me.picDate
    Me.picDate.BackColor = Item.BackColor
        
    Set Group = tplThis.Groups.Add(0, "常见任务")
    Group.Tooltip = "常见任务列表"
    Group.Items.Add conMenu_File_Excel, "导出到Excel", xtpTaskItemTypeLink, 1
    Group.Items.Add conMenu_File_Preview, "打印预览", xtpTaskItemTypeLink, 2
    Group.Items.Add conMenu_File_Print, "打印...", xtpTaskItemTypeLink, 3
    
    Set Group = tplThis.Groups.Add(0, "统计信息")
    Group.Tooltip = "统计结果汇总"
    Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
    Set Item.Control = Me.picInfo
    Me.picInfo.BackColor = Item.BackColor
    
    With Me.tbcThis
        .RemoveAll
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = False
            .ShowIcons = True
        End With
        .InsertItem(0, "门诊病历", vfg门诊病历.hWnd, 0).Tag = "门诊病历"
        .InsertItem(1, "住院病历", vfg住院病历.hWnd, 0).Tag = "住院病历"
        .InsertItem(2, "护理病历", vfg护理病历.hWnd, 0).Tag = "护理病历"
    End With
    
    tplThis.SetImageList imlTaskPanelIcons
    Call RestoreWinState(Me)
    
    '-----------------------------------------------------
    '基本数据装入:
    Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select Sysdate From Dual"
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption)
    With Me.dtpDateTo
        .Value = Format(rsTemp.Fields(0).Value, "yyyy-MM-dd")
        .MaxDate = .Value: .MinDate = Format("1990-01-01", "yyyy-MM-dd")
    End With
    With Me.dtpDateFrom
        .Value = Me.dtpDateTo.Value - 7
        .MaxDate = Me.dtpDateTo.MaxDate: .MinDate = Me.dtpDateTo.MinDate
    End With
    rsTemp.Close
    Set rsTemp = Nothing
    
    mvar病历种类 = 门诊病历
    
    Call cmdSearch_Click        '填入初始数据
End Sub

Private Sub dtpDateTo_Validate(Cancel As Boolean)
    Me.dtpDateFrom.MaxDate = Me.dtpDateTo.Value
    If Me.dtpDateFrom.Value > Me.dtpDateFrom.MaxDate Then Me.dtpDateFrom.Value = Me.dtpDateFrom.MaxDate
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsThis_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me)
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If tbcThis.Tag <> "" Then Exit Sub
    Select Case Item.Tag
    Case "门诊病历"
        mvar病历种类 = 门诊病历
    Case "住院病历"
        mvar病历种类 = 住院病历
    Case "护理病历"
        mvar病历种类 = 护理病历
    End Select
    Call cmdSearch_Click
End Sub

Private Sub tplThis_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Select Case Item.ID
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    End Select
End Sub

Private Sub vfg护理病历_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tbcThis.Tag = "Moving"
    tbcThis.Item(0).Selected = True
    tbcThis.Item(2).Selected = True
    tbcThis.Tag = ""
End Sub

Private Sub vfg门诊病历_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tbcThis.Tag = "Moving"
    tbcThis.Item(2).Selected = True
    tbcThis.Item(0).Selected = True
    tbcThis.Tag = ""
End Sub

Private Sub vfg住院病历_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tbcThis.Tag = "Moving"
    tbcThis.Item(0).Selected = True
    tbcThis.Item(1).Selected = True
    tbcThis.Tag = ""
End Sub

Private Sub vfg住院病历_RowColChange()
    Dim i As Long
    With vfg住院病历
        i = .Row
        If i > 1 And i < .Rows - 1 Then
            lblMC = .TextMatrix(i, 1)
            lblBM = .TextMatrix(i, 0)
            lblRYRS = Val(.TextMatrix(i, 6)) & " 人"
            lblZZSX = Val(.TextMatrix(i, 2)) & " 份"
            lblSXCS = Val(.TextMatrix(i, 3)) & " 份"
            lblSXWC = Val(.TextMatrix(i, 4)) & " 份"
            lblZZXD = Val(.TextMatrix(i, 5)) & " 份"
        Else
            lblMC = "-"
            lblBM = "-"
            lblRYRS = "-" & " 人"
            lblZZSX = "-" & " 份"
            lblSXCS = "-" & " 份"
            lblSXWC = "-" & " 份"
            lblZZXD = "-" & " 份"
        End If
    End With
    If mvar病历种类 = 门诊病历 Then
        lblZZXD.Visible = False
    Else
        lblZZXD.Visible = True
    End If
End Sub
