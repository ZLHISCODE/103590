VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMediPriceBatch 
   Caption         =   "批量执行调价"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "frmMediPriceBatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   8415
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic提示 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   8175
      TabIndex        =   1
      Top             =   840
      Width           =   8175
      Begin VB.Label Label1 
         Caption         =   "以下内容为执行了调价，但是没有生效的调价记录，通过工具栏中“批量执行调价”功能可以使以下列表中的现价立即生效！"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   8055
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetails 
      Height          =   1575
      Left            =   1080
      TabIndex        =   0
      Top             =   2640
      Width           =   6255
      _cx             =   11033
      _cy             =   2778
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   0
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
   Begin XtremeCommandBars.ImageManager imgPicture 
      Left            =   1320
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMediPriceBatch.frx":6852
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   360
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMediPriceBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mconMenuEecute As Integer = 101
Private Const mconMenuExit As Integer = 102

'从参数表中取药品价格小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mint药库单位 As Integer

Private Enum Mcolumn
    mintcol序号 = 0
    mintcolid = 1
    mintcol药品id = 2
    mintcol编码
    mintcol名称
    mintcol规格
    mintcol原价
    mintcol现价
    mintcol调价人
    mintcol执行日期
    mintcol剂量系数
    mintcol药库包装
    mintcolCOUNT = 12
End Enum

Public Sub ShowMe(ByVal objfrm As frmMediLists)
    Me.Show vbModal, objfrm
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case 101
            Call ExecuteSave
        Case 102
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    pic提示.Move lngLeft, lngTop, lngRight - lngLeft
    vsfDetails.Move lngLeft, pic提示.Height + pic提示.Top, lngRight - lngLeft, lngBottom - lngTop - pic提示.Height
End Sub

Private Sub Form_Load()
    Me.Width = 12000
    Me.Height = 8000
    
    '判断是否以药库单位显示
    mint药库单位 = Val(zlDatabase.GetPara(29, glngSys))
    
    mintPriceDigit = GetDigit(1, 2, IIf(mint药库单位 = 0, 1, 4))
    
    Call InitComandBars
    Call InitVsf
    Call setVSF
    Call getData
End Sub

Private Sub InitComandBars()
    '初始化工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

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
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = imgPicture.Icons
    
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
    
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenuEecute, "批量执行调价")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mconMenuExit, "退出")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.BeginGroup = True
    End With
    
    cbsMain.Item(1).Delete
End Sub

Private Sub InitVsf()
    '初始化表格位置和大小
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    pic提示.Move lngLeft, lngTop, lngRight - lngLeft
    vsfDetails.Move lngLeft, pic提示.Height + pic提示.Top, lngRight - lngLeft, lngBottom - lngTop - pic提示.Height
End Sub

Private Sub setVSF()
    
    With vsfDetails
        .Editable = flexEDNone
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
'        .AllowSelection = True '不能多选单元格
        .AllowBigSelection = True '整行选择
        .SelectionMode = flexSelectionByRow '整行选择
        .Rows = 1
        
        .Cols = Mcolumn.mintcolCOUNT
        
        VsfGridColFormat vsfDetails, Mcolumn.mintcol序号, "序号", 600, flexAlignCenterCenter, "序号"
        VsfGridColFormat vsfDetails, Mcolumn.mintcolid, "id", 1000, flexAlignRightCenter, "id"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol药品id, "药品id", 1000, flexAlignRightCenter, "药品id"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol编码, "编码", 1500, flexAlignLeftCenter, "编码"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol名称, "名称", 2000, flexAlignLeftCenter, "名称"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol规格, "规格", 1500, flexAlignLeftCenter, "规格"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol调价人, "调价人", 1000, flexAlignLeftCenter, "调价人"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol执行日期, "生效日期", 2000, flexAlignLeftCenter, "生效日期"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol原价, "原价", 1000, flexAlignRightCenter, "原价"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol现价, "现价", 1000, flexAlignRightCenter, "现价"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol剂量系数, "剂量系数", 1000, flexAlignRightCenter, "剂量系数"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol药库包装, "药库包装", 1000, flexAlignRightCenter, "药库包装"
        
        .ColHidden(Mcolumn.mintcolid) = True
        .ColHidden(Mcolumn.mintcol药品id) = True
        .ColHidden(Mcolumn.mintcol剂量系数) = True
        .ColHidden(Mcolumn.mintcol药库包装) = True
        
    End With
End Sub

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf列设置：列名，列宽，列对齐方式，固定列对齐方式（默认为居中对齐）
    
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub

Private Sub getData()
    '获取数据的过程
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSql = "Select Distinct n.Id, i.Id As 药品id, i.编码, i.名称, i.规格, n.调价人, n.执行日期, n.终止日期, n.原价, n.现价, i.计算单位, p.药库单位, p.剂量系数, p.药库包装" & _
               " From 收费项目目录 I, 收费价目 N, 药品规格 P" & _
               " Where i.Id = n.收费细目id And i.Id = p.药品id And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & _
                   " n.变动原因 = 0 And Sysdate>n.执行日期" & _
                GetPriceClassString("N") & _
               " Order By n.id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    If rsTemp Is Nothing Then
        Exit Sub
    Else
        Call setColumn(rsTemp)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub setColumn(ByVal rsRecord As ADODB.Recordset)
    Dim i As Integer
    Dim dblPrice As Double
    
    With vsfDetails
        .Rows = rsRecord.RecordCount + 1
        For i = 1 To rsRecord.RecordCount
            If mint药库单位 = 1 Then
                dblPrice = IIf(IsNull(rsRecord!药库包装), 0, rsRecord!药库包装)
            Else
                dblPrice = 1
            End If
            
            .TextMatrix(i, Mcolumn.mintcol序号) = i
            .TextMatrix(i, Mcolumn.mintcolid) = rsRecord!ID
            .TextMatrix(i, Mcolumn.mintcol药品id) = rsRecord!药品id
            .TextMatrix(i, Mcolumn.mintcol编码) = rsRecord!编码
            .TextMatrix(i, Mcolumn.mintcol名称) = rsRecord!名称
            .TextMatrix(i, Mcolumn.mintcol规格) = rsRecord!规格
            .TextMatrix(i, Mcolumn.mintcol调价人) = IIf(IsNull(rsRecord!调价人), "", rsRecord!调价人)
            .TextMatrix(i, Mcolumn.mintcol执行日期) = Format(rsRecord!执行日期, "yyyy-mm-dd hh:mm:ss")

            .TextMatrix(i, Mcolumn.mintcol原价) = FormatEx(rsRecord!原价 * dblPrice, mintPriceDigit, , True)
            .TextMatrix(i, Mcolumn.mintcol现价) = FormatEx(rsRecord!现价 * dblPrice, mintPriceDigit, , True)

            .TextMatrix(i, Mcolumn.mintcol剂量系数) = IIf(IsNull(rsRecord!剂量系数), 0, rsRecord!剂量系数)
            .TextMatrix(i, Mcolumn.mintcol药库包装) = IIf(IsNull(rsRecord!药库包装), 0, rsRecord!药库包装)
            .RowHeight(i) = 350
            rsRecord.MoveNext
        Next
    End With
End Sub

Private Sub ExecuteSave()
    '执行批量调价
    Dim i As Integer
    On Error GoTo ErrHand
    
    If vsfDetails.Rows <= 1 Then Exit Sub
    For i = 1 To vsfDetails.Rows - 1
        gstrSql = ""
        gstrSql = "Zl_药品收发记录_Adjust(" & vsfDetails.TextMatrix(i, Mcolumn.mintcolid) & ")"
        zlDatabase.ExecuteProcedure gstrSql, Me.Caption
    Next
    MsgBox "批量执行调价成功,以下内容现价生效！", vbInformation, gstrSysName
    Call getData
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



