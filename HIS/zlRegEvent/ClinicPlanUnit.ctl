VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.UserControl ClinicPlanUnit 
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ScaleHeight     =   9225
   ScaleWidth      =   12000
   Begin VB.CheckBox chkOnlyOneUse 
      Caption         =   "独占方式"
      Height          =   300
      Left            =   5310
      TabIndex        =   4
      Top             =   50
      Width           =   1035
   End
   Begin VB.PictureBox picFun 
      BorderStyle     =   0  'None
      Height          =   4065
      Left            =   6840
      ScaleHeight     =   4065
      ScaleWidth      =   765
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1110
      Width           =   765
      Begin VB.CommandButton cmdFun 
         Caption         =   "<<"
         Enabled         =   0   'False
         Height          =   360
         Index           =   3
         Left            =   105
         TabIndex        =   11
         Top             =   1935
         Width           =   555
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "<"
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         Left            =   105
         TabIndex        =   10
         Top             =   1465
         Width           =   555
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   ">>"
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   995
         Width           =   555
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   525
         Width           =   555
      End
   End
   Begin VB.PictureBox picUnit 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   0
      Left            =   7650
      ScaleHeight     =   4050
      ScaleWidth      =   2760
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1110
      Visible         =   0   'False
      Width           =   2760
      Begin VB.CheckBox chkForbidBespeak 
         Caption         =   "禁止预约"
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   1110
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSelNum 
         Height          =   3285
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   360
         Width           =   2175
         _cx             =   3836
         _cy             =   5794
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"ClinicPlanUnit.ctx":0000
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
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   930
      Left            =   7770
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1080
      _Version        =   589884
      _ExtentX        =   1905
      _ExtentY        =   1640
      _StockProps     =   64
   End
   Begin VB.OptionButton optBespeakMode 
      Caption         =   "按总量预约"
      Height          =   300
      Index           =   1
      Left            =   2460
      TabIndex        =   2
      Top             =   50
      Width           =   1200
   End
   Begin VB.OptionButton optBespeakMode 
      Caption         =   "按比例预约"
      Height          =   300
      Index           =   0
      Left            =   1215
      TabIndex        =   1
      Top             =   50
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsUnit 
      Height          =   2865
      Left            =   90
      TabIndex        =   5
      Top             =   405
      Width           =   3120
      _cx             =   5503
      _cy             =   5054
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ClinicPlanUnit.ctx":0070
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
   Begin VB.OptionButton optBespeakMode 
      Caption         =   "按序号控制预约"
      Height          =   300
      Index           =   2
      Left            =   3690
      TabIndex        =   3
      Top             =   50
      Width           =   1560
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfNotSelNum 
      Height          =   4065
      Left            =   4440
      TabIndex        =   6
      Top             =   1110
      Width           =   2385
      _cx             =   4207
      _cy             =   7170
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ClinicPlanUnit.ctx":00F1
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
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "预约控制方式"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   110
      Width           =   1080
   End
End
Attribute VB_Name = "ClinicPlanUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobj所有合作单位 As 合作单位控制集
Private mobj合作单位集 As 合作单位控制集
Private mobj所有号序集 As 号序信息集
Private mblnNotClick As Boolean
Private mblnEdit As Boolean
Private mblnValiedCanSave As Boolean

Private Enum COL_Index
    Col_合作单位 = 0
    Col_禁止预约 = 1
    
    COL_序号 = 0
    Col_时间段 = 1
    COL_数量 = 2
End Enum

'属性变量:
Dim m_EditMode As gRegistPlanEditMode
Dim m_IsDataChanged As Boolean

'缺省属性值:
Const m_def_EditMode = 0
Const m_def_IsDataChanged = False
'事件声明:
Event DataIsChanged()


Public Function LoadData(ByVal obj合作单位集 As 合作单位控制集, ByVal obj所有号序集 As 号序信息集, _
    ByVal obj所有合作单位 As 合作单位控制集, Optional ByVal blnChanged As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载出诊安排
    '入参:
    '     obj合作单位集-合作单位分配信息
    '     obj所有合作单位 - 所有合作单位控制集 ,不传表示查看
    '     obj所有号序集 - 所有备选号序集
    '返回:加载成功，返回true,否则返回false
    '编制:刘兴洪
    '日期:2016-01-12 12:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj合作单位集 = obj合作单位集
    Set mobj所有号序集 = obj所有号序集
    Set mobj所有合作单位 = obj所有合作单位

    If mobj合作单位集 Is Nothing Then Set mobj合作单位集 = New 合作单位控制集
    If mobj所有号序集 Is Nothing Then Set mobj所有号序集 = New 号序信息集
    m_IsDataChanged = blnChanged
    LoadData = InitData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitFace()
    Err = 0: On Error GoTo Errhand
    With tbPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub UnitPageVisible(ByVal blnVisible As Boolean)
    '隐藏三方合作单位
    Dim i As Integer
    Dim blnDo As Boolean
    
    Err = 0: On Error GoTo Errhand
    'List
    For i = 1 To vsUnit.Rows - 1
        vsUnit.RowHidden(i) = False
        If vsUnit.RowData(i) = 1 Then vsUnit.RowHidden(i) = blnVisible = False
    Next
    'TabPage
    blnDo = False
    For i = 0 To tbPage.ItemCount - 1
        tbPage(i).Visible = True
        If Val(tbPage(i).Tag) = 1 Then tbPage(i).Visible = blnVisible
        If Val(tbPage(i).Tag) <> 1 And blnVisible = False And blnDo = False Then
            tbPage.Enabled = False
            tbPage(i).Selected = True: blnDo = True
            tbPage.Enabled = True
        End If
    Next
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetGridColVisible(ByVal bln分时段 As Boolean, ByVal bln序号控制 As Boolean)
    '设置网格列可见状态
    Dim i As Integer, j As Integer
    
    Err = 0: On Error GoTo Errhand:
    vsfNotSelNum.ColHidden(-1) = False
    vsfNotSelNum.AllowSelection = False
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        vsfSelNum(i).ColHidden(-1) = False
        vsfSelNum(i).Editable = flexEDNone '允许编辑
        vsfSelNum(i).FocusRect = flexFocusNone
        vsfSelNum(i).AllowSelection = False
    Next
    If bln分时段 Then
        If bln序号控制 Then
            '分时段序号控制"数量"列不可见
            vsfNotSelNum.ColHidden(COL_数量) = True
            vsfNotSelNum.AllowSelection = True
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                vsfSelNum(i).ColHidden(COL_数量) = True
                vsfSelNum(i).AllowSelection = True
            Next
        Else
            '分时段不序号控制"序号"列不可见
            vsfNotSelNum.ColHidden(COL_序号) = True
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                vsfSelNum(i).Editable = flexEDKbdMouse  '允许编辑
                vsfSelNum(i).FocusRect = flexFocusLight
                vsfSelNum(i).ColHidden(COL_序号) = True
            Next
        End If
    Else
        If bln序号控制 Then
            '不分时段序号控制只有"序号"列可见
            vsfNotSelNum.ColHidden(Col_时间段) = True
            vsfNotSelNum.ColHidden(COL_数量) = True
            vsfNotSelNum.AllowSelection = True
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                vsfSelNum(i).ColHidden(Col_时间段) = True
                vsfSelNum(i).ColHidden(COL_数量) = True
                vsfSelNum(i).AllowSelection = True
            Next
        End If
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2016-01-12 12:48:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj号序 As 号序信息, obj号序集 As 号序信息集
    Dim objVsfGrid As VSFlexGrid, obj合作单位 As 合作单位控制
    Dim bln分时段 As Boolean, bln序号控制 As Boolean, byt预约控制 As Byte
    Dim blnFind As Boolean, i As Long, lngRow As Long
    
    Err = 0: On Error GoTo Errhand:
    
    '============================================
    '先加载所有合作单位，初始化网格
    picFun.Tag = ""
    If mobj所有合作单位 Is Nothing Then
        vsUnit.Clear 1: vsUnit.Rows = 1
    Else
        With vsUnit
            .Clear 1
            .Rows = mobj所有合作单位.Count + 1
            lngRow = 1
            For Each obj合作单位 In mobj所有合作单位
                .TextMatrix(lngRow, Col_合作单位) = obj合作单位.合作单位名称
                .RowData(lngRow) = obj合作单位.类型 '1-三方机构;2-预约方式
                lngRow = lngRow + 1
            Next
        End With
    End If
    '加载页面
    Call InitUnitPage
    
    For i = 1 To vsUnit.Rows - 1
        vsUnit.TextMatrix(i, Col_禁止预约) = 0
        vsUnit.TextMatrix(i, COL_数量) = ""
        vsUnit.Cell(flexcpBackColor, i, COL_数量) = vsUnit.BackColor
    Next
    
    vsfNotSelNum.Clear 1: vsfNotSelNum.Rows = 1
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        vsfSelNum(i).Clear 1: vsfSelNum(i).Rows = 1
    Next
    '============================================
    
    bln分时段 = mobj所有号序集.是否分时段
    bln序号控制 = mobj所有号序集.是否序号控制
    '0-禁止预约(或挂号);1-按比例控制预约(或挂号);2-按总量控制预约(或挂号);3-按序号控制预约(或挂号);4-不作限制
    byt预约控制 = mobj合作单位集.预约控制方式
    
    '0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
    Call UnitPageVisible(mobj所有号序集.预约控制 <> 2)
    Call SetGridColVisible(bln分时段, bln序号控制)
    mblnEdit = bln分时段 And Not bln序号控制
    
    If bln分时段 = False And bln序号控制 = False And byt预约控制 = 3 Then byt预约控制 = 0
    mblnNotClick = True
    optBespeakMode(IIf(byt预约控制 = 0 Or byt预约控制 = 4, 0, byt预约控制 - 1)).Value = True
    chkOnlyOneUse.Value = IIf(mobj合作单位集.是否独占, vbChecked, vbUnchecked)
    mblnNotClick = False
    
    '标记按序号控制预约(或挂号)是否可见
    optBespeakMode(2).Tag = IIf(bln分时段 Or bln序号控制, "", "1")
    picFun.Tag = IIf(bln序号控制, "", "1")
    
    If byt预约控制 <> 3 Then
        For Each obj合作单位 In mobj合作单位集
            With vsUnit
                .Redraw = flexRDNone
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, Col_合作单位) = obj合作单位.合作单位名称 Then
                        Select Case obj合作单位.预约控制方式
                        Case 0 '禁止预约
                            .TextMatrix(i, Col_禁止预约) = 1
                            .Cell(flexcpBackColor, i, COL_数量) = vbButtonFace
                        Case 1, 2
                            If Not obj合作单位.号序信息集 Is Nothing Then
                                For Each obj号序 In obj合作单位.号序信息集
                                    '序号:控制方式=0,1,2,4时，填为0;否则存储启用序号或分时段的序号
                                    '数量:控制方式=0,4时，填为0;控制方式=1时，存放比例,如20,代表20%,控制方式=2时，存储的是限约数量，比如：10表示只能预约10个号;控制方式=3时，存储限约数量，启用序号的，一般为1,不启用序号且分时段的，存储限约数量
                                    .TextMatrix(i, COL_数量) = FormatEx(obj号序.数量, 2, False)
                                    Exit For
                                Next
                            End If
                        End Select
                    End If
                Next
                .Redraw = flexRDBuffered
            End With
        Next
    End If

    '加载所有序号信息
    If bln分时段 Or bln序号控制 Then
        With vsfNotSelNum
            .Redraw = flexRDNone
            For Each obj号序 In mobj所有号序集
                If obj号序.是否预约 And obj号序.数量 > 0 Then
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                    .TextMatrix(lngRow, COL_序号) = obj号序.序号
                    .TextMatrix(lngRow, Col_时间段) = Format(obj号序.开始时间, "hh:mm") & "-" & Format(obj号序.终止时间, "hh:mm")
                    .Cell(flexcpData, lngRow, Col_时间段) = obj号序.开始时间 & "-" & obj号序.终止时间
                    .TextMatrix(lngRow, COL_数量) = obj号序.数量
                    .Cell(flexcpData, lngRow, COL_数量) = obj号序.数量
                End If
            Next
            .Redraw = flexRDBuffered
        End With
        
        If bln分时段 And bln序号控制 = False Then
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                With vsfSelNum(i)
                    .Redraw = flexRDNone
                    For Each obj号序 In mobj所有号序集
                        If obj号序.是否预约 And obj号序.数量 > 0 Then
                            .Rows = .Rows + 1
                            lngRow = .Rows - 1
                            .TextMatrix(lngRow, COL_序号) = obj号序.序号
                            .TextMatrix(lngRow, Col_时间段) = Format(obj号序.开始时间, "hh:mm") & "-" & Format(obj号序.终止时间, "hh:mm")
                            .Cell(flexcpData, lngRow, Col_时间段) = obj号序.开始时间 & "-" & obj号序.终止时间
                            .TextMatrix(lngRow, COL_数量) = 0
                        End If
                    Next
                    .Redraw = flexRDBuffered
                End With
            Next
        End If
        If vsfNotSelNum.Rows > 1 And vsfNotSelNum.Row < 1 Then vsfNotSelNum.Row = 1

        '加载合作单位已选择序号信息
        For Each obj合作单位 In mobj合作单位集
            Set objVsfGrid = GetUnitVsfGrid(obj合作单位.合作单位名称)
            If Not objVsfGrid Is Nothing Then
                Select Case obj合作单位.预约控制方式
                Case 0 '禁止预约
                    mblnNotClick = True
                    chkForbidBespeak(objVsfGrid.index).Value = vbChecked
                    mblnNotClick = False
                    objVsfGrid.Editable = flexEDNone
                Case 3
                    If Not obj合作单位.号序信息集 Is Nothing Then
                        vsfNotSelNum.Redraw = flexRDNone
                        objVsfGrid.Redraw = flexRDNone
                        For Each obj号序 In obj合作单位.号序信息集
                            '序号:控制方式=0,1,2,4时，填为0;否则存储启用序号或分时段的序号
                            '数量:控制方式=0,4时，填为0;控制方式=1时，存放比例,如20,代表20%,控制方式=2时，存储的是限约数量，比如：10表示只能预约10个号;控制方式=3时，存储限约数量，启用序号的，一般为1,不启用序号且分时段的，存储限约数量
                            If bln分时段 And bln序号控制 = False Then
                                RemoveItem vsfNotSelNum, objVsfGrid, obj号序.序号, True, obj号序.数量
                            Else
                                RemoveItem vsfNotSelNum, objVsfGrid, obj号序.序号
                            End If
                        Next
                        vsfNotSelNum.Redraw = flexRDBuffered
                        objVsfGrid.Redraw = flexRDBuffered
                    End If
                    objVsfGrid.Editable = IIf(m_EditMode = ED_RegistPlan_Edit And mblnEdit, flexEDKbdMouse, flexEDNone)
                End Select
            End If
        Next
    End If
    
Handler:
    Call SetUnitVisible
    If Not tbPage.Selected Is Nothing Then
        Call SetButtonEnable(tbPage.Selected.index)
    End If
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetUnitVsfGrid(ByVal str合作单位 As String) As VSFlexGrid
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据名称获取对应的VSFlexGrid控件
    '返回:返回合作单位对应的VSFlexGrid控件
    '编制:刘兴洪
    '日期:2016-01-12 13:43:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Err = 0: On Error GoTo Errhand:
    With tbPage
        For i = 0 To .ItemCount - 1
            If .Item(i).Caption = str合作单位 Then
                Set GetUnitVsfGrid = vsfSelNum(i): Exit Function
            End If
        Next
    End With
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub chkForbidBespeak_Click(index As Integer)
    Dim objVsfGrid As VSFlexGrid, i As Long
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    Set objVsfGrid = vsfSelNum(index)
    
    objVsfGrid.Redraw = flexRDNone
    vsfNotSelNum.Redraw = flexRDNone
    If Not mobj所有号序集 Is Nothing Then
        If mobj所有号序集.是否分时段 And mobj所有号序集.是否序号控制 = False Then
            For i = 1 To objVsfGrid.Rows - 1
                RemoveItem vsfNotSelNum, objVsfGrid, Val(objVsfGrid.TextMatrix(i, COL_序号)), True, 0
            Next
            objVsfGrid.Editable = IIf(m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(index).Value <> vbChecked And mblnEdit, flexEDKbdMouse, flexEDNone)
            Exit Sub
        End If
    End If
    
    For i = 1 To objVsfGrid.Rows - 1
        If i > objVsfGrid.Rows - 1 Then Exit For
        RemoveItem objVsfGrid, vsfNotSelNum, Val(objVsfGrid.TextMatrix(i, COL_序号))
        i = i - 1
    Next
    objVsfGrid.Redraw = flexRDBuffered
    vsfNotSelNum.Redraw = flexRDBuffered
    
    Call SetButtonEnable(objVsfGrid.index)
    objVsfGrid.Editable = IIf(m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(index).Value <> vbChecked And mblnEdit, flexEDKbdMouse, flexEDNone)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetButtonEnable(ByVal index As Integer)
    cmdFun(0).Enabled = m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(index).Value <> vbChecked And vsfNotSelNum.Row > 0
    cmdFun(1).Enabled = m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(index).Value <> vbChecked And vsfNotSelNum.Rows > 1
    cmdFun(2).Enabled = m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(index).Value <> vbChecked And vsfSelNum(index).Row > 0
    cmdFun(3).Enabled = m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(index).Value <> vbChecked And vsfSelNum(index).Rows > 1
End Sub

Private Sub RemoveItem(ByVal objVsfGridFrom As VSFlexGrid, ByVal objVsfGridTo As VSFlexGrid, ByVal lngSN As Long, _
    Optional ByVal blnChangeNum As Boolean, Optional lngNum As Long)
    '移动项目或更改数量
    '参数：
    '   lngSN 序号
    '   blnChangeNum 仅改变数量,分时段，不序号控制时
    '   lngNum 改变的数量
    Dim blnFind As Boolean, i As Integer, j As Integer
    Dim lngRow As Long
    Dim intLow As Integer, intHigh As Integer, intMid As Integer
    
    On Error GoTo Errhand
    If objVsfGridFrom.Rows > 1 Then
        If Val(objVsfGridFrom.TextMatrix(1, COL_序号)) = lngSN Then
            lngRow = 1
        ElseIf Val(objVsfGridFrom.TextMatrix(objVsfGridFrom.Rows - 1, COL_序号)) = lngSN Then
            lngRow = objVsfGridFrom.Rows - 1
        End If
    End If
    '二分法查找
    If lngRow = 0 Then
        intLow = 1
        intHigh = objVsfGridFrom.Rows - 1
        Do While intLow <= intHigh
            intMid = (intLow + intHigh) \ 2
            If Val(objVsfGridFrom.TextMatrix(intMid, COL_序号)) < lngSN Then '在后面
                intLow = intMid + 1
            ElseIf Val(objVsfGridFrom.TextMatrix(intMid, COL_序号)) > lngSN Then '在前面
                intHigh = intMid - 1
            Else
                lngRow = intMid: Exit Do
            End If
        Loop
    End If
    If lngRow = 0 Then Exit Sub
    
    If blnChangeNum Then
        For i = 1 To objVsfGridTo.Rows - 1
            If Val(objVsfGridTo.TextMatrix(i, COL_序号)) = lngSN Then
                objVsfGridTo.TextMatrix(lngRow, COL_数量) = lngNum
                Exit For
            End If
        Next
        '计算剩余数量
        lngNum = Val(objVsfGridFrom.Cell(flexcpData, lngRow, COL_数量))
        For i = vsfSelNum.LBound To vsfSelNum.UBound
            For j = 1 To vsfSelNum(i).Rows - 1
                If Val(vsfSelNum(i).TextMatrix(j, COL_序号)) = lngSN Then
                    lngNum = lngNum - Val(vsfSelNum(i).TextMatrix(j, COL_数量))
                    Exit For
                End If
            Next
        Next
        objVsfGridFrom.TextMatrix(lngRow, COL_数量) = lngNum
    Else
        '按顺序插入
        blnFind = False
        If objVsfGridTo.Rows <= 1 Then
            With objVsfGridFrom
                objVsfGridTo.AddItem .TextMatrix(lngRow, COL_序号) & vbTab & .TextMatrix(lngRow, Col_时间段) & _
                    vbTab & .TextMatrix(lngRow, COL_数量)
                objVsfGridTo.Cell(flexcpData, objVsfGridTo.Rows - 1, Col_时间段) = .Cell(flexcpData, lngRow, Col_时间段)
            End With
            blnFind = True
        Else
            If Val(objVsfGridTo.TextMatrix(1, COL_序号)) >= lngSN Then
                With objVsfGridFrom
                    objVsfGridTo.AddItem .TextMatrix(lngRow, COL_序号) & vbTab & .TextMatrix(lngRow, Col_时间段) & _
                        vbTab & .TextMatrix(lngRow, COL_数量), 1
                    objVsfGridTo.Cell(flexcpData, 1, Col_时间段) = .Cell(flexcpData, lngRow, Col_时间段)
                End With
                blnFind = True
            ElseIf Val(objVsfGridTo.TextMatrix(objVsfGridTo.Rows - 1, COL_序号)) <= lngSN Then
                With objVsfGridFrom
                    objVsfGridTo.AddItem .TextMatrix(lngRow, COL_序号) & vbTab & .TextMatrix(lngRow, Col_时间段) & _
                        vbTab & .TextMatrix(lngRow, COL_数量)
                    objVsfGridTo.Cell(flexcpData, objVsfGridTo.Rows - 1, Col_时间段) = .Cell(flexcpData, lngRow, Col_时间段)
                End With
                blnFind = True
            End If
        End If
        
        '二分法查找
        If blnFind = False Then
            intLow = 1
            intHigh = objVsfGridTo.Rows - 1
            Do While intLow <= intHigh
                intMid = (intLow + intHigh) \ 2
                If Val(objVsfGridTo.TextMatrix(intMid - 1, COL_序号)) < lngSN _
                    And Val(objVsfGridTo.TextMatrix(intMid, COL_序号)) > lngSN Then   '找到位置了，且肯定能找到
                    With objVsfGridFrom
                        objVsfGridTo.AddItem .TextMatrix(lngRow, COL_序号) & vbTab & .TextMatrix(lngRow, Col_时间段) & _
                            vbTab & .TextMatrix(lngRow, COL_数量), intMid
                        objVsfGridTo.Cell(flexcpData, intMid, Col_时间段) = .Cell(flexcpData, lngRow, Col_时间段)
                    End With
                    Exit Do
                ElseIf Val(objVsfGridTo.TextMatrix(intMid, COL_序号)) < lngSN _
                    And Val(objVsfGridTo.TextMatrix(intMid + 1, COL_序号)) > lngSN Then '找到位置了，且肯定能找到
                    With objVsfGridFrom
                        objVsfGridTo.AddItem .TextMatrix(lngRow, COL_序号) & vbTab & .TextMatrix(lngRow, Col_时间段) & _
                            vbTab & .TextMatrix(lngRow, COL_数量), intMid + 1
                        objVsfGridTo.Cell(flexcpData, intMid + 1, Col_时间段) = .Cell(flexcpData, lngRow, Col_时间段)
                    End With
                    Exit Do
                End If
                
                If Val(objVsfGridTo.TextMatrix(intMid, COL_序号)) < lngSN Then '在后面
                    intLow = intMid + 1
                ElseIf Val(objVsfGridTo.TextMatrix(intMid, COL_序号)) > lngSN Then '在前面
                    intHigh = intMid - 1
                End If
            Loop
        End If
        objVsfGridFrom.RemoveItem lngRow
        
        If objVsfGridFrom.Rows > 1 And objVsfGridFrom.Row < 1 Then objVsfGridFrom.Row = 1
        If objVsfGridTo.Rows > 1 And objVsfGridTo.Row < 1 Then objVsfGridTo.Row = 1
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chkForbidBespeak_GotFocus(index As Integer)
    chkForbidBespeak(index).BackColor = GCTRL_SELBACK_COLOR
End Sub
 
Private Sub chkForbidBespeak_LostFocus(index As Integer)
     chkForbidBespeak(index).BackColor = Me.BackColor
End Sub


Private Sub chkForbidBespeak_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkOnlyOneUse_Click()
    Dim i As Integer
    
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    '清除数据
    For i = 1 To vsUnit.Rows - 1
        vsUnit.TextMatrix(i, Col_禁止预约) = 0
        vsUnit.TextMatrix(i, COL_数量) = ""
        vsUnit.Cell(flexcpBackColor, i, COL_数量) = vsUnit.BackColor
    Next
End Sub

Private Sub chkOnlyOneUse_GotFocus()
    chkOnlyOneUse.BackColor = GCTRL_SELBACK_COLOR
End Sub
Private Sub chkOnlyOneUse_LostFocus()
     chkOnlyOneUse.BackColor = Me.BackColor
End Sub
Private Sub chkOnlyOneUse_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdFun_Click(index As Integer)
    Dim objVsfGrid As VSFlexGrid
    Dim blnFind As Boolean, i As Integer
    Dim intStartRow As Integer, intEndRow As Integer
    
    On Error GoTo Errhand
    If Not tbPage.Selected Is Nothing Then
        Set objVsfGrid = GetUnitVsfGrid(tbPage.Selected.Caption)
    End If
    If objVsfGrid Is Nothing Then Exit Sub
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    vsfNotSelNum.Redraw = flexRDNone
    objVsfGrid.Redraw = flexRDNone
    Select Case index
    Case 0 '选进
        '批量设置
        intStartRow = vsfNotSelNum.RowSel: intEndRow = vsfNotSelNum.Row
        If vsfNotSelNum.Row < vsfNotSelNum.RowSel Then
            intStartRow = vsfNotSelNum.Row: intEndRow = vsfNotSelNum.RowSel
        End If
        Do While True
            If intStartRow > intEndRow Then Exit Do
            RemoveItem vsfNotSelNum, objVsfGrid, Val(vsfNotSelNum.TextMatrix(intStartRow, COL_序号))
            intEndRow = intEndRow - 1
        Loop
        If intStartRow > 0 And intStartRow < vsfNotSelNum.Rows Then vsfNotSelNum.Select intStartRow, 0
    Case 1 '全选进
        For i = 1 To vsfNotSelNum.Rows - 1
            If i > vsfNotSelNum.Rows - 1 Then Exit For
            RemoveItem vsfNotSelNum, objVsfGrid, Val(vsfNotSelNum.TextMatrix(i, COL_序号))
            i = i - 1
        Next
    Case 2 '移除
        '批量设置
        intStartRow = objVsfGrid.RowSel: intEndRow = objVsfGrid.Row
        If objVsfGrid.Row < objVsfGrid.RowSel Then
            intStartRow = objVsfGrid.Row: intEndRow = objVsfGrid.RowSel
        End If
        Do While True
            If intStartRow > intEndRow Then Exit Do
            RemoveItem objVsfGrid, vsfNotSelNum, Val(objVsfGrid.TextMatrix(intStartRow, COL_序号))
            intEndRow = intEndRow - 1
        Loop
        If intStartRow > 0 And intStartRow < objVsfGrid.Rows Then objVsfGrid.Select intStartRow, 0
    Case 3 '全移除
        For i = 1 To objVsfGrid.Rows - 1
            If i > objVsfGrid.Rows - 1 Then Exit For
            RemoveItem objVsfGrid, vsfNotSelNum, Val(objVsfGrid.TextMatrix(i, COL_序号))
            i = i - 1
        Next
    End Select
    vsfNotSelNum.Redraw = flexRDBuffered
    objVsfGrid.Redraw = flexRDBuffered
    
    Call SetButtonEnable(objVsfGrid.index)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdFun_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optBespeakMode_Click(index As Integer)
    Dim i As Integer, j As Integer
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    '清除数据
    For i = 1 To vsUnit.Rows - 1
        vsUnit.TextMatrix(i, Col_禁止预约) = 0
        vsUnit.TextMatrix(i, COL_数量) = ""
        vsUnit.Cell(flexcpBackColor, i, COL_数量) = vsUnit.BackColor
    Next
    If Not mobj所有号序集 Is Nothing Then
        For i = 0 To tbPage.ItemCount - 1
            chkForbidBespeak(i).Value = vbUnchecked
            For j = 1 To vsfSelNum(i).Rows - 1
                If mobj所有号序集.是否序号控制 Then
                    If j > vsfSelNum(i).Rows - 1 Then Exit For
                    RemoveItem vsfSelNum(i), vsfNotSelNum, Val(vsfSelNum(i).TextMatrix(j, COL_序号))
                    j = j - 1
                Else
                    RemoveItem vsfNotSelNum, vsfSelNum(i), Val(vsfSelNum(i).TextMatrix(j, COL_序号)), True, 0
                End If
            Next
        Next
    End If
    Call SetUnitVisible
    If Not tbPage.Selected Is Nothing Then
        Call SetButtonEnable(tbPage.Selected.index)
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optBespeakMode_GotFocus(index As Integer)
    optBespeakMode(index).BackColor = GCTRL_SELBACK_COLOR
End Sub
 
Private Sub optBespeakMode_LostFocus(index As Integer)
     optBespeakMode(index).BackColor = Me.BackColor
End Sub

Private Sub optBespeakMode_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub picFun_Resize()
    Err = 0: On Error Resume Next
    cmdFun(0).Top = (picFun.ScaleHeight - (cmdFun(0).Height + 100) * 4) / 2
    cmdFun(1).Top = cmdFun(0).Top + cmdFun(0).Height + 100
    cmdFun(2).Top = cmdFun(1).Top + cmdFun(1).Height + 100
    cmdFun(3).Top = cmdFun(2).Top + cmdFun(2).Height + 100
End Sub

Private Sub PicUnit_Resize(index As Integer)
    Err = 0: On Error Resume Next
    With picUnit(index)
        chkForbidBespeak(index).Left = .ScaleLeft + 30
        chkForbidBespeak(index).Top = .ScaleTop + 30
        
        vsfSelNum(index).Left = .ScaleLeft
        vsfSelNum(index).Width = .ScaleWidth
        vsfSelNum(index).Top = chkForbidBespeak(index).Top + chkForbidBespeak(index).Height
        vsfSelNum(index).Height = .ScaleHeight - vsfSelNum(index).Top
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNotClick Then Exit Sub
    
    If Val(tbPage.Tag) < tbPage.ItemCount Then
        mblnNotClick = True
        tbPage.Enabled = False
        tbPage.Item(Val(tbPage.Tag)).Selected = True
        tbPage.Enabled = True
        
        mblnValiedCanSave = True
        vsfSelNum(Val(tbPage.Tag)).FinishEditing False
        If mblnValiedCanSave = False Then mblnNotClick = False: Exit Sub
        
        tbPage.Enabled = False
        tbPage.Item(Item.index).Selected = True
        tbPage.Enabled = True
        mblnNotClick = False
    End If
    
    Call SetButtonEnable(Item.index)
    tbPage.Tag = Item.index
End Sub

Private Sub UserControl_Initialize()
    Call InitFace
    Call SetUnitVisible
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With vsUnit
        .Left = 0
        .Top = optBespeakMode(0).Top + optBespeakMode(0).Height + 50
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - .Top
    End With
    If m_EditMode = ED_RegistPlan_Edit Then
        With vsfNotSelNum
            .Left = vsUnit.Left
            .Top = vsUnit.Top
            .Height = vsUnit.Height
        End With
        With picFun
            .Left = vsfNotSelNum.Left + vsfNotSelNum.Width
            .Top = vsUnit.Top
            .Height = vsUnit.Height
        End With
        With tbPage
            .Left = IIf(picFun.Tag = "", picFun.Left + picFun.Width, picFun.Left)
            .Top = vsUnit.Top
            .Height = vsUnit.Height
            .Width = ScaleWidth - .Left
        End With
    Else
        With tbPage
            .Left = 0
            .Top = vsUnit.Top
            .Height = vsUnit.Height
            .Width = ScaleWidth - .Left
        End With
    End If
End Sub

Private Sub InitUnitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2016-01-11 14:23:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Dim objUnit As 合作单位控制, lngRow As Long
    Dim intPageCount As Integer
    Dim intSelectedPageIndex As Integer
    
    Err = 0: On Error GoTo Errhand:
    If Not tbPage.Selected Is Nothing Then
        intSelectedPageIndex = tbPage.Selected.index
    End If
    
    '减少控件的加载
'    If Not mobj所有合作单位 Is Nothing Then
'        intPageCount = mobj所有合作单位.Count
'        If intPageCount > tbPage.ItemCount Then
'            intPageCount = tbPage.ItemCount
'        End If
'    End If
'    If intPageCount = 0 Then intPageCount = 1
    
    tbPage.RemoveAll
'    For i = intPageCount To picUnit.UBound
'        Unload chkForbidBespeak(i)
'        Unload vsfSelNum(i)
'        Unload picUnit(i)
'    Next
    intPageCount = picUnit.Count
    
    If Not mobj所有合作单位 Is Nothing Then
        For Each objUnit In mobj所有合作单位
            If lngRow >= intPageCount Then
                Load chkForbidBespeak(lngRow): chkForbidBespeak(lngRow).Visible = True
                Load vsfSelNum(lngRow): vsfSelNum(lngRow).Visible = True
                Load picUnit(lngRow): picUnit(lngRow).Visible = True
                Set chkForbidBespeak(lngRow).Container = picUnit(lngRow)
                Set vsfSelNum(lngRow).Container = picUnit(lngRow)
                picUnit(lngRow).TabStop = False
            End If
            
            picUnit(lngRow).Visible = True
            Set ObjItem = tbPage.InsertItem(lngRow + 1, objUnit.合作单位名称, picUnit(lngRow).Hwnd, 0)
            ObjItem.Tag = objUnit.类型 '1-三方机构;2-预约方式
            lngRow = lngRow + 1
        Next
    End If
    
    If tbPage.ItemCount = 0 Then
        lngRow = 0
        picUnit(lngRow).Visible = True
        Set ObjItem = tbPage.InsertItem(lngRow + 1, "无合作单位", picUnit(lngRow).Hwnd, 0)
        ObjItem.Tag = "无合作单位"
    End If
    
    If intSelectedPageIndex = 0 Or intSelectedPageIndex > tbPage.ItemCount - 1 Then
        intSelectedPageIndex = tbPage.ItemCount - 1
    End If
    '手动触发SelectedChanged事件
    Call tbPage_SelectedChanged(tbPage.Item(intSelectedPageIndex))
    tbPage.Enabled = False
    tbPage.Item(intSelectedPageIndex).Selected = True
    tbPage.Enabled = True
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitUnitGrid()
    '初始化合作单位序号网格
    Dim i As Integer
    
    Err = 0: On Error GoTo Errhand:
    With vsfNotSelNum
        .Clear 1
        .Rows = 1
        .HighLight = flexHighlightAlways
        .ColHidden(-1) = False
    End With
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        With vsfSelNum(i)
            .Clear 1
            .Rows = 1
            .Editable = flexEDNone
            .ColHidden(-1) = False
            .HighLight = flexHighlightAlways
            .FocusRect = flexFocusNone
        End With
        mblnNotClick = True
        chkForbidBespeak(i).Value = vbUnchecked
        mblnNotClick = False
    Next
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetUnitVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据合作单位的预约控制方式，设置对应的控件显示
    '编制:刘兴洪
    '日期:2016-01-12 11:23:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    vsfNotSelNum.Visible = m_EditMode = ED_RegistPlan_Edit And optBespeakMode(2).Value
    picFun.Visible = m_EditMode = ED_RegistPlan_Edit And optBespeakMode(2).Value And Val(picFun.Tag) = 0
    tbPage.Visible = optBespeakMode(2).Value
    vsUnit.Visible = Not optBespeakMode(2).Value
    chkOnlyOneUse.Visible = Not optBespeakMode(2).Value
    optBespeakMode(2).Visible = Val(optBespeakMode(2).Tag) = 0
    If Val(optBespeakMode(2).Tag) = 0 Then
        chkOnlyOneUse.Left = optBespeakMode(2).Left + optBespeakMode(2).Width + 50
    Else
        chkOnlyOneUse.Left = optBespeakMode(2).Left
    End If
    vsUnit.TextMatrix(0, COL_数量) = IIf(optBespeakMode(0).Value, "比例(%)", "限约数")
    Call UserControl_Resize
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function Get合作单位控制集() As 合作单位控制集
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取合作单信息信息数据
    '返回:号序信息集
    '编制:刘兴洪
    '日期:2016-01-13 12:34:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, intCol As Integer
    Dim objUnits As New 合作单位控制集, objUnit As 合作单位控制
    Dim lngSum As Double, varTemp As Variant
    Dim strUnitName As String
    Dim objNums As 号序信息集, objNum As 号序信息
    Dim bln禁止预约 As Boolean
    
    Err = 0: On Error GoTo Errhand:
    '数据未改变，直接返回原集合的副本
    If m_IsDataChanged = False Then
        If mobj合作单位集.Count = 0 And mobj所有合作单位.Count > 0 Then
            '第一次加载，没有改变，应该是全部不限制
            
        Else
            Set Get合作单位控制集 = mobj合作单位集.Clone
            Exit Function
        End If
    End If
    
    '数据已改变，重新构造集合对象
    With objUnits
        .预约控制方式 = GetSelectedIndex(optBespeakMode) + 1
        .是否独占 = chkOnlyOneUse.Value = vbChecked
        .是否修改 = True
    End With
    
    If optBespeakMode(0).Value Or optBespeakMode(1).Value Then
        '按比例控制或按总量控制
        With vsUnit
            For lngRow = 1 To .Rows - 1
                Set objUnit = New 合作单位控制
                objUnit.合作单位名称 = .TextMatrix(lngRow, Col_合作单位)
                objUnit.类型 = .RowData(lngRow)
                
                If .RowHidden(lngRow) Then '隐藏的就是禁止预约
                    bln禁止预约 = True
                    lngSum = 0
                Else
                    bln禁止预约 = Abs(Val(.TextMatrix(lngRow, Col_禁止预约))) = 1
                    lngSum = Val(.TextMatrix(lngRow, COL_数量))
                End If
                '0-禁止预约;1-按比例控制预约;2-按总量控制预约;3-按序号控制预约;4-不作限制
                objUnit.预约控制方式 = IIf(bln禁止预约, 0, _
                                        IIf(lngSum = 0, 4, IIf(optBespeakMode(0).Value, 1, 2)))
                Set objNums = New 号序信息集
                If lngSum > 0 Or bln禁止预约 Then
                    Set objNum = New 号序信息
                    objNum.序号 = 0
                    objNum.数量 = lngSum
                    objNums.AddItem objNum
                End If
                Set objUnit.号序信息集 = objNums
                objUnits.AddItem objUnit, "K" & objUnit.合作单位名称
            Next
        End With
    Else
        '0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
        For lngRow = 0 To tbPage.ItemCount - 1
            If tbPage(lngRow).Caption = "无合作单位" Then Exit For
            
            If GetLocaleUnit(lngRow, objUnit) Then
                objUnits.AddItem objUnit, "K" & objUnit.合作单位名称
            End If
        Next
    End If
    Set Get合作单位控制集 = objUnits
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetLocaleUnit(ByVal intPage As Integer, ByRef objUnit As 合作单位控制) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的合作单信息
    '入参:intPage-指定的页
    '出参:objUnit-合作单位信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-01-13 18:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objNums As 号序信息集, objNum As 号序信息
    Dim varTemp As Variant, lngCount As Long
    
    Set objUnit = New 合作单位控制
    Err = 0: On Error GoTo Errhand:
    objUnit.合作单位名称 = tbPage.Item(intPage).Caption
    objUnit.类型 = Val(tbPage.Item(intPage).Tag)
    '0-禁止预约;1-按比例控制预约;2-按总量控制预约;3-按序号控制预约;4-不作限制
    If chkForbidBespeak(intPage).Value = vbChecked _
        Or mobj所有号序集.预约控制 = 2 And Val(tbPage(intPage).Tag) = 1 Then
        '仅禁止三方合作单位
        objUnit.预约控制方式 = 0
    Else
        objUnit.预约控制方式 = 3
    End If

    Set objNums = New 号序信息集
    If objUnit.预约控制方式 = 3 Then
        With vsfSelNum(intPage)
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_数量)) <> 0 Then
                    lngCount = lngCount + Val(.TextMatrix(i, COL_数量))
                    
                    Set objNum = New 号序信息
                    objNum.序号 = Val(.TextMatrix(i, COL_序号))
                    If .TextMatrix(i, Col_时间段) <> "" Then
                        varTemp = Split(.Cell(flexcpData, i, Col_时间段), "-")
                        objNum.开始时间 = varTemp(0)
                        objNum.终止时间 = varTemp(1)
                    End If
                    objNum.数量 = Val(.TextMatrix(i, COL_数量))
                    objNums.AddItem objNum
                End If
            Next
        End With
        '一个序号都没有设置数量,则表示不限制
        If lngCount = 0 Then objUnit.预约控制方式 = 4
    End If
    If objUnit.预约控制方式 = 0 Or objUnit.预约控制方式 = 4 Then
        '禁止预约或不限制时添加一个记录，以便保存
        Set objNum = New 号序信息
        objNum.序号 = 0
        objNum.数量 = 0
        objNums.AddItem objNum
    End If
    Set objUnit.号序信息集 = objNums
    GetLocaleUnit = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Property Get Get合作单位控制信息集() As 合作单位控制集
   Set Get合作单位控制信息集 = Get合作单位控制集
End Property

Public Function IsValied(Optional ByVal blnChanged As Boolean) As Boolean
    '检查数据
    '外面一层是否改变，若改变则本层也要检查
    Dim lngSum As Double, lng限约数 As Long, lngSN As Long
    Dim i As Long, j As Integer, k As Long
    
    Err = 0: On Error GoTo ErrHandler
    '数据未改变不检查
    If m_IsDataChanged = False And blnChanged = False Then IsValied = True: Exit Function
    
    mblnValiedCanSave = True
    vsUnit.FinishEditing False
    If mblnValiedCanSave = False Then Exit Function
    
    mblnValiedCanSave = True
    If Not tbPage.Selected Is Nothing Then
        vsfSelNum(tbPage.Selected.index).FinishEditing False
    End If
    If mblnValiedCanSave = False Then Exit Function

    If optBespeakMode(0).Value Then '按比例
        If chkOnlyOneUse.Value = vbChecked Then
            For i = 1 To vsUnit.Rows - 1
                lngSum = lngSum + Val(vsUnit.TextMatrix(i, COL_数量))
            Next
            If lngSum > 100 Then
                MsgBox "独占方式时，限约比例之和不能超过100！", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        Else
            For i = 1 To vsUnit.Rows - 1
                lngSum = Val(vsUnit.TextMatrix(i, COL_数量))
                If lngSum > 100 Then
                    MsgBox vsUnit.TextMatrix(i, Col_合作单位) & "的预约比例不能超过100！", vbInformation + vbOKOnly, gstrSysName
                    vsUnit.Row = i: vsUnit.Col = COL_数量
                    Exit Function
                End If
            Next
        End If
    ElseIf optBespeakMode(1).Value Then '按总量
        If Not mobj所有号序集 Is Nothing Then lng限约数 = mobj所有号序集.限约数
        If lng限约数 > 0 Then '不限约时不用检查
            If chkOnlyOneUse.Value = vbChecked Then
                For i = 1 To vsUnit.Rows - 1
                    lngSum = lngSum + Val(vsUnit.TextMatrix(i, COL_数量))
                Next
                If lngSum > lng限约数 Then
                    MsgBox "独占方式时，限约数之和不能超过限约数(" & lng限约数 & ")！", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            Else
                For i = 1 To vsUnit.Rows - 1
                    lngSum = Val(vsUnit.TextMatrix(i, COL_数量))
                    If lngSum > lng限约数 Then
                        MsgBox vsUnit.TextMatrix(i, Col_合作单位) & "的限约数不能超过限约数(" & lng限约数 & ")！", vbInformation + vbOKOnly, gstrSysName
                        vsUnit.Row = i: vsUnit.Col = COL_数量
                        Exit Function
                    End If
                Next
            End If
        End If
    Else '按序号
        If Not mobj所有号序集 Is Nothing Then
            If mobj所有号序集.是否分时段 And mobj所有号序集.是否序号控制 = False Then
                For k = 1 To vsfNotSelNum.Rows - 1
                    lngSum = Val(vsfNotSelNum.Cell(flexcpData, k, COL_数量))
                    lngSN = Val(vsfNotSelNum.TextMatrix(k, COL_序号))
                    For i = vsfSelNum.LBound To vsfSelNum.UBound
                        For j = 1 To vsfSelNum(i).Rows - 1
                            If Val(vsfSelNum(i).TextMatrix(j, COL_序号)) = lngSN Then
                                lngSum = lngSum - Val(vsfSelNum(i).TextMatrix(j, COL_数量))
                            End If
                        Next
                    Next
                    If lngSum < 0 Then
                        MsgBox vsfNotSelNum.Cell(flexcpData, k, Col_时间段) & " 分配的预约数超过了该时段的可预约数量(" & Val(vsfNotSelNum.Cell(flexcpData, k, COL_数量)) & ")！", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                Next
            End If
        End If
    End If
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    On Error Resume Next
    lblEdit.BackColor = UserControl.BackColor
    optBespeakMode(0).BackColor = UserControl.BackColor
    optBespeakMode(1).BackColor = UserControl.BackColor
    optBespeakMode(2).BackColor = UserControl.BackColor
    chkOnlyOneUse.BackColor = UserControl.BackColor
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "指出 Label 或 Shape 的背景样式是透明的还是不透明的。"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_IsDataChanged = m_def_IsDataChanged
    m_EditMode = m_def_EditMode
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    m_IsDataChanged = PropBag.ReadProperty("IsDataChanged", m_def_IsDataChanged)
    m_EditMode = PropBag.ReadProperty("EditMode", m_def_EditMode)
End Sub

Private Sub UserControl_Terminate()
    Set mobj合作单位集 = Nothing
    Set mobj所有号序集 = Nothing
    Set mobj所有合作单位 = Nothing
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("IsDataChanged", m_IsDataChanged, m_def_IsDataChanged)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
End Sub

Private Sub vsfNotSelNum_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If tbPage.Selected Is Nothing Then Exit Sub
    Call SetButtonEnable(tbPage.Selected.index)
End Sub

Private Sub vsfNotSelNum_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If m_EditMode <> ED_RegistPlan_Edit Then Cancel = True: Exit Sub
End Sub

Private Sub vsfNotSelNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsfNotSelNum_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
End Sub

Private Sub vsfSelNum_AfterEdit(index As Integer, ByVal Row As Long, ByVal Col As Long)
    RemoveItem vsfNotSelNum, vsfSelNum(index), _
        Val(vsfSelNum(index).TextMatrix(Row, COL_序号)), _
        True, Val(vsfSelNum(index).TextMatrix(Row, COL_数量))
    m_IsDataChanged = True: RaiseEvent DataIsChanged
End Sub

Private Sub vsfSelNum_AfterRowColChange(index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    Call SetButtonEnable(index)
    If vsfSelNum(index).Editable = flexEDKbdMouse Then
        vsfNotSelNum.Row = NewRow
        vsfNotSelNum.TopRow = vsfSelNum(index).TopRow
    End If
End Sub

Private Sub vsfSelNum_BeforeEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If m_EditMode <> ED_RegistPlan_Edit Then Cancel = True: Exit Sub
    If COL_数量 <> Col Then Cancel = True: Exit Sub
End Sub

Private Sub vsfSelNum_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And vsfSelNum(index).Editable = flexEDKbdMouse Then
        If vsfSelNum(index).Row = vsfSelNum(index).Rows - 1 And vsfSelNum(index).Col = vsfSelNum(index).Cols - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlVsMoveGridCell(vsfSelNum(index), 2)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub vsfSelNum_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsfSelNum_KeyPressEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    '输入位数限制，整数位长度不能大于9
    If InStr(vsfSelNum(index).EditText, ".") > 0 Then
        If InStr(vsfSelNum(index).EditText, ".") > 9 Then KeyAscii = 0
    Else
        If Len(vsfSelNum(index).EditText) >= 9 Then KeyAscii = 0
    End If
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsfSelNum_ValidateEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngSum As Long, lngSN As Long
    Dim i As Integer, j As Integer, lngRow As Long
    
    On Error GoTo Errhand
    '整数位多余9位的直接截掉,防止溢出
    If InStr(vsfSelNum(index).EditText, ".") > 0 Then
        If InStr(vsfSelNum(index).EditText, ".") > 9 Then
            vsfSelNum(index).EditText = Left(vsfSelNum(index).EditText, 9)
        End If
    Else
        vsfSelNum(index).EditText = Left(vsfSelNum(index).EditText, 9)
    End If
    
    lngSN = Val(vsfSelNum(index).TextMatrix(Row, COL_序号))
    For i = 1 To vsfNotSelNum.Rows - 1
        If Val(vsfNotSelNum.TextMatrix(i, COL_序号)) = lngSN Then
            lngRow = i
        End If
    Next
    '计算剩余数量
    lngSum = Val(vsfNotSelNum.Cell(flexcpData, lngRow, COL_数量))
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        If i <> index Then
            For j = 1 To vsfSelNum(i).Rows - 1
                If Val(vsfSelNum(i).TextMatrix(j, COL_序号)) = lngSN Then
                    lngSum = lngSum - Val(vsfSelNum(i).TextMatrix(j, COL_数量))
                    Exit For
                End If
            Next
        End If
    Next
    
    If Val(vsfSelNum(index).EditText) > lngSum Then
        MsgBox tbPage(index).Caption & " 预约数(" & Val(vsfSelNum(index).EditText) & ")不能超过剩余预约数量(" & lngSum & ")！", vbInformation + vbOKOnly, gstrSysName
        Cancel = True: mblnValiedCanSave = False: Exit Sub
    End If
    vsfSelNum(index).EditText = FormatEx(Val(vsfSelNum(index).EditText), 0)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsUnit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        If vsUnit.TextMatrix(Row, Col_禁止预约) = True Then
            vsUnit.TextMatrix(Row, COL_数量) = ""
            vsUnit.Cell(flexcpBackColor, Row, COL_数量) = vbButtonFace
        Else
            vsUnit.Cell(flexcpBackColor, Row, COL_数量) = vsUnit.BackColor
        End If
    End If
End Sub

Private Sub vsUnit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If m_EditMode <> ED_RegistPlan_Edit Then Cancel = True: Exit Sub
    If Col = Col_合作单位 Then Cancel = True: Exit Sub
    If Col = COL_数量 Then
        If Abs(Val(vsUnit.TextMatrix(Row, Col_禁止预约))) = 1 Then Cancel = True: Exit Sub
    End If
    '由事件AfterEdit调到这里，因为当正在编辑时直接按保存，检查不到
    m_IsDataChanged = True: RaiseEvent DataIsChanged
End Sub

Private Sub vsUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If vsUnit.Row = vsUnit.Rows - 1 And vsUnit.Col = vsUnit.Cols - 1 Then
            'Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlVsMoveGridCell(vsUnit, 1)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub vsUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsUnit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    '输入位数限制，整数位长度不能大于9
    If InStr(vsUnit.EditText, ".") > 0 Then
        If InStr(vsUnit.EditText, ".") > 9 Then KeyAscii = 0
    Else
        If Len(vsUnit.EditText) >= 9 Then KeyAscii = 0
    End If
    If optBespeakMode(0).Value Then
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Else
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Public Property Let 预约控制(ByVal vNewValue As Byte)
    Dim i As Integer, j As Integer
    
    On Error GoTo Errhand
    If mobj所有号序集 Is Nothing Then Set mobj所有号序集 = New 号序信息集
    mobj所有号序集.预约控制 = vNewValue
    '0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
    Call UnitPageVisible(mobj所有号序集.预约控制 <> 2)
    
    '清除数据
    If mobj所有号序集.预约控制 = 2 Then
        For i = 1 To vsUnit.Rows - 1
            If vsUnit.RowData(i) = 1 Then
                vsUnit.TextMatrix(i, Col_禁止预约) = 1
                vsUnit.TextMatrix(i, COL_数量) = ""
                vsUnit.Cell(flexcpBackColor, i, COL_数量) = vsUnit.BackColor
            End If
        Next
        For i = 0 To tbPage.ItemCount - 1
            If Val(tbPage(i).Tag) = 1 Then
                chkForbidBespeak(i).Value = vbChecked
                For j = 1 To vsfSelNum(i).Rows - 1
                    If mobj所有号序集.是否序号控制 Then
                        If j > vsfSelNum(i).Rows - 1 Then Exit For
                        RemoveItem vsfSelNum(i), vsfNotSelNum, Val(vsfSelNum(i).TextMatrix(j, COL_序号))
                        j = j - 1
                    Else
                        RemoveItem vsfNotSelNum, vsfSelNum(i), Val(vsfSelNum(i).TextMatrix(j, COL_序号)), True, 0
                    End If
                Next
            End If
        Next
        If Not tbPage.Selected Is Nothing Then
            Call SetButtonEnable(tbPage.Selected.index)
        End If
    End If
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Public Property Let 所有号序信息集(ByVal vNewValue As 号序信息集)
    Err = 0: On Error GoTo Errhand
    Set mobj所有号序集 = vNewValue
    If mobj所有号序集 Is Nothing Then Set mobj所有号序集 = New 号序信息集
    Set mobj合作单位集 = Get合作单位控制集
    Call InitData
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Private Sub vsUnit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngSum As Double, lng限约数 As Long
    Dim i As Long
    
    On Error GoTo Errhand
    '编辑禁止预约列时不检查
    If Col = Col_禁止预约 Then Exit Sub
    '整数位多余9位的直接截掉,防止溢出
    If InStr(vsUnit.EditText, ".") > 0 Then
        If InStr(vsUnit.EditText, ".") > 9 Then
            vsUnit.EditText = Left(vsUnit.EditText, 9)
        End If
    Else
        vsUnit.EditText = Left(vsUnit.EditText, 9)
    End If
    
    If chkOnlyOneUse.Value = vbChecked Then
        For i = 1 To vsUnit.Rows - 1
            If i <> vsUnit.Row Then
                lngSum = lngSum + Val(vsUnit.TextMatrix(i, COL_数量))
            End If
        Next
        lngSum = lngSum + Val(vsUnit.EditText)
        If optBespeakMode(0).Value Then '按比例
            If lngSum > 100 Then
                MsgBox "独占方式时，合作单位控制限约比例之和不能超过100！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True: mblnValiedCanSave = False: Exit Sub
            End If
        ElseIf optBespeakMode(1).Value Then '按总量
            If Not mobj所有号序集 Is Nothing Then lng限约数 = mobj所有号序集.限约数
            If lng限约数 > 0 Then '不限约时不用检查
                If lngSum > lng限约数 Then
                    MsgBox "独占方式时，合作单位控制限约数之和不能超过限约数(" & lng限约数 & ")！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True: mblnValiedCanSave = False: Exit Sub
                End If
            End If
        End If
    Else
        lngSum = Val(vsUnit.EditText)
        If optBespeakMode(0).Value Then '按比例
            If lngSum > 100 Then
                MsgBox vsUnit.TextMatrix(vsUnit.Row, Col_合作单位) & " 预约比例不能超过100！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True: mblnValiedCanSave = False: Exit Sub
            End If
        ElseIf optBespeakMode(1).Value Then  '按总量
            If Not mobj所有号序集 Is Nothing Then lng限约数 = mobj所有号序集.限约数
            If lng限约数 > 0 Then '不限约时不用检查
                If lngSum > lng限约数 Then
                    MsgBox vsUnit.TextMatrix(vsUnit.Row, Col_合作单位) & " 限约数不能超过限约数(" & lng限约数 & ")！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True: mblnValiedCanSave = False: Exit Sub
                End If
            End If
        End If
    End If
    vsUnit.EditText = FormatEx(Val(vsUnit.EditText), 2)
    vsUnit.EditText = IIf(vsUnit.EditText = "0", "", vsUnit.EditText)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,false
Public Property Get IsDataChanged() As Boolean
    IsDataChanged = m_IsDataChanged
End Property

Public Property Let IsDataChanged(ByVal New_IsDataChanged As Boolean)
    m_IsDataChanged = New_IsDataChanged
    PropertyChanged "IsDataChanged"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14,0,0,0
Public Property Get EditMode() As gRegistPlanEditMode
    EditMode = m_EditMode
End Property

Public Property Let EditMode(ByVal New_EditMode As gRegistPlanEditMode)
    Dim i As Integer
    
    m_EditMode = IIf(New_EditMode = ED_RegistPlan_UpdateUnit, ED_RegistPlan_Edit, New_EditMode)
    If mobj所有号序集 Is Nothing Then
        m_EditMode = ED_RegistPlan_View
    ElseIf m_EditMode = ED_RegistPlan_Edit And mobj所有号序集.预约控制 = Val("1-禁止预约") Then
        m_EditMode = ED_RegistPlan_View
    End If
    PropertyChanged "EditMode"
    
    For i = optBespeakMode.LBound To optBespeakMode.UBound
        optBespeakMode(i).Enabled = m_EditMode = ED_RegistPlan_Edit
    Next
    chkOnlyOneUse.Enabled = m_EditMode = ED_RegistPlan_Edit
    vsUnit.Editable = IIf(m_EditMode = ED_RegistPlan_Edit, flexEDKbdMouse, flexEDNone)
    picFun.Enabled = m_EditMode = ED_RegistPlan_Edit
    For i = 0 To tbPage.ItemCount - 1
        chkForbidBespeak(i).Enabled = m_EditMode = ED_RegistPlan_Edit
        vsfSelNum(i).Editable = IIf(m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(i).Value = Unchecked And mblnEdit, flexEDKbdMouse, flexEDNone)
    Next
    
    '隐藏
    SetUnitVisible
    UserControl_Resize
End Property

