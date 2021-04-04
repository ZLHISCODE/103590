VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm卫材缺料清单 
   BorderStyle     =   0  'None
   Caption         =   "卫材缺料清单"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   4125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7320
      _cx             =   12912
      _cy             =   7276
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm卫材缺料清单.frx":0000
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
      ExplorerBar     =   7
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
End
Attribute VB_Name = "frm卫材缺料清单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsNotPayStuff As ADODB.Recordset
Private mintUnit As Integer
Private mlngModule As Long
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Sub InitVsGrid()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始网格控件
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-12 10:27:06
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        '0-可选,1-必选,-1-隐藏
        .ColData(.ColIndex("状态")) = 1
        .ColData(.ColIndex("单据类型")) = 1
        .ColData(.ColIndex("单据号")) = 1
        .ColData(.ColIndex("材料名称")) = 1
        .ColData(.ColIndex("数量")) = 1
    End With
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With vsGrid
        .Top = ScaleTop
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = ScaleHeight
    End With
End Sub
Public Function zlFullData(ByVal intUnit As Integer, ByVal rsNotPayStuff As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:填充汇总数据到Vss控件中
    '入参:rsNotPayStuff-未发料清单
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 17:11:13
    '-----------------------------------------------------------------------------------------------------------
    If mintUnit <> intUnit Then
        '需要初始化相关的数字格式化参数
        Call Form_Load
    End If
    mintUnit = intUnit
    
    Set mrsNotPayStuff = rsNotPayStuff
    With vsGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .Clear (1)
        '填充数据
        zlFullData = LoadDataToVssGrid
        .Redraw = flexRDBuffered
    End With
End Function
 
Private Sub Form_Load()
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "缺料清单"
    Call InitVsGrid
    
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
End Sub
Private Function LoadDataToVssGrid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:将相关的数据填充到指定的网格控件中
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 11:06:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    LoadDataToVssGrid = False
    
    err = 0: On Error GoTo ErrHand:

    '填充数据到控件中
    mrsNotPayStuff.Filter = 0
    If mrsNotPayStuff.RecordCount <> 0 Then mrsNotPayStuff.MoveFirst
    
    With vsGrid
        If mrsNotPayStuff.EOF Then '
            LoadDataToVssGrid = True
            Exit Function
        End If
        lngRow = .FixedRows
        Do While Not mrsNotPayStuff.EOF
            If mrsNotPayStuff!执行状态 = 0 Then
                .RowData(lngRow) = Val(mrsNotPayStuff!Id)
                .TextMatrix(lngRow, .ColIndex("科室")) = NVL(mrsNotPayStuff!科室)
                .Cell(flexcpData, lngRow, .ColIndex("单据号")) = Val(NVL(mrsNotPayStuff!位置))
                .TextMatrix(lngRow, .ColIndex("开单医生")) = NVL(mrsNotPayStuff!开单医生)
                .TextMatrix(lngRow, .ColIndex("状态")) = NVL(mrsNotPayStuff!状态)
                '24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料；
                .TextMatrix(lngRow, .ColIndex("单据类型")) = NVL(mrsNotPayStuff!类型)
                .TextMatrix(lngRow, .ColIndex("单据号")) = NVL(mrsNotPayStuff!NO)
                .TextMatrix(lngRow, .ColIndex("记帐员")) = NVL(mrsNotPayStuff!记帐员)
                .TextMatrix(lngRow, .ColIndex("床号")) = NVL(mrsNotPayStuff!床号)
                .TextMatrix(lngRow, .ColIndex("病人姓名")) = NVL(mrsNotPayStuff!姓名)
                .TextMatrix(lngRow, .ColIndex("住院号")) = NVL(mrsNotPayStuff!住院号)
                .TextMatrix(lngRow, .ColIndex("材料名称")) = NVL(mrsNotPayStuff!材料名称)
                .TextMatrix(lngRow, .ColIndex("规格")) = NVL(mrsNotPayStuff!规格)
                .TextMatrix(lngRow, .ColIndex("产地")) = NVL(mrsNotPayStuff!产地)
                .TextMatrix(lngRow, .ColIndex("批号")) = NVL(mrsNotPayStuff!批号)
                '.TextMatrix(lngRow, .ColIndex("付")) = Format(Val(NVL(mrsNotPayStuff!付)), "###")
                .TextMatrix(lngRow, .ColIndex("数量")) = NVL(mrsNotPayStuff!数量)
                .TextMatrix(lngRow, .ColIndex("单价")) = Format(Val(NVL(mrsNotPayStuff!单价)) * mrsNotPayStuff!换算系数, mFMT.FM_零售价)
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(NVL(mrsNotPayStuff!金额)), mFMT.FM_金额)
                .TextMatrix(lngRow, .ColIndex("说明")) = NVL(mrsNotPayStuff!说明)
                .TextMatrix(lngRow, .ColIndex("记帐时间")) = NVL(mrsNotPayStuff!记帐时间)
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            mrsNotPayStuff.MoveNext
         Loop
    End With
    LoadDataToVssGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Property Get zlHaveData() As Boolean
    Dim i As Integer
    With vsGrid
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("材料名称")) <> "" Then zlHaveData = True: Exit Function
        Next
    End With
    zlHaveData = False
End Property

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "缺料清单"
End Sub

Public Sub zlSetFontSize(ByVal curFontSize As Currency)
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-05-06 17:00:44
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        .Font.Size = curFontSize
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("刘") + 120
        .RowHeightMax = TextHeight("刘") + 120
        .Refresh
    End With
End Sub


