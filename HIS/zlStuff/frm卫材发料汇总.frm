VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm卫材发料汇总 
   BorderStyle     =   0  'None
   Caption         =   "卫材汇总发料"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   5535
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2160
      Width           =   5535
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7320
      _cx             =   12912
      _cy             =   3678
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm卫材发料汇总.frx":0000
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
   Begin VSFlex8Ctl.VSFlexGrid vsfChargeOff 
      Height          =   2085
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   7320
      _cx             =   12912
      _cy             =   3678
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm卫材发料汇总.frx":0191
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
Attribute VB_Name = "frm卫材发料汇总"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsNotPayStuff As ADODB.Recordset
Private mrsChargeOff As New ADODB.Recordset                   '用于显示销帐申请记录
Private mintUnit As Integer
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
Private mbln分科室 As Boolean
Private mlngModule As Long
Private mbln发料时汇总销账 As Boolean

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
        .ColData(.ColIndex("科室")) = IIf(mbln分科室, 1, -1)
        .ColData(.ColIndex("材料名称")) = 1
        .ColData(.ColIndex("实发数量")) = 1
    End With
End Sub
Private Sub LoadDataToChargeOffList(ByVal lng科室id As Long, ByVal lng材料ID As Long)
    With vsfChargeOff
        .Rows = 1
        
        mrsChargeOff.Filter = "领料部门id=" & lng科室id & " And 材料id=" & lng材料ID & " And 销帐数量>0 "
        If mrsChargeOff.RecordCount = 0 Then Exit Sub
    
        .Redraw = flexRDNone
        
        Do While Not mrsChargeOff.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("科室")) = mrsChargeOff!领料部门
            .TextMatrix(.Rows - 1, .ColIndex("NO")) = mrsChargeOff!NO
            .TextMatrix(.Rows - 1, .ColIndex("材料名称")) = mrsChargeOff!材料名称
            .TextMatrix(.Rows - 1, .ColIndex("产地")) = IIf(NVL(mrsChargeOff!产地) = "", "", mrsChargeOff!产地)
            .TextMatrix(.Rows - 1, .ColIndex("批号")) = IIf(NVL(mrsChargeOff!批号) = "", "", mrsChargeOff!批号)
            .TextMatrix(.Rows - 1, .ColIndex("准退数量")) = Format(mrsChargeOff!准退数量 / mrsChargeOff!包装, mFMT.FM_数量)
            .TextMatrix(.Rows - 1, .ColIndex("销帐数量")) = Format(mrsChargeOff!销帐数量 / mrsChargeOff!包装, mFMT.FM_数量)
            .TextMatrix(.Rows - 1, .ColIndex("单位")) = mrsChargeOff!单位
            
            .Cell(flexcpFontBold, .Rows - 1, .ColIndex("销帐数量")) = True
            
            mrsChargeOff.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
End Sub


Public Function zlFullData(ByVal intUnit As Integer, ByVal rsNotPayStuff As ADODB.Recordset, ByVal rsChargeOff As ADODB.Recordset) As Boolean
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
    
    mbln发料时汇总销账 = (Val(zlDataBase.GetPara("发料时汇总退料销帐记录", glngSys, 1723, , , True)) = 1)
    mbln分科室 = mbln发料时汇总销账
    
    Set mrsNotPayStuff = rsNotPayStuff
    Set mrsChargeOff = rsChargeOff
    With vsfChargeOff
        .Rows = 1
    End With
    
    With vsGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .Clear (1)
        '填充数据
        zlFullData = LoadDataToVssGrid
        .Redraw = flexRDBuffered
    End With
    
    Call Form_Resize
End Function
 
Private Sub Form_Load()
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "汇总清单"
    
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

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With vsGrid
        .Top = ScaleTop
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = IIf(mbln发料时汇总销账 = False, ScaleHeight, ScaleHeight / 4 * 3)
    End With

    With picHsc
        .Visible = mbln发料时汇总销账
        .Top = vsGrid.Top + vsGrid.Height
        .Width = ScaleWidth
    End With
    
    With vsfChargeOff
        .Visible = mbln发料时汇总销账
        .Top = picHsc.Top + picHsc.Height
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = ScaleHeight - picHsc.Top - picHsc.Height - 50
    End With
End Sub
Private Function LoadDataToVssGrid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:将数据填充到网格控件中(主要是按品种统计)
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 17:13:18
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, bln分批 As Boolean, strKey As String, strTemp As String
    Dim strSort As String
    Dim dbl销帐数量 As Double
    Dim lng科室id As Long
    Dim lng材料ID As Long
    
    LoadDataToVssGrid = False
    
    If mbln分科室 = False Then
        vsGrid.ColHidden(vsGrid.ColIndex("科室")) = True
        vsGrid.ColHidden(vsGrid.ColIndex("应发数量")) = True
        vsGrid.ColHidden(vsGrid.ColIndex("销帐数量")) = True
    Else
        vsGrid.ColHidden(vsGrid.ColIndex("科室")) = False
        vsGrid.ColHidden(vsGrid.ColIndex("应发数量")) = False
        vsGrid.ColHidden(vsGrid.ColIndex("销帐数量")) = False
    End If
    
    If mrsNotPayStuff.RecordCount = 0 Then
         LoadDataToVssGrid = True
        Exit Function
    End If
        
    mrsNotPayStuff.Filter = 0
    mrsNotPayStuff.MoveFirst
    strSort = IIf(mbln分科室, "科室id Asc,", "") & "材料名称 Asc,规格 Asc,批次 Asc"
    
    mrsNotPayStuff.Sort = strSort

    With vsGrid
        .Subtotal flexSTClear
        .Rows = 2
        .Clear 1
        lngRow = .FixedRows - 1
        '手工处理后显示
        strKey = ""
        Do While Not mrsNotPayStuff.EOF
            strTemp = IIf(mbln分科室, NVL(mrsNotPayStuff!科室id) & "_", "") & NVL(mrsNotPayStuff!材料ID) & IIf(bln分批, "_" & NVL(mrsNotPayStuff!批次), "")
            
            If strKey <> strTemp And mrsNotPayStuff!执行状态 = 1 Then
                .Rows = .Rows + 1
                lngRow = lngRow + 1
                .TextMatrix(lngRow, .ColIndex("科室")) = IIf(mbln分科室, mrsNotPayStuff!科室, "")
                .TextMatrix(lngRow, .ColIndex("材料名称")) = mrsNotPayStuff!材料名称
                .TextMatrix(lngRow, .ColIndex("材料id")) = mrsNotPayStuff!材料ID
                .TextMatrix(lngRow, .ColIndex("规格")) = mrsNotPayStuff!规格
                .TextMatrix(lngRow, .ColIndex("产地")) = mrsNotPayStuff!产地
                .TextMatrix(lngRow, .ColIndex("批号")) = IIf(bln分批, mrsNotPayStuff!批号, "")
                .TextMatrix(lngRow, .ColIndex("单价")) = Format(mrsNotPayStuff!单价 * mrsNotPayStuff!换算系数, mFMT.FM_零售价)
                .TextMatrix(lngRow, .ColIndex("单位")) = NVL(mrsNotPayStuff!单位)

                .Cell(flexcpData, lngRow, .ColIndex("批号")) = IIf(bln分批, NVL(mrsNotPayStuff!批次), 0)
                .Cell(flexcpData, lngRow, .ColIndex("科室")) = IIf(mbln分科室, NVL(mrsNotPayStuff!科室id), "")
                .Cell(flexcpData, lngRow, .ColIndex("材料名称")) = mrsNotPayStuff!材料ID
                strKey = strTemp
            End If
            If mrsNotPayStuff!执行状态 = 1 Then
                '只具备发料的才汇总
                .Cell(flexcpData, lngRow, .ColIndex("应发数量")) = Val(.Cell(flexcpData, lngRow, .ColIndex("应发数量"))) + (mrsNotPayStuff!实际数量 * mrsNotPayStuff!付)
                .TextMatrix(lngRow, .ColIndex("应发数量")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("应发数量"))) / mrsNotPayStuff!换算系数, mFMT.FM_数量)

                .Cell(flexcpData, lngRow, .ColIndex("实发数量")) = Val(.Cell(flexcpData, lngRow, .ColIndex("实发数量"))) + (mrsNotPayStuff!实际数量 * mrsNotPayStuff!付)
                .TextMatrix(lngRow, .ColIndex("实发数量")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("实发数量"))) / mrsNotPayStuff!换算系数, mFMT.FM_数量)
                .Cell(flexcpFontBold, lngRow, .ColIndex("实发数量")) = True
                
                .Cell(flexcpData, lngRow, .ColIndex("金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("金额"))) + (mrsNotPayStuff!金额)
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("金额"))), mFMT.FM_金额)
            End If
            mrsNotPayStuff.MoveNext
        Loop
        
        '合并销账数据
        mrsChargeOff.Filter = "执行状态=1"
        If mbln分科室 = True And mrsChargeOff.RecordCount > 0 Then
            For lngRow = 1 To .Rows - 1
                If Val(.Cell(flexcpData, lngRow, .ColIndex("科室"))) > 0 Then
                    .TextMatrix(lngRow, .ColIndex("应发数量")) = .TextMatrix(lngRow, .ColIndex("实发数量"))

                    lng科室id = Val(.Cell(flexcpData, lngRow, .ColIndex("科室")))
                    lng材料ID = Val(.TextMatrix(lngRow, .ColIndex("材料id")))

                    mrsChargeOff.Filter = " 执行状态=1 And 领料部门id=" & lng科室id & " And 材料ID=" & lng材料ID
                    If mrsChargeOff.RecordCount > 0 Then
                        dbl销帐数量 = 0
                        Do While Not mrsChargeOff.EOF
                            dbl销帐数量 = dbl销帐数量 + mrsChargeOff!销帐数量
                            mrsChargeOff.MoveNext
                        Loop
                        
                        .TextMatrix(lngRow, .ColIndex("销帐数量")) = Format(dbl销帐数量, mFMT.FM_数量)
                        .TextMatrix(lngRow, .ColIndex("实发数量")) = Format(Val(.TextMatrix(lngRow, .ColIndex("应发数量"))) - Val(.TextMatrix(lngRow, .ColIndex("销帐数量"))), mFMT.FM_数量)
                    End If
                End If
            Next
        End If

        If mrsNotPayStuff.RecordCount <> 0 Then mrsNotPayStuff.MoveFirst
        If .Rows > 2 Then .Rows = .Rows - 1
        If .Rows = 2 And Val(.Cell(flexcpData, 1, .ColIndex("材料名称"))) = 0 Then
        Else
            Call SetTotalRowData(mbln分科室)
        End If
        
        If Val(.Cell(flexcpData, .Row, .ColIndex("科室"))) = 0 Then
            vsfChargeOff.Rows = 1
        Else
            Call LoadDataToChargeOffList(Val(.Cell(flexcpData, .Row, .ColIndex("科室"))), Val(.TextMatrix(.Row, .ColIndex("材料id"))))
        End If
    End With
    
    LoadDataToVssGrid = True
End Function
Private Function SetTotalRowData(ByVal bln科室汇总 As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置行的汇总属性
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-22 10:22:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long
    With vsGrid
        .Redraw = flexRDNone
        .OutlineBar = flexOutlineBarComplete
        .SubtotalPosition = flexSTBelow
        If bln科室汇总 = True Then
            .Subtotal flexSTSum, .ColIndex("科室"), .ColIndex("金额"), mFMT.FM_金额, , vbBlue, True, "小计"
        End If
        .Subtotal flexSTSum, -1, .ColIndex("金额"), mFMT.FM_金额, , vbBlue, True, "合计"
        If bln科室汇总 = False Then .TextMatrix(.Rows - 1, .ColIndex("材料名称")) = "合计"
        .Redraw = flexRDBuffered
        
    End With
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
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "汇总清单"
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
    
    With vsfChargeOff
        .Font.Size = curFontSize
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("刘") + 120
        .RowHeightMax = TextHeight("刘") + 120
        .Refresh
    End With
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow Then
        With vsGrid
            If Val(.Cell(flexcpData, NewRow, .ColIndex("科室"))) = 0 Then
                vsfChargeOff.Rows = 1
                Exit Sub
            End If
            
            Call LoadDataToChargeOffList(Val(.Cell(flexcpData, NewRow, .ColIndex("科室"))), Val(.TextMatrix(NewRow, .ColIndex("材料id"))))
        End With
    End If
End Sub

Private Sub vsGrid_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    With vsGrid
        If Position <= .ColIndex("材料名称") Then
            ShowMsgBox "不能将列移动到材料名称以前的列!"
            Position = Col
        End If
    End With
End Sub
