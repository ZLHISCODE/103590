VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPurchaseInputReturn 
   Caption         =   "按入库单据退货"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12315
   Icon            =   "frmPurchaseInputReturn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   12315
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboInputDate 
      Height          =   300
      Left            =   4935
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   60
      Width           =   1440
   End
   Begin VB.TextBox txtNO 
      Height          =   300
      Left            =   2400
      TabIndex        =   13
      Top             =   60
      Width           =   1440
   End
   Begin VB.TextBox txtLeechdom 
      Height          =   300
      Left            =   480
      TabIndex        =   12
      Top             =   60
      Width           =   1440
   End
   Begin VB.CheckBox chkAllSelect 
      Height          =   200
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   1320
      TabIndex        =   10
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8280
      TabIndex        =   9
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6960
      TabIndex        =   8
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "提取(&G)"
      Height          =   300
      Left            =   11400
      TabIndex        =   6
      Top             =   60
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtp开始时间 
      Height          =   315
      Left            =   7440
      TabIndex        =   2
      Top             =   53
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   214630403
      CurrentDate     =   40848
   End
   Begin MSComCtl2.DTPicker dtp结束时间 
      Height          =   315
      Left            =   9960
      TabIndex        =   3
      Top             =   53
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   214630403
      CurrentDate     =   40848
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   2925
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   9615
      _cx             =   16960
      _cy             =   5159
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   44
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPurchaseInputReturn.frx":000C
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
      VirtualData     =   0   'False
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
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "入库时间"
      Height          =   180
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lbl结束日期 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "结束日期"
      Height          =   180
      Left            =   9120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lbl开始日期 
      Caption         =   "开始日期"
      Height          =   255
      Left            =   6645
      TabIndex        =   4
      Top             =   83
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "NO"
      Height          =   180
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   180
   End
   Begin VB.Label lblLeechdom 
      AutoSize        =   -1  'True
      Caption         =   "药品"
      Height          =   180
      Left            =   50
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmPurchaseInputReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngProvider As Long
Private mintUnit As Integer   '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mrsData As ADODB.Recordset  '要返回的数据集
Private mlngStoreroom As Long   '库房id
Private mstrSelectInfo As String '已经选择的药品

Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数

Public Sub ShowMe(ByVal frmParent As Form, ByVal lngStoreroom As Long, ByVal lngProvider As Long, ByVal intUnit As Integer, ByVal intCostDigit As Integer, ByVal intPricedigit As Integer, ByVal intNumberDigit As Integer, ByVal intMoneyDigit As Integer, ByRef rsData As ADODB.Recordset)
    mlngStoreroom = lngStoreroom
    mintUnit = intUnit
    mlngProvider = lngProvider
    mintCostDigit = intCostDigit
    mintPriceDigit = intPricedigit
    mintNumberDigit = intNumberDigit
    mintMoneyDigit = intMoneyDigit
    
    Me.Show vbModal, frmParent
    Set rsData = mrsData
End Sub

Private Sub cboInputDate_Click()
    With cboInputDate
        If .Text = "自定义" Then
            lbl开始日期.Visible = True
            dtp开始时间.Visible = True
            Lbl结束日期.Visible = True
            dtp结束时间.Visible = True
        Else
            lbl开始日期.Visible = False
            dtp开始时间.Visible = False
            Lbl结束日期.Visible = False
            dtp结束时间.Visible = False
        End If
    End With
End Sub

Private Sub chkAllSelect_Click()
    Dim lngRow As Long
    
    With vsfList
        If .rows = 1 Then Exit Sub
        If chkAllSelect.Value = 1 Then
            For lngRow = 1 To .rows - 1
                If InStr(1, mstrSelectInfo, .TextMatrix(lngRow, .ColIndex("药品id")) & "," & Val(.TextMatrix(lngRow, .ColIndex("批次"))) & "|") = 0 Then
                    .TextMatrix(lngRow, 0) = "√"
                    mstrSelectInfo = mstrSelectInfo & .TextMatrix(lngRow, .ColIndex("药品id")) & "," & Val(.TextMatrix(lngRow, .ColIndex("批次"))) & "|"
                End If
            Next
        Else
            For lngRow = 1 To .rows - 1
                .TextMatrix(lngRow, 0) = ""
            Next
            mstrSelectInfo = ""
        End If
    End With
End Sub

Private Sub cmdAllCls_Click()
    With vsfList
        .rows = 1
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetData_Click()
    Dim dbBeginDate As Date
    Dim dbEndDate As Date
    Dim rsTemp As ADODB.Recordset
    Dim rsSumNum As ADODB.Recordset
    Dim lngLeechdom As Long
    Dim strNo As String
    Dim lngRow As Long
    Dim strCurUnit As String
    Dim int换算系数 As Integer
    
    On Error GoTo errHandle
    
    If cboInputDate.Text = "今日内" Then
        dbBeginDate = CDate(Format(Date, "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "一周内" Then
        dbBeginDate = CDate(Format(DateAdd("d", -7, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "一月内" Then
        dbBeginDate = CDate(Format(DateAdd("M", -1, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "三个月内" Then
        dbBeginDate = CDate(Format(DateAdd("M", -3, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "一年内" Then
        dbBeginDate = CDate(Format(DateAdd("yyyy", -1, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "自定义" Then
        dbBeginDate = CDate(Format(dtp开始时间.Value, "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(dtp结束时间.Value, "yyyy-mm-dd") & " 23:59:59")
    End If
    
    strNo = Trim(txtNo.Text)
    
    gstrSQL = ""
    If Trim(txtLeechdom.Text) <> "" And txtLeechdom.Tag <> "" Then
        lngLeechdom = txtLeechdom.Tag
        gstrSQL = " And d.id = [5]"
    End If
    If Trim(txtNo.Text) <> "" Then
        gstrSQL = gstrSQL & " And a.no = [6]"
    End If
    '超过30天就提示用户
    If dbEndDate - dbBeginDate > 30 Then
        If MsgBox("查询时间范围太长，是否继续？", vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    gstrSQL = "Select Distinct a.No, a.药品id, d.编码, d.名称, d.规格, Nvl(d.是否变价, 0) As 是否变价, d.计算单位, e.门诊单位, e.门诊包装, e.住院单位, e.住院包装, e.药库单位," & vbNewLine & _
            "                e.药库包装, e.指导批发价, e.指导差价率, e.药品来源, e.基本药物, e.药价级别, e.药库分批, e.药房分批, e.最大效期, Nvl(e.差价让利比, 0) As 差价让利比," & vbNewLine & _
            "                Nvl(e.招标药品, 0) As 招标药品, a.产地, a.批号, nvl(a.批次,0) as 批次, a.外观, a.批准文号, a.生产日期, a.实际数量, a.成本价, a.零售价, a.成本金额, a.零售金额," & vbNewLine & _
            "                a.差价, a.生产日期, a.效期, c.随货单号, c.发票号, c.发票代码, c.发票金额, c.发票日期" & vbNewLine & _
            "From 药品收发记录 A, 药品库存 B, 应付记录 C, 收费项目目录 D, 药品规格 E" & vbNewLine & _
            "Where a.药品id + 0 = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) And a.库房id = b.库房id And a.Id = c.收发id(+) And b.药品id = d.Id And" & vbNewLine & _
            "      d.Id = e.药品id And b.库房id = [1] And a.供药单位id + 0 = [2] And b.性质 = 1 And a.单据 = 1 And Nvl(a.发药方式, 0) = 0 And" & vbNewLine & _
            "      (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0) And a.审核日期 Between [3] And" & vbNewLine & _
            "      [4]" & gstrSQL & vbNewLine & _
            "order by no"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询外购记录", mlngStoreroom, mlngProvider, dbBeginDate, dbEndDate, lngLeechdom, strNo)
    mstrSelectInfo = ""
    vsfList.rows = 1
    chkAllSelect.Value = 0
    
    If rsTemp.RecordCount > 0 Then
        vsfList.rows = rsTemp.RecordCount + 1
        lngRow = 1
    End If
    Do While Not rsTemp.EOF
        With vsfList
            .TextMatrix(lngRow, .ColIndex("no")) = rsTemp!NO
            .TextMatrix(lngRow, .ColIndex("编码名称")) = "[" & rsTemp!编码 & "]" & IIf(IsNull(rsTemp!名称), "", rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("规格")) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
            .TextMatrix(lngRow, .ColIndex("产地")) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
            .TextMatrix(lngRow, .ColIndex("批号")) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
            .TextMatrix(lngRow, .ColIndex("批次")) = IIf(IsNull(rsTemp!批次), "", rsTemp!批次)
            .TextMatrix(lngRow, .ColIndex("药品来源")) = IIf(IsNull(rsTemp!药品来源), "", rsTemp!药品来源)
            .TextMatrix(lngRow, .ColIndex("基本药物")) = IIf(IsNull(rsTemp!基本药物), "", rsTemp!基本药物)
            .TextMatrix(lngRow, .ColIndex("批准文号")) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
            .TextMatrix(lngRow, .ColIndex("生产日期")) = IIf(IsNull(rsTemp!生产日期), "", rsTemp!生产日期)
            '单位系数：1-售价;2-门诊;3-住院;4-药库
            Select Case mintUnit
            Case 1
                strCurUnit = rsTemp!计算单位
                int换算系数 = 1
            Case 2
                strCurUnit = rsTemp!门诊单位
                int换算系数 = rsTemp!门诊包装
            Case 3
                strCurUnit = rsTemp!住院单位
                int换算系数 = rsTemp!住院包装
            Case 4
                strCurUnit = rsTemp!药库单位
                int换算系数 = rsTemp!药库包装
            End Select
            .TextMatrix(lngRow, .ColIndex("单位")) = strCurUnit
            
            gstrSQL = "Select Sum(实际数量) As 实际数量" & vbNewLine & _
                "From 药品库存" & vbNewLine & _
                "Where 性质 = 1 And 库房id = [1] And 药品id = [2] And Nvl(批次, 0) = [3]"
            Set rsSumNum = zlDatabase.OpenSQLRecord(gstrSQL, "查询库存", mlngStoreroom, rsTemp!药品id, rsTemp!批次)
            If rsSumNum.RecordCount > 0 Then
                .TextMatrix(lngRow, .ColIndex("库存数量")) = GetFormat(rsSumNum!实际数量, mintNumberDigit)
            Else
                .TextMatrix(lngRow, .ColIndex("库存数量")) = 0
            End If
            .TextMatrix(lngRow, .ColIndex("入库数量")) = GetFormat(rsTemp!实际数量, mintNumberDigit)
            .TextMatrix(lngRow, .ColIndex("成本价")) = GetFormat(IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价 * int换算系数), mintCostDigit)
            .TextMatrix(lngRow, .ColIndex("售价")) = GetFormat(IIf(IsNull(rsTemp!零售价), 0, rsTemp!零售价 * int换算系数), mintPriceDigit)
            .TextMatrix(lngRow, .ColIndex("成本金额")) = GetFormat(rsTemp!成本金额, mintMoneyDigit)
            .TextMatrix(lngRow, .ColIndex("售价金额")) = GetFormat(rsTemp!零售金额, mintMoneyDigit)
            .TextMatrix(lngRow, .ColIndex("差价")) = GetFormat(rsTemp!差价, mintMoneyDigit)
            .TextMatrix(lngRow, .ColIndex("随货单号")) = IIf(IsNull(rsTemp!随货单号), "", rsTemp!随货单号)
            .TextMatrix(lngRow, .ColIndex("发票号")) = IIf(IsNull(rsTemp!发票号), "", rsTemp!发票号)
            .TextMatrix(lngRow, .ColIndex("发票代码")) = IIf(IsNull(rsTemp!发票代码), "", rsTemp!发票代码)
            .TextMatrix(lngRow, .ColIndex("发票金额")) = GetFormat(IIf(IsNull(rsTemp!发票金额), 0, rsTemp!发票金额), mintMoneyDigit)
            .TextMatrix(lngRow, .ColIndex("发票日期")) = Format(IIf(IsNull(rsTemp!发票日期), "", rsTemp!发票日期), "yyyy-mm-dd")
            .TextMatrix(lngRow, .ColIndex("药品id")) = rsTemp!药品id
            .TextMatrix(lngRow, .ColIndex("效期")) = Format(IIf(IsNull(rsTemp!效期), "", rsTemp!效期), "yyyy-mm-dd")
            .TextMatrix(lngRow, .ColIndex("指导批发价")) = GetFormat(rsTemp!指导批发价, mintCostDigit)
            .TextMatrix(lngRow, .ColIndex("指导差价率")) = GetFormat(rsTemp!指导差价率, mintCostDigit)
            .TextMatrix(lngRow, .ColIndex("是否变价")) = rsTemp!是否变价
            .TextMatrix(lngRow, .ColIndex("药库分批")) = IIf(IsNull(rsTemp!药库分批), 0, rsTemp!药库分批)
            .TextMatrix(lngRow, .ColIndex("药房分批")) = IIf(IsNull(rsTemp!药房分批), 0, rsTemp!药房分批)
            .TextMatrix(lngRow, .ColIndex("最大效期")) = IIf(IsNull(rsTemp!最大效期), 0, rsTemp!最大效期)
            .TextMatrix(lngRow, .ColIndex("差价让利比")) = IIf(IsNull(rsTemp!差价让利比), 0, rsTemp!差价让利比)
            .TextMatrix(lngRow, .ColIndex("招标药品")) = IIf(IsNull(rsTemp!招标药品), 0, rsTemp!招标药品)
            .TextMatrix(lngRow, .ColIndex("比例系数")) = int换算系数
            .TextMatrix(lngRow, .ColIndex("计算单位")) = IIf(IsNull(rsTemp!计算单位), "", rsTemp!计算单位)
            .TextMatrix(lngRow, .ColIndex("门诊单位")) = IIf(IsNull(rsTemp!门诊单位), "", rsTemp!门诊单位)
            .TextMatrix(lngRow, .ColIndex("门诊包装")) = IIf(IsNull(rsTemp!门诊包装), 1, rsTemp!门诊包装)
            .TextMatrix(lngRow, .ColIndex("住院单位")) = IIf(IsNull(rsTemp!住院单位), "", rsTemp!住院单位)
            .TextMatrix(lngRow, .ColIndex("住院包装")) = IIf(IsNull(rsTemp!住院包装), 1, rsTemp!住院包装)
            .TextMatrix(lngRow, .ColIndex("药库单位")) = IIf(IsNull(rsTemp!药库单位), "", rsTemp!药库单位)
            .TextMatrix(lngRow, .ColIndex("药库包装")) = IIf(IsNull(rsTemp!药库包装), 1, rsTemp!药库包装)
            .TextMatrix(lngRow, .ColIndex("药价级别")) = IIf(IsNull(rsTemp!药价级别), "", rsTemp!药价级别)
            .TextMatrix(lngRow, .ColIndex("外观")) = IIf(IsNull(rsTemp!外观), "", rsTemp!外观)
            lngRow = lngRow + 1
        End With
        rsTemp.MoveNext
    Loop
    
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CmdSave_Click()
    Dim lngRow As Long
    
    Call GetAssembled   '初始化数据集
    
    With vsfList
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) = "√" And .TextMatrix(lngRow, .ColIndex("药品id")) <> "" Then
                mrsData.AddNew
                
                mrsData!药品id = .TextMatrix(lngRow, .ColIndex("药品id"))
                mrsData!编码名称 = .TextMatrix(lngRow, .ColIndex("编码名称"))
                mrsData!规格 = .TextMatrix(lngRow, .ColIndex("规格"))
                mrsData!产地 = .TextMatrix(lngRow, .ColIndex("产地"))
                mrsData!批号 = .TextMatrix(lngRow, .ColIndex("批号"))
                mrsData!批次 = .TextMatrix(lngRow, .ColIndex("批次"))
                mrsData!药品来源 = .TextMatrix(lngRow, .ColIndex("药品来源"))
                mrsData!基本药物 = .TextMatrix(lngRow, .ColIndex("基本药物"))
                mrsData!批准文号 = .TextMatrix(lngRow, .ColIndex("批准文号"))
                mrsData!生产日期 = .TextMatrix(lngRow, .ColIndex("生产日期"))
                mrsData!单位 = .TextMatrix(lngRow, .ColIndex("单位"))
                mrsData!库存数量 = .TextMatrix(lngRow, .ColIndex("库存数量"))
                mrsData!入库数量 = .TextMatrix(lngRow, .ColIndex("入库数量"))
                mrsData!售价 = .TextMatrix(lngRow, .ColIndex("售价"))
                mrsData!成本价 = .TextMatrix(lngRow, .ColIndex("成本价"))
                mrsData!成本金额 = .TextMatrix(lngRow, .ColIndex("成本金额"))
                mrsData!售价金额 = .TextMatrix(lngRow, .ColIndex("售价金额"))
                mrsData!差价 = .TextMatrix(lngRow, .ColIndex("差价"))
                mrsData!比例系数 = .TextMatrix(lngRow, .ColIndex("比例系数"))
                mrsData!随货单号 = .TextMatrix(lngRow, .ColIndex("随货单号"))
                mrsData!发票号 = .TextMatrix(lngRow, .ColIndex("发票号"))
                mrsData!发票代码 = .TextMatrix(lngRow, .ColIndex("发票代码"))
                mrsData!发票金额 = .TextMatrix(lngRow, .ColIndex("发票金额"))
                mrsData!发票日期 = .TextMatrix(lngRow, .ColIndex("发票日期"))
                mrsData!效期 = .TextMatrix(lngRow, .ColIndex("效期"))
                mrsData!指导批发价 = .TextMatrix(lngRow, .ColIndex("指导批发价"))
                mrsData!指导差价率 = .TextMatrix(lngRow, .ColIndex("指导差价率"))
                mrsData!是否变价 = .TextMatrix(lngRow, .ColIndex("是否变价"))
                mrsData!药库分批 = .TextMatrix(lngRow, .ColIndex("药库分批"))
                mrsData!药房分批 = .TextMatrix(lngRow, .ColIndex("药房分批"))
                mrsData!最大效期 = .TextMatrix(lngRow, .ColIndex("最大效期"))
                mrsData!差价让利比 = .TextMatrix(lngRow, .ColIndex("差价让利比"))
                mrsData!招标药品 = .TextMatrix(lngRow, .ColIndex("招标药品"))
                mrsData!药价级别 = .TextMatrix(lngRow, .ColIndex("药价级别"))
                mrsData!外观 = .TextMatrix(lngRow, .ColIndex("外观"))
                mrsData!计算单位 = .TextMatrix(lngRow, .ColIndex("计算单位"))
                mrsData!门诊单位 = .TextMatrix(lngRow, .ColIndex("门诊单位"))
                mrsData!门诊包装 = .TextMatrix(lngRow, .ColIndex("门诊包装"))
                mrsData!住院单位 = .TextMatrix(lngRow, .ColIndex("住院单位"))
                mrsData!住院包装 = .TextMatrix(lngRow, .ColIndex("住院包装"))
                mrsData!药库单位 = .TextMatrix(lngRow, .ColIndex("药库单位"))
                mrsData!药库包装 = .TextMatrix(lngRow, .ColIndex("药库包装"))
                
                mrsData.Update
            End If
        Next
        Unload Me
    End With
End Sub

Private Sub GetAssembled()
    '初始化数据集
    Set mrsData = New ADODB.Recordset

    With mrsData
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "编码名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 60, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "批次", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "药品来源", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "基本药物", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "批准文号", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "生产日期", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "库存数量", adDouble, 18, adFldIsNullable
        .Fields.Append "入库数量", adDouble, 18, adFldIsNullable
        .Fields.Append "成本价", adDouble, 16, adFldIsNullable
        .Fields.Append "售价", adDouble, 16, adFldIsNullable
        .Fields.Append "成本金额", adDouble, 18, adFldIsNullable
        .Fields.Append "售价金额", adDouble, 18, adFldIsNullable
        .Fields.Append "差价", adDouble, 18, adFldIsNullable
        .Fields.Append "比例系数", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "随货单号", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "发票号", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "发票代码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "发票金额", adDouble, 18, adFldIsNullable
        .Fields.Append "发票日期", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "效期", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "指导批发价", adDouble, 16, adFldIsNullable
        .Fields.Append "指导差价率", adDouble, 16, adFldIsNullable
        .Fields.Append "是否变价", adLongVarChar, 2, adFldIsNullable
        .Fields.Append "药库分批", adLongVarChar, 2, adFldIsNullable
        .Fields.Append "药房分批", adLongVarChar, 2, adFldIsNullable
        .Fields.Append "最大效期", adLongVarChar, 5, adFldIsNullable
        .Fields.Append "差价让利比", adLongVarChar, 5, adFldIsNullable
        .Fields.Append "招标药品", adDouble, 16, adFldIsNullable
        .Fields.Append "药价级别", adLongVarChar, 2, adFldIsNullable
        .Fields.Append "外观", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "计算单位", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "门诊单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "门诊包装", adDouble, 11, adFldIsNullable
        .Fields.Append "住院单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "住院包装", adDouble, 11, adFldIsNullable
        .Fields.Append "药库单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "药库包装", adDouble, 11, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
End Sub

Private Sub Form_Load()
    Set mrsData = Nothing
    
    dtp开始时间.Value = DateAdd("d", -2, zlDatabase.Currentdate)
    dtp结束时间.Value = DateAdd("d", -0, zlDatabase.Currentdate)
    Call InitComBox
End Sub

Private Sub Form_Resize()
    cmdGetData.Left = Me.ScaleWidth - cmdGetData.Width - 100
    With cmdAllCls
        .Top = Me.ScaleHeight - .Height - 100
        .Left = 200
    End With
    
    With cmdCancel
        .Top = cmdAllCls.Top
        .Left = Me.ScaleWidth - .Width - 200
    End With
    
    With CmdSave
        .Top = cmdAllCls.Top
        .Left = cmdCancel.Left - .Width - 200
    End With
    
    With vsfList
        .Left = 50
        .Top = txtLeechdom.Top + txtLeechdom.Height + 50
        .Width = Me.ScaleWidth - 50
        .Height = Me.ScaleHeight - .Top - cmdAllCls.Height - 200
    End With
    
    With chkAllSelect
        .Left = vsfList.Left + 70
        .Top = vsfList.Top + 60
    End With
End Sub

Private Sub InitComBox()
    '初始化下拉列表
    With cboInputDate
        .AddItem "今日内"
        .AddItem "一周内"
        .AddItem "一月内"
        .AddItem "三个月内"
        .AddItem "一年内"
        .AddItem "自定义"
        .ListIndex = 0
    End With
End Sub

Private Sub txtLeechdom_Change()
    If txtLeechdom.Text = "" Then
        txtLeechdom.Tag = ""
    End If
End Sub

Private Sub txtLeechdom_GotFocus()
    zlControl.TxtSelAll txtLeechdom
End Sub

Private Sub txtLeechdom_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecReturn As ADODB.Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtLeechdom.Text) = "" Then Exit Sub
    sngLeft = Me.Left + txtLeechdom.Left
    sngTop = Me.Top + txtLeechdom.Top + txtLeechdom.Height + 500
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - txtLeechdom.Height - 3630
    End If
    
    strkey = Trim(txtLeechdom.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "药品外购入库管理", mlngStoreroom, mlngStoreroom)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 1, 1, txtLeechdom.Text, sngLeft, sngTop, mlngStoreroom, mlngStoreroom, mlngStoreroom, , , , , , False)
    If RecReturn.RecordCount > 0 Then
        txtLeechdom.Tag = RecReturn!药品id
        txtLeechdom.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    Else
        txtLeechdom.Tag = ""
    End If
End Sub



Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNo
End Sub


Private Sub txtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txtNo) < 8 And Len(txtNo) > 0 Then
            txtNo.Text = GetFullNO(txtNo.Text, 21, mlngStoreroom)
        End If
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub


Private Sub vsfList_Click()
    With vsfList
        If .Row = 0 Or .rows = 1 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then
            If InStr(1, mstrSelectInfo, .TextMatrix(.Row, .ColIndex("药品id")) & "," & Val(.TextMatrix(.Row, .ColIndex("批次"))) & "|") = 0 Then
                .TextMatrix(.Row, 0) = "√"
                mstrSelectInfo = mstrSelectInfo & .TextMatrix(.Row, .ColIndex("药品id")) & "," & Val(.TextMatrix(.Row, .ColIndex("批次"))) & "|"
            Else
                MsgBox "此批次药品已经选择了，不需要选择多次！", vbInformation, gstrSysName
            End If
        Else
            .TextMatrix(.Row, 0) = ""
            If InStr(1, mstrSelectInfo, .TextMatrix(.Row, .ColIndex("药品id")) & "," & Val(.TextMatrix(.Row, .ColIndex("批次"))) & "|") > 0 Then
                mstrSelectInfo = Replace(mstrSelectInfo, .TextMatrix(.Row, .ColIndex("药品id")) & "," & Val(.TextMatrix(.Row, .ColIndex("批次"))) & "|", "")
            End If
        End If
    End With
End Sub


