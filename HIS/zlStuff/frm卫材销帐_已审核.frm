VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frm卫材销帐_已审核 
   BorderStyle     =   0  'None
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   12615
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   12615
   End
   Begin VB.PictureBox picBatHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   12735
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4155
      Width           =   12735
   End
   Begin VSFlex8Ctl.VSFlexGrid vsHeadGrid 
      Height          =   2055
      Left            =   75
      TabIndex        =   0
      Tag             =   "待处理"
      Top             =   105
      Width           =   10800
      _cx             =   19050
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm卫材销帐_已审核.frx":0000
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
   Begin VSFlex8Ctl.VSFlexGrid vsDetail 
      Height          =   1935
      Left            =   15
      TabIndex        =   1
      Tag             =   "明细"
      Top             =   2205
      Width           =   10800
      _cx             =   19050
      _cy             =   3413
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm卫材销帐_已审核.frx":00B8
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
   Begin VSFlex8Ctl.VSFlexGrid vsBatch 
      Height          =   1815
      Left            =   15
      TabIndex        =   2
      Tag             =   "明细"
      Top             =   4245
      Width           =   10800
      _cx             =   19050
      _cy             =   3201
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm卫材销帐_已审核.frx":01F7
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
End
Attribute VB_Name = "frm卫材销帐_已审核"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mstrPrivs As String
Private mlngModule As Long
Private mArrFilter As Variant   '过滤条件
Private mrsVerifyBatch As ADODB.Recordset
Private mrsVerifyDetail As ADODB.Recordset      '已审核明细记录数据集

Private mint审核标志 As Integer
Private mintUnit As Integer '0-散装单位,1-包装单位
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
 
Private Sub initPara()
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
    With vsHeadGrid
       ' .Editable = flexEDKbdMouse
    End With
    
End Sub
Private Function RefreshData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:重新获取数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-03 21:09:18
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As ADODB.Recordset, strFields As String, strWere As String, lngRow As Long
    Dim str费用ID As String
    
    On Error GoTo ErrHandle
    Call InitRsStruct
    mint审核标志 = 1
    ''''1、提取汇总数据
    '单位，包装换算
    Select Case mintUnit
    Case 0
        strFields = "X.计算单位 单位,1 换算系数,A.数量 As 销帐数量 "
    Case Else
        strFields = "D.包装单位 单位,d.换算系数 ,A.数量 As 销帐数量 "
    End Select
    
    strWere = strWere & " And A.审核人 Is Not Null And A.状态 <> 0 And A.审核时间 Between [3] And [4] "
      
    '病区/医技科室
    If Val(mArrFilter("申请科室ID")) > 0 Then strWere = strWere & " And A.申请部门id = [2] "
    '申请人
    If Trim(mArrFilter("申请人")) <> "" Then strWere = strWere & " And A.申请人=[7] "
    '病人姓名
    If Trim(mArrFilter("病人姓名")) <> "" Then strWere = strWere & " And nvl(F.姓名,B.姓名)=[8] "
    strWere = strWere & IIf(Val(mArrFilter("住院号")) = 0, "", "             AND b.标识号=[9] and b.门诊标志=2 ")
    strWere = strWere & IIf(Val(mArrFilter("病人ID")) = 0, "", "             AND b.病人iD=[10]  ")
    strWere = strWere & IIf(Trim(mArrFilter("床号")) = "", "", "             AND b.病人iD=[10]  ")

    
    gstrSQL = "" & _
    "   Select Distinct A.状态,A.收费细目id,'['||X.编码||']'||X.名称 as 材料名称, X.规格, " & strFields & _
    "   From (  Select A.状态,A.收费细目id, Sum(A.数量) As 数量 " & _
    "           From 住院费用记录 B, 病案主页 F, 病人费用销帐 A " & _
    "           Where A.申请类别=1 And A.费用id = B.ID And A.审核部门id = [1] And B.病人id = F.病人id(+)  " & _
    "                 And B.主页id = F.主页id(+)  And F.出院日期(+)  Is Null And F.状态(+)<> 3 " & _
    "                 " & vbCrLf & strWere & _
    "                 And Exists (Select 1 From 药品收发记录 C  Where C.费用id = A.费用id And C.审核人 Is Not Null And (C.记录状态 = 1 Or Mod(C.记录状态, 3) = 0))" & _
    "           Group By A.收费细目id,A.状态) A,材料特性 D, 收费项目别名 E, 收费项目目录 X " & _
    " Where A.收费细目id = D.材料id And A.收费细目id = X.ID And X.ID = E.收费细目id(+) And E.性质(+) = 3 " & _
    " Order By 材料名称"
     
    '[5]和[6]参数为发料时间,现取消
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取退料申请", _
        Val(mArrFilter("发料部门id")), Val(mArrFilter("申请科室id")), _
        CDate(mArrFilter("审核日期")(0)), CDate(mArrFilter("审核日期")(1)), _
        CDate(mArrFilter("审核日期")(0)), CDate(mArrFilter("审核日期")(1)), _
        Trim(mArrFilter("申请人")), Trim(mArrFilter("病人姓名")), _
        Val(mArrFilter("住院号")), Val(mArrFilter("病人ID")), Trim(mArrFilter("床号")))
    
    With vsHeadGrid
        .Clear 1
        .Rows = 2
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 0
        Do While Not rsTemp.EOF
            lngRow = lngRow + 1
            .TextMatrix(lngRow, .ColIndex("审核")) = IIf(Val(NVL(rsTemp!状态)) = 1, "√", "×")
            .TextMatrix(lngRow, .ColIndex("材料名称")) = NVL(rsTemp!材料名称)
            .TextMatrix(lngRow, .ColIndex("规格")) = NVL(rsTemp!规格)
            .TextMatrix(lngRow, .ColIndex("销帐数量")) = Format(Val(NVL(rsTemp!销帐数量)) / rsTemp!换算系数, mFMT.FM_数量)
            .TextMatrix(lngRow, .ColIndex("单位")) = NVL(rsTemp!单位)
            .Cell(flexcpData, lngRow, .ColIndex("材料名称")) = NVL(rsTemp!收费细目id)
            .Cell(flexcpData, lngRow, .ColIndex("销帐数量")) = NVL(rsTemp!销帐数量)
            rsTemp.MoveNext
        Loop

    End With
    
    
    ''''2、提取明细数据
    '单位字串
    Select Case mintUnit
    Case 0
        strFields = "X.计算单位 单位,1 换算系数, A.数量 "
    Case Else
        strFields = "D.包装单位 单位,D.换算系数, A.数量 "
    End Select
    
    gstrSQL = "" & _
        "   Select 单据, NO, 药品ID as 材料ID, 申请时间, 标识号, 姓名, 床号, 单位, 换算系数,开单科室, Sum(数量) As 销帐数量 " & _
        "   From (  Select Distinct C.单据, C.NO, C.药品ID, A.申请时间, B.标识号, nvl(F.姓名,B.姓名) 姓名, B.床号,P.名称 开单科室, " & strFields & " " & _
        "           From 病人费用销帐 A, 住院费用记录 B,药品收发记录 C, 材料特性 D, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E " & _
        "           Where   A.申请类别=1 And A.费用id = B.ID And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.材料id And B.收费细目id = X.ID  " & _
        "                   And  B.病人id = F.病人id(+) And B.主页id = F.主页id(+) And F.出院日期(+) Is Null And F.状态(+) <> 3 " & _
        "                   And A.申请部门id = E.ID And B.执行部门id = [1]  " & _
        "                   And C.审核人 Is Not Null And C.单据 In (24, 25, 26) And (C.记录状态 = 1 Or Mod(C.记录状态, 3) = 0)   " & strWere & ")" & _
        "           Group By 单据, NO, 药品ID, 申请时间, 标识号, 姓名, 床号, 单位, 换算系数,开单科室 "
        
           
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取单据明细", _
          Val(mArrFilter("发料部门id")), Val(mArrFilter("申请科室id")), _
        CDate(mArrFilter("审核日期")(0)), CDate(mArrFilter("审核日期")(1)), _
        CDate(mArrFilter("审核日期")(0)), CDate(mArrFilter("审核日期")(1)), _
          Trim(mArrFilter("申请人")), Trim(mArrFilter("病人姓名")), _
          Val(mArrFilter("住院号")), Val(mArrFilter("病人ID")), Trim(mArrFilter("床号")))
      
    Do While Not rsTemp.EOF
        With mrsVerifyDetail
            .AddNew
    
            !单据 = rsTemp!单据
            !NO = rsTemp!NO
            !材料ID = rsTemp!材料ID
            !申请时间 = Format(rsTemp!申请时间, "yyyy-mm-dd hh:mm:ss")
            !标识号 = rsTemp!标识号
            !姓名 = rsTemp!姓名
            !床号 = rsTemp!床号
            !销帐数量 = rsTemp!销帐数量
            !换算系数 = rsTemp!换算系数
            !单位 = rsTemp!单位
            !开单科室 = rsTemp!开单科室
            .Update
            rsTemp.MoveNext
        End With
    Loop
     
    ''''3、提取批次明细数据
    '单位，包装换算
    Select Case mintUnit
    Case 0
        strFields = "X.计算单位 单位,1 换算系数,C.实际数量 As 准退数量,c.实际数量 * c.入出系数 As 销帐数量"
    Case Else
        strFields = "D.包装单位 单位,D.换算系数  ,C.实际数量 As 准退数量,c.实际数量 * c.入出系数 As 销帐数量"
    End Select
    
'    gstrSQL = "" & _
'        "   Select C.ID As 收发ID, C.药品ID, C.单据, C.NO, C.序号 As 收发序号, C.产地, C.批号, C.效期, C.填制日期 As 开单时间, " & _
'        "           F.险类, P.名称 As 开单科室, A.费用id, B.序号 As 费用序号, B.记录性质, B.主页ID, A.申请时间, " & strFields & " " & _
'        "   From 病人费用销帐 A, 住院费用记录 B,药品收发记录 C, 材料特性 D, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E " & _
'        "   Where A.申请类别=1 And A.费用id = B.ID And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.材料id And B.收费细目id = X.ID  " & _
'        "       And B.病人id = F.病人id(+) And B.主页id = F.主页id(+) And F.出院日期(+) Is Null And F.状态(+) <> 3 And A.申请部门id = E.ID " & _
'        "       And B.执行部门id = [1]  " & _
'        "       And C.审核人 Is Not Null And C.单据 In (24, 25, 26) And (C.记录状态 = 1 Or Mod(C.记录状态, 3) = 0) " & strWere & _
'        "   Order By A.审核时间, C.单据, C.NO, C.序号"

        gstrSQL = "Select C.ID As 收发ID, C.药品ID, C.单据, C.NO, C.序号 As 收发序号, C.产地, C.批号, C.效期, F.险类, P.名称 As 开单科室,C.批次, " & _
            " A.费用id, B.序号 As 费用序号, B.记录性质, B.主页ID, A.申请时间, A.审核时间, C.零售价 As 单价, " & strFields & " " & _
            " From 病人费用销帐 A, 住院费用记录 B,药品收发记录 C, 材料特性 D, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E " & _
            " Where A.申请类别=1 And A.费用id = B.ID And B.No = C.No And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.材料id And B.收费细目id = X.ID And B.病人id = F.病人id(+) And B.主页id = F.主页id(+) And A.申请部门id = E.ID " & _
            " And B.执行部门id = [1]  " & strWere & _
            " And C.审核日期 Is Not Null " & _
            " And ((A.状态 = 1 And Mod(C.记录状态, 3) = 2 And A.审核时间 = C.审核日期) Or (A.状态 = 2 And (C.记录状态 = 1 Or Mod(C.记录状态, 3) = 0))) "
        
        gstrSQL = gstrSQL & " Order By A.审核时间, C.单据, C.NO, C.序号"

     
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取批次明细", _
          Val(mArrFilter("发料部门id")), Val(mArrFilter("申请科室id")), _
          CDate(mArrFilter("审核日期")(0)), CDate(mArrFilter("审核日期")(1)), _
          CDate(mArrFilter("审核日期")(0)), CDate(mArrFilter("审核日期")(1)), _
          Trim(mArrFilter("申请人")), Trim(mArrFilter("病人姓名")), _
          Val(mArrFilter("住院号")), Val(mArrFilter("病人ID")), Trim(mArrFilter("床号")))
    Do While Not rsTemp.EOF
        With mrsVerifyBatch
            .AddNew
            !单据 = rsTemp!单据
            !NO = rsTemp!NO
            !材料ID = rsTemp!药品ID
            !申请时间 = Format(rsTemp!申请时间, "yyyy-mm-dd hh:mm:ss")
            !收发序号 = rsTemp!收发序号
            !产地 = rsTemp!产地
            !批号 = rsTemp!批号
            !效期 = rsTemp!效期
            !准退数量 = rsTemp!准退数量
            !销帐数量 = rsTemp!销帐数量
            !换算系数 = rsTemp!换算系数
            !单位 = rsTemp!单位
            !收发ID = rsTemp!收发ID
            !主页id = Val(NVL(rsTemp!主页id))
            !费用序号 = rsTemp!费用序号
            !险类 = rsTemp!险类
            !费用ID = rsTemp!费用ID
            !记录性质 = rsTemp!记录性质
            !审核标志 = 1
            .Update
            rsTemp.MoveNext
        End With
    Loop
    Call AutoExpendQuantity
    ''''''4、定位到汇总第一行，并提取第一行明细数据
    With vsHeadGrid
        If .Rows > 1 Then
            .Row = 1: .TopRow = 1
            Call LoadDetailList(Val(vsHeadGrid.TextMatrix(1, vsHeadGrid.ColIndex("材料名称"))))
        End If
    End With
    With vsDetail
        '提取第一行批次明细数据
        If .Rows > 1 Then
            Call LoadBatchList(Val(.TextMatrix(1, .ColIndex("单据"))), .TextMatrix(1, .ColIndex("NO")), .RowData(1), .TextMatrix(1, .ColIndex("申请时间")), False)
        End If
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function InitRsStruct() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化内部记录集
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-03 21:08:22
    '-----------------------------------------------------------------------------------------------------------
    Set mrsVerifyDetail = New ADODB.Recordset
    With mrsVerifyDetail
        If .State = 1 Then .Close
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
        .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "标识号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "销帐数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "换算系数", adDouble, 18, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "开单科室", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set mrsVerifyBatch = New ADODB.Recordset
    With mrsVerifyBatch
        If .State = 1 Then .Close
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
        .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "收发序号", adDouble, 18, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "准退数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "销帐数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "换算系数", adDouble, 18, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable
        .Fields.Append "主页ID", adDouble, 18, adFldIsNullable
        .Fields.Append "费用序号", adDouble, 18, adFldIsNullable
        .Fields.Append "险类", adDouble, 18, adFldIsNullable
        .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "审核标志", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Function

Private Sub Form_Load()
    zl_vsGrid_Para_Restore mlngModule, vsHeadGrid, Me.Caption, "销帐已审_Head"
    zl_vsGrid_Para_Restore mlngModule, vsDetail, Me.Caption, "销帐已审_Detail"
    zl_vsGrid_Para_Restore mlngModule, vsBatch, Me.Caption, "销帐已审_Batch"
    Call initPara
End Sub
Private Sub AutoExpendQuantity()
    '考虑到同一费用ID对应多个收发ID的情况，需要将销帐数量分解到多个收发记录上
    '分解的原则是按序号大的优先分配（已按序号降序排序）
    Dim n As Integer
    Dim dbl准退数量 As Double
    Dim dbl剩余数量 As Double
    Dim int收发序号 As Integer
    Dim lng费用ID As Long
    Dim str申请时间 As String
    
    With mrsVerifyBatch
        .Sort = "收发序号 desc"
        If mrsVerifyBatch.RecordCount <> 0 Then .MoveFirst
        For n = 1 To .RecordCount
            dbl准退数量 = !准退数量
            
            If lng费用ID = !费用ID And str申请时间 = !申请时间 Then

            Else
                dbl剩余数量 = !销帐数量
            End If
            
            If dbl剩余数量 >= dbl准退数量 Then
                dbl剩余数量 = dbl剩余数量 - dbl准退数量
                !销帐数量 = dbl准退数量
            Else
                !销帐数量 = dbl剩余数量
                dbl剩余数量 = 0
            End If
            
            lng费用ID = !费用ID
            str申请时间 = !申请时间
            
            .Update
            .MoveNext
        Next
    End With
End Sub
Private Sub LoadDetailList(ByVal lng材料ID As Long)
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载明细数据
    '入参:lng材料ID-材料ID
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-03 22:25:04
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    With vsDetail
        .Clear 1
        .Rows = 2
        mrsVerifyDetail.Filter = "材料ID=" & lng材料ID
        If mrsVerifyDetail.RecordCount = 0 Then Exit Sub
        .Rows = mrsVerifyDetail.RecordCount + 1
        lngRow = 0
        Do While Not mrsVerifyDetail.EOF
            lngRow = lngRow + 1
            .RowData(lngRow) = Val(NVL(mrsVerifyDetail!材料ID))
            .TextMatrix(lngRow, .ColIndex("单据")) = NVL(mrsVerifyDetail!单据)
            .TextMatrix(lngRow, .ColIndex("NO")) = NVL(mrsVerifyDetail!NO)
            .TextMatrix(lngRow, .ColIndex("申请时间")) = Format(mrsVerifyDetail!申请时间, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(lngRow, .ColIndex("门诊(住院)号")) = NVL(mrsVerifyDetail!标识号)
            .TextMatrix(lngRow, .ColIndex("姓名")) = NVL(mrsVerifyDetail!姓名)
            .TextMatrix(lngRow, .ColIndex("床号")) = NVL(mrsVerifyDetail!床号)
            .TextMatrix(lngRow, .ColIndex("销帐数量")) = Format(Val(NVL(mrsVerifyDetail!销帐数量)) / mrsVerifyDetail!换算系数, mFMT.FM_数量)
            .TextMatrix(lngRow, .ColIndex("单位")) = NVL(mrsVerifyDetail!单位)
            .TextMatrix(lngRow, .ColIndex("开单科室")) = NVL(mrsVerifyDetail!开单科室)
            
            .Cell(flexcpData, lngRow, .ColIndex("销帐数量")) = Val(NVL(mrsVerifyDetail!销帐数量))
            .Cell(flexcpData, lngRow, .ColIndex("NO")) = Val(NVL(mrsVerifyDetail!换算系数))
            mrsVerifyDetail.MoveNext
        Loop
        .Cell(flexcpForeColor, 1, .ColIndex("销帐数量"), .Rows - 1, .ColIndex("销帐数量")) = vbBlue
    End With
End Sub
Private Sub LoadBatchList(ByVal int单据 As Integer, _
                ByVal strNo As String, ByVal lng材料ID As Long, _
                ByVal str申请时间 As String, ByVal bln更新标志 As Boolean)
                
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载批次信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-03 22:42:02
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
     With vsBatch
        .Clear 1
        .Rows = 2
        If mrsVerifyBatch Is Nothing Then
            Exit Sub
        End If
        
        mrsVerifyBatch.Filter = "单据=" & int单据 & _
                " And No='" & strNo & "' " & _
                " And 材料ID=" & lng材料ID & _
                " And 申请时间='" & str申请时间 & "' "
        mrsVerifyBatch.Sort = "收发序号 Desc"
        If mrsVerifyBatch.RecordCount = 0 Then Exit Sub
        
        If mrsVerifyBatch.RecordCount = 0 Then
            vsDetail.Height = Me.ScaleHeight - vsDetail.Top
            picBatHsc.Visible = False
        Else
            picBatHsc.Top = Me.ScaleHeight - 1935
            vsDetail.Height = picBatHsc.Top - vsDetail.Top
            vsBatch.Top = picBatHsc.Top + picBatHsc.Height
            vsBatch.Height = Me.ScaleHeight - Me.vsBatch.Top
            picBatHsc.Visible = True
        End If
        
        .Rows = mrsVerifyBatch.RecordCount + 1
        mrsVerifyBatch.MoveFirst
        lngRow = 0
        Do While Not mrsVerifyBatch.EOF
            lngRow = lngRow + 1
            .RowData(lngRow) = Val(NVL(mrsVerifyBatch!材料ID))
            vsBatch.TextMatrix(lngRow, .ColIndex("单据")) = NVL(mrsVerifyBatch!单据)
            vsBatch.TextMatrix(lngRow, .ColIndex("NO")) = NVL(mrsVerifyBatch!NO)
            vsBatch.TextMatrix(lngRow, .ColIndex("申请时间")) = Format(mrsVerifyBatch!申请时间, "yyyy-mm-dd hh:mm:ss")
            vsBatch.TextMatrix(lngRow, .ColIndex("产地")) = NVL(mrsVerifyBatch!产地)
            vsBatch.TextMatrix(lngRow, .ColIndex("批号")) = NVL(mrsVerifyBatch!批号)
            vsBatch.TextMatrix(lngRow, .ColIndex("效期")) = Format(mrsVerifyBatch!效期, "yyyy-mm-dd")
            vsBatch.TextMatrix(lngRow, .ColIndex("准退数量")) = Format(Val(NVL(mrsVerifyBatch!准退数量)) / mrsVerifyBatch!换算系数, mFMT.FM_数量)
            vsBatch.TextMatrix(lngRow, .ColIndex("销帐数量")) = Abs(Format(Val(NVL(mrsVerifyBatch!销帐数量)) / mrsVerifyBatch!换算系数, mFMT.FM_数量))
            vsBatch.TextMatrix(lngRow, .ColIndex("单位")) = NVL(mrsVerifyBatch!单位)
            
            vsBatch.Cell(flexcpData, lngRow, .ColIndex("准退数量")) = Val(NVL(mrsVerifyBatch!准退数量))
            vsBatch.Cell(flexcpData, lngRow, .ColIndex("销帐数量")) = Abs(Val(NVL(mrsVerifyBatch!销帐数量)))
            vsBatch.Cell(flexcpData, lngRow, .ColIndex("NO")) = Val(NVL(mrsVerifyBatch!换算系数))
            mrsVerifyBatch.MoveNext
        Loop
        .Cell(flexcpForeColor, 1, .ColIndex("销帐数量"), lngRow, .ColIndex("销帐数量")) = vbBlue
    End With
End Sub

Private Sub Form_Resize()
    With vsHeadGrid
        .Left = ScaleLeft
        .Width = ScaleWidth - .Left
        .Height = IIf(picHsc.Top - Top < 500, 500, picHsc.Top - Top)
        picHsc.Top = .Top + .Height
    End With
    With vsDetail
        .Top = picHsc.Top + picHsc.Height
        .Width = vsHeadGrid.Width
        If picBatHsc.Visible Then
            .Height = IIf(picBatHsc.Top - .Top < 500, 500, picBatHsc.Top - .Top)
            picBatHsc.Top = .Top + .Height
        Else
            .Height = Me.ScaleHeight - .Top
        End If
    End With
    
    If picBatHsc.Visible = True Then
        With vsBatch
            .Top = picBatHsc.Top + picBatHsc.Height
            .Width = vsHeadGrid.Width
            .Height = IIf(Me.ScaleHeight - .Top < 0, 0, Me.ScaleHeight - .Top)
            .Left = ScaleLeft
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsHeadGrid, Me.Caption, "销帐已审_Head"
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "销帐已审_Detail"
    zl_vsGrid_Para_Save mlngModule, vsBatch, Me.Caption, "销帐已审_Batch"
End Sub

Private Sub vsBatch_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBatch
        Select Case .Col
        Case .ColIndex("销帐数量")
            If .TextMatrix(Row, .ColIndex("单据")) = "" Then Cancel = True
        Case Else
            Cancel = True
        End Select
    End With
End Sub
  

Private Sub vsDetail_EnterCell()
    Dim lng材料ID As Long
    With vsDetail
        If .Row > 0 Then
            lng材料ID = IIf(IsNull(.RowData(.Row)), 0, .RowData(.Row))
            '提取批次明细数据
            Call LoadBatchList(Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), lng材料ID, .TextMatrix(.Row, .ColIndex("申请时间")), False)
        End If
    End With
End Sub

 
Private Sub vsHeadGrid_Click()
 
    With vsHeadGrid
        If .Row > 0 Then
            If .Cell(flexcpData, .Row, .ColIndex("材料名称")) = 0 Then Exit Sub
        End If
        If .Row > 0 Then
            '提取明细数据
            Call LoadDetailList(Val(.Cell(flexcpData, .Row, .ColIndex("材料名称"))))
        End If
    End With

    '提取批次明细数据
    With vsDetail
        Call LoadBatchList(Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), .RowData(.Row), .TextMatrix(.Row, .ColIndex("申请时间")), False)
    End With
End Sub


Private Sub vsBatch_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '只能输入数字
    With vsBatch
        If Col = .ColIndex("销帐数量") Then
            If InStr("1234567890" + Chr(46) + Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End With
End Sub

 


Public Function zlRefreshData(ByVal frmMain As Form, ByVal strPrivs As String, ByVal lngModule As Long, ByVal intUnit As Integer, _
    ByVal arrFilter As Variant) As Boolean
     '-----------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '入参:frmMain-父窗口
    '     strPrivs-权限串
    '     lngModule-模块号
    '     intUnit-显示单位(0-散装单位,1-包装单位)
    '     arrFilter-条件过滤
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-22 14:25:18
    '-----------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mstrPrivs = strPrivs: mlngModule = lngModule:
    Set mArrFilter = arrFilter
    mintUnit = intUnit
    '初始化值
    Call Form_Load
    zlRefreshData = RefreshData
End Function
Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsHeadGrid.Height + Y <= 500 Or vsDetail.Height - Y <= 500 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        
        vsHeadGrid.Height = vsHeadGrid.Height + Y
        If picBatHsc.Visible Then
            vsDetail.Height = vsDetail.Height - Y
            vsDetail.Top = vsDetail.Top + Y
        Else
            vsDetail.Top = vsDetail.Top + Y
            vsDetail.Height = Me.ScaleHeight - vsDetail.Top
        End If
        Me.Refresh
    End If
End Sub
Private Sub picBatHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsDetail.Height + Y <= 500 Then Exit Sub
        picBatHsc.Top = picBatHsc.Top + Y
        If Me.ScaleHeight - picBatHsc.Top < 500 Then picBatHsc.Top = Me.ScaleHeight - 500
        vsDetail.Height = picBatHsc.Top - vsDetail.Top
        vsBatch.Top = picBatHsc.Top + picBatHsc.Height
        vsBatch.Height = Me.ScaleHeight - vsBatch.Top
            
        Me.Refresh
    End If
End Sub



