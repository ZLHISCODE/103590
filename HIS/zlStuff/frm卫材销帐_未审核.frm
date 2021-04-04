VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frm卫材销帐_未审核 
   BorderStyle     =   0  'None
   Caption         =   "卫材销帐未审核"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10965
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2490
      Width           =   12615
   End
   Begin VB.PictureBox picBatHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   -330
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   12735
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4470
      Width           =   12735
   End
   Begin VB.CheckBox chk全选 
      Caption         =   "全审(&A)"
      Height          =   225
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1500
   End
   Begin VSFlex8Ctl.VSFlexGrid vsHeadGrid 
      Height          =   2055
      Left            =   75
      TabIndex        =   1
      Tag             =   "待处理"
      Top             =   435
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm卫材销帐_未审核.frx":0000
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
      TabIndex        =   2
      Tag             =   "明细"
      Top             =   2535
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm卫材销帐_未审核.frx":00DD
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
      TabIndex        =   3
      Tag             =   "明细"
      Top             =   4575
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm卫材销帐_未审核.frx":0242
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
Attribute VB_Name = "frm卫材销帐_未审核"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mstrPrivs As String
Private mlngModule As Long
Private mArrFilter As Variant   '过滤条件
Private mrsDetail As ADODB.Recordset
Private mrsBatch As ADODB.Recordset
Private mint审核标志 As Integer
Private mintUnit As Integer '0-散装单位,1-包装单位
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
'医保接口
Private gclsInsure As New clsInsure
Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    记帐作废上传 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
    support预交退个人帐户 = 2
    
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤消出院 = 17            '允许撤消病人出院
    support必须录入入出诊断 = 18    '病人入院与出院时，必须录入诊断名
    support记帐完成后上传 = 19      '要求上传在记帐数据提交后再进行
    support出院结算必须出院 = 20    '病人结帐时如果选择出院结帐，就检查必须出院才可以进行
    
    support挂号使用个人帐户 = 21    '使用医保挂号时是否使用个人帐户进行支付

    support门诊连续收费 = 22        '门诊在身份验证后，可进行多次收费操作
    support门诊收费完成后验证 = 23  '在门诊收费完成，是否再次调用身份验证
    
    support医嘱上传 = 24            '医嘱产生费用时是否实时传输
    support分币处理 = 25            '医保病人是否处理分币
    support中途结算仅处理已上传部分 = 26 '提供对已上传部分数据的结算功能
    support允许冲销已结帐的记帐单据 = 27 '是否允许冲销记帐单据，如果该单据已经结帐
    
    support允许部份冲销单据 = 28
    support出院无实际交易 = 29       '出院接口中是否要与接口商进行交易
    support允许部分冲销明细 = 32    '允许针对住院记帐处方的每笔明细进行部分冲销
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
    support住院结算作废 = 34        'HIS始终认为住院支持结算作废，如果不支持需医保接口内部处理，返回假即可；增加该参数是为了配合GetCapability交易来检查各种结算方式是否支持全退
    support负数记帐 = 35            '是否允许负数记帐，操作员首先要拥有负数记帐的权限。此参数缺省为真，不支持的接口需单独处理
    support结帐_指定住院次数 = 36   '是否支持指定住院次数进行医保结算
    support结帐_指定日期范围 = 37   '是否支持指定结帐日期范围进行医保结算
    support结帐_设置婴儿费条件 = 38 '是否允许设置婴儿费条件
    
    support门诊结帐 = 41            '是否支持门诊医保病人的记帐费用使用门诊结帐来完成
End Enum

Private mobjPlugIn As Object             '外挂接口对象

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
 
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
      '  .Editable = flexEDKbdMouse
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
    Dim strNOS As String
    
    On Error GoTo ErrHandle
    Call InitRsStruct
    vsHeadGrid.Rows = 1
    mint审核标志 = 1
        
    ''''1、提取汇总数据
    '单位，包装换算
    Select Case mintUnit
    Case 0
        strFields = "X.计算单位 单位,1 换算系数,A.数量 As 销帐数量 "
    Case Else
        strFields = "D.包装单位 单位,d.换算系数 换算系数,A.数量 As 销帐数量 "
    End Select
    If CDate(mArrFilter("日期范围")(0)) <= CDate("1949-02-01") Then
        strWere = strWere & " And A.审核人 Is Null And A.状态 = 0  "
    Else
        strWere = strWere & " And A.审核人 Is Null And A.状态 = 0 And A.申请时间 Between [3] And [4] "
    End If
    
    '病区/医技科室
    If Val(mArrFilter("申请科室ID")) > 0 Then strWere = strWere & " And A.申请部门id = [2] "
    '申请人
    If Trim(mArrFilter("申请人")) <> "" Then strWere = strWere & " And A.申请人=[7] "
    '病人姓名
    If Trim(mArrFilter("病人姓名")) <> "" Then strWere = strWere & " And nvl(F.姓名,B.姓名)=[8] "
    strWere = strWere & IIf(Val(mArrFilter("住院号")) = 0, "", "             AND b.标识号=[9] and b.门诊标志=2 ")
    strWere = strWere & IIf(Val(mArrFilter("病人ID")) = 0, "", "             AND b.病人iD=[10]  ")
    strWere = strWere & IIf(Trim(mArrFilter("床号")) = "", "", "             AND b.床号=[11]  ")

    '说明:
    '1.不包含出院病人: F.状态:0-正常住院；1-尚未入科；2-正在转科；3-已预出院
    '2.需要过滤掉未发料部份的申请:原则上在记帐的销帐审请时,是不会出院对未发部分的申请.
    gstrSQL = "" & _
    "   Select Distinct A.收费细目id,'['||X.编码||']'||X.名称 as 材料名称, X.规格,A.标准单价 as 零售价, " & strFields & _
    "   From (  Select A.收费细目id, Sum(A.数量) As 数量,B.标准单价  " & _
    "           From 住院费用记录 B, 病案主页 F, 病人费用销帐 A " & _
    "           Where A.申请类别=1 And A.费用id = B.ID And A.审核部门id = [1] And B.病人id = F.病人id  " & _
    "                 And B.主页id = F.主页id  And F.出院日期  Is Null And F.状态  <> 3 " & _
    "                 " & vbCrLf & strWere & _
    "                 And Exists (Select 1 From 药品收发记录 C  Where C.费用id = A.费用id And C.审核人 Is Not Null And (C.记录状态 = 1 Or Mod(C.记录状态, 3) = 0))" & _
    "           Group By A.收费细目id,b.标准单价) A,材料特性 D, 收费项目别名 E, 收费项目目录 X " & _
    " Where A.收费细目id = D.材料id And A.收费细目id = X.ID And X.ID = E.收费细目id(+) And E.性质(+) = 3 " & _
    " Order By 材料名称"
    
    '[5],[6]现在取消
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取退料申请", _
        Val(mArrFilter("发料部门id")), Val(mArrFilter("申请科室id")), _
        CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), _
        CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), _
        Trim(mArrFilter("申请人")), Trim(mArrFilter("病人姓名")), _
        Val(mArrFilter("住院号")), Val(mArrFilter("病人ID")), Trim(mArrFilter("床号")))
    
    With vsHeadGrid
        .Clear 1
        .Rows = 2
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 0
        Do While Not rsTemp.EOF
            lngRow = lngRow + 1
            .TextMatrix(lngRow, .ColIndex("审核")) = "√"
            .TextMatrix(lngRow, .ColIndex("材料名称")) = NVL(rsTemp!材料名称)
            .TextMatrix(lngRow, .ColIndex("规格")) = NVL(rsTemp!规格)
            .TextMatrix(lngRow, .ColIndex("销帐数量")) = Format(Val(NVL(rsTemp!销帐数量)) / rsTemp!换算系数, mFMT.FM_数量)
            .TextMatrix(lngRow, .ColIndex("销帐金额")) = Format(Val(NVL(rsTemp!销帐数量)) * rsTemp!零售价, mFMT.FM_金额)
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
        "   Select 单据, NO, 药品ID as 材料ID, 申请时间, 标识号, 姓名, 床号, 单位, 换算系数,开单科室, Sum(数量) As 销帐数量,零售价 " & _
        "   From (  Select Distinct C.单据, C.NO, C.药品ID, A.申请时间, B.标识号,B.标准单价 as 零售价, nvl(F.姓名,B.姓名) 姓名, B.床号,P.名称 开单科室, " & strFields & " " & _
        "           From 病人费用销帐 A, 住院费用记录 B,药品收发记录 C, 材料特性 D, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E " & _
        "           Where   A.申请类别=1 And A.费用id = B.ID And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.材料id And B.收费细目id = X.ID  " & _
        "                   And  B.病人id = F.病人id And B.主页id = F.主页id And F.出院日期 Is Null And F.状态 <> 3 " & _
        "                   And A.申请部门id = E.ID And B.执行部门id = [1]  " & _
        "                   And C.审核人 Is Not Null And C.单据 In (24, 25, 26) And (C.记录状态 = 1 Or Mod(C.记录状态, 3) = 0) " & strWere & ")" & _
        "           Group By 单据, NO, 药品ID, 申请时间, 标识号, 姓名, 床号, 单位, 换算系数,零售价,开单科室 "
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取单据明细", _
          Val(mArrFilter("发料部门id")), Val(mArrFilter("申请科室id")), _
          CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), _
          CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), _
          Trim(mArrFilter("申请人")), Trim(mArrFilter("病人姓名")), _
          Val(mArrFilter("住院号")), Val(mArrFilter("病人ID")), Trim(mArrFilter("床号")))
      
    Do While Not rsTemp.EOF
        With mrsDetail
            .AddNew
    
            !单据 = rsTemp!单据
            !NO = rsTemp!NO
            !材料ID = rsTemp!材料ID
            !申请时间 = Format(rsTemp!申请时间, "yyyy-mm-dd hh:mm:ss")
            !标识号 = rsTemp!标识号
            !姓名 = rsTemp!姓名
            !床号 = rsTemp!床号
            !销帐数量 = rsTemp!销帐数量
            !销帐金额 = rsTemp!销帐数量 * rsTemp!零售价
            !换算系数 = rsTemp!换算系数
            !单位 = rsTemp!单位
            !开单科室 = rsTemp!开单科室
            .Update
            
            If InStr(1, strNOS, rsTemp!NO) = 0 Then
                strNOS = IIf(strNOS = "", "", strNOS & ",") & rsTemp!NO
            End If
            rsTemp.MoveNext
        End With
    Loop
     
    ''''3、提取批次明细数据
    '单位，包装换算
    Select Case mintUnit
    Case 0
        strFields = "X.计算单位 单位,1 换算系数,C.实际数量 As 准退数量,A.数量 As 销帐数量"
    Case Else
        strFields = "D.包装单位 单位,D.换算系数 ,C.实际数量 As 准退数量,A.数量 As 销帐数量"
    End Select
    
    ' 'Having Sum(实际数量) > 0
    gstrSQL = "Select /*+ Rule*/ C.ID As 收发ID, C.药品ID, C.单据, C.NO, C.序号 As 收发序号, C.产地, C.批号, C.效期, F.险类, P.名称 As 开单科室,B.标准单价 as 零售价, " & _
        " A.费用id, B.序号 As 费用序号, B.记录性质, B.主页ID, A.申请时间, C.零售价 As 单价, " & strFields & " " & _
        " From 病人费用销帐 A, 住院费用记录 B, " & _
        " (Select A.ID, A.单据, A.NO, A.序号, A.药品id, A.产地, A.批号, A.效期, A.费用id, B.实际数量, A.零售价 " & _
        " From 药品收发记录 A, " & _
        " (Select a.单据, a.NO, a.序号, a.药品id, Sum(Nvl(a.付数, 1) * a.实际数量) As 实际数量 " & _
        " From 药品收发记录 a ,Table(Cast(f_Str2list([12]) As zlTools.t_Strlist)) b " & _
        " Where a.单据 In (24, 25,26) And a.审核日期 Is Not Null And a.No=b.Column_Value "
        
    gstrSQL = gstrSQL & " Group By 单据, NO, 序号, 药品id " & _
        " ) B" & _
        " Where A.NO = B.NO And A.单据 = B.单据 And A.药品id + 0 = B.药品id And A.序号 = B.序号 And A.审核人 Is Not Null " & _
        " And (A.记录状态 = 1 Or Mod(A.记录状态, 3) = 0))C, " & _
        " 材料特性 D, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E " & _
        " Where A.申请类别=1 And A.费用id = B.ID And B.No = C.No And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.材料id And B.收费细目id = X.ID And B.病人id = F.病人id And B.主页id = F.主页id And A.申请部门id = E.ID " & _
        " And B.执行部门id = [1] " & strWere

    gstrSQL = gstrSQL & " Order By A.申请时间, C.单据, C.NO, C.序号 Desc "
    
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取批次明细", _
          Val(mArrFilter("发料部门id")), Val(mArrFilter("申请科室id")), _
          CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), _
          CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), _
          Trim(mArrFilter("申请人")), Trim(mArrFilter("病人姓名")), _
          Val(mArrFilter("住院号")), Val(mArrFilter("病人ID")), Trim(mArrFilter("床号")), strNOS)
          
    Do While Not rsTemp.EOF
        With mrsBatch
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
            !销帐金额 = rsTemp!销帐数量 * rsTemp!零售价
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
    Set mrsDetail = New ADODB.Recordset
    With mrsDetail
        If .State = 1 Then .Close
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
        .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "标识号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "销帐数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "销帐金额", adDouble, 18, adFldIsNullable
        .Fields.Append "换算系数", adDouble, 18, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "开单科室", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set mrsBatch = New ADODB.Recordset
    With mrsBatch
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
        .Fields.Append "销帐金额", adDouble, 18, adFldIsNullable
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

Private Sub chk全选_Click()
    Dim n As Integer
    With vsHeadGrid
        If .Rows > 1 Then
            If Val(.Cell(flexcpData, 1, .ColIndex("材料名称"))) = 0 Then Exit Sub
        End If
        For n = 1 To .Rows - 1
            If Val(.Cell(flexcpData, n, .ColIndex("材料名称"))) <> 0 Then
                .TextMatrix(n, .ColIndex("审核")) = IIf(chk全选.Value = 1, "√", "")
            End If
        Next
    End With
    With mrsBatch
        .Filter = 0
        If mrsBatch.RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            !审核标志 = IIf(chk全选.Value = 1, 1, 0)
            .Update
            .MoveNext
        Loop
    End With
End Sub
Private Sub Form_Load()
    zl_vsGrid_Para_Restore mlngModule, vsHeadGrid, Me.Caption, "销帐未审_Head"
    zl_vsGrid_Para_Restore mlngModule, vsDetail, Me.Caption, "销帐未审_Detail"
    zl_vsGrid_Para_Restore mlngModule, vsBatch, Me.Caption, "销帐未审_Batch"
    
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
    
    With mrsBatch
        If mrsBatch.RecordCount <> 0 Then .MoveFirst
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
        mrsDetail.Filter = "材料ID=" & lng材料ID
        If mrsDetail.RecordCount = 0 Then Exit Sub
        .Rows = mrsDetail.RecordCount + 1
        lngRow = 0
        Do While Not mrsDetail.EOF
            lngRow = lngRow + 1
            .RowData(lngRow) = Val(NVL(mrsDetail!材料ID))
            .TextMatrix(lngRow, .ColIndex("单据")) = NVL(mrsDetail!单据)
            .TextMatrix(lngRow, .ColIndex("NO")) = NVL(mrsDetail!NO)
            .TextMatrix(lngRow, .ColIndex("申请时间")) = Format(mrsDetail!申请时间, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(lngRow, .ColIndex("门诊(住院)号")) = NVL(mrsDetail!标识号)
            .TextMatrix(lngRow, .ColIndex("姓名")) = NVL(mrsDetail!姓名)
            .TextMatrix(lngRow, .ColIndex("床号")) = NVL(mrsDetail!床号)
            .TextMatrix(lngRow, .ColIndex("销帐数量")) = Format(Val(NVL(mrsDetail!销帐数量)) / mrsDetail!换算系数, mFMT.FM_数量)
            .TextMatrix(lngRow, .ColIndex("销帐金额")) = Format(Val(NVL(mrsDetail!销帐金额)), mFMT.FM_金额)
            .TextMatrix(lngRow, .ColIndex("单位")) = NVL(mrsDetail!单位)
            .TextMatrix(lngRow, .ColIndex("开单科室")) = NVL(mrsDetail!开单科室)
            
            .Cell(flexcpData, lngRow, .ColIndex("销帐数量")) = Val(NVL(mrsDetail!销帐数量))
            .Cell(flexcpData, lngRow, .ColIndex("NO")) = Val(NVL(mrsDetail!换算系数))
            mrsDetail.MoveNext
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
        mrsBatch.Filter = "单据=" & int单据 & _
                " And No='" & strNo & "' " & _
                " And 材料ID=" & lng材料ID & _
                " And 申请时间='" & str申请时间 & "' "
        mrsBatch.Sort = "收发序号 Desc"
        .Clear 1
        .Rows = 2
        If mrsBatch.RecordCount = 0 Then Exit Sub
        
        If mrsBatch.RecordCount = 0 Then
            vsDetail.Height = Me.ScaleHeight - vsDetail.Top
            picBatHsc.Visible = False
        Else
            picBatHsc.Top = Me.ScaleHeight - 1935
            vsDetail.Height = picBatHsc.Top - vsDetail.Top
            vsBatch.Top = picBatHsc.Top + picBatHsc.Height
            vsBatch.Height = Me.ScaleHeight - Me.vsBatch.Top
            picBatHsc.Visible = True
        End If
            
        .Rows = mrsBatch.RecordCount + 1
        mrsBatch.MoveFirst
        lngRow = 0
        Do While Not mrsBatch.EOF
                lngRow = lngRow + 1
                .RowData(lngRow) = Val(NVL(mrsBatch!材料ID))
                vsBatch.TextMatrix(lngRow, .ColIndex("单据")) = NVL(mrsBatch!单据)
                vsBatch.TextMatrix(lngRow, .ColIndex("NO")) = NVL(mrsBatch!NO)
                vsBatch.TextMatrix(lngRow, .ColIndex("申请时间")) = Format(mrsBatch!申请时间, "yyyy-mm-dd hh:mm:ss")
                vsBatch.TextMatrix(lngRow, .ColIndex("产地")) = NVL(mrsBatch!产地)
                vsBatch.TextMatrix(lngRow, .ColIndex("批号")) = NVL(mrsBatch!批号)
                vsBatch.TextMatrix(lngRow, .ColIndex("效期")) = Format(mrsBatch!效期, "yyyy-mm-dd")
                vsBatch.TextMatrix(lngRow, .ColIndex("准退数量")) = Format(Val(NVL(mrsBatch!准退数量)) / mrsBatch!换算系数, mFMT.FM_数量)
                vsBatch.TextMatrix(lngRow, .ColIndex("销帐数量")) = Format(Val(NVL(mrsBatch!销帐数量)) / mrsBatch!换算系数, mFMT.FM_数量)
                vsBatch.TextMatrix(lngRow, .ColIndex("销帐金额")) = Format(Val(NVL(mrsBatch!销帐金额)), mFMT.FM_金额)
                vsBatch.TextMatrix(lngRow, .ColIndex("单位")) = NVL(mrsBatch!单位)
                
                vsBatch.Cell(flexcpData, lngRow, .ColIndex("准退数量")) = Val(NVL(mrsBatch!准退数量))
                vsBatch.Cell(flexcpData, lngRow, .ColIndex("销帐数量")) = Val(NVL(mrsBatch!销帐数量))
                vsBatch.Cell(flexcpData, lngRow, .ColIndex("NO")) = Val(NVL(mrsBatch!换算系数))
                If bln更新标志 Then
                    mrsBatch!审核标志 = mint审核标志
                    mrsBatch.Update
                End If
            mrsBatch.MoveNext
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
    zl_vsGrid_Para_Save mlngModule, vsHeadGrid, Me.Caption, "销帐未审_Head"
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "销帐未审_Detail"
    zl_vsGrid_Para_Save mlngModule, vsBatch, Me.Caption, "销帐未审_Batch"
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
    Dim bln更新标志 As Boolean
    
    With vsHeadGrid
        If .Row > 0 Then
            If .Cell(flexcpData, .Row, .ColIndex("材料名称")) = 0 Then Exit Sub
        End If
        
        
        If .Row > 0 And .Col = .ColIndex("审核") Then
            If .TextMatrix(.Row, .Col) = "√" Then
                .TextMatrix(.Row, .Col) = "×"
                mint审核标志 = 2
            ElseIf .TextMatrix(.Row, .Col) = "×" Then
                .TextMatrix(.Row, .Col) = ""
                mint审核标志 = 0
            Else
                .TextMatrix(.Row, .Col) = "√"
                mint审核标志 = 1
            End If
            bln更新标志 = True
        End If
        
        If .Row > 0 Then
            '提取明细数据
            Call LoadDetailList(Val(.Cell(flexcpData, .Row, .ColIndex("材料名称"))))
        End If
    End With
    
    '提取批次明细数据
    With vsDetail
        Call LoadBatchList(Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), .RowData(.Row), .TextMatrix(.Row, .ColIndex("申请时间")), bln更新标志)
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

Private Sub vsBatch_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dblKey As Double
    With vsBatch
        dblKey = Val(.EditText)

        If dblKey > Val(.TextMatrix(Row, .ColIndex("准退数量"))) Or dblKey < 0 Then
            dblKey = Val(.TextMatrix(Row, .ColIndex("准退数量")))
        End If
        .EditText = Format(dblKey, mFMT.FM_数量)
        .TextMatrix(Row, .ColIndex("销帐数量")) = Format(dblKey, mFMT.FM_数量)

        mrsBatch.Filter = "单据=" & Val(.TextMatrix(Row, .ColIndex("单据"))) & _
                        " And No='" & .TextMatrix(Row, .ColIndex("NO")) & "' " & _
                        " And 材料ID=" & .RowData(Row) & _
                        " And 收发序号=" & Val(.TextMatrix(Row, .ColIndex("收发序号"))) & _
                        " And 申请时间='" & Val(.TextMatrix(Row, .ColIndex("申请时间"))) & "' "
        If mrsBatch.EOF Then Exit Sub
        mrsBatch!销帐数量 = Val(.TextMatrix(Row, .ColIndex("销帐数量"))) * mrsBatch!换算系数
        mrsBatch.Update
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
Public Function zlVerifyData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:退料销帐
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-04 00:23:54
    '-----------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    zlVerifyData = SaveData()
    Screen.MousePointer = 0
End Function
Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:审核
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-04 00:08:26
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strCurDate As String
    
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim int审核标志 As Integer
    Dim bln是否有退料 As Boolean
    Dim str序号数量 As String
    Dim cllPro As Collection
    Dim strAudit As String  '记录已进行销帐的费用记录，避免重复执行
    Dim strReturnInfo As String
    Dim strReserve As String
    
    If vsHeadGrid.Rows = 1 Then Exit Function
    If Val(vsHeadGrid.Cell(flexcpData, 1, vsHeadGrid.ColIndex("材料名称"))) = 0 Then Exit Function
    strCurDate = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    Set cllPro = New Collection
    
    With mrsBatch
        .Filter = 0
        If .State = 0 Then Exit Function
        If .RecordCount = 0 Then Exit Function
        Do While Not .EOF
            If !审核标志 <> 0 And InStr("," & strAudit & ",", "," & !费用ID & !申请时间 & ",") = 0 Then
                strAudit = IIf(strAudit = "", !费用ID & !申请时间, strAudit & "," & !费用ID & !申请时间)
                
                'Zl_病人费用销帐_Audit
                gstrSQL = "Zl_病人费用销帐_Audit("
                '  Id_In       病人费用销帐.费用id%Type,
                gstrSQL = gstrSQL & "" & Val(NVL(!费用ID)) & ","
                '  申请时间_In 病人费用销帐.申请时间%Type,
                gstrSQL = gstrSQL & "To_Date('" & !申请时间 & "','YYYY-MM-DD HH24:MI:SS'),"
                '  审核人_In   病人费用销帐.审核人%Type,
                gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                '  审核时间_In 病人费用销帐.审核时间%Type,
                gstrSQL = gstrSQL & "To_Date('" & strCurDate & "','YYYY-MM-DD HH24:MI:SS'),"
                '  状态_In     病人费用销帐.状态%Type,
                gstrSQL = gstrSQL & "" & Val(NVL(!审核标志)) & ","
                '  int自动退料 Integer:=1
                gstrSQL = gstrSQL & "0)"
                AddArray cllPro, gstrSQL
            End If
            
            '退料处理
            If !审核标志 = 1 And !销帐数量 <> 0 Then
                    'Zl_材料收发记录_部门退料
                    gstrSQL = "Zl_材料收发记录_部门退料("
                    '    收发id_In   In 药品收发记录.ID%Type,
                    gstrSQL = gstrSQL & "" & NVL(!收发ID) & ","
                    '    审核人_In   In 药品收发记录.审核人%Type,
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    '    审核日期_In In 药品收发记录.审核日期%Type,
                    gstrSQL = gstrSQL & "to_date('" & strCurDate & "','yyyy-mm-dd HH24:mi:ss'),"
                    '    批号_In     In 药品库存.上次批号%Type := Null,
                    gstrSQL = gstrSQL & "'" & NVL(!批号) & "',"
                    '    效期_In     In 药品库存.效期%Type := Null,
                    gstrSQL = gstrSQL & "" & IIf(IsNull(!效期), "NULL", IIf(NVL(!效期) = "", "NULL", "To_Date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & ","
                    '    产地_In     In 药品库存.上次产地%Type := Null,
                    gstrSQL = gstrSQL & "'" & NVL(!产地) & "',"
                    '    退料数量_In In 药品收发记录.实际数量%Type := Null,
                    gstrSQL = gstrSQL & "" & NVL(!销帐数量) & ","
                    '    自动销帐_In Integer := 0,
                    gstrSQL = gstrSQL & "" & 0 & ","
                    '    退料人_In   In 药品收发记录.领用人%Type := Null
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    
                    '    是否销帐_In Integer := 1,
                    gstrSQL = gstrSQL & "" & 0 & ")"
                    AddArray cllPro, gstrSQL
                    bln是否有退料 = True
                    '销帐处理
                    str序号数量 = !费用序号 & ":" & !销帐数量
                    '--序号：格式如"1,3,5,7,8",或"1:2,3:2,5:2,7:2,8:2",冒号前面的数字表示行号,后面的数字表示退的数量,目前仅在销帐审核时非药品才传入
                    '--      为空表示冲销所有可冲销行

                    If !主页id = 0 Then
                        gstrSQL = "Zl_门诊记帐记录_Delete('" & !NO & "','" & !费用序号 & "','" & gstrUserCode & "','" & gstrUserName & "')"
                    Else
                        gstrSQL = "ZL_住院记帐记录_Delete('" & !NO & "','" & str序号数量 & "','" & gstrUserCode & "','" & gstrUserName & "'," & !记录性质 & ",1)"
                    End If
                    AddArray cllPro, gstrSQL
                    
                    '医保处理
                    If Not IsNull(!险类) And InStr(1, strMCNO, !NO) = 0 Then
                        MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(!险类))
                        MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(!险类))
                        strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !险类 & _
                                "," & IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
                    End If
                    
                    strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & NVL(!收发ID) & "," & NVL(!销帐数量)
            End If
            .MoveNext
        Loop
    End With
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllPro, Me.Caption, True
    '医保，记帐作废上传，作废时上传
    If strMCNO <> "" Then
        arrMCRec = Split(strMCNO, "|")
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    gcnOracle.RollbackTrans:  Exit Function
                End If
            End If
        Next
    End If
                            
    gcnOracle.CommitTrans
    
    '医保，记帐作废上传，完成后上传
    If strMCNO <> "" Then
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    MsgBox "单据""" & CStr(arrMCPar(0)) & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                End If
            End If
        Next
    End If
    Screen.MousePointer = 0
    err = 0: On Error GoTo ErrHandRpt:
    If bln是否有退料 = True Then
      If zlStr.IsHavePrivs(mstrPrivs, "退料通知单") Then
            If MsgBox("你需要打印退料清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_2", Me, "退料时间=" & strCurDate, "单位=" & mintUnit + 1, 2)
            End If
     End If
    End If
    
    '调用退药后的外挂接口
    If Not mobjPlugIn Is Nothing And bln是否有退料 Then
        mobjPlugIn.DrugReturnByID Val(mArrFilter("发料部门id")), strReturnInfo, CDate(strCurDate), strReserve
    End If
    
    SaveData = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Function
ErrHandRpt:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    SaveData = True
End Function
  

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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

Private Sub picBatHsc_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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

