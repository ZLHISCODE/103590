VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm不再发药处方标志 
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "Frm不再发药处方标志.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5581.047
   ScaleMode       =   0  'User
   ScaleWidth      =   9000
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   7830
      TabIndex        =   15
      Top             =   720
      Width           =   825
   End
   Begin VB.Frame fraSelect 
      Height          =   1110
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   8655
      Begin VB.TextBox txtBillNo 
         Height          =   300
         Left            =   3870
         TabIndex        =   2
         Top             =   630
         Width           =   1590
      End
      Begin VB.TextBox txtPati 
         Height          =   300
         Left            =   1290
         TabIndex        =   1
         Top             =   630
         Width           =   1590
      End
      Begin VB.Frame fraSelectFlag 
         Caption         =   "是否已标记不再发药"
         Height          =   825
         Left            =   5670
         TabIndex        =   7
         Top             =   180
         Width           =   1905
         Begin VB.OptionButton optUnFlag 
            Caption         =   "未标记"
            Height          =   195
            Left            =   450
            TabIndex        =   9
            Top             =   270
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton optFlag 
            Caption         =   "已标记"
            Height          =   195
            Left            =   450
            TabIndex        =   8
            Top             =   540
            Width           =   1005
         End
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   1290
         TabIndex        =   10
         Top             =   255
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   125370368
         CurrentDate     =   38455
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   3870
         TabIndex        =   17
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   125370368
         CurrentDate     =   38455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "结束时间"
         Height          =   180
         Left            =   3060
         TabIndex        =   16
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "开始时间"
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   315
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "病  人"
         Height          =   180
         Left            =   540
         TabIndex        =   12
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "处 方 号"
         Height          =   180
         Left            =   3060
         TabIndex        =   11
         Top             =   690
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   105
      TabIndex        =   5
      Top             =   5205
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   4
      Top             =   5205
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "标记(&O)"
      Height          =   350
      Left            =   6120
      TabIndex        =   3
      Top             =   5205
      Width           =   1275
   End
   Begin VB.Frame fraGrid 
      Height          =   3840
      Left            =   90
      TabIndex        =   0
      Top             =   1260
      Width           =   8655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBill 
         Height          =   3435
         Left            =   50
         TabIndex        =   14
         Top             =   180
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   6059
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "库存检查"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1440
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "Frm不再发药处方标志"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''常量

'控件或窗口标题
Private Const M_STR_FRASELECT_CAPTION As String = "查询条件"                             '查询框标题

'药品处方
Private Const M_STR_药品_FRM_CAPTION As String = "停止发药标志"                          '窗口标题
Private Const M_STR_药品_FRAGRID_CHECK_CAPTION As String = "已标记停止发药处方信息"      '处方信息框标题-已作标记时
Private Const M_STR_药品_FRAGRID_UNCHECK_CAPTION As String = "未标记停止发药处方信息"    '处方信息框标题-未作标记时
Private Const M_STR_药品_FRASELECTFLAG_CAPTION As String = "是否已标记停止发药"          '标记和未标记选择框标题

'卫材处方
Private Const M_STR_卫材_FRM_CAPTION As String = "停止发料标志"                          '窗口标题
Private Const M_STR_卫材_FRAGRID_CHECK_CAPTION As String = "已标记停止发料处方信息"      '处方信息框标题-已作标记时
Private Const M_STR_卫材_FRAGRID_UNCHECK_CAPTION As String = "未标记停止发料处方信息"    '处方信息框标题-未作标记时
Private Const M_STR_卫材_FRASELECTFLAG_CAPTION As String = "是否已标记停止发料"          '标记和未标记选择框标题


'确认按钮标题
Private Const M_STR_CMDOK_CHECK As String = "标记(&F)"
Private Const M_STR_CMDOK_UNCHECK As String = "恢复标记(&U)"

'标记名称
Private Const M_STR_CHECK_NAME As String = "已标记"
Private Const M_STR_UNCHECK_NAME As String = "未标记"

'已、未标记记录颜色
Const M_LNG_CHECKED_COLOR = &HC0C0C0
Const M_LNG_UNCHECKED_COLOR = &H8000000E

'固定单元背景色
Const M_LNG_FIXEDCOLS_COLOR = &H8000000F

'选中行背景色
Const M_LNG_SELECTEDCOLS_COLOR = &HFFC0C0

'默认行背景色
Const M_LNG_DEFAULTCOLS_COLOR = &H8000000E

'控件或窗口提示
Private Const M_STR_PATI_INPUT_DESC As String = "可以输入：*门诊号，+住院号，-病人ID来查找病人"                       '病人输入框提示
Private Const M_STR_FLAG_DESC As String = "选择未标记-查询未作标记的未发处方；选择已标记-查询已作标记的未发处方"        '标记和未标记选择框提示
Private Const M_STR_GRID_DESC As String = "在第一列打勾来标记未发处方，也可以取消已打勾的列"                              '处方记录列表提示
Private Const M_STR_BILLNO_DESC As String = "输入处方号，支持输入最后的数字来提取该处方信息"                                '处方号输入框提示

'窗口默认宽度、高度
Private Const M_STR_FRM_WIDTH As Long = 9000
Private Const M_STR_FRM_HEIGHT As Long = 6000

'各种字符、数字格式
Private Const M_STR_DEFAULT_ORA_NUMERIC_FORMAT As String = "9999990.00000"               '默认Oracle金额格式
Private Const M_STR_DEFAULT_ORA_DATE_FORMAT As String = "yyyy-mm-dd hh24:mi:ss"          '默认Oracle时间格式
Private Const M_STR_VB_DATE_FORMAT As String = "yyyy-mm-dd hh:mm:ss"                     '默认VB时间格式


''''变量

'窗口或控件标题变量
Private mstrFrmCaption As String
Private mstrFraGridCheckCaption As String
Private mstrFraGridUnCheckCaption As String
Private mstrFraSelectFlagCaption As String

Private mstr门诊收费与发药分离 As String
Private mstr住院记帐与发药分离 As String
Private mint单据类型 As Integer                  '1-处方发药单据 2-部门发药单据
Private mlng发药药房ID As Long
Private mstr药房 As String
Private mstrSystemNumericFormat As String            '金额格式
Private mstrSystemAmountFormat As String             '数量格式
Private mblnIsChecked As Boolean                      '判断记录原始状态是否为已作标记
Private mintBillType As Integer                      '处方类别 1-药品处方 2-卫材处方
Private mint请求类型 As Integer                       '区分门诊住院：0-门诊住院，1-门诊，2-住院
Private mbln启用审方 As Boolean
Private mlng药房ID As Long
Private mintNumberDigit As Integer          '数量小数位数


Private mstrPrivs As String                              '权限串

Private mIntCheckStock As Integer               '库存检查：0-不检查;1-检查,不足提醒;2-检查,不足禁止

Private mstr停止 As String
Private mstr恢复 As String

Private mBillCol As BILLCOL

Private mstrDeptNode As String

''''全局变量
Public gstrParentName As String                     '父窗体的名字

''''类型

'处方列属性
Private Type BILLCOL
    BillCols As Integer

    Flag  As String
    FlagCol  As Integer
    FlagColWidth As Long
    FlagColAlig As Integer
       
    Tag  As String
    TagCol  As Integer
    TagColWidth As Long
    TagColAlig As Integer
        
    Id As String
    IdCol As Integer
    IdColWidth As Long
    IdColAlig As Integer
        
    Pati As String
    PatiCol As Integer
    PatiColWidth As Long
    PatiColAlig As Integer
    
    Bill As String
    BILLCOL As Integer
    BillColWidth As Long
    BillColAlig As Integer
    
    NO As String
    NoCol As Integer
    NoColWidth As Long
    NoColAlig As Integer
    
    Drug As String
    DrugCol As Integer
    DrugColWidth As Long
    DrugColAlig As Integer
    
    Spec As String
    SpecCol As Integer
    SpecColWidth As Long
    SpecColAlig As Integer
    
    Unit As String
    UnitCol As Integer
    UnitColWidth As Long
    UnitColAlig As Integer
    
    UnitPrice As String
    UnitPriceCol As Integer
    UnitPriceColWidth As Long
    UnitPriceColAlig As Integer
    UnitPriceFormat As String
    
    count As String
    CountCol As Integer
    CountColWidth As Long
    CountColAlig As Integer
    
    Amount As String
    AmountCol As Integer
    AmountColWidth As Long
    AmountColAlig As Integer
        
    Price As String
    PriceCol As Integer
    PriceColWidth As Long
    PriceColAlig As Integer
    PriceColFormat As String
    
    Group As String
    GroupCol As Integer
    GroupColWidth As Long
    GroupColAlig As Integer
    
    BillDate As String
    BillDateCol As Integer
    BillDateColWidth As Long
    BillDateColAlig As Integer
    
    Category As String
    CategoryCol As Integer
    CategoryColWidth As Integer
    CategoryColAlig As Integer
    
    记录性质 As String
    记录性质Col As Integer
    记录性质ColWidth As Integer
    记录性质ColAlig As Integer
    
    门诊标志 As String
    门诊标志Col As Integer
    门诊标志ColWidth As Integer
    门诊标志ColAlig As Integer
    
    缺药 As String
    缺药Col As Integer
    缺药ColWidth As Integer
    缺药ColAlig As Integer
End Type

Public Property Get In_库存检查() As Integer
    In_库存检查 = mIntCheckStock
End Property

Public Property Let In_库存检查(ByVal vNewValue As Integer)
    mIntCheckStock = vNewValue
End Property

'--50313（zdt）:添加属性接受发料部门id的值
Public Property Get In_库房id() As Long
    In_库房id = mlng发药药房ID
End Property

Public Property Let In_库房id(ByVal vNewValue As Long)
    mlng发药药房ID = vNewValue
End Property

Public Property Let In_请求类型(ByVal vNewValue As Integer)
    mint请求类型 = vNewValue
End Property

Public Property Get In_请求类型() As Integer
    In_请求类型 = mint请求类型
End Property

'检查是否为日期
Private Function CheckIsDate(dtInput As Date) As Boolean
    CheckIsDate = IsDate(dtInput)
End Function

'检查是否为数字
Private Function CheckIsNumber(bytInput As Byte) As Boolean
    Dim strTmp As String
    strTmp = "0123456789"
    CheckIsNumber = (InStr(strTmp, bytInput) > 0)
    
End Function


Private Function GetPatiName(ByVal strInput As String) As String
    Dim intLen As Integer
    Dim strTmp As String
    Dim strsql As String
    Dim blnTmp As Boolean
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    
    blnTmp = True
    strTmp = Trim(strInput)
    intLen = Len(strTmp)
    
    strsql = "Select distinct 姓名 From 病人信息 "
    
    If InStr("*-+", Mid(strTmp, 1, 1)) > 0 Then
        For n = 2 To intLen
            If InStr("0123456789", Mid(strTmp, n, 1)) = 0 Then
                blnTmp = False
                Exit For
            End If
        Next
        If blnTmp = True Then
            Select Case Mid(strTmp, 1, 1)
                Case "*"
                    strsql = strsql & " where 门诊号=[1]"
                Case "+"
                    strsql = strsql & " where 住院号=[1]"
                Case "-"
                    strsql = strsql & " where 病人ID=[1]"
                Case Else
            End Select
        Else
            GetPatiName = strInput
            Exit Function
        End If
    Else
        GetPatiName = strInput
        Exit Function
    End If
    
    On Error GoTo err

    Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, Val(Mid(txtPati.Text, 2)))
    
    If rs.RecordCount > 0 Then
        GetPatiName = rs!姓名
    Else
        GetPatiName = strInput
    End If
    
    rs.Close
    Exit Function
   
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Get本机参数()
    Dim rsData As ADODB.Recordset
    Dim int处方审查 As Integer
    Dim int审查时机 As Integer
    
    On Error GoTo errHandle
    If gstrParentName = "frm药品处方发药New" Then
        mlng发药药房ID = Val(zlDatabase.GetPara("发药药房", glngSys, 1341, 0))
    ElseIf gstrParentName = "Frm部门发药管理New" Then
        mlng发药药房ID = Val(zlDatabase.GetPara("发药药房", glngSys, 1342, 0))
        If mlng发药药房ID = 0 Then
            mlng发药药房ID = mlng药房ID
        End If
    End If
    
    int处方审查 = zlDatabase.GetPara(241, glngSys, , 0)
    int审查时机 = zlDatabase.GetPara(242, glngSys, , 0)
    mstrDeptNode = GetDeptStationNode(mlng发药药房ID)
    mbln启用审方 = ((int处方审查 = 1 Or int处方审查 = 3) And int审查时机 = 2)
        
    
    If mlng发药药房ID > 0 Then
        gstrSQL = "Select 名称 From 部门表 Where ID = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Get本机参数", mlng发药药房ID)
        
        If rsData.RecordCount > 0 Then
            mstr药房 = rsData!名称
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'取系统参数
Private Sub Get系统参数()
    Dim rs As New ADODB.Recordset
    Dim intTmp As Integer
    Dim n As Integer
    
   '门诊收费与发药分离
    On Error GoTo errHandle
    mstr门诊收费与发药分离 = zlDatabase.GetPara(15, glngSys, , "0")
    '住院记帐与发药分离
    mstr住院记帐与发药分离 = zlDatabase.GetPara(16, glngSys, , "0")
    
    '费用金额保留位数
    intTmp = CInt(zlDatabase.GetPara(9, glngSys))
    mstrSystemNumericFormat = "0."
    
    For n = 1 To intTmp
        mstrSystemNumericFormat = mstrSystemNumericFormat & "0"
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub IniControls()
    '窗口默认宽度、高度
    Me.Width = M_STR_FRM_WIDTH
    Me.Height = M_STR_FRM_HEIGHT
    
    '窗口或控件标题
    Me.Caption = mstrFrmCaption
    fraSelect.Caption = M_STR_FRASELECT_CAPTION
    fraGrid.Caption = mstrFraGridUnCheckCaption
    fraSelectFlag.Caption = mstrFraSelectFlagCaption
    CmdOK.Caption = M_STR_CMDOK_CHECK
    
    '控件提示信息
    txtPati.ToolTipText = M_STR_PATI_INPUT_DESC
    txtBillNo.ToolTipText = M_STR_BILLNO_DESC
    mshBill.ToolTipText = M_STR_GRID_DESC
    fraSelectFlag.ToolTipText = M_STR_FLAG_DESC
    
    If zlStr.IsHavePrivs(mstrPrivs, mstr停止) = False Then
        optUnFlag.Enabled = False
        optFlag.Value = True
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, mstr恢复) = False Then
        optFlag.Enabled = False
        optUnFlag.Value = True
    End If
    
End Sub

Private Sub IniGrid()
    With mBillCol
        .BillCols = 19
    
        .Flag = ""
        .FlagCol = 0
        .FlagColWidth = 400
        .FlagColAlig = 1
        
        .Tag = "标记"
        .TagCol = 1
        .TagColWidth = 0
        .TagColAlig = 1
           
        .Pati = "病人姓名"
        .PatiCol = 2
        .PatiColWidth = 1000
        .PatiColAlig = 1
        
        .Bill = "处方号"
        .BILLCOL = 3
        .BillColWidth = 800
        .BillColAlig = 1
        
        .NO = "序号"
        .NoCol = 4
        .NoColWidth = 600
        .NoColAlig = 1
        
        Select Case mintBillType
            Case 1
                .Drug = "药品名称"
            Case 2
                .Drug = "卫材名称"
            Case Else
        End Select
        .DrugCol = 5
        .DrugColWidth = 1500
        .DrugColAlig = 1
        
        .Spec = "规格"
        .SpecCol = 6
        .SpecColWidth = 1000
        .SpecColAlig = 1
        
        .Unit = "单位"
        .UnitCol = 7
        .UnitColWidth = 600
        .UnitColAlig = 1
        
        .UnitPrice = "单价"
        .UnitPriceCol = 8
        .UnitPriceColWidth = 800
        .UnitPriceColAlig = 7
        .UnitPriceFormat = mstrSystemNumericFormat
        
        .count = "付数"
        .CountCol = 9
        Select Case mintBillType
            Case 1
                .CountColWidth = 600
            Case 2
                .CountColWidth = 0
            Case Else
        End Select
        .CountColAlig = 7
        
        .Amount = "数量"
        .AmountCol = 10
        .AmountColWidth = 800
        .AmountColAlig = 7
            
        .Price = "金额"
        .PriceCol = 11
        .PriceColWidth = 1000
        .PriceColAlig = 7
        .PriceColFormat = mstrSystemNumericFormat
        
        .Group = "批次"
        .GroupCol = 12
        .GroupColWidth = 600
        .GroupColAlig = 7
        
        .BillDate = "处方日期"
        .BillDateCol = 13
        .BillDateColWidth = 2000
        .BillDateColAlig = 1
        
        .Category = "发药方式"
        .CategoryCol = 14
        .CategoryColWidth = 0
        .CategoryColAlig = 1
        
        .Id = "Id"
        .IdCol = 15
        .IdColWidth = 0
        .IdColAlig = 7
    
        .记录性质 = "记录性质"
        .记录性质Col = 16
        .记录性质ColWidth = 0
        .记录性质ColAlig = 7
        
        .门诊标志 = "门诊标志"
        .门诊标志Col = 17
        .门诊标志ColWidth = 0
        .门诊标志ColAlig = 7
        
        .缺药 = "缺药"
        .缺药Col = 18
        .缺药ColWidth = 0
        .缺药ColAlig = 7
    End With
    
    With mshBill
        .Clear
        .Cols = mBillCol.BillCols
        .rows = 1
        .rows = 2
        
        .TextMatrix(0, mBillCol.AmountCol) = mBillCol.Amount
        .TextMatrix(0, mBillCol.BILLCOL) = mBillCol.Bill
        .TextMatrix(0, mBillCol.BillDateCol) = mBillCol.BillDate
        .TextMatrix(0, mBillCol.CategoryCol) = mBillCol.Category
        .TextMatrix(0, mBillCol.CountCol) = mBillCol.count
        .TextMatrix(0, mBillCol.DrugCol) = mBillCol.Drug
        .TextMatrix(0, mBillCol.GroupCol) = mBillCol.Group
        .TextMatrix(0, mBillCol.NoCol) = mBillCol.NO
        .TextMatrix(0, mBillCol.PatiCol) = mBillCol.Pati
        .TextMatrix(0, mBillCol.PriceCol) = mBillCol.Price
        .TextMatrix(0, mBillCol.SpecCol) = mBillCol.Spec
        .TextMatrix(0, mBillCol.UnitCol) = mBillCol.Unit
        .TextMatrix(0, mBillCol.UnitPriceCol) = mBillCol.UnitPrice
        .TextMatrix(0, mBillCol.FlagCol) = mBillCol.Flag
        .TextMatrix(0, mBillCol.TagCol) = mBillCol.Tag
        .TextMatrix(0, mBillCol.IdCol) = mBillCol.Id
        .TextMatrix(0, mBillCol.记录性质Col) = mBillCol.记录性质
        .TextMatrix(0, mBillCol.门诊标志Col) = mBillCol.门诊标志
        .TextMatrix(0, mBillCol.缺药Col) = mBillCol.缺药
        
        .ColWidth(mBillCol.AmountCol) = mBillCol.AmountColWidth
        .ColWidth(mBillCol.BILLCOL) = mBillCol.BillColWidth
        .ColWidth(mBillCol.BillDateCol) = mBillCol.BillDateColWidth
        .ColWidth(mBillCol.CategoryCol) = mBillCol.CategoryColWidth
        .ColWidth(mBillCol.CountCol) = mBillCol.CountColWidth
        .ColWidth(mBillCol.DrugCol) = mBillCol.DrugColWidth
        .ColWidth(mBillCol.GroupCol) = mBillCol.GroupColWidth
        .ColWidth(mBillCol.NoCol) = mBillCol.NoColWidth
        .ColWidth(mBillCol.PatiCol) = mBillCol.PatiColWidth
        .ColWidth(mBillCol.PriceCol) = mBillCol.PriceColWidth
        .ColWidth(mBillCol.SpecCol) = mBillCol.SpecColWidth
        .ColWidth(mBillCol.UnitCol) = mBillCol.UnitColWidth
        .ColWidth(mBillCol.UnitPriceCol) = mBillCol.UnitPriceColWidth
        .ColWidth(mBillCol.FlagCol) = mBillCol.FlagColWidth
        .ColWidth(mBillCol.TagCol) = mBillCol.TagColWidth
        .ColWidth(mBillCol.IdCol) = mBillCol.IdColWidth
        .ColWidth(mBillCol.记录性质Col) = mBillCol.记录性质ColWidth
        .ColWidth(mBillCol.门诊标志Col) = mBillCol.门诊标志ColWidth
        .ColWidth(mBillCol.缺药Col) = mBillCol.缺药ColWidth
        
        .ColAlignment(mBillCol.AmountCol) = mBillCol.AmountColAlig
        .ColAlignment(mBillCol.BILLCOL) = mBillCol.BillColAlig
        .ColAlignment(mBillCol.BillDateCol) = mBillCol.BillDateColAlig
        .ColAlignment(mBillCol.CategoryCol) = mBillCol.CategoryColAlig
        .ColAlignment(mBillCol.CountCol) = mBillCol.CountColAlig
        .ColAlignment(mBillCol.DrugCol) = mBillCol.DrugColAlig
        .ColAlignment(mBillCol.GroupCol) = mBillCol.GroupColAlig
        .ColAlignment(mBillCol.NoCol) = mBillCol.NoColAlig
        .ColAlignment(mBillCol.PatiCol) = mBillCol.PatiColAlig
        .ColAlignment(mBillCol.PriceCol) = mBillCol.PriceColAlig
        .ColAlignment(mBillCol.SpecCol) = mBillCol.SpecColAlig
        .ColAlignment(mBillCol.UnitCol) = mBillCol.UnitColAlig
        .ColAlignment(mBillCol.UnitPriceCol) = mBillCol.UnitPriceColAlig
        .ColAlignment(mBillCol.FlagCol) = mBillCol.FlagColAlig
        .ColAlignment(mBillCol.TagCol) = mBillCol.TagColAlig
        .ColAlignment(mBillCol.IdCol) = mBillCol.IdColAlig
        .ColAlignment(mBillCol.记录性质Col) = mBillCol.记录性质ColAlig
        .ColAlignment(mBillCol.门诊标志Col) = mBillCol.门诊标志ColAlig
        .ColAlignment(mBillCol.缺药Col) = mBillCol.缺药ColAlig
    End With
    
    Dim n As Long
    With mshBill
        .Row = 0
        For n = 0 To .Cols - 1
            .Col = n
            .CellBackColor = M_LNG_FIXEDCOLS_COLOR
        Next
    End With
    
End Sub
Public Function GetUnit(ByVal lng药房ID As Long) As String
    '返回指定库房、单据、NO适用的药品单位
    Dim intUnit As Integer
    Dim rstemp As New ADODB.Recordset
    
    '根据系统参数设定的单位显示数据
    Select Case mintBillType
        Case 1      '取药品单位
            intUnit = Val(zlDatabase.GetPara("药房属性", glngSys, 1341))
                If intUnit = 0 Then
                    '取当前处方的病人来源
                    intUnit = mint单据类型
                End If
                If intUnit = 1 Then
                    GetUnit = GetSpecUnit1(lng药房ID, 2)
                Else
                    GetUnit = GetSpecUnit1(lng药房ID, 3)
                End If
        Case 2      '取卫材单位
            intUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, 1723, "0"))
            GetUnit = IIf(intUnit = 1, "包装单位", "售价单位")
        Case Else
    End Select
    
End Function

Public Function GetSpecUnit1(ByVal lng库房id As Long, ByVal int范围 As Integer) As String
    Dim strobjTemp As String                    '保存服务对象字符串
    Dim strWorkTemp As String                   '保存工作性质字符串
    Dim strUnit As String
    Dim rsProperty As New ADODB.Recordset
    Dim strsql As String
    
    
    '返回指定库房指定适用范围的单位
    On Error GoTo ErrHand
    
    gstrSQL = "Select Nvl(性质,1) AS 单位 From 药品库房单位 Where 库房ID=[1] And 适用范围=[2]"
    
    Set rsProperty = zlDatabase.OpenSQLRecord(gstrSQL, "提取单位", lng库房id, int范围)
    
   
    If rsProperty.RecordCount = 1 Then
        strUnit = rsProperty!单位
    Else
'        MsgBox "该库房未设置库房单位，根据部门性质以及服务对象取缺省单位！" & _
'            vbCrLf & "缺省单位的规则：" & _
'            vbCrLf & "  服务对象是住院或门诊和住院的，取住院单位" & _
'            vbCrLf & "  仅服务于门诊的，取门诊单位" & _
'            vbCrLf & "  具有药库属性的，取药库单位" & _
'            vbCrLf & "  其他取售价单位", vbInformation, gstrSysName
        
        gstrSQL = "SELECT distinct 服务对象,工作性质 From 部门性质说明 Where 部门ID =[1]"
        
        Set rsProperty = zlDatabase.OpenSQLRecord(gstrSQL, "读取药品单位", lng库房id)
            
        '取服务对象及部门性质
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            '住院单位
            strUnit = 3
        ElseIf InStr(strobjTemp, "1") <> 0 Then
            '门诊单位
            strUnit = 2
        ElseIf InStr(strWorkTemp, "药库") <> 0 Then
            '药库单位
            strUnit = 4
        Else
            '售价单位：主要是制剂室
            strUnit = 1
        End If
    End If
    
    '转换为真实的单位返回给调用者
    GetSpecUnit1 = Switch(strUnit = 1, "售价单位", strUnit = 2, "门诊单位", strUnit = 3, "住院单位", strUnit = 4, "药库单位")
    If glngSys / 100 = 8 Then
        '药店只有售价单位与药库单位
        GetSpecUnit1 = IIf(strUnit = 1, "售价单位", "药库单位")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub GetDetailBill()
    Dim strUnit As String           '药品单位
    Dim rs As New ADODB.Recordset
    Dim strInputFlag As String
    Dim strSubSql As String
    Dim strsql As String
    Dim n As Long
    Dim strStartDate As String
    Dim strEndDate As String
    Dim strTmp As String
    
    On Error GoTo err
    
    strStartDate = Format(dtpStartDate, "yyyy-mm-dd 00:00:01")
    strEndDate = Format(dtpEndDate, "yyyy-mm-dd 23:59:59")
    
    Call IniGrid
    
    '''''''''''''构造SQL语句
    
    ''''select 子句
    strsql = "select distinct "
    strsql = strsql & " b.姓名 as " & mBillCol.Pati & ","
    strsql = strsql & " a.No as " & mBillCol.Bill & ","
    strsql = strsql & " a.序号 as " & mBillCol.NO & ","
    strsql = strsql & " Decode(e.名称,NULL,d.名称,e.名称) as " & mBillCol.Drug & ","
    strsql = strsql & " d.规格 as " & mBillCol.Spec & ","
    strsql = strsql & " a.零售价 as " & mBillCol.UnitPrice & ","
    strsql = strsql & " a.零售金额 as " & mBillCol.Price & ","
    strsql = strsql & " NVL(a.批次,0) as " & mBillCol.Group & ","
    strsql = strsql & " a.填制日期 as " & mBillCol.BillDate & ","
    strsql = strsql & " NVL(a.发药方式,-999) as " & mBillCol.Category & ","
    strsql = strsql & " NVL(a.付数,1) as " & mBillCol.count & ","
    strsql = strsql & " a.id as " & mBillCol.Id & ","
    strsql = strsql & " b.记录性质 as " & mBillCol.记录性质 & ","
    strsql = strsql & " b.门诊标志 as " & mBillCol.门诊标志 & ","
    
        
    '药品单位、单价、数量
    strUnit = GetUnit(mlng发药药房ID)
'    strUnit = "住院单位"
    Select Case strUnit
    Case "售价单位"
        strSubSql = "1"
        strsql = strsql & " D.计算单位 as " & mBillCol.Unit & ","
        strsql = strsql & " ltrim(to_char(A.零售价,'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.UnitPrice & ","
        strsql = strsql & " ltrim(to_char(A.实际数量,'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.Amount
    Case "门诊单位"
        strSubSql = "Decode(门诊包装,Null,1,0,1,门诊包装)"
        strsql = strsql & " F.门诊单位 as " & mBillCol.Unit & ","
        strsql = strsql & " ltrim(to_char(A.零售价*Decode(F.门诊包装,Null,1,0,1,F.门诊包装),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.UnitPrice & ","
        strsql = strsql & " ltrim(to_char(A.实际数量/Decode(F.门诊包装,Null,1,0,1,F.门诊包装),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.Amount
    Case "住院单位"
        strSubSql = "Decode(住院包装,Null,1,0,1,住院包装)"
        strsql = strsql & " F.住院单位 as " & mBillCol.Unit & ","
        strsql = strsql & " ltrim(to_char(A.零售价*Decode(F.住院包装,Null,1,0,1,F.住院包装),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.UnitPrice & ","
        strsql = strsql & " ltrim(to_char(A.实际数量/Decode(F.住院包装,Null,1,0,1,F.住院包装),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.Amount
    Case "药库单位"
        strSubSql = "Decode(药库包装,Null,1,0,1,药库包装)"
        strsql = strsql & " F.药库单位 as " & mBillCol.Unit & ","
        strsql = strsql & " ltrim(to_char(A.零售价*Decode(F.药库包装,Null,1,0,1,F.药库包装),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.UnitPrice & ","
        strsql = strsql & " ltrim(to_char(A.实际数量/Decode(F.药库包装,Null,1,0,1,F.药库包装),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.Amount
    Case "包装单位"         '卫材
        strSubSql = "Decode(包装单位,Null,1,0,1,包装单位)"
        strsql = strsql & " F.包装单位 as " & mBillCol.Unit & ","
        strsql = strsql & " ltrim(to_char(A.零售价*Decode(F.换算系数,Null,1,0,1,F.换算系数),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.UnitPrice & ","
        strsql = strsql & " ltrim(to_char(A.实际数量/Decode(F.换算系数,Null,1,0,1,F.换算系数),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.Amount
    End Select
    
    If mIntCheckStock > 0 And mblnIsChecked = True Then
        strsql = strsql & ",Decode(Sign(Nvl(K.可用数量, 0) - A.实际数量 * Nvl(A.付数, 1)), -1, 1, 0) " & mBillCol.缺药
    Else
        strsql = strsql & ",0 " & mBillCol.缺药
    End If
    
    ''''from子句
    strsql = strsql & " From 药品收发记录 A,门诊费用记录 B,部门表 C,收费项目目录 D,收费项目别名 E,部门表 P " & IIf(mbln启用审方, ",处方审查记录 Q,处方审查明细 K ", "")
    Select Case mintBillType
        Case 1
            strsql = strsql & ",药品规格 F"
        Case 2
            strsql = strsql & ",材料特性 F"
        Case Else
    End Select
    
    '要库存检查时，关联药品库存表
    If mIntCheckStock > 0 Then
        strsql = strsql & ",(Select 药品id, Nvl(批次, 0) 批次, Nvl(可用数量, 0) 可用数量 From 药品库存 Where 性质 = 1 And 库房id = [1]) K "
    End If
    
    ''''where 子句
    strsql = strsql & " where A.费用id=B.Id And A.药品id=D.Id And D.ID=E.收费细目ID(+) And B.收费细目id=D.Id " & IIf(mbln启用审方, " and b.医嘱序号=k.医嘱id(+) and Q.id(+)=K.审方id and K.最后提交(+)=1 And (b.医嘱序号 is null or nvl(q.审查结果,0) = 1)", "") & _
         " And A.库房id+0=C.Id  AND E.性质(+)=3 And A.对方部门ID = P.ID And Nvl(B.费用状态,0)<>1 "
    
    If mstrDeptNode <> "" Then
        strsql = strsql & " And (P.站点 = [6] Or P.站点 Is Null)"
    End If
    
    Select Case mintBillType
        Case 1
            strsql = strsql & " AND A.药品id=F.药品id "
        Case 2
            strsql = strsql & " AND A.药品id=F.材料id "
        Case Else
    End Select
    
    '必须的查询条件
    strsql = strsql & " and Mod(A.记录状态,3)=1 and A.审核人 is null "                 '未发药处方
        
    '根据参数和父窗口传递的查询条件
    If mint单据类型 = 1 Then
        strsql = strsql & " and A.单据 in(8,9) "
        If mstr门诊收费与发药分离 = "0" Then
            strsql = strsql & " and A.库房id+0=[1]"
        End If
    ElseIf mint单据类型 = 2 Then
        '如果是部门发药，则只查询住院费用记录
        strsql = strsql & " and A.单据 in(9,10) "
        If mstr住院记帐与发药分离 = "0" Then
            strsql = strsql & " and A.库房id+0=[1]"
        End If
    Else
        strsql = strsql & " and A.库房id+0=[1]"
    End If
        
    '用户选择的查询条件
    strsql = strsql & " and A.填制日期>=[2]  and A.填制日期<=[3] "
    
    If txtPati.Text <> "" Then
        strsql = strsql & " and B.姓名=[4] "
    End If
    
    If txtBillNo.Text <> "" Then
        strsql = strsql & " and A.no=[5] "
    End If
    
    If Not mblnIsChecked Then
        strsql = strsql & " and NVL(A.发药方式,-999)<>-1"
    Else
        strsql = strsql & " and A.发药方式=-1"
        
        If mIntCheckStock > 0 Then
            strsql = strsql & " And A.药品id = K.药品id(+) And Nvl(A.批次, 0) = K.批次(+) "
        End If
    End If
    
    If mint单据类型 = 1 Then
        '如果是处方发药，则合并门诊费用记录和住院费用记录
        strTmp = Replace(strsql, "b.姓名", "nvl(R.姓名,b.姓名)")
        strTmp = Replace(strTmp, "B.姓名", "nvl(R.姓名,b.姓名)")
        strTmp = Replace(strTmp, "门诊费用记录 B", "住院费用记录 B,病案主页 R")
        strTmp = Replace(strTmp, "And Nvl(B.费用状态,0)<>1", " And B.病人id=R.病人id And B.主页ID=R.主页ID ")
        strTmp = Replace(strTmp, "in(8,9)", "in(9,10)")
        
        strsql = strsql & " Union All " & strTmp
        
        ''''order子句
        strsql = strsql & " order by 处方号,序号 "
    ElseIf mint单据类型 = 0 Or mint单据类型 = 2 Then
        If mint请求类型 = 0 Then
            strTmp = Replace(strsql, "b.姓名", "nvl(R.姓名,b.姓名)")
            strTmp = Replace(strTmp, "B.姓名", "nvl(R.姓名,b.姓名)")
            strTmp = Replace(strTmp, "门诊费用记录 B", "住院费用记录 B,病案主页 R")
            strTmp = Replace(strTmp, "And Nvl(B.费用状态,0)<>1", " And B.病人id=R.病人id And B.主页ID=R.主页ID ")
            
            strsql = strsql & " Union All " & strTmp
        ElseIf mint请求类型 = 2 Then
            strsql = Replace(strsql, "b.姓名", "nvl(R.姓名,b.姓名)")
            strsql = Replace(strsql, "B.姓名", "nvl(R.姓名,b.姓名)")
            strsql = Replace(strsql, "门诊费用记录 B", "住院费用记录 B,病案主页 R")
            strsql = Replace(strsql, "And Nvl(B.费用状态,0)<>1", " And B.病人id=R.病人id And B.主页ID=R.主页ID ")
        End If
        
        ''''order子句
        strsql = strsql & " order by 处方号,序号 "
    Else
        ''''order子句
        strsql = strsql & " order by A.No,A.序号 "
    End If
    
    
    
    '''''''''''''以上构造SQL语句
    
    
    ''''查询结果
    Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mlng发药药房ID, CDate(strStartDate), CDate(strEndDate), txtPati.Text, txtBillNo.Text, mstrDeptNode)
    
        
    DoEvents
    Me.MousePointer = 11
        
    With mshBill
        .Redraw = False

        If (Not rs.EOF) And (Not rs.BOF) Then
            CmdOK.Enabled = True
        End If
        Do While Not rs.EOF
        
            If .rows >= 2 And .TextMatrix(1, 1) <> "" Then
                .rows = .rows + 1
            End If
            .TextMatrix(.rows - 1, mBillCol.AmountCol) = FormatEx(rs.Fields(mBillCol.Amount).Value, mintNumberDigit)
            .TextMatrix(.rows - 1, mBillCol.BILLCOL) = rs.Fields(mBillCol.Bill).Value
            .TextMatrix(.rows - 1, mBillCol.BillDateCol) = rs.Fields(mBillCol.BillDate).Value
            .TextMatrix(.rows - 1, mBillCol.CategoryCol) = rs.Fields(mBillCol.Category).Value
            .TextMatrix(.rows - 1, mBillCol.CountCol) = rs.Fields(mBillCol.count).Value
            .TextMatrix(.rows - 1, mBillCol.DrugCol) = rs.Fields(mBillCol.Drug).Value
            .TextMatrix(.rows - 1, mBillCol.GroupCol) = rs.Fields(mBillCol.Group).Value
            .TextMatrix(.rows - 1, mBillCol.NoCol) = rs.Fields(mBillCol.NO).Value
            .TextMatrix(.rows - 1, mBillCol.PatiCol) = IIf(IsNull(rs.Fields(mBillCol.Pati).Value), "", rs.Fields(mBillCol.Pati).Value)
            .TextMatrix(.rows - 1, mBillCol.PriceCol) = Format(rs.Fields(mBillCol.Price).Value, mBillCol.PriceColFormat)
            .TextMatrix(.rows - 1, mBillCol.SpecCol) = IIf(IsNull(rs.Fields(mBillCol.Spec).Value), "", rs.Fields(mBillCol.Spec).Value)
            .TextMatrix(.rows - 1, mBillCol.UnitCol) = IIf(IsNull(rs.Fields(mBillCol.Unit).Value), "", rs.Fields(mBillCol.Unit).Value)
            .TextMatrix(.rows - 1, mBillCol.UnitPriceCol) = Format(rs.Fields(mBillCol.UnitPrice).Value, mBillCol.UnitPriceFormat)
            .TextMatrix(.rows - 1, mBillCol.IdCol) = rs.Fields(mBillCol.Id).Value
            .TextMatrix(.rows - 1, mBillCol.记录性质Col) = rs.Fields(mBillCol.记录性质).Value
            .TextMatrix(.rows - 1, mBillCol.门诊标志Col) = rs.Fields(mBillCol.门诊标志).Value
            .TextMatrix(.rows - 1, mBillCol.缺药Col) = rs.Fields(mBillCol.缺药).Value
            
            .TextMatrix(.rows - 1, mBillCol.FlagCol) = ""
            
            .Col = 0
            .Row = .rows - 1
            
            .TextMatrix(.rows - 1, mBillCol.TagCol) = M_STR_UNCHECK_NAME
            Set .CellPicture = LoadResPicture("unchecked", vbResBitmap)
            
            '缺药药品用红色字体标记
            If mIntCheckStock > 0 And Val(.TextMatrix(.rows - 1, mBillCol.缺药Col)) = 1 Then
                For n = 0 To .Cols - 1
                    .Col = n
                    .CellForeColor = vbRed
                Next
            End If
            
            rs.MoveNext
        Loop
        .Col = 0
        .Row = 0
        
        .TextMatrix(.rows - 1, mBillCol.TagCol) = M_STR_UNCHECK_NAME
        Set mshBill.CellPicture = LoadResPicture("unchecked", vbResBitmap)
        
        .Redraw = True
    End With
    
    DoEvents
    Me.MousePointer = 0
    
    Exit Sub

err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cmdCancel_Click()
'    Call SaveFlexState(mshBill, Me.Name)
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim n As Long
    Dim int收费与发药分离 As Integer
    Dim int门诊 As Integer
    Dim blnBeginTrans As Boolean
    Dim arrSql As Variant
    
    If mshBill.rows < 2 Then
        Exit Sub
    End If
    
    Select Case mintBillType
        Case 1
            If mint单据类型 = 1 Then
                int收费与发药分离 = mstr门诊收费与发药分离
            Else
                int收费与发药分离 = mstr住院记帐与发药分离
            End If
        Case 2
            int收费与发药分离 = 0
        Case Else
    End Select
    
    mblnIsChecked = optFlag.Value
    arrSql = Array()
    
    On Error GoTo err
    For n = 1 To mshBill.rows - 1
        If mblnIsChecked = True Then
            If mshBill.TextMatrix(n, mBillCol.TagCol) = M_STR_CHECK_NAME Then
                If Val(mshBill.TextMatrix(n, mBillCol.记录性质Col)) = 1 Or (Val(mshBill.TextMatrix(n, mBillCol.记录性质Col)) = 2 And (Val(mshBill.TextMatrix(n, mBillCol.门诊标志Col))) = 1 Or (Val(mshBill.TextMatrix(n, mBillCol.门诊标志Col))) = 4) Then
                    int门诊 = 1
                Else
                    int门诊 = 2
                End If
                
                gstrSQL = "Zl_不发药处方标记_Unchecked(" & mshBill.TextMatrix(n, mBillCol.IdCol) & "," & int收费与发药分离 & "," & int门诊 & ")"
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        End If
        If mblnIsChecked = False Then
            If mshBill.TextMatrix(n, mBillCol.TagCol) = M_STR_CHECK_NAME Then
                gstrSQL = "Zl_不发药处方标记_Checked(" & mshBill.TextMatrix(n, mBillCol.IdCol) & "," & int收费与发药分离 & ")"
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        End If
    Next
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    For n = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(n)), Me.Caption & "-作标记")
    Next
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    Call GetDetailBill
    
    Exit Sub
            
err:
    '如果已开启事务，并且未提交，则出错时回滚事务
    If blnBeginTrans Then
        gcnOracle.RollbackTrans
    End If
    
    MsgBox "提示：更新失败。"
    Call SaveErrLog
End Sub



Private Sub cmdRefresh_Click()
    If mblnIsChecked Then
        CmdOK.Caption = M_STR_CMDOK_UNCHECK
        fraGrid.Caption = mstrFraGridCheckCaption
    Else
        CmdOK.Caption = M_STR_CMDOK_CHECK
        fraGrid.Caption = mstrFraGridUnCheckCaption
    End If

    Call GetDetailBill
End Sub

'
Private Sub dtpStartDate_Change()
'         Call GetDetailBill
End Sub


Private Sub Form_Load()
    Dim intUnit As Integer
    Dim rstemp As Recordset
    
    mstrPrivs = gstrprivs
    
    dtpStartDate.Value = Format(Date, "yyyy-mm-01")
    dtpEndDate.Value = Format(Date, "yyyy-mm-dd")
    
    mblnIsChecked = False
    
    mstrSystemAmountFormat = "0"
    
   
      
    Select Case mintBillType
        Case 1
            mstrFrmCaption = M_STR_药品_FRM_CAPTION
            mstrFraGridCheckCaption = M_STR_药品_FRAGRID_CHECK_CAPTION
            mstrFraGridUnCheckCaption = M_STR_药品_FRAGRID_UNCHECK_CAPTION
            mstrFraSelectFlagCaption = M_STR_药品_FRASELECTFLAG_CAPTION
        Case 2
            mstrFrmCaption = M_STR_卫材_FRM_CAPTION
            mstrFraGridCheckCaption = M_STR_卫材_FRAGRID_CHECK_CAPTION
            mstrFraGridUnCheckCaption = M_STR_卫材_FRAGRID_UNCHECK_CAPTION
            mstrFraSelectFlagCaption = M_STR_卫材_FRASELECTFLAG_CAPTION
        Case Else
    End Select
    
    Call Get系统参数
    Call Get本机参数
    
    Call IniGrid
    
    Me.Caption = Me.Caption & "-" & mstr药房
    If mIntCheckStock = 1 Then
        lblComment.Caption = "可用数量不足药品用红色字体标识！"
    Else
        lblComment.Caption = "可用数量不足药品用红色字体标识，不能恢复标志！"
    End If
    
     '根据父窗体的名称来判断处方类型，如果父窗体名称改变，这里也要作相应改变
    Select Case gstrParentName
        Case "frm药品处方发药New"
            mint单据类型 = 1
            mintBillType = 1
            mstr停止 = "停止发药"
            mstr恢复 = "恢复发药"
            
            Select Case GetDrugUnit(mlng发药药房ID, Me.Caption)
                Case "售价单位"             '售价单位：主要是制剂室
                    intUnit = 1
                Case "门诊单位"
                    intUnit = 2
                Case "住院单位"
                    intUnit = 3
                Case "药库单位"
                    intUnit = 4
            End Select
    
            gstrSQL = "select 精度 from 药品卫材精度 where 性质=0 and 类别 = 1 And 内容 = 3 And 单位 = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询数量精度", intUnit)
            
        Case "Frm部门发药管理New"
            mint单据类型 = 2
            mintBillType = 1
            mstr停止 = "停止发药"
            mstr恢复 = "恢复发药"
    
            Select Case GetDrugUnit(mlng发药药房ID, Me.Caption)
                Case "售价单位"             '售价单位：主要是制剂室
                    intUnit = 1
                Case "门诊单位"
                    intUnit = 2
                Case "住院单位"
                    intUnit = 3
                Case "药库单位"
                    intUnit = 4
            End Select
    
            gstrSQL = "select 精度 from 药品卫材精度 where 性质=0 and 类别 = 1 And 内容 = 3 And 单位 = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询数量精度", intUnit)
        Case "frm卫材发放管理"
            mbln启用审方 = False
            mintBillType = 2
            mstr停止 = "停止发料"
            mstr恢复 = "恢复发料"
            
            '获取数量精度
            intUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, 1723))
            
            gstrSQL = "select 精度 from 药品卫材精度 where 性质=0 and 类别 = 2 And 内容 = 3 And 单位 = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询数量精度", intUnit)
            
            
        Case "frmStuffRxSend"
            mbln启用审方 = False
            mintBillType = 2
            mstr停止 = "停止发料"
            mstr恢复 = "恢复发料"
        Case "frmStuffDeptSend"
            mbln启用审方 = False
            mintBillType = 2
            mstr停止 = "停止发料"
            mstr恢复 = "恢复发料"
        Case Else
    End Select
    
    If Not rstemp.EOF Then
        mintNumberDigit = rstemp!精度
    Else
        mintNumberDigit = 5
    End If
    
    Call IniControls
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    If Me.Width < M_STR_FRM_WIDTH Then Me.Width = M_STR_FRM_WIDTH
    If Me.Height < M_STR_FRM_HEIGHT Then Me.Height = M_STR_FRM_HEIGHT

'    With CmdHelp
'        .Top = Me.ScaleHeight - .Height - 100
'    End With
'
'    With CmdCancel
'        .Top = CmdHelp.Top
'        .Left = Me.ScaleWidth - .Width - 100
'    End With
'    With CmdOK
'        .Top = CmdHelp.Top
'        .Left = CmdCancel.Left - .Width - 100
'    End With
'
'    With fraSelect
'        .Width = Me.ScaleWidth - .Left - 50
'    End With
'
'    With fraGrid
'        .Height = CmdOK.Top - 1400
'        .Width = Me.ScaleWidth - .Left - 50
'    End With
'
'    With mshBill
'        .Left = 50
'        .Height = CmdOK.Top - 1600
'        .Width = fraGrid.Width - .Left - 150
'    End With
    
    With CmdHelp
        .Move .Left, Me.ScaleHeight - .Height - 100
    End With

    With CmdCancel
        .Move Me.ScaleWidth - .Width - 100, CmdHelp.Top
    End With
    
    With CmdOK
        .Move CmdCancel.Left - .Width - 100, CmdHelp.Top
    End With

    With fraSelect
        .Move .Left, .Top, Me.ScaleWidth - .Left - 50
    End With

    With fraGrid
        .Move .Left, .Top, Me.ScaleWidth - .Left - 50, CmdOK.Top - 1400
    End With

    With mshBill
        .Move 50, .Top, fraGrid.Width - .Left - 150, CmdOK.Top - 1600
    End With
    
    With lblComment
        .Top = CmdHelp.Top + 100
    End With
End Sub

Private Sub mshBill_Click()
    Dim n As Long
    Dim i As Long
    Dim lngColor As Long
    Dim lngCurRow As Long
    Dim lngCurCol As Long
    Dim blnWarn As Boolean
    Dim blnWarnDo As Boolean
    
'    Debug.Print "row:" & mshBill.Row & " col:" & mshBill.Col
    
    With mshBill
        .Redraw = False
        lngCurRow = .Row
        lngCurCol = .Col
        If .rows > 1 And .TextMatrix(.rows - 1, mBillCol.BILLCOL) <> "" Then
            '如果选择的是第一列，则作标记或取消标记处理
            If .Col = 0 Then
                If .Row = 0 Then
                    If .TextMatrix(.Row, mBillCol.TagCol) = M_STR_CHECK_NAME Then
                        For n = 0 To .rows - 1
                            .TextMatrix(n, mBillCol.TagCol) = M_STR_UNCHECK_NAME
                            .Row = n
                            .Col = mBillCol.FlagCol
                            Set .CellPicture = LoadResPicture("unchecked", vbResBitmap)
                            If .Row > 0 Then
                                For i = 0 To .Cols - 1
                                    .Row = n
                                    .Col = i
                                    .CellBackColor = M_LNG_UNCHECKED_COLOR
                                Next
                            End If
                        Next
                    Else
                        For n = 0 To .rows - 1
                            blnWarnDo = True
                            
                            '缺药情况处理
                            If n > 0 And mIntCheckStock > 0 And Val(.TextMatrix(n, mBillCol.缺药Col)) = 1 Then
                                If mIntCheckStock = 1 Then
                                    '库存不足提醒时，对缺药记录提醒，并按操作结果进行处理
                                    If blnWarn = False Then
                                        blnWarn = True
                                        If MsgBox("存在恢复标志后可用库存不足的药品，是否恢复标志？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                            blnWarnDo = False
                                        End If
                                     End If
                                ElseIf mIntCheckStock = 2 Then
                                    '严格控制库存时不处理缺药记录
                                    blnWarnDo = False
                                End If
                            End If
                            
                            If blnWarnDo = True Then
                                .TextMatrix(n, mBillCol.TagCol) = M_STR_CHECK_NAME
                                .Row = n
                                .Col = mBillCol.FlagCol
                                Set .CellPicture = LoadResPicture("checked", vbResBitmap)
                                If .Row > 0 Then
                                    For i = 0 To .Cols - 1
                                        .Row = n
                                        .Col = i
                                        .CellBackColor = M_LNG_CHECKED_COLOR
                                    Next
                                End If
                            End If
                        Next
                    End If
                ElseIf .Row > 0 Then
                    If .TextMatrix(.Row, mBillCol.TagCol) = M_STR_CHECK_NAME Then
                        .TextMatrix(.Row, mBillCol.TagCol) = M_STR_UNCHECK_NAME
                        .Col = mBillCol.FlagCol
                        Set .CellPicture = LoadResPicture("unchecked", vbResBitmap)
                        For i = 0 To .Cols - 1
                            .Col = i
                            .CellBackColor = M_LNG_UNCHECKED_COLOR
                        Next
                    Else
                        '缺药情况处理
                        If mIntCheckStock > 0 And Val(.TextMatrix(.Row, mBillCol.缺药Col)) = 1 Then
                            If mIntCheckStock = 1 Then
                                '库存不足提醒时，对缺药记录提醒，并按操作结果进行处理
                                If MsgBox("存在恢复标志后可用库存不足的药品，是否恢复标志？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Exit Sub
                                End If
                            ElseIf mIntCheckStock = 2 Then
                                '严格控制库存时不处理缺药记录
                                Exit Sub
                            End If
                        End If
                            
                        .TextMatrix(.Row, mBillCol.TagCol) = M_STR_CHECK_NAME
                        .Col = mBillCol.FlagCol
                        Set .CellPicture = LoadResPicture("checked", vbResBitmap)
                        For i = 0 To .Cols - 1
                            .Col = i
                            .CellBackColor = M_LNG_CHECKED_COLOR
                        Next
                    End If
                
                End If
            '如果选择的不是第一列，则作选择行处理
            ElseIf .Row > 0 Then
                For n = 1 To .rows - 1
                    .Row = n
                    .Col = 0
                    If .CellBackColor = M_LNG_SELECTEDCOLS_COLOR Then
                        If .TextMatrix(.Row, mBillCol.TagCol) = M_STR_CHECK_NAME Then
                            lngColor = M_LNG_CHECKED_COLOR
                        Else
                            lngColor = M_LNG_DEFAULTCOLS_COLOR
                        End If
                        For i = 0 To .Cols - 1
                            .Col = i
                            .CellBackColor = lngColor
                        Next
                    End If
                Next
                .Row = lngCurRow
                .Col = lngCurCol
                lngColor = M_LNG_SELECTEDCOLS_COLOR
                For i = 0 To .Cols - 1
                    .Col = i
                    .CellBackColor = lngColor
                Next
            End If
        End If
        
        .Redraw = True
    End With
                        
End Sub


Private Sub optFlag_Click()
    mblnIsChecked = True
    CmdOK.Enabled = False
    
    If mIntCheckStock = 0 Then
        lblComment.Visible = False
    Else
        lblComment.Visible = True
    End If
End Sub

Private Sub optUnFlag_Click()
    mblnIsChecked = False
    CmdOK.Enabled = False
    
    lblComment.Visible = False
End Sub

Private Sub txtBillNo_GotFocus()
    Call zlControl.TxtSelAll(txtBillNo)

End Sub

Private Sub txtBillNo_KeyPress(KeyAscii As Integer)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '第一位可以输入字母,其它位不行
    If KeyAscii <> 13 Then
        If Not (Len(Trim(txtBillNo.Text)) = 0 Or txtBillNo.SelLength = Len(txtBillNo.Text)) And _
            InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0:  Exit Sub
        End If
    End If
    
    If KeyAscii = 13 And txtBillNo.Text <> "" Then
        
        txtBillNo.Text = zlCommFun.GetFullNO(txtBillNo.Text, 13)
'        Call GetDetailBill
        
    End If

End Sub


Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim(txtPati.Text)) <> 0 Then
'        If InStr(1, "*-+", Mid(Trim(txtPati.Text), 1, 1)) = 0 Then
'            txtPati.Text = "*" & Trim(txtPati.Text)
'        End If
        txtPati.Text = GetPatiName(txtPati.Text)
        Call OS.PressKey(vbKeyTab)
    End If
End Sub


Private Sub txtPati_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(Trim((txtPati.Text))) > 0 Then
        If InStr("*-+", Mid(txtPati.Text, 1, 1)) > 0 Then
            If InStr("0123456789", Chr(KeyCode)) = 0 Then
                Exit Sub
            End If
        End If
    End If
End Sub

Public Sub showMe(ByVal frmParent As Form, ByVal lng药房ID As Long)
    mlng药房ID = lng药房ID
    Me.Show 1, frmParent
End Sub




